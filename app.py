from dotenv import load_dotenv

# Must run before anything else. DOTENV_PATH is optional; if unset, python-dotenv
# searches for a .env file starting from this file's folder and walking upward.
import os
load_dotenv(os.environ.get("DOTENV_PATH"))

import base64
import hmac
import io
import json
import re
import time
from datetime import datetime

import anthropic
import streamlit as st
from google import genai
from google.genai import errors as genai_errors
from google.genai import types
from pptx import Presentation

try:
    import fitz  # PyMuPDF, used only to render a source page with its passage highlighted
except ImportError:
    fitz = None

CLAUDE_MODEL = "claude-sonnet-5"
# Override with GEMINI_MODEL in .env if a newer Flash model is available to you.
GEMINI_MODEL = os.environ.get("GEMINI_MODEL") or "gemini-3.6-flash"  # `or` also covers an empty GEMINI_MODEL=

REQUIRED_VARS = ["ANTHROPIC_API_KEY", "GEMINI_API_KEY", "APP_PASSWORD"]

st.set_page_config(page_title="Content Lineage Panel", layout="wide")

# Study condition and participant ID come from URL parameters, not visible controls, so a
# participant looking at the screen sees no indication of either (open ?condition=without or
# ?condition=with, with &participant=P1, before handing the device over -- researcher_links.html
# builds these links for you). Condition defaults to "with" and participant to "unknown" if absent.
show_panel = st.query_params.get("condition", "with") != "without"
PARTICIPANT = (st.query_params.get("participant", "") or "unknown").strip()
print(f"[content-lineage-panel] participant={PARTICIPANT!r} "
      f"condition={'with' if show_panel else 'without'}")


# ---- Hidden, timestamped interaction log (never rendered in the UI) --------

LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")


def _log_path() -> str:
    if "log_file" not in st.session_state:
        safe_id = re.sub(r"[^A-Za-z0-9_-]", "_", PARTICIPANT) or "unknown"
        stamp = datetime.now().strftime("%Y%m%dT%H%M%S")
        os.makedirs(LOG_DIR, exist_ok=True)
        st.session_state["log_file"] = os.path.join(LOG_DIR, f"{safe_id}_{stamp}.jsonl")
    return st.session_state["log_file"]


def log_event(event: str, **details):
    """Appends one timestamped JSON line to this session's log file. Never shown in the UI;
    a logging failure is swallowed so it can never interrupt a study session."""
    try:
        record = {
            "time": datetime.now().astimezone().isoformat(timespec="seconds"),
            "participant": PARTICIPANT,
            "condition": "with" if show_panel else "without",
            "event": event,
            **details,
        }
        with open(_log_path(), "a") as f:
            f.write(json.dumps(record) + "\n")
    except Exception:
        pass


def _save_log_image(key: str, image_bytes: bytes):
    """Saves the exact source-page image a participant was shown, alongside the log file that
    references it, so the researcher can open precisely what the participant saw."""
    try:
        stem = os.path.splitext(os.path.basename(_log_path()))[0]
        images_dir = os.path.join(LOG_DIR, "images")
        os.makedirs(images_dir, exist_ok=True)
        path = os.path.join(images_dir, f"{stem}_{re.sub(r'[^A-Za-z0-9_-]', '_', key)}.png")
        with open(path, "wb") as f:
            f.write(image_bytes)
        return path
    except Exception:
        return None


# ---- Startup checks and password gate --------------------------------------

missing = [v for v in REQUIRED_VARS if not os.environ.get(v)]
if missing:
    st.error(
        "Missing required environment variable(s): " + ", ".join(missing)
        + ". Add them to the server's .env file (or set DOTENV_PATH to its location) and restart."
    )
    st.stop()

st.title("Content Lineage Panel")

if not st.session_state.get("authed"):
    entered = st.text_input("Password", type="password")
    if entered:
        if hmac.compare_digest(entered.encode(), os.environ["APP_PASSWORD"].encode()):
            st.session_state["authed"] = True
            log_event("session_start")
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()

claude = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
gemini = genai.Client(
    api_key=os.environ["GEMINI_API_KEY"],
    http_options=types.HttpOptions(
        # Gemini occasionally returns a transient 503 ("high demand") or just hangs; bound each
        # attempt so a stuck request fails fast into a retry instead of blocking the UI for minutes,
        # and let the SDK's own retry logic (it retries 408/429/5xx by default) handle the rest.
        timeout=30_000,  # milliseconds per attempt
        retry_options=types.HttpRetryOptions(attempts=3, initial_delay=2.0, max_delay=10.0),
    ),
)


# ---- Reading the uploaded slide deck -----------------------------------------

def parse_slides(upload):
    """Returns (slides, prompt_text).
    slides is a list of {"number", "title", "body"} for a PPTX, for stepping through
    slide by slide in the viewer, or None for a plain text outline (shown as one block).
    prompt_text is the flattened text sent to the models either way."""
    if upload.name.lower().endswith(".pptx"):
        slides = []
        for i, slide in enumerate(Presentation(io.BytesIO(upload.getvalue())).slides, 1):
            title_shape = slide.shapes.title
            title = title_shape.text.strip() if title_shape and title_shape.has_text_frame else ""
            body_lines = []
            for shape in slide.shapes:
                if shape is title_shape:
                    continue
                if shape.has_text_frame and shape.text_frame.text.strip():
                    body_lines.append(shape.text_frame.text)
                if getattr(shape, "has_table", False) and shape.has_table:
                    for row in shape.table.rows:
                        body_lines.append(" | ".join(c.text for c in row.cells))
            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame.text
                if notes.strip():
                    body_lines.append("Notes: " + notes)
            slides.append({
                "number": i, "title": title or f"Slide {i}",
                "body": "\n".join(l for l in body_lines if l.strip()),
            })
        prompt_text = "\n\n".join(f"--- Slide {s['number']} ---\n{s['title']}\n{s['body']}" for s in slides)
        return slides, prompt_text
    text = upload.getvalue().decode("utf-8", errors="replace")
    return None, text


@st.fragment  # isolates Prev/Next reruns from the rest of the page, including the panel
def render_slide_viewer(slides, raw_text):
    if slides:
        n = len(slides)
        st.session_state.setdefault("slide_idx", 1)
        st.session_state["slide_idx"] = min(st.session_state["slide_idx"], n)
        prev_col, mid_col, next_col = st.columns([1, 3, 1])
        if prev_col.button("Prev", disabled=st.session_state["slide_idx"] <= 1, width="stretch"):
            st.session_state["slide_idx"] -= 1
            log_event("slide_view", slide=st.session_state["slide_idx"])
            st.rerun()
        mid_col.markdown(
            f"<div style='text-align:center;color:#667085'>Slide {st.session_state['slide_idx']} of {n}</div>",
            unsafe_allow_html=True,
        )
        if next_col.button("Next", disabled=st.session_state["slide_idx"] >= n, width="stretch"):
            st.session_state["slide_idx"] += 1
            log_event("slide_view", slide=st.session_state["slide_idx"])
            st.rerun()
        cur = slides[st.session_state["slide_idx"] - 1]
        st.markdown(f"#### {cur['title']}")
        st.text(cur["body"] or "(no text on this slide)")
    else:
        st.caption("Text outline: shown as one block, since it has no slide-by-slide structure.")
        st.text(raw_text)


# ---- Check 1: text points (Claude, with citations) --------------------------

TEXT_PROMPT = """Below is the text of a slide deck that was generated from the attached source document.

<slides>
{slides}
</slides>

Identify points from the source document that are MISSING from the slides (dropped) or that
appear only in a much shorter, weaker or less specific form (condensed). Only report points of
real substance, not trivia. Report at most 12, most important first.

Use exactly this format for every finding, with no other text before, between or after them:

@@@ <short title>
STATUS: dropped OR condensed - <one line saying what was lost>
SLIDE: <for condensed points, the slide number(s) from the "--- Slide N ---" labels where the shortened version appears; NONE if dropped or the deck has no slide labels>
ORIGINAL: <quote the relevant source passage>

Quote the source passage so it is cited. If nothing important is missing, reply with only: NONE"""


def first_page_int(display: str):
    m = re.search(r"\d+", display or "")
    return int(m.group()) if m else None


def text_check(name: str, data: bytes, slides: str) -> list[dict]:
    doc = {
        "type": "document",
        "source": {
            "type": "base64",
            "media_type": "application/pdf",
            "data": base64.standard_b64encode(data).decode(),
        },
        "title": name,
        "citations": {"enabled": True},
    }

    response = claude.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=16000,
        messages=[{
            "role": "user",
            "content": [doc, {"type": "text", "text": TEXT_PROMPT.format(slides=slides)}],
        }],
    )

    # The model's text arrives in blocks; cited blocks carry a `citations` list.
    # Split on the @@@ marker, attaching each block's citations to the finding it belongs to.
    findings = []
    for block in response.content:
        if block.type != "text":
            continue
        pieces = block.text.split("@@@")
        for n, piece in enumerate(pieces):
            if n > 0:
                findings.append({"text": "", "citations": []})
            if not findings:
                continue
            findings[-1]["text"] += piece
            if n == len(pieces) - 1 and block.citations:
                findings[-1]["citations"].extend(block.citations)

    results = []
    for f in findings:
        title = f["text"].strip().splitlines()[0].strip() if f["text"].strip() else "Untitled"
        status = re.search(r"STATUS:\s*(dropped|condensed)\s*[-–—:]?\s*(.*)", f["text"], re.I)
        slide = re.search(r"SLIDE:\s*(.*)", f["text"], re.I)
        slide = slide.group(1).strip() if slide else ""
        # Page and wording come from the API's citations, not from the model's own text.
        pages, quotes = [], []
        for c in f["citations"]:
            quotes.append(c.cited_text)
            if c.type == "page_location":
                start, end = c.start_page_number, c.end_page_number - 1
                pages.append(str(start) if start >= end else f"{start}-{end}")
        page_display = ", ".join(dict.fromkeys(pages)) or "not cited"
        results.append({
            "title": title,
            "page_display": page_display,
            "page_num": first_page_int(page_display),
            "status": status.group(1).lower() if status else "unclear",
            "note": status.group(2).strip() if status else "",
            "slide": "" if slide.upper().startswith("NONE") else slide,
            "quotes": quotes,
            "source": "Text check",
        })
    return results


# ---- Check 2: figures (Gemini reads PDF pages as images) --------------------

FIGURE_PROMPT = """You are given a source PDF and the text of a slide deck generated from it.

<slides>
{slides}
</slides>

List every figure, chart or diagram in the PDF. For each, decide whether its content (the data or
the main takeaway) made it into the slides. Return a JSON array where each item has:
  "title": short title of the figure,
  "page": the PDF page number (1 = first page),
  "status": one of "present", "condensed", "dropped",
  "note": one line explaining the status,
  "slide": the slide number(s) from the "--- Slide N ---" labels where it appears, or "" if none,
  "description": what the figure shows, including key numbers or labels.
Return [] if the PDF has no figures."""


def figure_check(pdf_data: bytes, slides: str) -> list[dict]:
    # Gemini occasionally returns a transient 503 ("high demand"); Google's own advice is to
    # retry shortly after. A 4xx (bad model name, bad request) won't fix itself, so don't retry those.
    for attempt in range(3):
        try:
            response = gemini.models.generate_content(
                model=GEMINI_MODEL,
                contents=[
                    types.Part.from_bytes(data=pdf_data, mime_type="application/pdf"),
                    FIGURE_PROMPT.format(slides=slides),
                ],
                config=types.GenerateContentConfig(response_mime_type="application/json"),
            )
            break
        except genai_errors.ServerError:
            if attempt == 2:
                raise
            time.sleep(5)
    results = []
    for item in json.loads(response.text):
        if item.get("status") == "present":
            continue
        page_display = str(item.get("page", "")) or "not cited"
        results.append({
            "title": item.get("title", "Untitled figure"),
            "page_display": page_display,
            "page_num": first_page_int(page_display),
            "status": item.get("status", "unclear"),
            "note": item.get("note", ""),
            "slide": str(item.get("slide", "")),
            "quotes": [item.get("description", "")],
            "source": "Figure check",
        })
    return results


# ---- Rendering the marked source page ("Show in source") --------------------

@st.cache_data(show_spinner=False)  # PyMuPDF rendering is the expensive step; never redo it for the same finding
def render_source_page(pdf_data: bytes, page_num: int, quotes: list[str]):
    """Best-effort: renders the page as an image and highlights any quote it can find
    on it. Returns None (not an error) if PyMuPDF isn't installed or the page can't be read."""
    if fitz is None or page_num is None:
        return None
    try:
        with fitz.open(stream=pdf_data, filetype="pdf") as pdf:
            page = pdf[max(0, min(page_num - 1, len(pdf) - 1))]
            for q in quotes:
                q = re.sub(r"\s+", " ", q or "").strip()
                if not q:
                    continue
                rects = page.search_for(q) or page.search_for(q[:80])
                for rect in rects:
                    page.add_highlight_annot(rect)
            return page.get_pixmap(dpi=140).tobytes("png")
    except Exception:
        return None


# ---- The lineage panel (Sketch 1: tabs by status, one card per finding) -----

GROUP_LABELS = {"dropped": "Left out", "condensed": "Shortened", "unclear": "Unclear"}


def _add_note(entry: dict):
    # Runs before the script redraws, so every widget sees the updated list in this same pass.
    st.session_state.setdefault("notes", []).append(entry)
    log_event("note_added", title=entry["title"], page=entry["page_display"])


def _clear_notes():
    count = len(st.session_state.get("notes", []))
    st.session_state["notes"] = []
    log_event("notes_cleared", count=count)


def _show_source(key: str, r: dict, pdf_data: bytes):
    st.session_state[f"open-{key}"] = True
    seen = st.session_state.setdefault("seen", set())
    if key in seen:
        return  # log the page image once per item per session, not on every reopen
    seen.add(key)
    image = render_source_page(pdf_data, r["page_num"], r["quotes"])
    image_path = _save_log_image(key, image) if image else None
    log_event("item_seen", title=r["title"], page=r["page_display"], status=r["status"],
               source=r["source"], image=image_path)


def _hide_source(key: str):
    st.session_state[f"open-{key}"] = False


def render_card(r: dict, pdf_data: bytes, key: str):
    with st.container(border=True):
        tag = " · from figure check, less precise" if r["source"] == "Figure check" else ""
        st.markdown(f"**{r['title']}**")
        st.caption(f"p. {r['page_display']}{tag}")
        st.write(r["note"])
        if r["slide"]:
            st.caption(f"Now on slide {r['slide']}")
        if st.session_state.get(f"open-{key}"):
            with st.container(border=True):
                image = render_source_page(pdf_data, r["page_num"], r["quotes"])
                if image:
                    st.image(image, width="stretch")
                elif r["page_num"] is None:
                    st.caption("No page citation was returned for this item.")
                else:
                    st.caption("Page preview unavailable; showing the quoted text only.")
                for q in r["quotes"]:
                    if q.strip():
                        st.markdown(f"> {q.strip()}")
                st.button("Hide source", key=f"hide-{key}", on_click=_hide_source, args=(key,))
        else:
            # A plain button, not st.expander: this version of Streamlit gives expanders no way
            # to detect that a participant actually opened one, which "seen" tracking needs.
            st.button("Show in source", key=f"show-{key}", on_click=_show_source, args=(key, r, pdf_data))

        notes = st.session_state.get("notes", [])
        already_added = any(n["title"] == r["title"] and n["page_display"] == r["page_display"] for n in notes)
        if already_added:
            st.caption("Added to notes ✓")
        else:
            entry = {
                "title": r["title"], "page_display": r["page_display"],
                "quote": next((q.strip() for q in r["quotes"] if q.strip()), ""),
            }
            st.button("Add to notes", key=f"addnote-{key}", on_click=_add_note, args=(entry,))


def render_notes():
    notes = st.session_state.get("notes", [])
    if not notes:
        return
    st.markdown(f"##### My notes ({len(notes)})")
    text = "\n\n".join(f"{n['title']} (p. {n['page_display']})\n{n['quote']}" for n in notes)
    st.text_area("Copy these into your speaker notes", text, height=120, key="notes_text_area")
    st.button("Clear notes", on_click=_clear_notes)


@st.fragment  # isolates panel reruns from the slide viewer, and from any future widget inside it
def render_panel(results: list[dict], pdf_data: bytes):
    st.markdown(f"#### What is missing ({len(results)})")
    st.caption("Page references link to the source.")
    if not results:
        st.success("No omissions found.")
        return

    groups = {"dropped": [], "condensed": [], "unclear": []}
    for r in results:
        groups.get(r["status"], groups["unclear"]).append(r)

    order = [k for k in ("dropped", "condensed", "unclear") if groups[k]]
    tabs = st.tabs([f"{GROUP_LABELS[k]} ({len(groups[k])})" for k in order])
    for tab, group_key in zip(tabs, order):
        with tab:
            for i, r in enumerate(groups[group_key]):
                render_card(r, pdf_data, key=f"{group_key}-{i}")

    render_notes()


# ---- UI ---------------------------------------------------------------------

source = st.file_uploader(
    "Source document (PDF only)", type=["pdf"],
    help="Word users: File > Save as PDF first. Page numbers in the results count from the "
         "first page of the PDF file, which may differ from any numbers printed on the pages.",
)
slides_upload = st.file_uploader(
    "Generated slide deck", type=["pptx", "txt", "md"],
    help="PPTX slides are numbered automatically. A text outline has no slide numbers "
         "unless you label them yourself, e.g. 'Slide 1', 'Slide 2'.",
)

pdf_data = source.getvalue() if source else None

if slides_upload:
    sig = f"{slides_upload.name}:{slides_upload.size}"
    if st.session_state.get("slides_sig") != sig:
        st.session_state["slides_sig"] = sig
        st.session_state["slide_idx"] = 1
        st.session_state.pop("results", None)
        st.session_state.pop("errors", None)
        log_event("deck_uploaded", filename=slides_upload.name, slide=1)
    slides_list, slides_prompt_text = parse_slides(slides_upload)
else:
    slides_list, slides_prompt_text = None, ""

if source and pdf_data and not pdf_data.startswith(b"%PDF"):
    st.error("That file is not a valid PDF. Please export your document as a PDF and upload it again.")
    st.stop()

if not show_panel:
    # True no-op for the study's baseline condition: no button, no API calls, no spinner,
    # no possible error -- so a participant in this condition sees no sign the tool exists.
    st.session_state.pop("results", None)
    st.session_state.pop("errors", None)
elif source and slides_upload and st.button("Check lineage"):
    log_event("check_started")
    t0 = time.time()
    results, errors = [], []
    with st.spinner("Running text check..."):
        try:
            results += text_check(source.name, pdf_data, slides_prompt_text)
        except Exception as e:
            errors.append(f"Text check failed: {e}")
            log_event("check_error", check="text", error=str(e))
    with st.spinner("Running figure check..."):
        try:
            results += figure_check(pdf_data, slides_prompt_text)
        except Exception as e:
            errors.append(f"Figure check failed: {e}")
            log_event("check_error", check="figure", error=str(e))
    st.session_state["results"] = results
    st.session_state["errors"] = errors
    log_event("check_completed", findings=len(results), seconds=round(time.time() - t0, 1))

if show_panel:
    for err in st.session_state.get("errors", []):
        st.error(err)

if slides_upload:
    with st.container(border=True):
        if show_panel:
            col_slide, col_panel = st.columns([3, 2])
        else:
            col_slide, col_panel = st.container(), None

        with col_slide:
            st.markdown("**Slide review window**")
            render_slide_viewer(slides_list, slides_prompt_text)

        if col_panel is not None:
            with col_panel:
                if "results" in st.session_state:
                    render_panel(st.session_state["results"], pdf_data)
                else:
                    st.markdown("#### What is missing")
                    st.caption("Upload a source PDF and click Check lineage to see results here.")
