# Content Lineage Panel

A small Streamlit app that shows what a generated slide deck left out of its source document.

Upload a source PDF and a slide deck (PPTX or a text outline). The app lists each point that was dropped or shortened, with its page number, a one-line note, and the original wording.

## What it checks

| Check | Model | What it does | Precision |
|---|---|---|---|
| Text | Claude (`claude-sonnet-5`) with citations | Finds points from the PDF that are missing or condensed in the slides. Page numbers and quotes come from the API's citation data, not from the model. | Higher |
| Figures | Gemini (Flash) | Reads the PDF pages as images and checks whether each figure or chart made it into the slides. | Lower |

Each result says which check it came from. Gemini occasionally returns a temporary "high demand" error; the app retries automatically before giving up.

## Layout

Two columns once a deck is uploaded:

- **Slide review window** (left): steps through the generated deck slide by slide.
- **Lineage panel** (right): tabs for **Left out** and **Shortened**, each item showing its page and a one-line note. **Show in source** reveals the highlighted source page and the original wording. **Add to notes** on a card saves it into a running list you can copy into your own speaker notes.

## Study condition (for running a With-panel / Without-panel study)

The panel can be hidden entirely, so the two conditions differ in exactly one thing. There is no visible toggle in the app — a participant looking at the screen sees no sign of which condition is active. It's set purely by a URL parameter:

- `?condition=with` (or no parameter) — panel shows as normal.
- `?condition=without` — the panel never renders, and no checks run at all (no API calls, no spinner, no possible error) — a true no-op baseline, not just a hidden result.

A `&participant=P1` parameter is recorded in the session log; omit it and the session logs as `unknown`.

Open `researcher_links.html` in your own browser (never on the participant's screen) to build and open these two links without typing them from memory.

## Session log

Every session writes a timestamped, append-only log to `logs/<participant>_<timestamp>.jsonl` (one JSON object per line) — never shown anywhere in the app, so it can't leak the study condition or bias behaviour. Logged events: `session_start`, `deck_uploaded`, `slide_view`, `check_started` / `check_completed` / `check_error`, `item_seen` (first time a finding's source is opened — includes a path to the exact page image shown, saved under `logs/images/`), `note_added`, `notes_cleared`. `logs/` is gitignored; it holds real participant data and must never be committed.

**Retrieving the logs.** Running locally or on your own server, just read the `logs/` folder directly. On a cloud host (e.g. Streamlit Community Cloud), there is no file browser or SSH access, and — more importantly — **the filesystem is ephemeral**: it resets to a fresh copy of the repo on every restart (inactivity sleep, a new deploy, platform maintenance), silently losing anything not yet downloaded. Open `?view=logs` (password-protected, same as the app) to download everything collected so far as a zip; `researcher_links.html` has a button for it. Download after every session, not just at the end of the study.

## Files

- `app.py`: the whole app
- `requirements.txt`: pinned dependencies
- `researcher_links.html`: a local, researcher-only page for opening the With/Without panel URLs, with a participant ID field
- `samples/`: a made-up source PDF, a slide deck and the results they produce

## Setup

Python 3.10 or newer is recommended.

```
git clone https://github.com/Evans-Junior/content-linage-panel.git
cd content-linage-panel
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

No system packages are required — `pymupdf` (used to render and highlight the source page) ships its own PDF engine.

### Secrets

Create a `.env` file. Never commit it, and on a server keep it outside the folder that is served publicly. The app finds it in its own folder or in any parent folder (or set `DOTENV_PATH` to point at it directly).

```
ANTHROPIC_API_KEY=your-claude-api-key
GEMINI_API_KEY=your-gemini-api-key
APP_PASSWORD=choose-a-password
GEMINI_MODEL=
```

- `GEMINI_MODEL` is optional. Leave it empty to use the default (`gemini-3.6-flash`); `gemini-flash-latest` is a reasonable alternative if that default model is retired later.
- The app stops at startup with an error if a required variable is missing.
- The password gate is only there to keep casual visitors out. It is not real authentication.

API keys: [Anthropic console](https://console.anthropic.com/settings/keys) (needs prepaid credit, a Claude Pro plan does not include API access) and [Google AI Studio](https://aistudio.google.com/apikey).

## Run

```
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```

The port and address come from the command line, not from the app.

## Using it

1. Enter the password.
2. Upload the source as a PDF only. Word users should export with File > Save as PDF first.
3. Upload the slide deck. PPTX slides are numbered automatically. For a text outline, label the slides yourself (Slide 1, Slide 2).
4. Click **Check lineage**. Step through the slide viewer on the left; review findings in the panel on the right.

Try it with `samples/sample_source.pdf` and `samples/sample_slides.pptx`. The expected output is in `samples/sample_results.pdf`, though the wording can vary from run to run.

## Limits

- Page numbers are PDF page numbers, counted from the first page of the file. They can differ from numbers printed on the pages.
- Scanned PDFs without a text layer give weak text results. The figure check still reads them as images.
- Every check sends the full PDF to the APIs and costs money. A 30 page PDF is roughly $0.25 with Claude Sonnet 5. Gemini has a free tier with rate limits.
- Uploaded documents are sent to Anthropic and Google. Do not use confidential material, as I am not liable for any consequences after that.
