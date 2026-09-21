# Content Lineage Panel

A small Streamlit app that shows what a generated slide deck left out of its source document.

Upload a source PDF and a slide deck (PPTX or a text outline). The app lists each point that was dropped or shortened, with the PDF page and the original wording.

## What it checks

| Check | Model | What it does | Precision |
|---|---|---|---|
| Text | Claude (`claude-sonnet-5`) with citations | Finds points from the PDF that are missing or condensed in the slides. Page numbers and quotes come from the API's citation data, not from the model. | Higher |
| Figures | Gemini (Flash) | Reads the PDF pages as images and checks whether each figure or chart made it into the slides. | Lower |

Each result says which check it came from. Slide numbers in the results are reported by the model, so treat them as a pointer.

## Files

- `app.py`: the whole app
- `requirements.txt`: pinned dependencies
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

### Secrets

Create a `.env` file. Never commit it, and on a server keep it outside the folder that is served publicly. The app finds it in its own folder or in any parent folder.

```
ANTHROPIC_API_KEY=your-claude-api-key
GEMINI_API_KEY=your-gemini-api-key
APP_PASSWORD=choose-a-password
GEMINI_MODEL=
```

- `GEMINI_MODEL` is optional. Leave it empty to use the default (`gemini-3.6-flash`).
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
2. Upload the source as a PDF. Word users should export with File > Save as PDF.
3. Upload the slide deck. PPTX slides are numbered automatically. For a text outline, label the slides yourself (Slide 1, Slide 2).
4. Click **Check lineage**.

Try it with `samples/sample_source.pdf` and `samples/sample_slides.pptx`. The expected output is in `samples/sample_results.pdf`, though the wording can vary from run to run.

## Limits

- Page numbers are PDF page numbers, counted from the first page of the file. They can differ from numbers printed on the pages.
- Scanned PDFs without a text layer give weak text results. The figure check still reads them as images.
- Every check sends the full PDF to the APIs and costs money. A 30 page PDF is roughly $0.25 with Claude Sonnet 5. Gemini has a free tier with rate limits.
- Gemini sometimes returns an overload error (503). The panel then shows the text results and an error for the figure check.
- Uploaded documents are sent to Anthropic and Google. Do not use confidential material unless that is acceptable for your study.
