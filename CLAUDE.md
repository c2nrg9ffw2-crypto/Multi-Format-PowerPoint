# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

CLI tool that converts `.xlsx` and/or `.pdf` files into a polished `.pptx` presentation styled with Anthropic's brand guidelines.

## Setup

```bash
# Python dependencies (Node 18+ required for PptxGenJS)
pip install openpyxl pdfplumber markitdown

# Node dependency (PptxGenJS — used as the PPTX renderer)
npm install pptxgenjs

# Only needed when using --template
pip install python-pptx
```

## Running

```bash
python main.py report.xlsx                              # xlsx only
python main.py doc.pdf                                  # pdf only
python main.py data.xlsx summary.pdf -o deck.pptx       # both, custom output
python main.py a.xlsx b.xlsx c.pdf -o combined.pptx     # multiple files merged
python main.py data.xlsx --template brand.pptx          # edit existing template
python main.py report.xlsx --no-qa                      # skip QA step
```

### CLI arguments (`main.py`)

| Argument | Description |
|---|---|
| `inputs` (positional, one or more) | `.xlsx` and/or `.pdf` source files; multiple files of each type are merged |
| `-o / --output` | Output `.pptx` path (default: `output.pptx`) |
| `--template TEMPLATE.pptx` | Existing `.pptx` to use as a base; triggers the python-pptx path instead of PptxGenJS |
| `--no-qa` | Skip the QA validation step |

### Generating test data

`_make_sample_pdf.py` generates `sample_report.pdf` (a multi-page business report with tables). Run once to get a realistic test PDF:

```bash
pip install reportlab   # one-time; only needed for this script
python _make_sample_pdf.py
```

## Architecture

Python orchestrates the pipeline with two rendering paths:

- **No `--template`** (default): `builder.py` generates a PptxGenJS JS script, writes it to `_pptx_render_tmp.js` in the project root, shells out to `node`, then deletes the temp file. Full Anthropic brand styling. Requires Node 18+.
- **With `--template`**: `builder.py:build_pptx_from_template()` uses `python-pptx` to open the supplied `.pptx`, maps slide types to the template's existing layouts by name substring match, and appends slides. Inherits the template's theme and fonts. Requires `pip install python-pptx`.

### Pipeline stages

| Stage | File | Role |
|---|---|---|
| Parse | `src/parsers.py` | `parse_xlsx()` → sheets/rows dict; `parse_pdf()` → pages/text/tables dict |
| Summarize | `src/summarizer.py` | `build_slide_data()` converts parsed content into a `slide_data` list (typed slide dicts) |
| Build | `src/builder.py` | `build_pptx()` → PptxGenJS via `node`; `build_pptx_from_template()` → python-pptx editing path |
| QA | `src/qa.py` | Runs `markitdown` on output; checks for placeholder text and empty slides |
| Brand | `src/brand.py` | Anthropic color and font constants — imported by `builder.py` only |

### slide_data contract

`summarizer.py` produces a list of dicts; `builder.py` consumes it. Each dict has a `type` key that controls dispatch:

```python
{"type": "title",   "title": str, "subtitle": str}
{"type": "agenda",  "items": [str, ...]}           # max 8 items
{"type": "content", "title": str, "bullets": [str, ...], "accent": str}  # max 6 bullets
{"type": "data",    "title": str, "headers": [str], "rows": [[...]], "chart": {...} | None, "accent": str}  # max 12 rows
{"type": "summary", "points": [str, ...]}           # max 6 points
```

Chart auto-generation: a bar chart (`chart` dict with `labels`/`values`) is created only when a data slide has **exactly 2 columns** and the second column is fully numeric. Otherwise `chart` is `None` and only the table renders.

### Slide order

`title` → `agenda` → content slides → `summary`.  
xlsx sheets → one `data` slide each.  
pdf page groups (1–2 pages) → one `content` slide + additional `data` slides for any tables found.  
pdf groups capped at 8 sections; xlsx sheets capped at 500 rows each.

### Visual design notes

- Title, agenda, and summary slides use a light (`#faf9f5`) or dark (`#141413`) background with an orange top/bottom bar.
- Summary slide inverts the palette: dark background, light text.
- Content and data slides use a vertical colored left bar (0.18" wide) as the accent, with the color cycling orange → blue → green.
- Slide canvas: 13.3" × 7.5" (`LAYOUT_WIDE`).

### Brand constants (`src/brand.py`)

Colors: `DARK` (#141413), `LIGHT` (#faf9f5), `ORANGE` (#d97757), `BLUE` (#6a9bcc), `GREEN` (#788c5d), `MID_GRAY` (#6b6b6b), `LIGHT_GRAY` (#d4d4d0).  
Fonts: `FONT_HEADING = "Poppins"`, `FONT_BODY = "Lora"`.  
Accent colors cycle via `ACCENT_CYCLE = [ORANGE, BLUE, GREEN]` — imported by `summarizer.py` and `builder.py`.

## Claude Code skills used in this project

| Skill | Purpose |
|---|---|
| `example-skills:brand-guidelines` | Anthropic brand colors, fonts, and visual identity applied in `src/brand.py` and `src/builder.py` |
| `document-skills:pptx` | Context for PptxGenJS slide generation patterns and layout decisions |
| `document-skills:xlsx` | openpyxl-based parsing conventions used in `src/parsers.py` |
| `document-skills:pdf` | pdfplumber text/table extraction patterns used in `src/parsers.py` |

## Constraints

- Inputs: `.xlsx` and `.pdf` only; unlimited files of each type per invocation.
- Multiple xlsx files: sheets are merged; sheet names are prefixed with `{stem} – ` when more than one xlsx is given.
- Multiple pdf files: pages are concatenated in input order before summarization.
- Corrupted or unreadable files raise `ParseError` with a clear message and exit code 1.
- Output: `.pptx` only.
- No AI/LLM call — summarization is purely rule-based (heuristics in `summarizer.py`).
- xlsx sheets capped at 500 rows; pdf groups capped at 8 sections of 1–2 pages.
- Slides capped: 6 bullets per content/summary slide, 8 agenda items, 12 table rows (data slide).
- Bullet text truncated at 120 chars; titles at 80 chars.
