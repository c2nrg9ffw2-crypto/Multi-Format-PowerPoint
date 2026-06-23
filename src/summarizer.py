"""Convert parsed content into a slide_data list (rule-based, no LLM)."""
import re
from src.brand import ACCENT_CYCLE


# ── helpers ──────────────────────────────────────────────────────────────────

def _truncate(text: str, max_chars: int = 120) -> str:
    text = text.strip()
    return text if len(text) <= max_chars else text[: max_chars - 1] + "…"


def _clean_bullets(lines: list[str], max_bullets: int = 6) -> list[str]:
    seen, out = set(), []
    for line in lines:
        line = line.strip(" •\t-–—")
        if not line or len(line) < 4:
            continue
        key = line.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(_truncate(line))
        if len(out) >= max_bullets:
            break
    return out


def _sheet_to_data_slide(sheet_name: str, rows: list[list[str]], accent: str) -> list[dict]:
    slides = []
    if not rows:
        return slides

    headers = rows[0]
    data_rows = rows[1:13]  # max 12

    chart = None
    # Attempt a simple bar chart if exactly two columns (label + numeric)
    if len(headers) == 2:
        try:
            chart_data = {
                "labels": [r[0] for r in data_rows if r],
                "values": [float(r[1]) for r in data_rows if r and r[1]],
            }
            if chart_data["labels"] and chart_data["values"]:
                chart = chart_data
        except (ValueError, IndexError):
            pass

    slides.append({
        "type": "data",
        "title": sheet_name,
        "headers": headers,
        "rows": data_rows,
        "chart": chart,
        "accent": accent,
    })
    return slides


def _pages_to_slides(pages: list[dict]) -> list[dict]:
    slides = []
    # Group pages into sections of 1–2, capped at 8 sections
    groups = []
    i = 0
    while i < len(pages) and len(groups) < 8:
        group = [pages[i]]
        if i + 1 < len(pages) and not pages[i]["tables"] and not pages[i + 1]["tables"]:
            group.append(pages[i + 1])
            i += 2
        else:
            i += 1
        groups.append(group)

    accent_idx = 0
    for group in groups:
        combined_text = "\n".join(p["text"] for p in group if p["text"])
        lines = [l for l in combined_text.splitlines() if l.strip()]

        title = _truncate(lines[0], 80) if lines else "Section"
        bullet_lines = lines[1:] if len(lines) > 1 else lines
        bullets = _clean_bullets(bullet_lines)

        accent = ACCENT_CYCLE[accent_idx % len(ACCENT_CYCLE)]
        accent_idx += 1

        slides.append({
            "type": "content",
            "title": title,
            "bullets": bullets,
            "accent": accent,
        })

        # Any tables within the group become data slides
        for page in group:
            for table in page["tables"]:
                if table:
                    data_accent = ACCENT_CYCLE[accent_idx % len(ACCENT_CYCLE)]
                    accent_idx += 1
                    slides.append({
                        "type": "data",
                        "title": f"Table – {title}",
                        "headers": table[0],
                        "rows": table[1:13],
                        "chart": None,
                        "accent": data_accent,
                    })

    return slides


# ── public API ────────────────────────────────────────────────────────────────

def build_slide_data(
    xlsx_data: dict | None,
    pdf_data: dict | None,
    source_label: str = "Report",
) -> list[dict]:
    slides: list[dict] = []

    # ── Title slide ───────────────────────────────────────────────────────────
    subtitle_parts = []
    if xlsx_data:
        subtitle_parts.append(f"{len(xlsx_data)} sheet(s)")
    if pdf_data:
        subtitle_parts.append(f"{len(pdf_data.get('pages', []))} page(s)")
    subtitle = "Compiled from " + " · ".join(subtitle_parts) if subtitle_parts else ""
    slides.append({"type": "title", "title": source_label, "subtitle": subtitle})

    # ── Agenda slide ──────────────────────────────────────────────────────────
    agenda_items: list[str] = []
    if xlsx_data:
        for sheet in list(xlsx_data.keys())[:8]:
            agenda_items.append(sheet)
    if pdf_data:
        pages = pdf_data.get("pages", [])
        for page in pages[:4]:
            first_line = next(
                (l.strip() for l in page["text"].splitlines() if l.strip()), None
            )
            if first_line:
                agenda_items.append(_truncate(first_line, 60))
    if agenda_items:
        slides.append({"type": "agenda", "items": agenda_items[:8]})

    # ── Content from xlsx ─────────────────────────────────────────────────────
    if xlsx_data:
        accent_idx = 0
        for sheet_name, rows in xlsx_data.items():
            accent = ACCENT_CYCLE[accent_idx % len(ACCENT_CYCLE)]
            accent_idx += 1
            slides.extend(_sheet_to_data_slide(sheet_name, rows, accent))

    # ── Content from pdf ──────────────────────────────────────────────────────
    if pdf_data:
        slides.extend(_pages_to_slides(pdf_data.get("pages", [])))

    # ── Summary slide ─────────────────────────────────────────────────────────
    summary_points: list[str] = []
    if xlsx_data:
        for sheet_name, rows in xlsx_data.items():
            summary_points.append(
                f"{sheet_name}: {max(0, len(rows) - 1)} row(s) of data"
            )
    if pdf_data:
        n_pages = len(pdf_data.get("pages", []))
        n_tables = sum(len(p["tables"]) for p in pdf_data.get("pages", []))
        summary_points.append(f"PDF: {n_pages} page(s), {n_tables} table(s) extracted")

    slides.append({"type": "summary", "points": summary_points[:6]})

    return slides
