"""Generate a PptxGenJS script from slide_data and execute it via node."""
import json
import subprocess
import sys
from pathlib import Path

from src import brand

# LAYOUT_WIDE dimensions: 13.3" wide × 7.5" tall
W = 13.3
H = 7.5
MARGIN = 0.5
CONTENT_X = MARGIN
CONTENT_W = W - MARGIN * 2   # 12.3"
FOOTER_Y = 7.38


def _hex(color: str) -> str:
    return color.lstrip("#")


def _js_str(s) -> str:
    s = str(s)
    s = s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n").replace("\r", "")
    return f'"{s}"'


def _bullets_js(items: list[str], font_face: str, color: str, size: int = 18) -> str:
    if not items:
        return ""
    parts = []
    for i, text in enumerate(items):
        is_last = i == len(items) - 1
        break_line = "false" if is_last else "true"
        parts.append(
            f'{{text: {_js_str(text)}, options: {{bullet:true, breakLine:{break_line}, '
            f'fontSize:{size}, color:"{color}", fontFace:"{font_face}"}}}}'
        )
    return ", ".join(parts)


def _bar(x, y, w, h, color) -> str:
    c = _hex(color)
    return f's.addShape(pptx.shapes.RECTANGLE, {{x:{x}, y:{y}, w:{w}, h:{h}, fill:{{color:"{c}"}}, line:{{color:"{c}"}}}});'


def _build_title_slide(slide: dict) -> str:
    title = _js_str(slide.get("title", "Presentation"))
    subtitle = _js_str(slide.get("subtitle", ""))
    dark = _hex(brand.DARK)
    light = _hex(brand.LIGHT)
    mid = _hex(brand.MID_GRAY)

    return f"""
  var s = pptx.addSlide();
  s.background = {{color: "{light}"}};
  {_bar(0, 0, W, 0.08, brand.ORANGE)}
  s.addText({title}, {{x:{CONTENT_X}, y:1.6, w:{CONTENT_W}, h:1.6,
    fontSize:44, bold:true, color:"{dark}", fontFace:"{brand.FONT_HEADING}",
    align:"left", wrap:true, margin:0}});
  s.addText({subtitle}, {{x:{CONTENT_X}, y:3.4, w:{CONTENT_W}, h:0.8,
    fontSize:20, color:"{mid}", fontFace:"{brand.FONT_BODY}",
    align:"left", wrap:true, margin:0}});
  {_bar(0, FOOTER_Y, W, 0.12, brand.ORANGE)}
"""


def _build_agenda_slide(slide: dict) -> str:
    items = slide.get("items", [])
    dark = _hex(brand.DARK)
    light = _hex(brand.LIGHT)

    rows_js = ", ".join(
        f'[{{text: {_js_str(item)}, options: {{color:"{dark}", fontFace:"{brand.FONT_BODY}", fontSize:18}}}}]'
        for item in items
    )

    return f"""
  var s = pptx.addSlide();
  s.background = {{color: "{light}"}};
  {_bar(0, 0, W, 0.08, brand.ORANGE)}
  s.addText("Agenda", {{x:{CONTENT_X}, y:0.25, w:{CONTENT_W}, h:0.7,
    fontSize:30, bold:true, color:"{dark}", fontFace:"{brand.FONT_HEADING}", margin:0}});
  s.addTable([{rows_js}], {{x:{CONTENT_X}, y:1.1, w:{CONTENT_W},
    colW:[{CONTENT_W}], rowH:0.52, border:{{type:"none"}},
    fill:{{color:"{light}"}}, autoPage:false}});
  {_bar(0, FOOTER_Y, W, 0.12, brand.ORANGE)}
"""


def _build_content_slide(slide: dict) -> str:
    title = _js_str(slide.get("title", ""))
    bullets = slide.get("bullets", [])
    accent = slide.get("accent", brand.ORANGE)
    dark = _hex(brand.DARK)
    light = _hex(brand.LIGHT)
    acc = _hex(accent)

    body_x = 0.55
    body_w = W - body_x - MARGIN

    bullet_items = _bullets_js(bullets, brand.FONT_BODY, dark)

    return f"""
  var s = pptx.addSlide();
  s.background = {{color: "{light}"}};
  {_bar(0, 0, 0.18, H, accent)}
  s.addText({title}, {{x:{body_x}, y:0.25, w:{body_w}, h:0.75,
    fontSize:26, bold:true, color:"{dark}", fontFace:"{brand.FONT_HEADING}",
    wrap:true, margin:0}});
  s.addText([{bullet_items}], {{x:{body_x}, y:1.2, w:{body_w}, h:{H - 1.2 - 0.3},
    valign:"top", wrap:true, autoFit:true}});
  {_bar(0, FOOTER_Y, W, 0.12, accent)}
"""


def _build_data_slide(slide: dict) -> str:
    title = _js_str(slide.get("title", "Data"))
    headers = slide.get("headers", [])
    rows = slide.get("rows", [])
    accent = slide.get("accent", brand.BLUE)
    dark = _hex(brand.DARK)
    light = _hex(brand.LIGHT)
    light_gray = _hex(brand.LIGHT_GRAY)
    acc = _hex(accent)

    def cell(text, bold=False, is_header=False):
        fill = acc if is_header else light
        color = light if is_header else dark
        return (
            f'{{text: {_js_str(text)}, '
            f'options: {{bold:{"true" if bold else "false"}, fontSize:14, '
            f'color:"{color}", fontFace:"{brand.FONT_BODY}", '
            f'fill:{{color:"{fill}"}}, align:"left", valign:"middle"}}}}'
        )

    header_row = "[" + ", ".join(cell(h, bold=True, is_header=True) for h in headers) + "]"
    data_rows_js = ", ".join(
        "[" + ", ".join(cell(c) for c in row) + "]"
        for row in rows
    )
    all_rows = header_row + (", " + data_rows_js if data_rows_js else "")

    body_x = 0.55
    chart_js = ""
    chart = slide.get("chart")
    if chart:
        labels_js = json.dumps(chart["labels"])
        values_js = json.dumps(chart["values"])
        chart_js = f"""
  s.addChart(pptx.charts.BAR, [{{name:"Values", labels:{labels_js}, values:{values_js}}}], {{
    x:{body_x}, y:1.15, w:5.8, h:3.6, barDir:"col",
    chartColors:["{acc}"],
    chartArea:{{fill:{{color:"{light}"}}, roundedCorners:false}},
    showLegend:false,
    showValue:true, dataLabelPosition:"outEnd", dataLabelColor:"{dark}",
    catAxisLabelColor:"{dark}", valAxisLabelColor:"{dark}",
    valGridLine:{{color:"{light_gray}", size:0.5}},
    catGridLine:{{style:"none"}},
  }});"""
        table_x = 6.7
        table_w = W - table_x - MARGIN
    else:
        table_x = body_x
        table_w = W - body_x - MARGIN

    return f"""
  var s = pptx.addSlide();
  s.background = {{color: "{light}"}};
  {_bar(0, 0, 0.18, H, accent)}
  s.addText({title}, {{x:{body_x}, y:0.25, w:{W - body_x - MARGIN}, h:0.75,
    fontSize:26, bold:true, color:"{dark}", fontFace:"{brand.FONT_HEADING}",
    wrap:true, margin:0}});{chart_js}
  s.addTable([{all_rows}], {{x:{table_x}, y:1.15, w:{table_w},
    rowH:0.32, autoPage:false,
    border:{{type:"solid", color:"{light_gray}", pt:0.5}}}});
  {_bar(0, FOOTER_Y, W, 0.12, accent)}
"""


def _build_summary_slide(slide: dict) -> str:
    points = slide.get("points", [])
    dark = _hex(brand.DARK)
    light = _hex(brand.LIGHT)

    bullet_items = _bullets_js(points, brand.FONT_BODY, light)

    return f"""
  var s = pptx.addSlide();
  s.background = {{color: "{dark}"}};
  {_bar(0, 0, W, 0.08, brand.ORANGE)}
  s.addText("Summary", {{x:{CONTENT_X}, y:0.4, w:{CONTENT_W}, h:0.8,
    fontSize:30, bold:true, color:"{light}", fontFace:"{brand.FONT_HEADING}", margin:0}});
  s.addText([{bullet_items}], {{x:{CONTENT_X}, y:1.4, w:{CONTENT_W}, h:{H - 1.4 - 0.3},
    valign:"top", wrap:true, autoFit:true}});
  {_bar(0, FOOTER_Y, W, 0.12, brand.ORANGE)}
"""


def build_pptx(slide_data: list[dict], output_path: str) -> None:
    output_path = str(Path(output_path).resolve())

    parts = [
        "const PptxGenJS = require('pptxgenjs');",
        "const pptx = new PptxGenJS();",
        "pptx.layout = 'LAYOUT_WIDE';",
    ]

    for slide in slide_data:
        stype = slide.get("type")
        if stype == "title":
            parts.append(_build_title_slide(slide))
        elif stype == "agenda":
            parts.append(_build_agenda_slide(slide))
        elif stype == "content":
            parts.append(_build_content_slide(slide))
        elif stype == "data":
            parts.append(_build_data_slide(slide))
        elif stype == "summary":
            parts.append(_build_summary_slide(slide))

    js_output = output_path.replace("\\", "\\\\")
    parts.append(
        f'pptx.writeFile({{fileName: "{js_output}"}}).then(() => console.log("OK")).catch(e => {{ console.error(e); process.exit(1); }});'
    )

    js_source = "\n".join(parts)

    project_root = Path(__file__).parent.parent
    tmp_file = project_root / "_pptx_render_tmp.js"
    tmp_file.write_text(js_source, encoding="utf-8")

    try:
        result = subprocess.run(
            ["node", str(tmp_file)],
            capture_output=True,
            text=True,
            timeout=60,
            cwd=str(project_root),
        )
        if result.returncode != 0:
            print(result.stderr, file=sys.stderr)
            raise RuntimeError(f"node exited with code {result.returncode}")
    finally:
        tmp_file.unlink(missing_ok=True)
