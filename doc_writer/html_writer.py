from __future__ import annotations

import html as _html
from pathlib import Path

from .base_writer import BaseDocumentWriter

_PAGE_CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
    font-family: 'Calibri', 'Segoe UI', Arial, sans-serif;
    max-width: 210mm;
    margin: auto;
    background: #ffffff;
}
.combined-block { margin-bottom: 2em; }
.row { margin-bottom: 0.75em; }
.row.multi-col { display: flex; gap: 1.2em; align-items: flex-start; }
ul.bullet-list { list-style: none; }
@media print {
    body { margin: 0; max-width: none; }
    .combined-block { page-break-inside: avoid; }
}
"""


class HtmlDocumentWriter(BaseDocumentWriter):
    """Generates a self-contained .html document from doc_writer templates."""

    def __init__(self, custom_templates_dir: Path | str | None = None):
        super().__init__(custom_templates_dir)
        self._blocks: list[str] = []
        self._body_style = "padding: 12mm 15mm;"
        self._pending_images: list[dict] = []

    # ------------------------------------------------------------------
    # Page setup
    # ------------------------------------------------------------------

    def _setup_page(self, page_def: dict) -> None:
        mm = page_def.get("margins_mm", {})
        top = mm.get("top", 25)
        bottom = mm.get("bottom", 25)
        left = mm.get("left", 30)
        right = mm.get("right", 30)
        self._body_style = f"padding: {top}mm {right}mm {bottom}mm {left}mm;"

    # ------------------------------------------------------------------
    # Layout rendering
    # ------------------------------------------------------------------

    def _render_combined(self, resolved_rows: list[dict]) -> None:
        parts = ['<div class="combined-block">']
        for row in resolved_rows:
            cols = row["columns"]
            if len(cols) == 1:
                parts.append('<div class="row single-col">')
                for comp in cols[0]["components"]:
                    parts.append(self._render_component(comp))
                parts.append("</div>")
            else:
                parts.append('<div class="row multi-col">')
                for col in cols:
                    pct = col["width_pct"]
                    parts.append(f'<div style="width:{pct}%;flex-shrink:0;min-width:0;">')
                    for comp in col["components"]:
                        parts.append(self._render_component(comp))
                    parts.append("</div>")
                parts.append("</div>")
        parts.append("</div>")
        self._blocks.append("\n".join(parts))

    def _render_component(self, comp: dict) -> str:
        tpl = comp["template"]
        content = comp["content"]
        dispatch = {
            "title": self._html_title,
            "paragraph": self._html_paragraph,
            "bullet_list": self._html_bullet_list,
            "image": self._html_image,
            "table": self._html_table,
            "callout": self._html_callout,
            "separator": self._html_separator,
            "chart": self._html_chart,
            "flowchart": self._html_flowchart,
        }
        handler = dispatch.get(tpl["type"])
        return handler(tpl, content) if handler else ""

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _font_style(font_def: dict) -> str:
        parts = []
        if name := font_def.get("name"):
            parts.append(f"font-family:'{name}',sans-serif")
        if size := font_def.get("size"):
            parts.append(f"font-size:{size}pt")
        if font_def.get("bold"):
            parts.append("font-weight:bold")
        if font_def.get("italic"):
            parts.append("font-style:italic")
        if color := font_def.get("color"):
            parts.append(f"color:{color}")
        return ";".join(parts)

    # ------------------------------------------------------------------
    # Singular renderers
    # ------------------------------------------------------------------

    def _html_title(self, tpl: dict, content: dict) -> str:
        text = content.get("text", "")
        level = content.get("level", 1)
        level_def = tpl["levels"].get(str(level), next(iter(tpl["levels"].values())))
        font_def = level_def["font"]
        sp = level_def.get("spacing", {})
        style = (
            self._font_style(font_def)
            + f";margin:{sp.get('before_pt', 12)}pt 0 {sp.get('after_pt', 6)}pt"
        )
        tag = f"h{min(max(level, 1), 6)}"
        return f'<{tag} style="{style}">{_html.escape(text)}</{tag}>'

    def _html_paragraph(self, tpl: dict, content: dict) -> str:
        text = content.get("text", "")
        font_def = tpl["font"]
        sp = tpl.get("spacing", {})
        align = tpl.get("alignment", "left")
        style = (
            self._font_style(font_def)
            + f";margin:{sp.get('before_pt', 6)}pt 0 {sp.get('after_pt', 6)}pt"
            + f";line-height:{sp.get('line_spacing', 1.15)}"
            + f";text-align:{align}"
        )
        return f'<p style="{style}">{_html.escape(text)}</p>'

    def _html_bullet_list(self, tpl: dict, content: dict) -> str:
        items = content.get("items", [])
        bullet_def = tpl["bullet"]
        font_def = tpl["font"]
        sp = tpl.get("spacing", {})
        indent = tpl.get("indent_pt", 18)
        bullet_color = bullet_def.get("color", "#333")
        symbol = _html.escape(bullet_def.get("symbol", "•"))

        li_style = (
            self._font_style(font_def)
            + f";margin:{sp.get('before_pt', 2)}pt 0"
            + f";line-height:{sp.get('line_spacing', 1.15)}"
            + f";text-align:{tpl.get('alignment', 'left')}"
            + ";display:flex;align-items:baseline;gap:6px"
        )
        ul_style = f"padding-left:{indent}pt;margin:6pt 0"

        lis = "".join(
            f'<li style="{li_style}">'
            f'<span style="color:{bullet_color};flex-shrink:0">{symbol}</span>'
            f'<span>{_html.escape(str(item))}</span>'
            f"</li>"
            for item in items
        )
        return f'<ul class="bullet-list" style="{ul_style}">{lis}</ul>'

    def _html_image(self, tpl: dict, content: dict) -> str:
        frame = tpl.get("frame", {})
        caption_def = tpl.get("caption", {})
        caption_text = content.get("caption", "")

        border_color = frame.get("border_color", "#CCCCCC")
        border_w = frame.get("border_width_pt", 1.5)
        padding = frame.get("padding_pt", 8)
        bg = frame.get("background_color", "#FAFAFA")

        wrapper_style = (
            f"border:{border_w}pt solid {border_color}"
            f";padding:{padding}pt"
            f";background:{bg}"
            f";text-align:center"
            f";margin:4pt 0"
        )

        # Token replaced at save() time with an embedded base64 <img>
        idx = len(self._pending_images)
        self._pending_images.append({"tpl": tpl, "content": content})
        inner = f"<!--IMG:{idx}-->"

        caption_html = ""
        if caption_text:
            cap_style = (
                self._font_style(caption_def.get("font", {}))
                + ";display:block;margin-top:4pt;text-align:center"
            )
            caption_html = f'<span style="{cap_style}">{_html.escape(caption_text)}</span>'

        return f'<div style="{wrapper_style}">{inner}{caption_html}</div>'

    def _html_table(self, tpl: dict, content: dict) -> str:
        headers = content.get("headers", [])
        rows_data = content.get("rows", [])
        header_def = tpl.get("header", {})
        cell_def = tpl.get("cell", {})
        border_def = tpl.get("border", {})

        border_color = border_def.get("color", "#AAAAAA")
        border_w = border_def.get("width_pt", 0.5)
        border_css = f"{border_w}pt solid {border_color}"

        th_style = (
            self._font_style(header_def.get("font", {}))
            + f";background:{header_def.get('background_color', '#333')}"
            + f";padding:{cell_def.get('padding_pt', 4)}pt 8pt"
            + f";text-align:{header_def.get('alignment', 'center')}"
            + f";border:{border_css}"
        )
        td_style = (
            self._font_style(cell_def.get("font", {}))
            + f";padding:{cell_def.get('padding_pt', 4)}pt 8pt"
            + f";text-align:{cell_def.get('alignment', 'left')}"
            + f";border:{border_css}"
        )

        thead = ""
        if headers:
            ths = "".join(f'<th style="{th_style}">{_html.escape(str(h))}</th>' for h in headers)
            thead = f"<thead><tr>{ths}</tr></thead>"

        tbody_rows = "".join(
            "<tr>" + "".join(f'<td style="{td_style}">{_html.escape(str(v))}</td>' for v in row) + "</tr>"
            for row in rows_data
        )
        tbody = f"<tbody>{tbody_rows}</tbody>" if tbody_rows else ""

        table_style = "width:100%;border-collapse:collapse;margin:8pt 0"
        return f'<table style="{table_style}">{thead}{tbody}</table>'

    def _html_callout(self, tpl: dict, content: dict) -> str:
        label = content.get("label", "")
        text = content.get("text", "")
        border_color = tpl.get("border_left_color", "#2980B9")
        bg_color = tpl.get("background_color", "#EBF5FB")
        border_w = tpl.get("border_left_width_pt", 4)
        padding = tpl.get("padding_pt", 10)
        sp = tpl.get("spacing", {})

        wrapper_style = (
            f"border-left:{border_w}pt solid {border_color}"
            f";background:{bg_color}"
            f";padding:{padding}pt {padding}pt {padding}pt {padding + 4}pt"
            f";margin:{sp.get('before_pt', 8)}pt 0 {sp.get('after_pt', 8)}pt"
        )

        label_html = ""
        if label:
            label_style = (
                self._font_style(tpl.get("label_font", {}))
                + ";display:block;margin-bottom:4pt;letter-spacing:0.05em;text-transform:uppercase"
            )
            label_html = f'<span style="{label_style}">{_html.escape(label)}</span>'

        text_style = self._font_style(tpl.get("font", {}))
        text_html = f'<span style="{text_style}">{_html.escape(text)}</span>'

        return f'<div style="{wrapper_style}">{label_html}{text_html}</div>'

    def _html_separator(self, tpl: dict, content: dict) -> str:
        color = tpl.get("line_color", "#CCCCCC")
        width = tpl.get("line_width_pt", 0.5)
        margin_top = tpl.get("margin_top_pt", 12)
        margin_bottom = tpl.get("margin_bottom_pt", 12)
        style = (
            f"border:none;border-top:{width}pt solid {color}"
            f";margin:{margin_top}pt 0 {margin_bottom}pt"
            f";opacity:0.4"
        )
        return f'<hr style="{style}">'

    def _html_chart(self, tpl: dict, content: dict) -> str:
        import base64
        from .chart_utils import render_chart_png
        buf = render_chart_png(tpl, content)
        b64 = base64.b64encode(buf.getvalue()).decode()
        width_px = int(tpl.get("figure", {}).get("width_inches", 6) * tpl.get("figure", {}).get("dpi", 150))
        return (
            f'<img src="data:image/png;base64,{b64}" '
            f'style="max-width:100%;width:{width_px}px;display:block;margin:4pt auto;">'
        )

    def _html_flowchart(self, tpl: dict, content: dict) -> str:
        from .flowchart_utils import build_simple_svg, build_complex_svg
        ftype = content.get("flowchart_type", "simple")
        nodes = content.get("nodes", [])
        if ftype == "simple":
            svg = build_simple_svg(nodes, tpl)
        else:
            svg = build_complex_svg(nodes, content.get("edges", []), tpl)
        return f'<div style="text-align:center;margin:8pt 0">{svg}</div>'

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------

    def save(self, output_path: str | Path) -> None:
        import base64
        from .image_utils import resolve_image_bytes

        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        assets_dir = out.parent / "assets"

        body_content = "\n".join(self._blocks)
        for i, item in enumerate(self._pending_images):
            img_bytes, filename = resolve_image_bytes(item["content"], assets_dir, i)
            b64 = base64.b64encode(img_bytes).decode()
            ext = Path(filename).suffix.lstrip(".").lower()
            mime = "image/jpeg" if ext in ("jpg", "jpeg") else f"image/{ext or 'png'}"
            img_tag = (
                f'<img src="data:{mime};base64,{b64}" '
                f'style="max-width:100%;height:auto;">'
            )
            body_content = body_content.replace(f"<!--IMG:{i}-->", img_tag)

        doc = f"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Document</title>
<style>
{_PAGE_CSS}
body {{ {self._body_style} }}
</style>
</head>
<body>
{body_content}
</body>
</html>"""
        out.write_text(doc, encoding="utf-8")
