"""
Flowchart generation.

Simple (linear sequence):
  - Word  : native DrawingML inline shapes (editable in Word)
  - HTML  : equivalent inline SVG

Complex (branching graph):
  - HTML  : rich SVG with gradients, shadows, bezier edges
  - Word  : PNG via cairosvg (if installed) or matplotlib fallback
"""
from __future__ import annotations

import html as _html
from io import BytesIO

# ─── Layout constants ────────────────────────────────────────────────────────
_INCH = 914_400                     # EMU per inch (Word)

# Word EMU node sizes
_NW, _NH   = int(1.9 * _INCH), int(0.58 * _INCH)   # process/start/end
_DW, _DH   = int(1.9 * _INCH), int(0.88 * _INCH)   # decision
_GAP       = int(0.28 * _INCH)
_CANW      = int(2.75 * _INCH)

# Simple SVG px
_S_NW, _S_NH = 180, 46
_S_DW, _S_DH = 172, 70
_S_GAP, _S_ARR, _S_W = 24, 20, 260

# Complex SVG px
_C_NW, _C_NH = 150, 44
_C_DW, _C_DH = 148, 68
_C_H_PAD, _C_V_GAP = 36, 54

# ─── Style maps ──────────────────────────────────────────────────────────────
_GEOM = {
    "start": "roundRect", "end": "roundRect",
    "process": "rect", "decision": "diamond", "io": "parallelogram",
}

# (fill_hex6, border_hex6)
_STYLE: dict[str, tuple[str, str]] = {
    "start":    ("27AE60", "1E8449"),
    "end":      ("C0392B", "922B21"),
    "process":  ("2980B9", "1A5276"),
    "decision": ("E67E22", "B7770D"),
    "io":       ("8E44AD", "6C3483"),
}

# ─── Word XML namespace URIs ─────────────────────────────────────────────────
_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
_WP  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
_WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
_W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ═══════════════════════════════════════════════════════════════════════════════
# WORD – native DrawingML shapes  (simple / linear only)
# ═══════════════════════════════════════════════════════════════════════════════

def _wshape(x: int, y: int, w: int, h: int,
            geom: str, fill: str, border: str,
            text: str, font_pt: int = 9) -> str:
    hz = int(font_pt * 2)
    t = _html.escape(text, quote=True)
    return (
        f'<wps:wsp xmlns:wps="{_WPS}" xmlns:a="{_A}" xmlns:w="{_W}">'
        '<wps:cNvSpPr txBx="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>'
        '<wps:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        f'<a:ln w="19050"><a:solidFill><a:srgbClr val="{border}"/></a:solidFill></a:ln>'
        '</wps:spPr>'
        '<wps:txbx><w:txbxContent>'
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:rPr><w:color w:val="FFFFFF"/>'
        f'<w:sz w:val="{hz}"/><w:szCs w:val="{hz}"/><w:b/></w:rPr>'
        f'<w:t xml:space="preserve">{t}</w:t></w:r></w:p>'
        '</w:txbxContent></wps:txbx>'
        '<wps:bodyPr anchor="ctr" lIns="91440" tIns="45720" rIns="91440" bIns="45720"/>'
        '</wps:wsp>'
    )


def _wconn(x1: int, y1: int, x2: int, y2: int) -> str:
    ox, oy = min(x1, x2), min(y1, y2)
    cw = max(abs(x2 - x1), 914)
    ch = max(abs(y2 - y1), 914)
    fh = "1" if x2 < x1 else "0"
    fv = "1" if y2 < y1 else "0"
    return (
        f'<wps:wsp xmlns:wps="{_WPS}" xmlns:a="{_A}">'
        '<wps:cNvCxnSpPr/>'
        '<wps:spPr>'
        f'<a:xfrm flipH="{fh}" flipV="{fv}">'
        f'<a:off x="{ox}" y="{oy}"/><a:ext cx="{cw}" cy="{ch}"/></a:xfrm>'
        '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
        '<a:ln w="19050">'
        '<a:solidFill><a:srgbClr val="555555"/></a:solidFill>'
        '<a:tailEnd type="triangle" w="sm" len="sm"/>'
        '</a:ln>'
        '</wps:spPr>'
        '<wps:bodyPr/>'
        '</wps:wsp>'
    )


def build_word_simple_flowchart(nodes: list[dict], tpl: dict):
    """Return lxml Element (w:drawing) with native Word inline grouped shapes."""
    from lxml import etree

    font_pt = tpl.get("font_size", 9)
    pos: list[tuple[int, int, int, int]] = []
    y = 0
    for node in nodes:
        is_d = node.get("type") == "decision"
        w, h = (_DW, _DH) if is_d else (_NW, _NH)
        pos.append((_CANW // 2 - w // 2, y, w, h))
        y += h + _GAP
    total_h = y - _GAP

    parts: list[str] = []
    for i, (node, (px, py, pw, ph)) in enumerate(zip(nodes, pos)):
        ntype = node.get("type", "process")
        fill, border = _STYLE.get(ntype, _STYLE["process"])
        parts.append(_wshape(px, py, pw, ph, _GEOM.get(ntype, "rect"),
                             fill, border, node.get("text", ""), font_pt))
        if i < len(nodes) - 1:
            _, ny, _, _ = pos[i + 1]
            mid = _CANW // 2
            parts.append(_wconn(mid, py + ph, mid, ny))

    inner = "".join(parts)
    xml = (
        f'<w:drawing xmlns:w="{_W}">'
        f'<wp:inline xmlns:wp="{_WP}" distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{_CANW}" cy="{total_h}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="1" name="Flowchart"/>'
        '<wp:cNvGraphicFramePr/>'
        f'<a:graphic xmlns:a="{_A}">'
        f'<a:graphicData uri="{_WPG}">'
        f'<wpg:wgp xmlns:wpg="{_WPG}">'
        '<wpg:cNvGrpSpPr/>'
        '<wpg:grpSpPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{_CANW}" cy="{total_h}"/>'
        f'<a:chOff x="0" y="0"/><a:chExt cx="{_CANW}" cy="{total_h}"/></a:xfrm>'
        '</wpg:grpSpPr>'
        f'{inner}'
        '</wpg:wgp></a:graphicData></a:graphic>'
        '</wp:inline></w:drawing>'
    )
    return etree.fromstring(xml)


# ═══════════════════════════════════════════════════════════════════════════════
# SIMPLE SVG  (HTML — visual equivalent of the Word shapes)
# ═══════════════════════════════════════════════════════════════════════════════

def build_simple_svg(nodes: list[dict], tpl: dict) -> str:
    # compute node widths/heights and dynamic canvas width so text is not clipped
    defs = (
        '<defs>'
        '<marker id="ah" markerWidth="8" markerHeight="6" refX="7" refY="3" orient="auto">'
        '<path d="M0,0 L8,3 L0,6 Z" fill="#555"/></marker>'
        '</defs>'
    )
    parts = [defs]
    y = 10
    pos: list[tuple[int, int, int, int]] = []
    # first pass: determine sizes and vertical layout
    sizes: list[tuple[int, int]] = []
    for node in nodes:
        is_d = node.get("type") == "decision"
        w, h = (_S_DW, _S_DH) if is_d else (_S_NW, _S_NH)
        sizes.append((w, h))
        y += 0  # keep same baseline progression below

    # compute total height and spacing
    y = 10
    for w, h in sizes:
        pos.append((0, y, w, h))
        y += h + _S_GAP + _S_ARR

    total_h = y - _S_GAP - _S_ARR + 10

    # compute canvas width dynamically to avoid clipping; center nodes
    max_w = max((w for w, _ in sizes), default=_S_NW)
    padding_x = 40
    canvas_w = max(_S_W, max_w + padding_x)
    CX = canvas_w // 2

    # update positions with centered x
    pos = [(CX - w // 2, py, w, h) for (_, py, w, h), (w, h) in zip(pos, sizes)]

    # connector lines between nodes
    for i in range(len(nodes) - 1):
        px, py, pw, ph = pos[i]
        pcx = px + pw // 2
        parts.append(
            f'<line x1="{pcx}" y1="{py + ph}" x2="{pcx}" y2="{pos[i+1][1] - 2}" '
            f'stroke="#555" stroke-width="1.5" marker-end="url(#ah)"/>'
        )

    for node, (x, py, w, h) in zip(nodes, pos):
        ntype = node.get("type", "process")
        fill_h, bord_h = _STYLE.get(ntype, _STYLE["process"])
        fill, stroke = f"#{fill_h}", f"#{bord_h}"
        cx_s, cy_s = x + w // 2, py + h // 2
        text = _html.escape(node.get("text", ""))

        if ntype == "decision":
            pts = f"{cx_s},{py} {x+w},{cy_s} {cx_s},{py+h} {x},{cy_s}"
            parts.append(f'<polygon points="{pts}" fill="{fill}" stroke="{stroke}" stroke-width="2"/>')
        elif ntype in ("start", "end"):
            rx = h // 2
            parts.append(f'<rect x="{x}" y="{py}" width="{w}" height="{h}" rx="{rx}" ry="{rx}" '
                         f'fill="{fill}" stroke="{stroke}" stroke-width="2"/>')
        elif ntype == "io":
            sk = 12
            pts = f"{x+sk},{py} {x+w},{py} {x+w-sk},{py+h} {x},{py+h}"
            parts.append(f'<polygon points="{pts}" fill="{fill}" stroke="{stroke}" stroke-width="2"/>')
        else:
            parts.append(f'<rect x="{x}" y="{py}" width="{w}" height="{h}" '
                         f'fill="{fill}" stroke="{stroke}" stroke-width="2"/>')

        # use foreignObject to allow wrapped, centered text inside the node
        fo_style = (
            "display:flex;align-items:center;justify-content:center;" 
            "text-align:center;color:white;" 
            "font-family:Calibri,Arial,sans-serif;" 
            "font-weight:bold;overflow:hidden;" 
            "word-wrap:break-word;white-space:normal;padding:4px;"
        )
        parts.append(
            f'<foreignObject x="{x}" y="{py}" width="{w}" height="{h}">'
            f'<div xmlns="http://www.w3.org/1999/xhtml" style="{fo_style};font-size:11px">{text}</div>'
            f'</foreignObject>'
        )
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'width="{canvas_w}" height="{total_h}" viewBox="0 0 {canvas_w} {total_h}">'
        + "\n".join(parts) + '</svg>'
    )


# ═══════════════════════════════════════════════════════════════════════════════
# COMPLEX SVG  (HTML only — graph layout, gradients, bezier edges)
# ═══════════════════════════════════════════════════════════════════════════════

def _bfs_layers(nodes: list[dict], edges: list[dict]) -> list[list[str]]:
    """Assign layers via Kahn's topological sort; back-edges (cycles) are ignored."""
    ids = [n["id"] for n in nodes]
    in_deg: dict[str, int] = {nid: 0 for nid in ids}
    out_e:  dict[str, list[str]] = {nid: [] for nid in ids}

    # Detect back-edges with a DFS; only forward/cross edges count for layering
    visited: set[str] = set()
    on_stack: set[str] = set()
    back_edges: set[tuple[str, str]] = set()

    def dfs(nid: str) -> None:
        visited.add(nid); on_stack.add(nid)
        for e in edges:
            if e["from"] != nid:
                continue
            tgt = e["to"]
            if tgt in on_stack:
                back_edges.add((nid, tgt))
            elif tgt not in visited:
                dfs(tgt)
        on_stack.discard(nid)

    for nid in ids:
        if nid not in visited:
            dfs(nid)

    for e in edges:
        if (e["from"], e["to"]) not in back_edges:
            in_deg[e["to"]] = in_deg.get(e["to"], 0) + 1
            out_e[e["from"]].append(e["to"])

    layer: dict[str, int] = {nid: -1 for nid in ids}
    queue = [nid for nid in ids if in_deg[nid] == 0] or [ids[0]]
    for nid in queue:
        layer[nid] = 0

    i = 0
    while i < len(queue):
        nid = queue[i]; i += 1
        for tgt in out_e.get(nid, []):
            if layer.get(tgt, -1) < layer[nid] + 1:
                layer[tgt] = layer[nid] + 1
                if tgt not in queue:
                    queue.append(tgt)

    for nid in ids:
        if layer[nid] < 0:
            layer[nid] = 0

    max_l = max(layer.values())
    layers: list[list[str]] = [[] for _ in range(max_l + 1)]
    for nid, l in layer.items():
        layers[l].append(nid)
    return [lyr for lyr in layers if lyr]  # drop empty layers from cycle resolution


def _node_dims_c(ntype: str) -> tuple[int, int]:
    return (_C_DW, _C_DH) if ntype == "decision" else (_C_NW, _C_NH)


def build_complex_svg(nodes: list[dict], edges: list[dict], tpl: dict) -> str:
    by_id = {n["id"]: n for n in nodes}
    layers = _bfs_layers(nodes, edges)

    def ndim(nid: str) -> tuple[int, int]:
        return _node_dims_c(by_id[nid].get("type", "process"))

    layer_widths = [
        sum(ndim(nid)[0] for nid in lyr) + _C_H_PAD * max(0, len(lyr) - 1)
        for lyr in layers
    ]
    canvas_w = max(layer_widths, default=200) + 60

    positions: dict[str, tuple[int, int]] = {}
    y = 40
    for lyr in layers:
        ws = [ndim(nid)[0] for nid in lyr]
        hs = [ndim(nid)[1] for nid in lyr]
        tw = sum(ws) + _C_H_PAD * max(0, len(lyr) - 1)
        x = (canvas_w - tw) // 2
        cy = y + max(hs, default=_C_NH) // 2
        for nid, w, h in zip(lyr, ws, hs):
            positions[nid] = (x + w // 2, cy)
            x += w + _C_H_PAD
        y += max(hs, default=_C_NH) + _C_V_GAP

    canvas_h = y - _C_V_GAP + 30

    # gradient defs
    grads: list[str] = []
    for ntype, (f_hex, _) in _STYLE.items():
        r, g, b = int(f_hex[0:2], 16), int(f_hex[2:4], 16), int(f_hex[4:6], 16)
        lt = f"{min(255,r+50):02X}{min(255,g+50):02X}{min(255,b+50):02X}"
        grads.append(
            f'<linearGradient id="g_{ntype}" x1="0" y1="0" x2="0" y2="1">'
            f'<stop offset="0%" stop-color="#{lt}"/>'
            f'<stop offset="100%" stop-color="#{f_hex}"/>'
            f'</linearGradient>'
        )

    defs = (
        '<defs>' + "".join(grads)
        + '<filter id="sh" x="-15%" y="-15%" width="130%" height="130%">'
        + '<feDropShadow dx="2" dy="2" stdDeviation="3" flood-color="rgba(0,0,0,0.18)"/>'
        + '</filter>'
        + '<marker id="arh" markerWidth="10" markerHeight="8" refX="9" refY="4" orient="auto">'
        + '<path d="M0,0 L10,4 L0,8 Z" fill="#4A4A4A"/></marker>'
        + '</defs>'
    )

    parts = [defs]

    # edges (bezier)
    for e in edges:
        s, t = e.get("from", ""), e.get("to", "")
        if s not in positions or t not in positions:
            continue
        sx, sy = positions[s]
        tx, ty = positions[t]
        sh = ndim(s)[1] // 2
        th = ndim(t)[1] // 2
        y1, y2 = sy + sh, ty - th
        my = (y1 + y2) // 2
        parts.append(
            f'<path d="M{sx},{y1} C{sx},{my} {tx},{my} {tx},{y2}" '
            f'fill="none" stroke="#4A4A4A" stroke-width="1.5" '
            f'opacity="0.75" marker-end="url(#arh)"/>'
        )
        lbl = e.get("label", "")
        if lbl:
            lx, ly = (sx + tx) // 2 + 7, (y1 + y2) // 2
            parts.append(
                f'<text x="{lx}" y="{ly}" fill="#333" '
                f'font-family="Calibri,Arial,sans-serif" font-size="9" '
                f'font-style="italic">{_html.escape(lbl)}</text>'
            )

    # nodes
    for node in nodes:
        nid = node["id"]
        if nid not in positions:
            continue
        cx, cy = positions[nid]
        ntype = node.get("type", "process")
        f_hex, b_hex = _STYLE.get(ntype, _STYLE["process"])
        w, h = ndim(nid)
        x, y = cx - w // 2, cy - h // 2
        gurl = f"url(#g_{ntype})"
        stroke = f"#{b_hex}"
        text = _html.escape(node.get("text", ""))
        filt = 'filter="url(#sh)"'

        if ntype == "decision":
            pts = f"{cx},{y} {x+w},{cy} {cx},{y+h} {x},{cy}"
            parts.append(f'<polygon points="{pts}" fill="{gurl}" stroke="{stroke}" stroke-width="2" {filt}/>')
        elif ntype in ("start", "end"):
            rx = h // 2
            parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" ry="{rx}" '
                         f'fill="{gurl}" stroke="{stroke}" stroke-width="2" {filt}/>')
        elif ntype == "io":
            sk = 10
            pts = f"{x+sk},{y} {x+w},{y} {x+w-sk},{y+h} {x},{y+h}"
            parts.append(f'<polygon points="{pts}" fill="{gurl}" stroke="{stroke}" stroke-width="2" {filt}/>')
        else:
            parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" '
                         f'fill="{gurl}" stroke="{stroke}" stroke-width="2" {filt}/>')

        parts.append(
            f'<text x="{cx}" y="{cy}" text-anchor="middle" dominant-baseline="middle" '
            f'fill="white" font-family="Calibri,Arial,sans-serif" '
            f'font-size="10" font-weight="bold">{text}</text>'
        )

    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'width="{canvas_w}" height="{canvas_h}" '
        f'viewBox="0 0 {canvas_w} {canvas_h}">'
        + "\n".join(parts) + '</svg>'
    )


# ═══════════════════════════════════════════════════════════════════════════════
# PNG conversion for complex flowchart (Word)
# ═══════════════════════════════════════════════════════════════════════════════

def build_complex_png_for_word(nodes: list[dict], edges: list[dict], tpl: dict) -> BytesIO:
    """SVG → PNG via cairosvg, else matplotlib fallback."""
    svg_str = build_complex_svg(nodes, edges, tpl)
    try:
        import cairosvg
        buf = BytesIO()
        cairosvg.svg2png(bytestring=svg_str.encode("utf-8"), write_to=buf, dpi=150)
        buf.seek(0)
        return buf
    except (ImportError, Exception):
        return _complex_matplotlib_png(nodes, edges, tpl)


def _complex_matplotlib_png(nodes: list[dict], edges: list[dict], tpl: dict) -> BytesIO:
    """Matplotlib fallback for complex flowchart PNG."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    import numpy as np

    by_id = {n["id"]: n for n in nodes}
    layers = _bfs_layers(nodes, edges)

    NW, NH = 2.0, 0.55
    DW, DH = 1.9, 0.80
    H_PAD, V_GAP = 0.45, 0.75

    def dims(nid: str) -> tuple[float, float]:
        return (DW, DH) if by_id[nid].get("type") == "decision" else (NW, NH)

    lw = [sum(dims(n)[0] for n in lyr) + H_PAD * max(0, len(lyr) - 1) for lyr in layers]
    total_w = max(lw, default=2.0) + 1.2
    total_h = sum(
        max(dims(n)[1] for n in lyr) for lyr in layers
    ) + V_GAP * len(layers) + 0.6

    fig_w = max(total_w * 1.1, 3.5)
    fig_h = max(total_h * 1.1, 2.0)
    dpi = 150
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.set_xlim(0, total_w)
    ax.set_ylim(0, total_h)
    ax.axis("off")
    ax.set_facecolor("white")
    fig.patch.set_facecolor("white")

    positions: dict[str, tuple[float, float]] = {}
    y = total_h - 0.4
    for lyr in layers:
        ws = [dims(n)[0] for n in lyr]
        hs = [dims(n)[1] for n in lyr]
        tw = sum(ws) + H_PAD * max(0, len(lyr) - 1)
        x = (total_w - tw) / 2
        cy = y - max(hs) / 2
        for nid, w, h in zip(lyr, ws, hs):
            positions[nid] = (x + w / 2, cy)
            x += w + H_PAD
        y -= max(hs) + V_GAP

    # edges
    for e in edges:
        s, t = e.get("from", ""), e.get("to", "")
        if s not in positions or t not in positions:
            continue
        sx, sy = positions[s]
        tx, ty = positions[t]
        sh = dims(s)[1] / 2
        th = dims(t)[1] / 2
        ax.annotate("", xy=(tx, ty + th + 0.02), xytext=(sx, sy - sh - 0.02),
                    arrowprops=dict(arrowstyle="->", color="#555", lw=1.2))
        lbl = e.get("label", "")
        if lbl:
            ax.text((sx + tx) / 2 + 0.08, (sy - sh + ty + th) / 2,
                    lbl, fontsize=6, color="#333", fontstyle="italic")

    # nodes
    for node in nodes:
        nid = node["id"]
        if nid not in positions:
            continue
        cx, cy = positions[nid]
        ntype = node.get("type", "process")
        f_hex, b_hex = _STYLE.get(ntype, _STYLE["process"])
        fill = f"#{f_hex}"
        border = f"#{b_hex}"
        w, h = dims(nid)

        if ntype == "decision":
            diamond = plt.Polygon(
                [[cx, cy + h/2], [cx + w/2, cy], [cx, cy - h/2], [cx - w/2, cy]],
                closed=True, facecolor=fill, edgecolor=border, linewidth=1.5,
                zorder=2
            )
            ax.add_patch(diamond)
        elif ntype in ("start", "end"):
            fancy = mpatches.FancyBboxPatch(
                (cx - w/2, cy - h/2), w, h,
                boxstyle=f"round,pad=0,rounding_size={h/2}",
                facecolor=fill, edgecolor=border, linewidth=1.5, zorder=2
            )
            ax.add_patch(fancy)
        else:
            rect = mpatches.FancyBboxPatch(
                (cx - w/2, cy - h/2), w, h,
                boxstyle="round,pad=0,rounding_size=0.04",
                facecolor=fill, edgecolor=border, linewidth=1.5, zorder=2
            )
            ax.add_patch(rect)

        ax.text(cx, cy, node.get("text", ""), ha="center", va="center",
                fontsize=7, color="white", fontweight="bold", zorder=3,
                wrap=True)

    fig.tight_layout(pad=0.3)
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, facecolor="white", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf
