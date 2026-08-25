import re
from pathlib import Path

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

W = NS["w"]
WP = NS["wp"]
A = NS["a"]
V = NS["v"]

WATERMARK_MARKER = "licensedto"
WATERMARK_PATTERNS = [
    re.compile(r"licensed\s*to.*?mecstudio\.com", re.IGNORECASE | re.DOTALL),
    re.compile(r"licensed\s*to", re.IGNORECASE),
    re.compile(r"drishtikhanna", re.IGNORECASE),
    re.compile(r"drishti\.khanna@mecstudio\.com", re.IGNORECASE),
]


def _norm_spans(full: str, pattern) -> list[tuple[int, int]]:
    idx = [i for i, ch in enumerate(full) if not ch.isspace()]
    norm = "".join(full[i] for i in idx)
    spans = []
    offset = 0
    while True:
        m = pattern.search(norm, offset)
        if m is None:
            break
        spans.append((idx[m.start()], idx[m.end() - 1] + 1))
        offset = m.end()
        if m.end() == m.start():
            offset += 1
    return spans

A4_W_TWIPS = 11906
A4_H_TWIPS = 16838
EMU_PER_TWIP = 635

REPLACEMENTS = [
    (re.compile(r"morningstar", re.IGNORECASE), "Company"),
    (re.compile(r"pitchbook", re.IGNORECASE), "Subsidiary"),
]

SHAPE_TAGS = {f"{{{W}}}pict", f"{{{W}}}drawing", f"{{{NS['mc']}}}AlternateContent"}


def _match_case(source: str, replacement: str) -> str:
    if source.isupper():
        return replacement.upper()
    if source.islower():
        return replacement.lower()
    return replacement


def _iter_parts(doc):
    yield doc.element
    for section in doc.sections:
        for part in (
            section.header,
            section.footer,
            section.first_page_header,
            section.first_page_footer,
            section.even_page_header,
            section.even_page_footer,
        ):
            try:
                yield part._element
            except Exception:
                continue


def remove_watermark_shapes(root) -> int:
    removed = 0
    for txbx in root.findall(f".//{{{W}}}txbxContent"):
        text = "".join(txbx.itertext())
        if not any(_norm_spans(text, p) for p in WATERMARK_PATTERNS):
            continue
        node = txbx
        target = None
        while node is not None:
            if node.tag in SHAPE_TAGS:
                target = node
            node = node.getparent()
        if target is not None and target.getparent() is not None:
            target.getparent().remove(target)
            removed += 1
    return removed


def _replace_in_text_nodes(t_nodes, full: str) -> int:
    count = 0
    for pattern, replacement in REPLACEMENTS:
        full = "".join(t.text or "" for t in t_nodes)
        matches = list(pattern.finditer(full))
        count += len(matches)
        for m in reversed(matches):
            rep = _match_case(m.group(0), replacement)
            _edit_span(t_nodes, full, m.start(), m.end(), rep)
    return count


def _edit_span(t_nodes, full: str, start: int, end: int, replacement: str):
    spans = []
    pos = 0
    for t in t_nodes:
        text = t.text or ""
        spans.append((t, pos, pos + len(text)))
        pos += len(text)
    first = True
    for t, s, e in spans:
        if e <= start or s >= end:
            continue
        ls = max(start, s) - s
        le = min(end, e) - s
        seg = t.text or ""
        if first:
            t.text = seg[:ls] + replacement + seg[le:]
            first = False
        else:
            t.text = seg[:ls] + seg[le:]


def _watermark_spans(full: str) -> list[tuple[int, int]]:
    spans = []
    for pattern in WATERMARK_PATTERNS:
        for span in _norm_spans(full, pattern):
            if not any(s <= span[0] < e or s < span[1] <= e for s, e in spans):
                spans.append(span)
    return sorted(spans)


def remove_watermark_paragraphs(root) -> int:
    count = 0
    for p in list(root.iter(f"{{{W}}}p")):
        t_nodes = p.findall(f".//{{{W}}}t")
        if not t_nodes:
            continue
        full = "".join(t.text or "" for t in t_nodes)
        spans = _watermark_spans(full)
        if not spans:
            continue
        for start, end in reversed(spans):
            _edit_span(t_nodes, full, start, end, "")
        count += len(spans)
        if not "".join(t.text or "" for t in p.findall(f".//{{{W}}}t")).strip():
            parent = p.getparent()
            if parent is not None and p.find(f".//{{{W}}}txbxContent") is None:
                parent.remove(p)
    return count


def _remove_empty_watermark_shapes(root) -> int:
    removed = 0
    for txbx in root.findall(f".//{{{W}}}txbxContent"):
        if "".join(txbx.itertext()).strip():
            continue
        node = txbx
        target = None
        while node is not None:
            if node.tag in SHAPE_TAGS:
                target = node
            node = node.getparent()
        if target is not None and target.getparent() is not None and len(target.findall(f".//{{{W}}}p")) <= 1:
            target.getparent().remove(target)
            removed += 1
    return removed


def replace_branded_words(root) -> int:
    count = 0
    for p in root.iter(f"{{{W}}}p"):
        t_nodes = p.findall(f".//{{{W}}}t")
        if not t_nodes:
            continue
        full = "".join(t.text or "" for t in t_nodes)
        joined = full.lower()
        if "morningstar" not in joined and "pitchbook" not in joined:
            continue
        count += _replace_in_text_nodes(t_nodes, full)
    return count


def _is_landscape(pg_sz) -> bool:
    orient = pg_sz.get(f"{{{W}}}orient")
    if orient == "landscape":
        return True
    if orient == "portrait":
        return False
    try:
        return int(pg_sz.get(f"{{{W}}}w", "0")) > int(pg_sz.get(f"{{{W}}}h", "1"))
    except ValueError:
        return False


def fix_orientation(root) -> int:
    fixed = 0
    for sect_pr in root.findall(f".//{{{W}}}sectPr"):
        pg_sz = sect_pr.find(f"{{{W}}}pgSz")
        if pg_sz is None or not _is_landscape(pg_sz):
            continue
        pg_sz.set(f"{{{W}}}w", str(A4_W_TWIPS))
        pg_sz.set(f"{{{W}}}h", str(A4_H_TWIPS))
        pg_sz.set(f"{{{W}}}orient", "portrait")
        fixed += 1
    return fixed


def _section_layout(sect_pr) -> dict:
    pg_sz = sect_pr.find(f"{{{W}}}pgSz")
    page_w = int(pg_sz.get(f"{{{W}}}w", str(A4_W_TWIPS))) if pg_sz is not None else A4_W_TWIPS
    pg_mar = sect_pr.find(f"{{{W}}}pgMar")
    left = right = 1440
    if pg_mar is not None:
        left = int(pg_mar.get(f"{{{W}}}left", "1440"))
        right = int(pg_mar.get(f"{{{W}}}right", "1440"))
    return {
        "page_w": page_w,
        "left": left,
        "usable": max(page_w - left - right, 1),
    }


def fit_content(root) -> int:
    elements = list(root.iter())
    index_of = {id(el): i for i, el in enumerate(elements)}
    sect_prs = [el for el in elements if el.tag == f"{{{W}}}sectPr"]
    if not sect_prs:
        return 0
    bounds = [index_of[id(s)] for s in sect_prs]
    layouts = [_section_layout(s) for s in sect_prs]

    import bisect

    def layout_for(el) -> dict:
        k = bisect.bisect_right(bounds, index_of[id(el)])
        return layouts[k] if k < len(layouts) else layouts[-1]

    scaled = 0
    for tag in ("inline", "anchor"):
        for drawing in root.findall(f".//{{{WP}}}{tag}"):
            extent = drawing.find(f"{{{WP}}}extent")
            if extent is None:
                continue
            cx = int(extent.get("cx", "0"))
            cy = int(extent.get("cy", "0"))
            usable = layout_for(drawing)["usable"] * EMU_PER_TWIP
            if cx <= usable or cx == 0:
                continue
            factor = usable / cx
            extent.set("cx", str(usable))
            extent.set("cy", str(int(cy * factor)))
            for ext in drawing.findall(f".//{{{A}}}ext"):
                ext.set("cx", str(int(int(ext.get("cx", "0")) * factor)))
                ext.set("cy", str(int(int(ext.get("cy", "0")) * factor)))
            scaled += 1

    def _in_group(el) -> bool:
        parent = el.getparent()
        while parent is not None:
            if parent.tag == f"{{{V}}}group":
                return True
            parent = parent.getparent()
        return False

    def _scale_style(el, usable_pt) -> bool:
        style = el.get("style", "")
        m_w = re.search(r"width:([\d.]+)pt", style)
        m_h = re.search(r"height:([\d.]+)pt", style)
        if not m_w or float(m_w.group(1)) <= usable_pt:
            return False
        factor = usable_pt / float(m_w.group(1))
        new_style = style
        if m_h:
            new_style = new_style.replace(m_h.group(0), f"height:{float(m_h.group(1)) * factor:.2f}pt")
        new_style = new_style.replace(m_w.group(0), f"width:{usable_pt:.2f}pt")
        el.set("style", new_style)
        return True

    def _fit_vml_position(el) -> bool:
        style = el.get("style", "")
        if "position:absolute" not in style:
            return False
        m_l = re.search(r"margin-left:([\d.]+)pt", style)
        m_w = re.search(r"width:([\d.]+)pt", style)
        m_h = re.search(r"height:([\d.]+)pt", style)
        if not (m_l and m_w and m_h) or float(m_w.group(1)) <= 0:
            return False
        lay = layout_for(el)
        left = float(m_l.group(1))
        width = float(m_w.group(1))
        height = float(m_h.group(1))
        page_relative = "mso-position-horizontal-relative:page" in style
        allowed_right = (lay["usable"] + (lay["left"] if page_relative else 0)) / 20
        if left + width <= allowed_right:
            return False
        avail = allowed_right - left
        if avail < width * 0.2:
            avail = allowed_right
            style = style.replace(m_l.group(0), "margin-left:0pt")
            left = 0.0
        factor = avail / width
        style = style.replace(m_w.group(0), f"width:{width * factor:.2f}pt")
        style = style.replace(m_h.group(0), f"height:{height * factor:.2f}pt")
        el.set("style", style)
        return True

    for group in root.findall(f".//{{{V}}}group"):
        if _in_group(group):
            continue
        changed = _fit_vml_position(group)
        if _scale_style(group, layout_for(group)["usable"] / 20):
            changed = True
        if changed:
            scaled += 1
    for shape in root.findall(f".//{{{V}}}shape"):
        if _in_group(shape):
            continue
        changed = _fit_vml_position(shape)
        if shape.find(f".//{{{V}}}imagedata") is not None and _scale_style(
            shape, layout_for(shape)["usable"] / 20
        ):
            changed = True
        if changed:
            scaled += 1

    for tbl in root.iter(f"{{{W}}}tbl"):
        cols = [int(c.get(f"{{{W}}}w", "0")) for c in tbl.findall(f".//{{{W}}}gridCol")]
        total = sum(cols)
        usable = layout_for(tbl)["usable"]
        if total <= usable or total == 0:
            continue
        factor = usable / total
        for c in tbl.findall(f".//{{{W}}}gridCol"):
            val = c.get(f"{{{W}}}w")
            if val:
                c.set(f"{{{W}}}w", str(int(int(val) * factor)))
        for tc_w in tbl.findall(f".//{{{W}}}tcW"):
            if tc_w.get(f"{{{W}}}type") == "dxa" and tc_w.get(f"{{{W}}}w"):
                tc_w.set(f"{{{W}}}w", str(int(int(tc_w.get(f"{{{W}}}w")) * factor)))
        tbl_w = tbl.find(f".//{{{W}}}tblW")
        if tbl_w is not None and tbl_w.get(f"{{{W}}}type") == "dxa" and tbl_w.get(f"{{{W}}}w"):
            tbl_w.set(f"{{{W}}}w", str(int(int(tbl_w.get(f"{{{W}}}w")) * factor)))
        scaled += 1
    return scaled


def sanitize_document(doc) -> dict:
    stats = {
        "watermarks_removed": 0,
        "words_replaced": 0,
        "sections_fixed": 0,
        "shapes_scaled": 0,
    }
    for root in _iter_parts(doc):
        stats["watermarks_removed"] += remove_watermark_shapes(root)
    for root in _iter_parts(doc):
        stats["watermarks_removed"] += remove_watermark_paragraphs(root)
        stats["watermarks_removed"] += _remove_empty_watermark_shapes(root)
    for root in _iter_parts(doc):
        stats["sections_fixed"] += fix_orientation(root)
    for root in _iter_parts(doc):
        stats["words_replaced"] += replace_branded_words(root)
        stats["shapes_scaled"] += fit_content(root)
    return stats
