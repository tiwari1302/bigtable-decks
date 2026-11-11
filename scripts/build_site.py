#!/usr/bin/env python3
"""Build the Bigtable decks site from the PowerPoint files in this folder.

Reads each .pptx, writes browser chapters under chapters/, and refreshes index.html.
Speaker notes become the text under each slide. The original files stay in place
for download. Re-run this after changing a deck; do not hand-edit chapters/.
"""

import colorsys
import html
import json
import os
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

SLIDE_W = 12192000
SLIDE_H = 6858000
DRAW = {"sp", "pic", "grpSp", "cxnSp", "graphicFrame"}


def local(tag):
    return tag.split("}")[-1]


def esc(text):
    return html.escape(text or "", quote=True)


def quote_path(name):
    return "/".join(html.escape(part, quote=True) for part in Path(name).parts).replace(" ", "%20")


class Frame:
    def __init__(self, ox=0.0, oy=0.0, sx=1.0, sy=1.0):
        self.ox, self.oy, self.sx, self.sy = ox, oy, sx, sy

    def nest(self, box):
        chw = box["chw"] or 1
        chh = box["chh"] or 1
        sx = box["w"] / chw
        sy = box["h"] / chh
        return Frame(
            self.ox + (box["x"] - box["chx"] * sx) * self.sx,
            self.oy + (box["y"] - box["chy"] * sy) * self.sy,
            self.sx * sx,
            self.sy * sy,
        )

    def rect(self, box):
        return (
            self.ox + box["x"] * self.sx,
            self.oy + box["y"] * self.sy,
            box["w"] * self.sx,
            box["h"] * self.sy,
        )


def xfrm_of(el):
    xfrm = None
    for path in (f"{P}spPr/{A}xfrm", f"{P}grpSpPr/{A}xfrm", f"{P}xfrm", f"{A}xfrm"):
        xfrm = el.find(path)
        if xfrm is not None:
            break
    if xfrm is None:
        return None
    off, ext = xfrm.find(f"{A}off"), xfrm.find(f"{A}ext")
    ch_off, ch_ext = xfrm.find(f"{A}chOff"), xfrm.find(f"{A}chExt")

    def num(node, attr, default=0):
        return int(node.get(attr, default)) if node is not None else default

    width, height = num(ext, "cx"), num(ext, "cy")
    return {
        "x": num(off, "x"),
        "y": num(off, "y"),
        "w": width,
        "h": height,
        "chx": num(ch_off, "x"),
        "chy": num(ch_off, "y"),
        "chw": num(ch_ext, "cx", width),
        "chh": num(ch_ext, "cy", height),
        "rot": int(xfrm.get("rot", "0")),
        "flipH": xfrm.get("flipH") == "1",
        "flipV": xfrm.get("flipV") == "1",
    }


def is_placeholder(el):
    return (
        el.find(f"{P}nvSpPr/{P}nvPr/{P}ph") is not None
        or el.find(f"{P}nvPicPr/{P}nvPr/{P}ph") is not None
    )


def load_scheme(zf):
    root = ET.fromstring(zf.read("ppt/theme/theme1.xml"))
    scheme = {}
    for child in root.find(f".//{A}clrScheme"):
        srgb = child.find(f"{A}srgbClr")
        sys = child.find(f"{A}sysClr")
        if srgb is not None:
            scheme[local(child.tag)] = srgb.get("val")
        elif sys is not None:
            scheme[local(child.tag)] = sys.get("lastClr")
    for alias, source in (("tx1", "dk1"), ("bg1", "lt1"), ("tx2", "dk2"), ("bg2", "lt2")):
        if source in scheme:
            scheme[alias] = scheme[source]
    return scheme


def rgb_of(hex_color):
    text = hex_color.strip().lstrip("#")
    return [int(text[i : i + 2], 16) / 255 for i in (0, 2, 4)]


def modulate(hex_color, node):
    r, g, b = rgb_of(hex_color)
    hue, light, sat = colorsys.rgb_to_hls(r, g, b)
    alpha = 1.0
    for child in list(node):
        tag = local(child.tag)
        val = int(child.get("val", "0")) / 100000
        if tag == "lumMod":
            light *= val
        elif tag == "lumOff":
            light += val
        elif tag == "satMod":
            sat *= val
        elif tag == "shade":
            r, g, b = [c * val for c in colorsys.hls_to_rgb(hue, light, sat)]
            hue, light, sat = colorsys.rgb_to_hls(r, g, b)
        elif tag == "tint":
            r, g, b = colorsys.hls_to_rgb(hue, light, sat)
            r, g, b = [c * val + (1 - val) for c in (r, g, b)]
            hue, light, sat = colorsys.rgb_to_hls(r, g, b)
        elif tag == "alpha":
            alpha = val
    light = max(0.0, min(1.0, light))
    sat = max(0.0, min(1.0, sat))
    r, g, b = colorsys.hls_to_rgb(hue, light, sat)
    color = "".join(f"{round(max(0, min(1, c)) * 255):02X}" for c in (r, g, b))
    return color, alpha


def solid_color(parent, scheme):
    if parent is None:
        return None
    node = parent if local(parent.tag) == "solidFill" else parent.find(f"{A}solidFill")
    if node is None:
        return None
    for child in list(node):
        tag = local(child.tag)
        if tag == "srgbClr":
            return modulate(child.get("val"), child)
        if tag == "schemeClr":
            base = scheme.get(child.get("val"), "17212B")
            return modulate(base, child)
    return None


def css_color(pair):
    if pair is None:
        return None
    color, alpha = pair
    if alpha >= 0.999:
        return f"#{color}"
    r, g, b = [int(color[i : i + 2], 16) for i in (0, 2, 4)]
    return f"rgba({r},{g},{b},{alpha:.3f})"


def pt_to_cqi(points):
    # 1cqi is 1% of the slide width. Slide width is 13.333in.
    return points / 72 / (SLIDE_W / 914400) * 100


def rels_of(zf, part):
    path = Path(part)
    rels_path = (path.parent / "_rels" / f"{path.name}.rels").as_posix()
    if rels_path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rels_path))
    return {rel.get("Id"): rel.get("Target") for rel in root}


def resolve(part, target):
    if target.startswith("/"):
        return target.lstrip("/")
    path = Path(part).parent
    for bit in Path(target).parts:
        if bit == "..":
            path = path.parent
        elif bit != ".":
            path = path / bit
    return path.as_posix()


def paragraphs(tx_body, scheme):
    if tx_body is None:
        return "", 0, "t", "left"
    body = tx_body.find(f"{A}bodyPr")
    anchor = body.get("anchor", "t") if body is not None else "t"
    font_scale = 1.0
    if body is not None:
        fit = body.find(f"{A}normAutofit")
        if fit is not None and fit.get("fontScale"):
            font_scale = int(fit.get("fontScale")) / 100000
    pads = []
    if body is not None:
        for attr in ("lIns", "tIns", "rIns", "bIns"):
            if body.get(attr):
                pads.append(f"{pt_to_cqi(int(body.get(attr)) / 12700):.3f}cqi")
            else:
                pads.append("0.4cqi")
    chunks = []
    max_size = 0
    plain = []
    for para in tx_body.findall(f"{A}p"):
        p_pr = para.find(f"{A}pPr")
        align = {"l": "left", "ctr": "center", "r": "right", "just": "justify"}.get(
            p_pr.get("algn", "l") if p_pr is not None else "l", "left"
        )
        before = 0
        bullet = ""
        if p_pr is not None:
            if p_pr.get("marL"):
                before = int(p_pr.get("marL"))
            bu = p_pr.find(f"{A}buChar")
            if bu is not None:
                bullet = (bu.get("char") or "•") + " "
            elif p_pr.find(f"{A}buAutoNum") is not None:
                bullet = "• "
        runs = []
        texts = []
        for child in list(para):
            tag = local(child.tag)
            if tag == "br":
                runs.append("<br>")
                continue
            if tag not in ("r", "fld"):
                continue
            r_pr = child.find(f"{A}rPr")
            size = 18
            if r_pr is not None and r_pr.get("sz"):
                size = int(r_pr.get("sz")) / 100 * font_scale
            max_size = max(max_size, size)
            text = "".join(node.text or "" for node in child.iter(f"{A}t"))
            texts.append(text)
            styles = [f"font-size:{pt_to_cqi(size):.3f}cqi"]
            if r_pr is not None:
                if r_pr.get("b") == "1":
                    styles.append("font-weight:700")
                if r_pr.get("i") == "1":
                    styles.append("font-style:italic")
                if r_pr.get("u") and r_pr.get("u") != "none":
                    styles.append("text-decoration:underline")
                if r_pr.get("spc"):
                    styles.append(f"letter-spacing:{pt_to_cqi(int(r_pr.get('spc')) / 100):.3f}cqi")
                face = r_pr.find(f"{A}latin")
                if face is not None:
                    styles.append(f"font-family:{font_family(face.get('typeface'))}")
                color = css_color(solid_color(r_pr, scheme))
                if color:
                    styles.append(f"color:{color}")
            runs.append(f'<span style="{";".join(styles)}">{esc(text)}</span>')
        if not runs:
            continue
        joined = "".join(texts).strip()
        if joined:
            plain.append(joined)
        margin = f"margin-top:{pt_to_cqi(before / 12700):.3f}cqi;" if before else ""
        prefix = esc(bullet)
        chunks.append(f'<p style="text-align:{align};{margin}">{prefix}{"".join(runs)}</p>')
    style = f"justify-content:{ {'t':'flex-start','ctr':'center','b':'flex-end'}.get(anchor,'flex-start') };padding:{' '.join(pads)}"
    return "".join(chunks), max_size, style, " ".join(plain)


def font_family(name):
    if not name or name.startswith("+"):
        return "Calibri,'Segoe UI',sans-serif"
    if name == "Cambria":
        return 'Cambria,Georgia,serif'
    if name == "Calibri":
        return "Calibri,'Segoe UI',sans-serif"
    return f"{name},'Segoe UI',sans-serif"


def shape_style(box, frame, extra):
    x, y, w, h = frame.rect(box)
    left, top, width, height = (100 * x / SLIDE_W, 100 * y / SLIDE_H, 100 * w / SLIDE_W, 100 * h / SLIDE_H)
    transforms = []
    if box["rot"]:
        transforms.append(f"rotate({box['rot'] / 60000:.2f}deg)")
    if box["flipH"]:
        transforms.append("scaleX(-1)")
    if box["flipV"]:
        transforms.append("scaleY(-1)")
    if transforms:
        extra.append("transform:" + " ".join(transforms))
    return left, top, width, height, extra


def geometry(sp_pr):
    if sp_pr is None:
        return "rect", ""
    geom = sp_pr.find(f"{A}prstGeom")
    if geom is None:
        return "rect", ""
    kind = geom.get("prst") or "rect"
    radius = ""
    if kind == "roundRect":
        adj = 16667
        guide = geom.find(f"{A}avLst/{A}gd")
        if guide is not None and guide.get("fmla", "").startswith("val "):
            adj = int(guide.get("fmla").split()[-1])
        radius = f"border-radius:{adj / 1000:.2f}%"
    elif kind == "ellipse":
        radius = "border-radius:50%"
    elif kind == "can":
        radius = "border-radius:50% / 18%"
    return kind, radius


def line_svg(box, color):
    x1, y1, x2, y2 = "0", "0", "100", "100"
    if box["flipH"]:
        x1, x2 = x2, x1
    if box["flipV"]:
        y1, y2 = y2, y1
    stroke = color or "#3D5163"
    return (
        '<svg viewBox="0 0 100 100" preserveAspectRatio="none" width="100%" height="100%">'
        f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{stroke}" stroke-width="2" vector-effect="non-scaling-stroke"/>'
        "</svg>"
    )


def chart_html(zf, part, rels, el, scheme):
    chart = el.find(f".//{C}chart")
    if chart is None:
        return ""
    target = rels.get(chart.get(f"{RID}id"))
    if not target:
        return ""
    root = ET.fromstring(zf.read(resolve(part, target)))
    plot = None
    kind = ""
    for name in ("barChart", "bar3DChart", "lineChart", "pieChart"):
        plot = root.find(f".//{C}{name}")
        if plot is not None:
            kind = name
            break
    if plot is None:
        return ""
    bar_dir = plot.find(f"{C}barDir")
    horizontal = bar_dir is None or bar_dir.get("val") != "col"
    rows = []
    color = "#2B53C8"
    for series in plot.findall(f"{C}ser"):
        fill = css_color(solid_color(series.find(f"{C}spPr"), scheme))
        if fill:
            color = fill
        cats = cached_points(series.find(f"{C}cat"))
        vals = cached_points(series.find(f"{C}val"))
        fmt = ""
        labels = series.find(f"{C}dLbls")
        if labels is not None and labels.find(f"{C}numFmt") is not None:
            fmt = labels.find(f"{C}numFmt").get("formatCode") or ""
        for index, raw in enumerate(vals):
            cat = cats[index] if index < len(cats) else ""
            try:
                number = float(raw)
            except (TypeError, ValueError):
                continue
            rows.append((cat, number, format_value(number, fmt)))
    if not rows:
        return ""
    peak = max(value for _, value, _ in rows) or 1
    font = max(0.85, min(1.35, 14 / max(len(rows), 1)))
    track = "0.9cqi" if horizontal else "70%"
    parts = [f'<div class="chart" style="font-size:{font:.2f}cqi">']
    for cat, value, label in rows:
        size = max(value / peak * 100, 0.8)
        if horizontal:
            parts.append(
                '<div class="bar-row">'
                f'<span class="cat">{esc(cat)}</span>'
                f'<span class="track" style="height:{track}"><span class="bar" style="width:{size:.2f}%;background:{color}"></span></span>'
                f"<span>{esc(label)}</span></div>"
            )
        else:
            parts.append(
                f'<div class="bar-row"><span class="cat">{esc(cat)}</span>'
                f'<span class="track" style="height:{track}"><span class="bar" style="width:{size:.2f}%;background:{color}"></span></span>'
                f"<span>{esc(label)}</span></div>"
            )
    parts.append("</div>")
    return "".join(parts)


def cached_points(node):
    if node is None:
        return []
    points = []
    for pt in node.iter(f"{C}pt"):
        idx = int(pt.get("idx", len(points)))
        while len(points) <= idx:
            points.append("")
        value = pt.find(f"{C}v")
        points[idx] = value.text if value is not None and value.text else ""
    return points


def format_value(number, fmt):
    if "TB" in fmt:
        return f"{number:.1f} TB" if number < 1 else f"{number:.0f} TB"
    if number.is_integer():
        return str(int(number))
    return f"{number:g}"


class DeckBuilder:
    def __init__(self, zf, deck_id, scheme):
        self.zf = zf
        self.deck_id = deck_id
        self.scheme = scheme
        self.copied = {}

    def image_src(self, part, rels, embed_id):
        target = rels.get(embed_id)
        if not target:
            return ""
        source = resolve(part, target)
        if source not in self.zf.namelist():
            return ""
        name = Path(source).name or "image.bin"
        data = self.zf.read(source)
        if not Path(name).suffix and data.startswith(b"\x89PNG"):
            name += ".png"
        dest = ROOT / "media" / self.deck_id / name
        if name not in self.copied:
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(data)
            self.copied[name] = True
        return f"../../media/{self.deck_id}/{quote_path(name)}"

    def render_tree(self, part, rels, tree, skip_placeholders):
        sp_tree = tree.find(f".//{P}spTree")
        if sp_tree is None:
            return [], ""
        html_parts = []
        titles = []

        def walk(el, frame):
            for child in list(el):
                if local(child.tag) not in DRAW:
                    continue
                if skip_placeholders and is_placeholder(child):
                    continue
                box = xfrm_of(child)
                if box is None:
                    continue
                if local(child.tag) == "grpSp":
                    walk(child, frame.nest(box))
                    continue
                piece, title = self.render_one(part, rels, child, box, frame)
                if piece:
                    html_parts.append(piece)
                if title:
                    titles.append(title)

        walk(sp_tree, Frame())
        best = ""
        if titles:
            titles.sort(key=lambda item: (-item[0], item[1]))
            best = titles[0][2]
        return html_parts, best

    def render_one(self, part, rels, el, box, frame):
        kind = local(el.tag)
        sp_pr = el.find(f"{P}spPr")
        styles = []
        fill = css_color(solid_color(sp_pr, self.scheme)) if sp_pr is not None else None
        if fill:
            styles.append(f"background:{fill}")
        geom, radius = geometry(sp_pr)
        if radius:
            styles.append(radius)
        if geom in ("chevron", "homePlate", "rightArrow"):
            styles.append("clip-path:" + clip_for(geom))
        line = sp_pr.find(f"{A}ln") if sp_pr is not None else None
        line_color = css_color(solid_color(line, self.scheme)) if line is not None else None
        inner = ""
        title = None
        if kind == "pic":
            blip = el.find(f".//{A}blip")
            src = self.image_src(part, rels, blip.get(f"{RID}embed")) if blip is not None else ""
            if src:
                inner = f'<img alt="" src="{src}">'
        elif kind == "graphicFrame":
            inner = chart_html(self.zf, part, rels, el, self.scheme)
        else:
            tx, size, text_style, plain = paragraphs(el.find(f"{P}txBody"), self.scheme)
            if tx:
                inner += f'<div class="txt" style="{text_style}">{tx}</div>'
                if plain and size >= 14 and len(plain) >= 8:
                    title = (size, box["y"], plain[:90])
            if geom == "line" or kind == "cxnSp":
                thin = min(box["w"], box["h"]) / max(box["w"], box["h"], 1) < 0.18
                if thin and line_color:
                    styles.append(f"background:{line_color}")
                elif not thin:
                    inner = line_svg(box, line_color) + inner
            elif line is not None and line_color and line.get("w"):
                width = pt_to_cqi(int(line.get("w")) / 12700)
                styles.append(f"border:{width:.3f}cqi solid {line_color}")
        if not inner and "background" not in ";".join(styles) and geom == "rect":
            return "", title
        left, top, width, height, styles = shape_style(box, frame, styles)
        if width <= 0 and height <= 0:
            return "", title
        style = f"left:{left:.3f}%;top:{top:.3f}%;width:{width:.3f}%;height:{height:.3f}%;" + ";".join(styles)
        return f'<div class="shape" style="{style}">{inner}</div>', title


def clip_for(kind):
    if kind == "rightArrow":
        return "polygon(0 22%, 62% 22%, 62% 0, 100% 50%, 62% 100%, 62% 78%, 0 78%)"
    if kind == "chevron":
        return "polygon(0 0, 72% 0, 100% 50%, 72% 100%, 0 100%, 28% 50%)"
    return "polygon(0 0, 78% 0, 100% 50%, 78% 100%, 0 100%)"


def background_css(tree, scheme):
    bg = tree.find(f"{P}cSld/{P}bg")
    if bg is None:
        return None
    solid = bg.find(f".//{A}solidFill")
    if solid is not None:
        return css_color(solid_color(solid, scheme))
    ref = bg.find(f"{P}bgRef")
    if ref is None:
        return None
    for child in list(ref):
        if local(child.tag) == "srgbClr":
            return css_color(modulate(child.get("val"), child))
        if local(child.tag) == "schemeClr":
            return css_color(modulate(scheme.get(child.get("val"), "FFFFFF"), child))
    return None


def notes_html(zf, part):
    if part not in zf.namelist():
        return "", ""
    root = ET.fromstring(zf.read(part))
    for sp in root.iter(f"{P}sp"):
        ph = sp.find(f"{P}nvSpPr/{P}nvPr/{P}ph")
        if ph is None or ph.get("type") != "body":
            continue
        tx = sp.find(f"{P}txBody")
        blocks = []
        first = ""
        for para in tx.findall(f"{A}p"):
            text = "".join(node.text or "" for node in para.iter(f"{A}t")).strip()
            if not text:
                continue
            if not first:
                first = text
            blocks.append(f"<p>{esc(text)}</p>")
        return "".join(blocks), first
    return "", ""


def slide_order(zf):
    rels = rels_of(zf, "ppt/presentation.xml")
    root = ET.fromstring(zf.read("ppt/presentation.xml"))
    order = []
    for sld in root.find(f"{P}sldIdLst"):
        target = rels.get(sld.get(f"{RID}id"))
        order.append(resolve("ppt/presentation.xml", target))
    size = root.find(f"{P}sldSz")
    return order, int(size.get("cx")), int(size.get("cy"))


def part_tree(zf, part, rels, builder, skip_placeholders):
    if part not in zf.namelist():
        return [], ""
    return builder.render_tree(part, rels, ET.fromstring(zf.read(part)), skip_placeholders)


def build_deck(pptx, deck_id, title, previous, nxt):
    with zipfile.ZipFile(pptx) as zf:
        scheme = load_scheme(zf)
        order, width, height = slide_order(zf)
        if (width, height) != (SLIDE_W, SLIDE_H):
            raise SystemExit(f"{pptx.name} has an unexpected slide size {width}x{height}")
        builder = DeckBuilder(zf, deck_id, scheme)
        pres_rels = rels_of(zf, "ppt/presentation.xml")
        master = next(target for rid, target in pres_rels.items() if "slideMaster" in (target or ""))
        master = resolve("ppt/presentation.xml", master)
        slides = []
        for index, part in enumerate(order, start=1):
            rels = rels_of(zf, part)
            layout = next(resolve(part, target) for target in rels.values() if target and "slideLayout" in target)
            pieces = []
            for source, skip in ((master, True), (layout, True), (part, False)):
                source_rels = rels if source == part else rels_of(zf, source)
                html_parts, heading = part_tree(zf, source, source_rels, builder, skip)
                pieces.extend(html_parts)
                if source == part:
                    slide_title = heading or f"Slide {index}"
            notes_target = next((resolve(part, target) for target in rels.values() if target and "notesSlide" in target), "")
            notes, _first = notes_html(zf, notes_target)
            slide_bg = background_css(ET.fromstring(zf.read(part)), scheme)
            if not slide_bg and layout in zf.namelist():
                slide_bg = background_css(ET.fromstring(zf.read(layout)), scheme)
            slides.append({
                "n": index,
                "title": slide_title,
                "html": "".join(pieces),
                "notes": notes,
                "bg": slide_bg or "#ffffff",
            })
        excerpt = ""
        if slides and slides[0]["notes"]:
            excerpt = ET.fromstring(f"<div>{slides[0]['notes']}</div>").findtext("p") or ""
        write_chapter(deck_id, title, pptx.name, slides, previous, nxt)
        return {"id": deck_id, "title": title, "slides": len(slides), "excerpt": excerpt, "download": pptx.name}


def write_chapter(deck_id, title, pptx_name, slides, previous, nxt):
    toc = []
    body = []
    for slide in slides:
        toc.append(f'<li><a href="#s{slide["n"]}">{slide["n"]}. {esc(slide["title"])}</a></li>')
        notes = f'<div class="notes">{slide["notes"]}</div>' if slide["notes"] else ""
        body.append(
            f'<section class="slide" id="s{slide["n"]}">'
            f'<p class="slabel">Slide {slide["n"]} of {len(slides)} · {esc(slide["title"])}</p>'
            f'<div class="stage" style="background:{esc(slide["bg"])}">{slide["html"]}</div>{notes}</section>'
        )
    prev_link = f'<a class="nav" href="../../chapters/{previous["id"]}/">{esc(previous["title"])}</a>' if previous else "<span></span>"
    next_link = f'<a class="nav" href="../../chapters/{nxt["id"]}/">{esc(nxt["title"])}</a>' if nxt else "<span></span>"
    page = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(deck_id)} {esc(title)} · Bigtable Decks</title>
<meta name="description" content="{esc(title)}, a chapter of Bigtable, Visually.">
<link rel="stylesheet" href="../../assets/site.css">
</head>
<body>
<header class="top">
  <div class="wrap">
    <div class="top-row">
      <div>
        <a class="nav" href="../../">All chapters</a>
        <h1>{esc(deck_id)} · {esc(title)}</h1>
      </div>
      <div class="actions">
        <a class="nav" href="../../{quote_path(pptx_name)}" download>Download PowerPoint</a>
      </div>
    </div>
    <ol class="toc">{''.join(toc)}</ol>
  </div>
</header>
<main class="wrap chapter-body">
  {''.join(body)}
  <div class="end">{prev_link}{next_link}</div>
</main>
<script src="../../assets/site.js"></script>
</body>
</html>
"""
    dest = ROOT / "chapters" / deck_id / "index.html"
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(page, encoding="utf-8")


def write_index(decks):
    cards = []
    for deck in decks:
        excerpt = esc(deck["excerpt"][:220] + ("…" if len(deck["excerpt"]) > 220 else ""))
        cards.append(
            '<li class="entry">'
            f'<a class="chapter" href="chapters/{deck["id"]}/">'
            f'<h2>{esc(deck["id"])} · {esc(deck["title"])}</h2>'
            f"<p>{excerpt}</p>"
            f'<div class="meta"><span>{deck["slides"]} slides</span></div></a>'
            f'<a class="dl" href="{quote_path(deck["download"])}" download>PowerPoint</a></li>'
        )
    page = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Bigtable Decks</title>
<meta name="description" content="Thirteen chapters on Google Bigtable. Read them in the browser, or download the PowerPoint.">
<link rel="stylesheet" href="assets/site.css">
</head>
<body>
<main class="wrap">
  <h1>Bigtable Decks</h1>
  <p class="lede">Thirteen chapters, in learning order. Open one to read it, or take the PowerPoint.</p>
  <ol class="chapters">
    {''.join(cards)}
  </ol>
</main>
</body>
</html>
"""
    (ROOT / "index.html").write_text(page, encoding="utf-8")
    (ROOT / "chapters.json").write_text(json.dumps(decks, indent=2) + "\n", encoding="utf-8")


def deck_title(path):
    stem = path.stem
    return stem.split(" ", 1)[1] if " " in stem else stem


def main():
    decks = sorted(path for path in ROOT.glob("[0-9]*.pptx") if not path.name.startswith("~$"))
    if len(decks) != 13:
        raise SystemExit(f"expected 13 decks, found {len(decks)}")
    for folder in (ROOT / "chapters", ROOT / "media"):
        if folder.exists():
            shutil.rmtree(folder)
    metas = []
    pending = [(f"{index:02d}", path, deck_title(path)) for index, path in enumerate(decks, start=1)]
    for index, (deck_id, path, title) in enumerate(pending):
        previous = {"id": pending[index - 1][0], "title": pending[index - 1][2]} if index else None
        nxt = {"id": pending[index + 1][0], "title": pending[index + 1][2]} if index + 1 < len(pending) else None
        meta = build_deck(path, deck_id, title, previous, nxt)
        metas.append(meta)
        print(f"{deck_id} {title}: {meta['slides']} slides")
    write_index(metas)


if __name__ == "__main__":
    main()
