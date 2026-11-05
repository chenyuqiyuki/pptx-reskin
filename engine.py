# -*- coding: utf-8 -*-
"""
北大红卡片风 PPT 自动排版引擎
流程：解析任意 pptx 的逐页内容 → 分类（封面/目录/章节/金句/表格/内容）→
     按参照件设计语言（红横幅+双徽皮肤、红竖条标题、圆角卡片、数字芯片、底部红条）
     自动生成成品。文字内容原样保留。
"""
import math, re, os
from copy import deepcopy
from io import BytesIO
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

EMU = 914400
def IN(v): return Emu(int(v * EMU))

RED   = RGBColor(0x8B, 0x00, 0x12)
DARK  = RGBColor(0x1F, 0x23, 0x29)
BODYC = RGBColor(0x5A, 0x5A, 0x5A)
SUB   = RGBColor(0x78, 0x78, 0x78)
GRAY  = RGBColor(0x8C, 0x8C, 0x8C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CARDC = RGBColor(0xFA, 0xF5, 0xF5)
FONT  = "Noto Sans SC"

SKIN_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "skin")
BASE = os.path.join(SKIN_DIR, "base.pptx")

# 浏览器版：base.pptx 以字节注入，不走文件系统
_base_bytes = None


def set_base_bytes(data):
    global _base_bytes
    _base_bytes = data


def _base_src():
    import io
    return io.BytesIO(_base_bytes) if _base_bytes is not None else BASE

FOOTER_RE = re.compile(r"^(来源|依据|资料来源|参考|注|文件编号|编号|生效日期)[:：]?")
PAGENO_RE = re.compile(r"^([Pp]\d{1,2}|\d{1,2}\s*/\s*\d{1,2})$")

# ---------------- 皮肤素材（从 base.pptx 提取一次，缓存到内存） ----------------
_skin = None
def skin():
    global _skin
    if _skin is not None:
        return _skin
    ref = Presentation(_base_src())
    sk = {"banner_xml": None, "badge_hos": None, "badge_dept": None,
          "band_top": None, "band_cloud": None, "ending_idx": len(ref.slides.__iter__.__self__._sldIdLst) - 1}
    s2 = ref.slides[1]
    for sh in s2.shapes:
        if sh.shape_type == 13:
            if (sh.left or 0) / EMU > 5:
                sk["badge_hos"] = (sh.image.blob, 7.93, 0.04, 2.09, 0.57)
            else:
                sk["badge_dept"] = (sh.image.blob, 0.13, 6.88, 1.97, 0.45)
        elif sh.shape_type == 5 and sk["banner_xml"] is None:
            sk["banner_xml"] = deepcopy(sh._element)
    s1 = ref.slides[0]
    for sh in s1.shapes:
        if sh.shape_type == 13:
            if (sh.top or 0) / EMU < 1:
                sk["band_top"] = (sh.image.blob, 0.0, 0.0, 10.0, 0.69)
            else:
                sk["band_cloud"] = (sh.image.blob, 0.0, 6.91, 10.0, 0.59)
    _skin = sk
    return _skin

# ---------------- 基础绘制组件（与参照件同款） ----------------
def set_font(r, size, color, bold):
    f = r.font
    f.size = Pt(size); f.bold = bold; f.color.rgb = color; f.name = FONT
    rPr = r._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {}); rPr.append(el)
        el.set('typeface', FONT)

def no_bullet(p):
    pPr = p._p.get_or_add_pPr()
    for tag in ('a:buNone', 'a:buChar', 'a:buAutoNum'):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    pPr.append(pPr.makeelement(qn('a:buNone'), {}))

def txt(slide, x, y, w, h, paras, anchor=None):
    tb = slide.shapes.add_textbox(IN(x), IN(y), IN(w), IN(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if anchor:
        tf.vertical_anchor = anchor
    first = True
    for pa in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        no_bullet(p)
        p.alignment = pa.get("a", PP_ALIGN.LEFT)
        p.line_spacing = pa.get("ls", 1.15)
        if pa.get("sb"):
            p.space_after = Pt(pa["sb"])
        r = p.add_run(); r.text = pa["t"] or ""
        set_font(r, pa.get("s", 13), pa.get("c", BODYC), pa.get("b", False))
    return tb

def shape(slide, kind, x, y, w, h, fill=None, line_hex=None, line_alpha=None, line_w=12700, adj=None):
    sp = slide.shapes.add_shape(kind, IN(x), IN(y), IN(w), IN(h))
    sp.shadow.inherit = False
    if adj is not None:
        try: sp.adjustments[0] = adj
        except Exception: pass
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line_hex is None:
        sp.line.fill.background()
    else:
        ln = sp._element.spPr.find(qn('a:ln'))
        if ln is None:
            ln = sp._element.spPr.makeelement(qn('a:ln'), {'w': str(line_w)})
            sp._element.spPr.append(ln)
        ln.set('w', str(line_w))
        for ch in list(ln):
            if ch.tag in (qn('a:solidFill'), qn('a:noFill'), qn('a:prstDash')):
                ln.remove(ch)
        sf = parse_xml(
            '<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:srgbClr val="{line_hex}"><a:alpha val="{int(line_alpha*1000)}"/></a:srgbClr></a:solidFill>')
        ln.insert(0, sf)
    return sp

def card(slide, x, y, w, h, strip=True):
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=CARDC, line_hex="8B0012", line_alpha=15, adj=0.10)
    if strip:
        shape(slide, MSO_SHAPE.RECTANGLE, x + 0.12, y, w - 0.24, 0.055, fill=RED)

def chip(slide, x, y, label, d=0.30, size=11.5):
    sp = shape(slide, MSO_SHAPE.OVAL, x, y, d, d, fill=RED)
    tf = sp.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.word_wrap = False
    p = tf.paragraphs[0]; no_bullet(p)
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = label
    set_font(r, size, WHITE, True)
    return sp

def bar(slide, text, y=5.94, h=0.58, size=13):
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, 0.69, y, 8.61, h, fill=RED, adj=0.22)
    txt(slide, 0.95, y + (h - 0.36) / 2, 8.1, 0.4,
        [{"t": text, "s": size, "c": WHITE, "b": True}], anchor=MSO_ANCHOR.MIDDLE)

def footer(slide, text, y=6.61):
    txt(slide, 0.69, y, 8.61, 0.28, [{"t": text, "s": 11, "c": GRAY}])

def chrome(slide, cover=False):
    sk = skin()
    if cover:
        slide.shapes.add_picture(BytesIO(sk["band_top"][0]), IN(0), IN(0), IN(10), IN(0.69))
        slide.shapes.add_picture(BytesIO(sk["band_cloud"][0]), IN(0), IN(6.91), IN(10), IN(0.59))
    else:
        slide.shapes._spTree.append(deepcopy(sk["banner_xml"]))
        b = sk["badge_hos"]
        slide.shapes.add_picture(BytesIO(b[0]), IN(b[1]), IN(b[2]), IN(b[3]), IN(b[4]))
        b = sk["badge_dept"]
        slide.shapes.add_picture(BytesIO(b[0]), IN(b[1]), IN(b[2]), IN(b[3]), IN(b[4]))

def head(slide, title, sub=None, title_size=26, title_y=0.84, title_h=0.52, sub_y=1.36, region_top=1.72):
    # 标题过长自动缩小
    ts = title_size
    while ts > 17 and _est_h([title], ts, 8.35) > title_h + (0.45 if not sub else 0.0):
        ts -= 1
    shape(slide, MSO_SHAPE.RECTANGLE, 0.78, title_y + 0.03, 0.045, 0.36, fill=RED)
    txt(slide, 0.95, title_y, 8.4, title_h, [{"t": title, "s": ts, "c": RED, "b": True, "ls": 1.05}])
    if sub:
        txt(slide, 0.95, sub_y, 8.4, 0.32, [{"t": sub, "s": 13, "c": SUB}])

def num_row(slide, y, h, num, title, desc=None, title_size=15.5, desc_size=12.5, tag=None):
    card(slide, 0.69, y, 8.61, h, strip=False)
    chip(slide, 0.92, y + (h - 0.30) / 2, num)
    tx = 1.42
    if desc:
        txt(slide, tx, y + 0.10, 6.6, 0.34, [{"t": title, "s": title_size, "c": DARK, "b": True}])
        txt(slide, tx, y + 0.10 + 0.36, 7.6, h - 0.5, [{"t": desc, "s": desc_size, "c": BODYC, "ls": 1.2}])
    else:
        txt(slide, tx, y + (h - 0.34) / 2, 7.6, 0.36, [{"t": title, "s": title_size, "c": DARK, "b": True}])
    if tag:
        txt(slide, 7.0, y + 0.12, 2.1, 0.3, [{"t": tag, "s": 11, "c": GRAY, "a": PP_ALIGN.RIGHT}])

def divider(slide, num, title, sub):
    chrome(slide)
    txt(slide, 3.0, 1.75, 4.0, 1.2, [{"t": num, "s": 64, "c": RED, "b": True, "a": PP_ALIGN.CENTER}])
    txt(slide, 1.0, 3.25, 8.0, 0.7, [{"t": title, "s": 36, "c": DARK, "b": True, "a": PP_ALIGN.CENTER}])
    if sub:
        txt(slide, 1.0, 4.15, 8.0, 0.4, [{"t": sub, "s": 15, "c": SUB, "a": PP_ALIGN.CENTER}])

def quote_page(slide, kicker, quote_lines, tail=None):
    chrome(slide)
    if kicker:
        txt(slide, 0.83, 2.05, 8.3, 0.35, [{"t": kicker, "s": 13, "c": SUB, "b": True}])
    qs = quote_lines
    size = 30
    while size > 18 and _est_h(qs, size, 8.3) > 2.6:
        size -= 1
    txt(slide, 0.83, 2.6, 8.4, 2.4, [{"t": l, "s": size, "c": RED, "b": True, "ls": 1.25} for l in qs])
    if tail:
        txt(slide, 0.83, 5.45, 8.4, 0.9, [{"t": tail, "s": 15, "c": BODYC, "ls": 1.3}])

# ---------------- 文本量估算（防溢出） ----------------
def _wlen(t):
    return sum(1.0 if ord(c) > 0x2E80 else 0.55 for c in t)

def _est_h(lines, size, w_in, ls=1.32, sb=0):
    cap = max(1.0, w_in * 72 / size)
    total = 0.0
    for t in lines:
        n = max(1, math.ceil(_wlen(t) / cap))
        total += n * size * ls / 72
    return total + sb * max(0, len(lines) - 1) / 72.0

# ---------------- 解析源 PPT ----------------
def _iter_blocks(shapes, dx=0.0, dy=0.0):
    for sh in shapes:
        st = sh.shape_type
        try: stn = int(st)
        except Exception: stn = None
        if stn == 6:  # group
            gx = gy = 0.0
            try:
                xfrm = sh._element.find(qn('p:grpSpPr')).find(qn('a:xfrm'))
                off = xfrm.find(qn('a:off')); ch = xfrm.find(qn('a:chOff'))
                gx = (int(off.get('x')) - int(ch.get('x'))) / EMU
                gy = (int(off.get('y')) - int(ch.get('y'))) / EMU
            except Exception:
                pass
            yield from _iter_blocks(sh.shapes, dx + gx, dy + gy)
            continue
        if getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip():
            paras, sizes, bolds = [], [], []
            for p in sh.text_frame.paragraphs:
                t = ("".join(r.text for r in p.runs) or p.text).strip()
                if not t: continue
                paras.append(t)
                ss = [r.font.size.pt for r in p.runs if r.font.size]
                sizes.append(max(ss) if ss else None)
                bolds.append(any(r.font.bold for r in p.runs))
            if not paras: continue
            szs = [s for s in sizes if s]
            yield {"kind": "text", "paras": paras, "sizes": sizes, "bolds": bolds,
                   "size": max(szs) if szs else 14.0,
                   "x": (sh.left or 0) / EMU + dx, "y": (sh.top or 0) / EMU + dy,
                   "w": (sh.width or 0) / EMU, "h": (sh.height or 0) / EMU}
        elif getattr(sh, "has_table", False):
            tbl = [[c.text.strip() for c in row.cells] for row in sh.table.rows]
            yield {"kind": "table", "table": tbl,
                   "x": (sh.left or 0) / EMU, "y": (sh.top or 0) / EMU,
                   "w": (sh.width or 0) / EMU, "h": (sh.height or 0) / EMU}
        elif stn == 13:  # picture
            w = (sh.width or 0) / EMU; h = (sh.height or 0) / EMU
            if w > 1.2 and h > 1.2:  # 忽略小装饰图
                yield {"kind": "image", "blob": sh.image.blob,
                       "x": (sh.left or 0) / EMU, "y": (sh.top or 0) / EMU,
                       "w": w, "h": h}

def parse_deck(path):
    prs = Presentation(path)
    slides = []
    for s in prs.slides:
        blocks = sorted(_iter_blocks(s.shapes), key=lambda b: (round(b["y"] * 2) / 2, b["x"]))
        slides.append(blocks)
    return slides, round(prs.slide_width / EMU, 2), round(prs.slide_height / EMU, 2)

# ---------------- 页面分类与内容整形 ----------------
def _norm_page(blocks):
    """把一页的 blocks 整形成统一结构"""
    texts = [b for b in blocks if b["kind"] == "text"]
    tables = [b for b in blocks if b["kind"] == "table"]
    images = [b for b in blocks if b["kind"] == "image"]
    footer_txt = None
    flat = []  # (size, bold, text, y, x) 按阅读顺序
    for b in texts:
        for i, t in enumerate(b["paras"]):
            size = b["sizes"][i] if i < len(b["sizes"]) and b["sizes"][i] else b["size"]
            bold = b["bolds"][i] if i < len(b["bolds"]) else False
            t = t.strip()
            if PAGENO_RE.match(t):
                continue
            if not re.search(r"[\w\u4e00-\u9fff]", t):  # 纯符号行（→ · 等）跳过
                continue
            if footer_txt is None and size <= 12 and FOOTER_RE.match(t):
                footer_txt = t
                continue
            flat.append((size, bold, t, b["y"], b["x"]))
    maxsize = max((f[0] for f in flat), default=0)
    return {"flat": flat, "maxsize": maxsize, "tables": tables, "images": images,
            "footer": footer_txt, "nchars": sum(len(f[2]) for f in flat)}

def classify(pages):
    kinds = []
    for i, pg in enumerate(pages):
        n = pg
        if i == 0:
            kinds.append("cover"); continue
        if n["nchars"] > 0 and re.search(r"感谢|谢谢|THANK", "".join(f[2] for f in n["flat"])) and len(n["flat"]) <= 5:
            kinds.append("ending"); continue
        if n["maxsize"] >= 36 and len(n["flat"]) <= 4:
            kinds.append("section"); continue
        if n["tables"]:
            kinds.append("table"); continue
        if len(n["flat"]) <= 6 and n["maxsize"] >= 24 and n["nchars"] < 140:
            kinds.append("quote"); continue
        kinds.append("content")
    return kinds

def _split_groups(flat, body_median):
    """把正文行分成组：组头=更大/加粗/以冒号结尾；返回 [(header, [items])]；无组头时 header=None"""
    if not flat:
        return [], []
    med = body_median
    groups, cur, headers = [], None, []
    def is_head(size, bold, t):
        if size >= med * 1.18: return True
        if bold and size >= med and len(t) <= 30: return True
        if t.endswith(("：", ":")) and len(t) < 30: return True
        return False
    heads = 0
    for size, bold, t, _y, _x in flat:
        if is_head(size, bold, t):
            heads += 1
            groups.append((t, []))
        else:
            if not groups:
                groups.append((None, []))
            groups[-1][1].append(t)
    # 若几乎全是组头（没有正文），当作纯列表——按原顺序保留所有行，避免丢掉夹在组头间的正文
    if heads >= len(flat) - 1:
        return [], [f[2] for f in flat]
    return groups, []

def plan_and_build(src_path, out_path, progress=None):
    """路径版入口：读文件 → 生成 → 存盘。浏览器版请用 build_bytes。"""
    with open(src_path, "rb") as f:
        data = f.read()
    return build_bytes(data, out_path, progress)


def build_bytes(src_data, out_path=None, progress=None):
    """字节版入口：src_data 为 pptx 字节；out_path 为 None 时返回输出字节。"""
    import io
    if isinstance(src_data, (bytes, bytearray)):
        src_data = io.BytesIO(src_data)
    raw, sw, sh_ = parse_deck(src_data)
    pages = [_norm_page(b) for b in raw]
    kinds = classify(pages)

    def tick(m):
        if progress: progress(m)

    tick("解析完成，规划版式…")

    # ---- 收集结构 ----
    cover = pages[0]
    section_pages = [(i, pages[i]) for i, k in enumerate(kinds) if k == "section"]

    def section_title(pg):
        """章节页：大数字行=num，标题=次大字号行（同字号相邻行合并），其余=sub"""
        flat = pg["flat"]
        big = max(flat, key=lambda f: f[0])
        if re.fullmatch(r"[0-9０-９]{1,2}|[一二三四五六七八九十]{1,3}|第?[0-9]+[章节部分]", big[2]):
            num = big[2]
            rest = [f for f in flat if f[2] != num]
            if not rest:
                return num, num, None
            t2 = max(rest, key=lambda f: f[0])[0]
            title = " ".join(t for s, b, t, _y, _x in flat if s >= t2 * 0.92 and t != num)[:30]
            subs = [t for s, b, t, _y, _x in flat if s < t2 * 0.92 and t != num]
            return num, title, " · ".join(subs)[:60] if subs else None
        return None, big[2], " · ".join(t for s, b, t, _y, _x in flat if t != big[2])[:60]

    toc_entries = []
    sec_idx = 0
    for i, pg in section_pages:
        sec_idx += 1
        num, title, sub = section_title(pg)
        toc_entries.append((num or f"{sec_idx:02d}", title, sub))
    if not toc_entries:  # 没有章节页 → 用内容页标题当目录
        n = 0
        for i, k in enumerate(kinds):
            if k in ("content", "table") and n < 8:
                flat = pages[i]["flat"]
                if not flat: continue
                title = max(flat, key=lambda f: f[0])[2]
                if len(title) <= 24:
                    n += 1
                    toc_entries.append((f"{n:02d}", title, None))
    # 原稿自带目录页 → 不再生成自动目录
    has_toc_page = any(
        k == "content" and pg["flat"] and any(
            re.search(r"目录|CONTENTS", t) and s >= pg["maxsize"] * 0.35
            for s, b, t, _y, _x in pg["flat"])
        for pg, k in zip(pages, kinds))
    if has_toc_page:
        toc_entries = []

    # ---- 开始生成 ----
    prs = Presentation(_base_src())
    blank = prs.slides[1].slide_layout
    add = prs.slides.add_slide
    n_made = 0

    def new_slide():
        nonlocal n_made
        n_made += 1
        return add(blank)

    # 封面
    tick("生成封面…")
    flat = sorted(cover["flat"], key=lambda f: -f[0])
    title_lines, rest = [], []
    if flat:
        tmax = flat[0][0]
        for s, b, t, _y, _x in flat:
            if s >= tmax * 0.72 and len(title_lines) < 3:
                title_lines.append(t)
            else:
                rest.append((s, b, t, _y, _x))
    rest.sort(key=lambda f: (-f[0], f[3]))
    subtitle = rest[0][2] if rest and rest[0][0] >= 15 else None
    if subtitle in title_lines: subtitle = None
    # 「文件编号…」小字并入上方同栏的制度/条目行，避免信息丢失
    rest_items = rest[1 if subtitle else 0:] or rest
    merged = []
    for s_, b_, t_, y_, x_ in rest_items:
        if t_.startswith(("文件编号", "编号")) and merged:
            best = None
            for k in range(len(merged)):
                if abs(merged[k][4] - x_) < 1.2 and y_ > merged[k][3]:
                    if best is None or merged[k][3] > merged[best][3]:
                        best = k
            if best is not None:
                merged[best] = (merged[best][0], merged[best][1],
                                f"{merged[best][2]}（{t_}）", merged[best][3], merged[best][4])
            else:
                merged.append((s_, b_, t_, y_, x_))
        else:
            merged.append((s_, b_, t_, y_, x_))

    def _lead_no(t_):
        m = re.match(r"^(\d{1,2})(?=\D|$)", t_)
        return int(m.group()) if m else 99

    merged.sort(key=lambda m: (_lead_no(m[2]), m[3], m[4]))
    infos = [m[2] for m in merged]
    kicker = next((t for s, b, t, _y, _x in cover["flat"]
                   if s <= 15 and t not in title_lines and t not in infos and t != subtitle
                   and not t.startswith(("文件编号", "编号"))
                   and not re.match(r"^\d{1,2}\s", t)), None)
    ft = cover.get("footer")
    if ft and ft.startswith("文件编号"):
        for k, m in enumerate(merged):
            if re.match(r"^01\s", m[2]) and "文件编号" not in m[2]:
                merged[k] = (m[0], m[1], f"{m[2]}（{ft}）", m[3], m[4])
                infos = [mm[2] for mm in merged]
                break
    s = new_slide(); chrome(s, cover=True)
    y = 1.45
    if kicker and kicker != subtitle:
        txt(s, 1.25, y, 7.5, 0.35, [{"t": kicker, "s": 14, "c": SUB, "b": True}]); y += 0.5
    tsize = 40
    while tsize > 26 and _est_h(title_lines, tsize, 7.5) > 2.2: tsize -= 1
    txt(s, 1.25, y, 7.5, 2.2, [{"t": t, "s": tsize, "c": RED, "b": True, "ls": 1.1} for t in title_lines])
    y += _est_h(title_lines, tsize, 7.5, ls=1.18) + 0.25
    if subtitle:
        txt(s, 1.25, y, 7.5, 0.45, [{"t": subtitle, "s": 17, "c": DARK}]); y += 0.6
    if infos:
        n_i = min(len(infos), 9)
        i_size = 12 if n_i <= 6 else 11
        txt(s, 1.25, min(y + 0.35, 5.05), 7.5, 1.9,
            [{"t": t, "s": i_size, "c": GRAY, "ls": 1.25, "sb": 3} for t in infos[:9]])

    # 目录
    if toc_entries:
        tick("生成目录…")
        s = new_slide(); chrome(s)
        head(s, "目录", "CONTENTS")
        y = 1.85
        gap = min(0.92, 4.4 / max(1, len(toc_entries)))
        rh = min(0.72, gap - 0.18)
        for num, t, d in toc_entries:
            if d:
                card(s, 0.69, y, 8.61, rh, strip=False)
                chip(s, 0.92, y + (rh - 0.30) / 2, num)
                txt(s, 1.42, y + 0.07, 7.7, 0.3, [{"t": t, "s": 15, "c": DARK, "b": True}])
                txt(s, 1.42, y + rh - 0.34, 7.7, 0.3, [{"t": d, "s": 11, "c": BODYC}])
            else:
                num_row(s, y, rh, num, t, title_size=16)
            y += gap

    # 逐页
    sec_no = 0
    total = len(pages)
    for i, (pg, kind) in enumerate(zip(pages, kinds)):
        if kind == "cover":
            continue
        tick(f"排版 {i+1}/{total} …")
        if kind == "ending":
            continue  # 用模板自带结束页
        if kind == "section":
            sec_no += 1
            num, title, sub = section_title(pg)
            s = new_slide()
            divider(s, num or f"{sec_no:02d}", title, sub)
            continue
        if kind == "quote":
            qs_txt, other = [], []
            for s_, b, t, _y, _x in pg["flat"]:  # 保持阅读顺序
                (qs_txt if s_ >= 22 else other).append(t)
            kicker = other[0] if other else None
            tail = "　".join(other[1:]) if len(other) > 1 else None
            s = new_slide()
            quote_page(s, kicker, qs_txt[:4], tail)
            continue
        # content / table → 内容页
        _content_slide(new_slide, pg, kind)
        continue

    # 收尾：删除 base 前 n_base-1 页（保留最后一页=模板结束页），再把结束页挪到最后
    lst = prs.slides._sldIdLst
    ids = list(lst)
    n_base = len(ids) - n_made
    for idx in range(n_base - 2, -1, -1):
        sid = ids[idx]
        rId = sid.get(qn('r:id'))
        prs.part.drop_rel(rId)
        lst.remove(sid)
    sid = list(lst)[0]          # 模板结束页挪到最后
    lst.remove(sid); lst.append(sid)
    if out_path:
        prs.save(out_path)
        tick("完成")
        return n_made + 1
    import io
    buf = io.BytesIO()
    prs.save(buf)
    tick("完成")
    return buf.getvalue()

# ---------------- 内容页排版（引擎核心判断） ----------------
BODY_TOP, BODY_BOTTOM = 1.72, 6.45
BODY_X, BODY_W = 0.69, 8.61

def _to_int(s_):
    s2 = s_.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    m = re.search(r"\d+", s2)
    if m:
        return int(m.group())
    cjk = "一二三四五六七八九十"
    m = re.search(r"[一二三四五六七八九十]{1,3}", s_)
    if m:
        t_ = m.group()
        if t_ == "十":
            return 10
        if "十" in t_:
            a, _, b_ = t_.partition("十")
            return (cjk.index(a) + 1 if a else 1) * 10 + (cjk.index(b_) + 1 if b_ else 0)
        return cjk.index(t_) + 1
    return 0

_NUM_ANCHOR_RE = re.compile(r"[0-9０-９]{1,2}|[一二三四五六七八九十]{1,3}")

def _analyze_numbered(flat):
    """检测原稿自带编号的条目结构：独立编号锚点（粗体、≥14pt）+ 右侧标题 [+ 下方描述]。
    返回 None 或 {"cols":[{x,header,notes,items:[{no,ti,desc,...}]}], "left":[未占用行]}"""
    idxs = [i for i, (sz, b, t, y, x) in enumerate(flat)
            if _NUM_ANCHOR_RE.fullmatch(t) and b and sz >= 14]
    if len(idxs) < 2:
        return None
    used = set(idxs)
    raw = []
    for i in idxs:
        sz, b, t, y, x = flat[i]
        best = None
        for j, (sz2, b2, t2, y2, x2) in enumerate(flat):
            if j in used or _NUM_ANCHOR_RE.fullmatch(t2):
                continue
            dx = x2 - x
            if 0.15 <= dx <= 1.3 and abs(y2 - y) <= 0.45:
                key = (round(abs(y2 - y), 2), round(dx, 2))
                if best is None or key < best[0]:
                    best = (key, j)
        if best is None:
            continue
        j = best[1]
        used.add(j)
        sz2, b2, t2, y2, x2 = flat[j]
        raw.append({"no": t, "ti": t2, "ti_sz": sz2, "ti_x": x2, "y": y, "x": x, "desc": ""})
    if len(raw) < 2:
        return None
    # 按锚点 x 聚栏
    xs = sorted(set(round(it["x"], 1) for it in raw))
    clusters = [[xs[0]]]
    for x in xs[1:]:
        if x - clusters[-1][-1] > 1.2:
            clusters.append([x])
        else:
            clusters[-1].append(x)
    centers = [sum(c) / len(c) for c in clusters]
    for it in raw:
        it["col"] = min(range(len(centers)), key=lambda ci: abs(centers[ci] - it["x"]))
    cols = []
    for ci in range(len(centers)):
        its = sorted([it for it in raw if it["col"] == ci], key=lambda it: it["y"])
        cols.append({"x": centers[ci], "items": its, "header": None, "notes": []})
    # 描述行：本栏区域内、位于本条与下一条之间的小字行（含子弹点）
    for c in cols:
        its = c["items"]
        for k, it in enumerate(its):
            y_end = its[k + 1]["y"] - 0.05 if k + 1 < len(its) else it["y"] + 1.9
            didx = [i2 for i2, f in enumerate(flat)
                    if i2 not in used
                    and c["x"] - 2.0 <= f[4] <= c["x"] + 2.0
                    and it["y"] + 0.3 < f[3] < y_end
                    and f[0] < max(it["ti_sz"], 17)
                    and not re.match(r"^(执行判断|结论|提示|注意|小结|易错点)", f[2])]
            if didx:
                ds = [flat[i2][2] for i2 in didx]
                it["desc"] = "".join(ds) if all(len(d) <= 22 for d in ds) else "；".join(ds)
                used.update(didx)
    leftovers = [f for i2, f in enumerate(flat) if i2 not in used]
    # 栏头：正文区内、首个条目之上、贴栏 x 的小字短行
    for c in cols:
        its = c["items"]
        if not its:
            continue
        for f in leftovers:
            if (abs(f[4] - c["x"]) <= 0.3 and 1.5 < f[3] < its[0]["y"] - 0.05
                    and f[0] <= 14 and len(f[2]) <= 14):
                c["header"] = (c["header"] + " " + f[2]) if c["header"] else f[2]
        taken = {c["header"]} if c["header"] else set()
        # 栏尾备注：最后一条之后的小字行（纯结论标签除外，留给红条逻辑）
        for f in leftovers:
            if (abs(f[4] - c["x"]) <= 0.3 and its[-1]["y"] + 0.3 < f[3]
                    and f[0] <= 14.5 and f[2] not in taken
                    and not re.fullmatch(r"(执行判断|结论|提示|注意|小结|易错点)", f[2].strip())):
                c["notes"].append(f[2])
    used_txt = ({c["header"] for c in cols if c["header"]} |
                {n for c in cols for n in c["notes"]} |
                {it["ti"] for c in cols for it in c["items"]})
    left = [f for f in leftovers if f[2] not in used_txt]
    # 消化率守门：未纳入结构的正文过多，说明版式判断不可信
    used_chars = sum(len(it["ti"]) + len(it.get("desc", "")) for c in cols for it in c["items"])
    left_chars = sum(len(f[2]) for f in left)
    if used_chars < left_chars * 1.8:
        return None
    return {"cols": cols, "left": left}

def _render_numbered(s, num, top, avail_h, body_w):
    """原稿自带编号页的渲染：左声明面板 + 编号行栈，或多栏卡片（芯片用原稿编号）"""
    cols = num["cols"]
    items = sorted((it for c in cols for it in c["items"]),
                   key=lambda it: _to_int(it["no"]))
    xmin = min(c["x"] for c in cols)
    lz = [f for f in num["left"] if f[4] < xmin - 1.5]
    stmt = sorted([f for f in lz if 11 < f[0] < 24], key=lambda f: f[3])
    kick = sorted([f for f in lz if f[0] <= 11 and f[3] < 3.0], key=lambda f: f[3])
    tag = [f for f in lz if f[0] <= 12 and f[3] >= 5.7]

    if len(cols) >= 2:
        # 多栏卡片：芯片保留原稿编号
        n = len(cols)
        gut = 0.17
        cw = (body_w - gut * (n - 1)) / n
        h = avail_h - 0.05
        for ci, c in enumerate(cols):
            cx = BODY_X + ci * (cw + gut)
            card(s, cx, top, cw, h)
            pad = 0.18
            yy = top + 0.16
            if c["header"]:
                hl = len(c["header"])
                fs_h = 13.5 if hl <= 16 else (12 if hl <= 26 else 11)
                cpl = max(4.0, (cw - 2 * pad) * 72 / (fs_h * 1.05))
                lines = max(1, math.ceil(hl / cpl))
                hh_ = lines * fs_h * 1.2 / 72 + 0.08
                txt(s, cx + pad, yy, cw - 2 * pad, hh_,
                    [{"t": c["header"], "s": fs_h, "c": RED, "b": True, "ls": 1.12}])
                yy += hh_ + 0.04
            # 自适应字号：条目总高超过卡片就逐级缩小
            w_text = cw - pad * 2 - 0.48
            notes_txt = " ".join(c["notes"])[:90] if c["notes"] else None
            avail = top + h - yy - (0.66 if notes_txt else 0.14)
            fs = 12.5
            while fs > 9.0:
                hs = []
                for it in c["items"]:
                    h_ = max(0.32, _est_h([it["ti"]], fs, w_text, ls=1.2) + 0.06)
                    if it.get("desc"):
                        h_ += _est_h([it["desc"]], fs - 1.5, w_text, ls=1.15) + 0.05
                    hs.append(max(0.32, h_))
                if sum(hs) + 0.16 * max(0, len(c["items"]) - 1) <= avail:
                    break
                fs -= 1
            for it, rh_ in zip(c["items"], hs):
                chip(s, cx + pad, yy + 0.01, it["no"])
                if it.get("desc"):
                    txt(s, cx + pad + 0.48, yy, w_text, rh_ + 0.1,
                        [{"t": it["ti"], "s": fs, "c": DARK, "b": True, "ls": 1.15},
                         {"t": it["desc"], "s": fs - 1.5, "c": BODYC, "ls": 1.15}])
                else:
                    txt(s, cx + pad + 0.48, yy, w_text, rh_ + 0.12,
                        [{"t": it["ti"], "s": fs, "c": DARK, "ls": 1.2}])
                yy += rh_ + 0.16
            if notes_txt:
                txt(s, cx + pad, top + h - 0.62, cw - 2 * pad, 0.58,
                    [{"t": notes_txt, "s": 10.5, "c": GRAY, "ls": 1.25}])
        return

    # 单栏：左声明面板（如有）+ 编号行栈（标题+描述）
    x0, w = BODY_X, body_w
    if stmt and xmin > 3.5:
        lw = 2.95
        card(s, BODY_X, top, lw, avail_h, strip=True)
        yy = top + 0.34
        if kick:
            txt(s, BODY_X + 0.26, yy, lw - 0.5, 0.3,
                [{"t": " ".join(f[2] for f in kick)[:24], "s": 10.5, "c": RED, "b": True}])
            yy += 0.5
        stmt_txt = "".join(f[2] for f in stmt)
        st_size = 15 if _est_h([stmt_txt], 15, lw - 0.52) <= avail_h - 1.6 else 13.5
        txt(s, BODY_X + 0.26, yy, lw - 0.52, avail_h - (yy - top) - 0.5,
            [{"t": stmt_txt, "s": st_size, "c": DARK, "b": True, "ls": 1.45}])
        if tag:
            txt(s, BODY_X + 0.26, top + avail_h - 0.55, lw - 0.5, 0.3,
                [{"t": tag[0][2][:20], "s": 10.5, "c": GRAY}])
        x0 = BODY_X + lw + 0.3
        w = body_w - lw - 0.3
    n = len(items)
    # 逐行按内容估高，整体缩放适配可用空间
    hs = []
    for it in items:
        h_ = 0.66
        if it.get("desc"):
            h_ += _est_h([it["desc"]], 12, w - 1.05, ls=1.2) * 1.3
        hs.append(min(max(h_, 0.62), 1.5))
    total = sum(hs) + 0.14 * (n - 1)
    scale = min(1.0, (avail_h - 0.08) / total) if total > 0 else 1.0
    y = top + max(0.0, (avail_h - total * scale) / 2)
    for it, h0 in zip(items, hs):
        rh = h0 * scale
        _num_row_at(s, x0, y, w, rh, it["no"], it["ti"],
                    desc=it["desc"] or None, title_size=15,
                    desc_size=max(9, round(12 * min(1.0, scale * 1.25), 1)))
        y += rh + 0.14 * scale


def _merge_num_items(items):
    """把被拆散的「纯编号 + 标题」合并成一个 item（"1" + "关闭病房门" → "1 关闭病房门"）"""
    out, i = [], 0
    while i < len(items):
        it = str(items[i]).strip()
        if re.fullmatch(r"[0-9０-９]{1,2}|[一二三四五六七八九十]{1,2}", it) and i + 1 < len(items):
            out.append(f"{it} {str(items[i+1]).strip()}")
            i += 2
        else:
            out.append(it)
            i += 1
    return out

def _detect_columns(body):
    """检测多栏网格：x 聚成 ≥2 簇（簇间距>1.2in）且至少两簇各≥2行 → 返回按栏分好的行组；否则 None"""
    if len(body) < 4:
        return None
    xs = sorted(set(round(f[4], 1) for f in body))
    clusters = [[xs[0]]]
    for x in xs[1:]:
        if x - clusters[-1][-1] > 1.2:
            clusters.append([x])
        else:
            clusters[-1].append(x)
    if len(clusters) < 2:
        return None
    centers = [sum(c) / len(c) for c in clusters]
    buckets = [[] for _ in centers]
    for f in body:
        ci = min(range(len(centers)), key=lambda i: abs(centers[i] - f[4]))
        buckets[ci].append(f)
    counts = [len(b) for b in buckets]
    if sum(1 for c in counts if c >= 2) < 2:
        return None
    if min(counts) == 0 or max(counts) > min(counts) * 5:
        return None
    for b in buckets:
        b.sort(key=lambda f: (round(f[3] * 2) / 2, f[4]))
    return buckets

def _groups_from_columns(buckets, body_median):
    """每栏一组：栏内首个短行(≤12字、无句末标点)作组头"""
    groups = []
    for b in buckets:
        if not b:
            continue
        header, items = None, []
        first = b[0]
        if (len(first[2]) <= 12 and not re.search(r"[。；！？.!?]$", first[2])
                and first[0] >= body_median * 0.55 and len(b) > 1):
            header = first[2]
            items = [f[2] for f in b[1:]]
        else:
            items = [f[2] for f in b]
        groups.append((header, items))
    return groups

def _content_slide(new_slide, pg, kind):
    flat = pg["flat"]
    if not flat and not pg["tables"] and not pg["images"]:
        return
    s = new_slide(); chrome(s)
    # ---- 原稿自带编号的页面走专用通道（编号芯片保留原数字，不重复编号）----
    num = _analyze_numbered(flat) if flat and not pg["images"] else None
    if num:
        tmax = max(f[0] for f in flat)
        tls = sorted([f for f in flat
                      if f[0] >= tmax * 0.9 and not re.fullmatch(r"[0-9０-９]{1,2}", f[2])],
                     key=lambda f: f[3])
        if not tls:
            cands = [f for f in flat if not re.fullmatch(r"[0-9０-９]{1,2}", f[2])]
            if cands:
                tls = [max(cands, key=lambda f: f[0])]
        title = "".join(f[2] for f in tls)[:40] or "内容"
        xmin = min(c["x"] for c in num["cols"])
        topz = sorted([f for f in num["left"]
                       if f[3] < 1.45 and f[4] >= xmin - 1.5
                       and f[0] < tmax * 0.9 and f[2] != title],
                      key=lambda f: (f[3], f[4]))
        sub = " ".join(f[2] for f in topz)[:42] or None
        # 独立的「执行判断/结论」金句行 → 底部红条；短标签+相邻金句也认
        bar_text = next((f[2] for f in num["left"]
                         if f not in topz and len(f[2]) > 10
                         and re.match(r"^(执行判断|结论|提示|注意|易错点)", f[2])), None)
        if not bar_text:
            for lb in [f for f in num["left"]
                       if re.fullmatch(r"(执行判断|结论|提示|注意|易错点|小结)", f[2].strip())]:
                stmt = next((g[2] for g in num["left"] if g is not lb and len(g[2]) > 10
                             and abs(g[3] - lb[3]) < 0.7), None)
                if stmt:
                    bar_text = f"{lb[2].strip()}：{stmt}"
                    break
        head(s, title, sub)
        top = BODY_TOP
        avail_h = BODY_BOTTOM - BODY_TOP - (0.42 if pg["footer"] else 0.12) - (0.78 if bar_text else 0)
        if pg["tables"]:
            tbl = pg["tables"][0]["table"]
            nrows = len(tbl); ncols = max(len(r) for r in tbl)
            rh = min(0.46, max(0.32, avail_h * 0.45 / nrows))
            th = rh * nrows
            gf = s.shapes.add_table(nrows, ncols, IN(BODY_X), IN(top), IN(BODY_W), IN(th))
            tb = gf.table
            tb.first_row = False; tb.horz_banding = False
            fs = 12 if nrows <= 4 else 10.5
            for ri in range(nrows):
                for ci in range(ncols):
                    cell = tb.cell(ri, ci)
                    cell.text = tbl[ri][ci] if ci < len(tbl[ri]) else ""
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RED if ri == 0 else CARDC
                    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    for p in cell.text_frame.paragraphs:
                        no_bullet(p)
                        for r in p.runs:
                            set_font(r, fs, WHITE if ri == 0 else DARK, ri == 0)
            top += th + 0.3
            avail_h -= th + 0.3
        _render_numbered(s, num, top, avail_h, BODY_W)
        if bar_text:
            bar(s, bar_text[:60], size=12.5)
        if pg["footer"]:
            footer(s, pg["footer"], y=6.58)
        return
    # 标题与副题：多行大字合并；过长时首行作标题、其余转副题；标题上方小字作眉题
    kicker, used_txt = None, set()
    if flat:
        # 字号基准排除装饰性纯数字行，避免 42pt 大数字压制真标题
        tmax = max((f[0] for f in flat if not re.fullmatch(r"[0-9０-９]{1,2}", f[2])), default=0)
        bigs = sorted([f for f in flat if f[0] >= tmax * 0.9
                       and not re.fullmatch(r"[0-9０-９]{1,2}", f[2])], key=lambda f: f[3])
        title_src_y = bigs[0][3] if bigs else 0.0
        extra_big = []
        if bigs:
            used_txt |= {f[2] for f in bigs}
            joined = "".join(f[2] for f in bigs)
            if len(bigs) > 1 and len(joined) > 22:
                title = re.sub(r"[，、；,;：:]+$", "", bigs[0][2])
                extra_big = [re.sub(r"[，、；,;：:]+$", "", f[2]) for f in bigs[1:]]
            else:
                title = joined
            if re.fullmatch(r"[0-9０-９]{1,2}|[一二三四五六七八九十]{1,3}", title) or len(title) <= 2:
                cands = [f for f in flat if f[0] >= tmax * 0.6 and len(f[2]) > 2 and f[2] not in used_txt]
                if cands:
                    title = min(cands, key=lambda f: f[3])[2]
        else:
            title = "内容"
        kick = [f for f in flat if f[3] < title_src_y - 0.2 and f[0] <= 14
                and len(f[2]) <= 34 and f[2] not in used_txt]
        if kick:
            kicker = " ".join(f[2] for f in kick)[:44]
            used_txt |= {f[2] for f in kick}
        subs = [f[2] for f in flat
                if f[2] not in used_txt and not f[1] and f[0] >= tmax * 0.55 and f[0] < tmax * 0.75
                and 0 < len(f[2]) <= 36 and (not bigs or f[3] >= title_src_y)
                and f[3] < 2.0]  # 副题只从标题下方近处挑，防止抢走组内正文
        sub = extra_big[0] if extra_big else (subs[0] if subs else None)
        if sub:
            used_txt.add(sub)
        # 眉题与副题同占一行区域：短则合并；放不下时两者都渲染（眉题移到副题下方一行）
        if kicker and sub and len(kicker) + len(sub) <= 42:
            sub = kicker + " ｜ " + sub
            kicker = None
    else:
        title, sub = "内容", None
    # 装饰性大数字（≥80% 最大字号的独立数字行）不进正文
    deco = {f[2] for f in flat
            if re.fullmatch(r"[0-9０-９]{1,2}", f[2]) and f[0] >= tmax * 0.8} if flat else set()
    body = [f for f in flat if f[2] not in used_txt and f[2] not in deco]
    head(s, title, sub)

    top = BODY_TOP
    if kicker:
        ky = 1.76 if sub else 1.44  # 有副题时眉题放副题下方，避免重叠
        txt(s, BODY_X, ky, BODY_W, 0.3, [{"t": kicker, "s": 11, "c": RED, "b": True}])
        top = BODY_TOP + (0.42 if sub else 0.36)
    if pg["footer"]:
        pass  # footer 放页底
    avail_h = BODY_BOTTOM - top - (0.42 if pg["footer"] else 0.12)
    # 表格
    if pg["tables"]:
        tbl = pg["tables"][0]["table"]
        nrows = len(tbl); ncols = max(len(r) for r in tbl)
        rh = min(0.46, max(0.32, avail_h * 0.5 / nrows))
        th = rh * nrows
        gf = s.shapes.add_table(nrows, ncols, IN(BODY_X), IN(top), IN(BODY_W), IN(th))
        tb = gf.table
        tb.first_row = False; tb.horz_banding = False
        fs = 12 if nrows <= 4 else 10.5
        for ri in range(nrows):
            for ci in range(ncols):
                cell = tb.cell(ri, ci)
                cell.text = tbl[ri][ci] if ci < len(tbl[ri]) else ""
                cell.fill.solid()
                cell.fill.fore_color.rgb = RED if ri == 0 else CARDC
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                for p in cell.text_frame.paragraphs:
                    no_bullet(p)
                    for r in p.runs:
                        set_font(r, fs, WHITE if ri == 0 else DARK, ri == 0)
        top += th + 0.3
        avail_h -= th + 0.3
    # 图片（大图放右侧）
    img = pg["images"][0] if pg["images"] else None
    body_w = BODY_W
    if img and img["w"] >= 2.5 and img["h"] >= 1.5:
        body_w = 5.3
        iw = 2.9
        ih = min(4.2, iw * img["h"] / img["w"])
        card(s, BODY_X + body_w + 0.25, top, iw + 0.35, ih + 0.35, strip=False)
        s.shapes.add_picture(BytesIO(img["blob"]), IN(BODY_X + body_w + 0.42), IN(top + 0.18), IN(iw), IN(ih))

    # 正文分组：先做栏检测（网格布局按栏读取），再分组
    body_median = sorted(sz for sz, b, t, _y, _x in body)[len(body) // 2] if body else 14
    col_buckets = _detect_columns(body)
    if col_buckets:
        groups = _groups_from_columns(col_buckets, body_median)
        flatlist = []
    else:
        groups, flatlist = _split_groups(body, body_median)
    groups = [(h, _merge_num_items(items)) for h, items in groups]
    # 末组长金句（无编号前缀、带句读、无正文）先转红条，避免被空组合并吃掉
    _bar_candidate = None
    if groups and len(groups[-1][1]) == 0 and groups[-1][0] \
            and len(groups[-1][0]) > 15 and re.search(r"[。；]", groups[-1][0]) \
            and not re.match(r"^[（(]?[0-9０-９一二三四五六七八九十]", groups[-1][0]):
        _bar_candidate = groups[-1][0]
        groups = groups[:-1]
    # 空组头（冒号引导句单独成组）并入下一组，避免空卡片
    merged = []
    for h_, its_ in groups:
        if (merged and its_ == [] and h_ and merged[-1][1]
                and not re.match(r"^(结论|注意|易错点|执行判断|小结)$", (h_ or "").strip())):
            merged[-1] = (merged[-1][0] + " " + h_, merged[-1][1])
        else:
            merged.append((h_, its_))
    groups = merged
    flatlist = _merge_num_items(flatlist)
    flatlist = [x for x in flatlist if x]
    if not groups and not flatlist:
        if pg["footer"]: footer(s, pg["footer"])
        return

    # 「结论/注意」短标签 + 相邻金句 → 底部红条（可出现在任何一组）
    bar_text = _bar_candidate
    for gi, (h_, its_) in enumerate(groups):
        for ii in range(len(its_)):
            if not re.fullmatch(r"结论|注意|小结", str(its_[ii]).strip()):
                continue
            prev_ok = ii > 0 and len(str(its_[ii - 1])) > 10
            next_ok = ii + 1 < len(its_) and len(str(its_[ii + 1])) > 10
            if prev_ok:
                bar_text = f"{str(its_[ii]).strip()}：{its_[ii - 1]}"
                groups[gi] = (h_, its_[:ii - 1] + its_[ii + 1:])
            elif next_ok:
                bar_text = f"{str(its_[ii]).strip()}：{its_[ii + 1]}"
                groups[gi] = (h_, its_[:ii] + its_[ii + 2:])
            break
        if bar_text:
            break
    if not bar_text and groups and len(groups[-1][1]) == 0 and re.match(r"^(结论|注意|易错点|执行|怎么用|一句话|要点)", groups[-1][0] or ""):
        bar_text = groups[-1][0]; groups = groups[:-1]
    elif not bar_text and groups and len(groups[-1][1]) == 0 and groups[-1][0] \
            and len(groups[-1][0]) > 15 and re.search(r"[。；]", groups[-1][0]) \
            and not re.match(r"^[（(]?[0-9０-９一二三四五六七八九十]", groups[-1][0]):
        # 末组长金句（无编号前缀、带句读）→ 底部红条
        bar_text = groups[-1][0]; groups = groups[:-1]
    elif groups and len(groups[-1][1]) <= 1 and re.match(r"^(结论|注意|易错点|执行判断|怎么用)", (groups[-1][0] or "")):
        bar_text = groups[-1][0] + "：" + ("；".join(groups[-1][1])); groups = groups[:-1]
    if bar_text:
        avail_h -= 0.78  # 底部红条占位

    n_g = len(groups)
    n_items = len(flatlist)
    # 组前的引导段（flatlist）不能丢：并成第一组
    if flatlist and groups:
        groups = [(None, flatlist)] + groups
        flatlist = []
        n_g = len(groups)
        n_items = 0
    # 布局选择
    if flatlist and n_g == 0:
        _layout_rows(s, top, avail_h, flatlist, body_w)
    elif n_g == 1:
        _layout_rows(s, top, avail_h, [groups[0][0]] + groups[0][1], body_w, first_is_head=True)
    elif 2 <= n_g <= 4 and all(len(g[1]) <= 7 and sum(len(x) for x in g[1]) < 320 for g in groups):
        _layout_columns(s, top, avail_h, groups, body_w)
    else:
        _layout_stacked(s, top, avail_h, groups, body_w)
    if bar_text:
        bar(s, bar_text[:60], size=12.5)
    if pg["footer"]:
        footer(s, pg["footer"], y=6.58)

def _layout_rows(s, top, avail_h, items, body_w, first_is_head=False):
    """编号行布局：每条一行卡片；行太多时拆成左右两栏"""
    n = len(items)
    if n == 0: return
    if n > 8 and not first_is_head:
        half = (n + 1) // 2
        colw = (body_w - 0.3) / 2
        for ci, chunk in enumerate([items[:half], items[half:]]):
            _rows_col(s, top, avail_h, chunk, colw, BODY_X + ci * (colw + 0.3), start_no=(1 if ci == 0 else half + 1))
        return
    _rows_col(s, top, avail_h, items, body_w, BODY_X, first_is_head=first_is_head)

def _rows_col(s, top, avail_h, items, w, x, first_is_head=False, start_no=1):
    # 合并被拆散的「纯编号行 + 标题行」→ (num, title) 对
    pairs = []
    i = 0
    while i < len(items):
        it = str(items[i])
        if re.fullmatch(r"[0-9０-９]{1,2}|[一二三四五六七八九十]{1,2}", it.strip()) and i + 1 < len(items):
            pairs.append((it.strip(), str(items[i + 1])))
            i += 2
        else:
            pairs.append((None, it))
            i += 1
    n = len(pairs)
    if n == 0: return
    gap = min(1.25, (avail_h - 0.1) / n)
    rh = max(0.42, min(gap - 0.14, 1.15))
    y = top
    for i, (pnum, it) in enumerate(pairs):
        no = start_no + i
        num = pnum or (f"{no:02d}" if n > 9 else str(no))
        if first_is_head and i == 0:
            h0 = min(rh, 0.72)
            card(s, x, y, w, h0, strip=False)
            chip(s, x + 0.23, y + (h0 - 0.30) / 2, num)
            txt(s, x + 0.73, y + (h0 - 0.34) / 2, w - 1.0, 0.36,
                [{"t": it, "s": 16, "c": DARK, "b": True}])
        elif pnum is None and (re.match(r"^(Q[0-9０-９]|[A-Z][0-9０-９]{1,2}\b)", it.strip())
                               or (it.strip().endswith("：") and len(it.strip()) <= 24)):
            # 自带编号（Q1…）或引导句 → 不再加芯片，避免双重编号
            _num_row_at(s, x, y, w, rh, None, it, title_size=14.5)
        else:
            _num_row_at(s, x, y, w, rh, num, it, title_size=14.5)
        y += gap

def _num_row_at(slide, x, y, w, h, num, title, desc=None, title_size=15.5, desc_size=12.5, tag=None):
    card(slide, x, y, w, h, strip=False)
    if num:
        chip(slide, x + 0.23, y + (h - 0.30) / 2, num)
        tx = x + 0.73
        tw = w - 1.0
    else:
        tx = x + 0.28
        tw = w - 0.56
    if desc:
        txt(slide, tx, y + 0.10, tw - 0.6, 0.34, [{"t": title, "s": title_size, "c": DARK, "b": True}])
        ds_ = desc_size
        while ds_ > 8 and _est_h([desc], ds_, tw, ls=1.2) > h - 0.5:
            ds_ -= 0.5
        txt(slide, tx, y + 0.10 + 0.36, tw, h - 0.42, [{"t": desc, "s": ds_, "c": BODYC, "ls": 1.2}])
    else:
        # 标题过长自动缩字号
        ts = title_size
        while ts > 9 and _est_h([title], ts, tw) > h - 0.1:
            ts -= 0.5
        txt(slide, tx, y + (h - 0.34) / 2, tw, 0.36, [{"t": title, "s": ts, "c": DARK, "b": True}])
    if tag:
        txt(slide, x + w - 2.3, y + 0.12, 2.1, 0.3, [{"t": tag, "s": 11, "c": GRAY, "a": PP_ALIGN.RIGHT}])

def _layout_columns(s, top, avail_h, groups, body_w):
    """2-4 列卡片布局"""
    n = len(groups)
    gut = 0.17
    cw = (body_w - gut * (n - 1)) / n
    h = avail_h - (0.15 if any(g[0] and re.match(r"^(结论|注意)", g[0]) for g in groups) else 0.05)
    for ci, (hname, items) in enumerate(groups):
        cx = BODY_X + ci * (cw + gut)
        card(s, cx, top, cw, h)
        pad = 0.18
        yy = top + 0.16
        if hname:
            # 组头不截断：按卡片宽度估算行数，缩字号+动态高度
            hl = len(hname)
            fs_h = 13.5 if hl <= 16 else (12 if hl <= 26 else 11)
            cpl = max(4.0, (cw - 2 * pad) * 72 / (fs_h * 1.05))
            lines = max(1, math.ceil(hl / cpl))
            hh_ = lines * fs_h * 1.2 / 72 + 0.08
            txt(s, cx + pad, yy, cw - 2 * pad, hh_,
                [{"t": hname, "s": fs_h, "c": RED, "b": True, "ls": 1.12}])
            yy += hh_ + 0.04
        paras, runs_line = [], []
        for j, it in enumerate(items):
            paras.append({"runs": [(f"{j+1}  ", 13, RED, True), (it, 12, DARK, False)], "sb": 7})
        # 估算需要高度，必要时缩字号
        base = 1.0
        total_lines = sum(max(1, math.ceil(_wlen(it) / ((cw - 2 * pad) * 72 / 12))) for it in items)
        need = total_lines * 12 * 1.55 / 72 + len(items) * 8 / 72
        scale = min(1.0, (h - (yy - top) - 0.15) / need) if need > 0 else 1.0
        k = max(0.62, scale)
        box = txt(s, cx + pad, yy, cw - 2 * pad, h - (yy - top) - 0.12, [])
        tf = box.text_frame
        first = True
        for j, it in enumerate(items):
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            no_bullet(p)
            p.line_spacing = 1.2
            p.space_after = Pt(7 * k)
            for t_, s_, c_, b_ in [(f"{j+1}  ", 13 * k, RED, True), (it, 12 * k, DARK, False)]:
                r = p.add_run(); r.text = t_
                set_font(r, round(s_, 1), c_, b_)

def _layout_stacked(s, top, avail_h, groups, body_w):
    """纵向堆叠大卡：组头+多行正文"""
    n = len(groups)
    gap = (avail_h - 0.12) / n
    y = top
    for hname, items in groups:
        ch = gap - 0.14
        card(s, BODY_X, y, body_w, ch)
        hn = hname or ""
        hsize = 14 if len(hn) <= 30 else 11.5
        txt(s, 0.95, y + 0.12, body_w - 0.5, 0.3, [{"t": hn[:60], "s": hsize, "c": DARK, "b": True}])
        lines = items if items else []
        # 行内拆分（一行里的 · 分隔）
        flat_lines = []
        for it in lines:
            flat_lines.extend([x.strip() for x in re.split(r"[·；;]", it) if x.strip()] if ("·" in it and len(it) > 40) else [it])
        size = 11.5
        while size > 8 and _est_h(flat_lines, size, body_w - 0.6, ls=1.25, sb=0.06) > ch - 0.55:
            size -= 0.5
        txt(s, 0.95, y + 0.5, body_w - 0.5, ch - 0.6,
            [{"t": t, "s": size, "c": BODYC, "ls": 1.22, "sb": 4} for t in flat_lines])
        y += gap
