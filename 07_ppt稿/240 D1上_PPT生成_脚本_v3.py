#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Step 4.5 内容深化生成：解析 240 D1上_PPT页面内容深化_v1.md，
克隆母版页型 → 填标题+正文要点+底部条 → 插示例图 → 出 240 D1上_PPT骨架_v4.pptx。
修复 v1-v3 "页面只有一行" 问题。
"""
import copy
import re
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches, Emu

BASE = "/Users/panda/Desktop/Vibe Coding 3天营/07_ppt稿"
TEMPLATE = BASE + "/课程PPT母版.pptx"
SCRIPT = BASE + "/240 D1上_PPT页面内容深化_v2.md"
OUTPUT = BASE + "/240 D1上_PPT骨架_v5.pptx"
MAT = BASE + "/素材"

# ---------- 1. 解析深化表 ----------
def parse_deep(md):
    """返回 {P01: {page,type,title,body[],footnote,subtitle,imgs[]}}"""
    pages = {}
    cur = None
    for line in md.splitlines():
        line = line.strip()
        m = re.match(r"^### (P\d{2})\s*｜\s*([^｜]+?)\s*｜\s*(.+)$", line)
        if m:
            if cur is not None:
                pages[cur["page"]] = cur
            cur = {"page": m.group(1), "type": m.group(2).strip(),
                   "title": m.group(3).strip(), "body": [], "footnote": "",
                   "subtitle": None, "imgs": []}
            continue
        if cur is None:
            continue
        if line.startswith("- "):
            cur["body"].append(line[2:].strip())
        elif line.startswith("> 底部："):
            cur["footnote"] = line[len("> 底部："):].strip()
        elif line.startswith("> 副标题："):
            cur["subtitle"] = line[len("> 副标题："):].strip()
        elif line.startswith("> 图："):
            for part in line[len("> 图："):].split("＋"):
                part = part.strip()
                mm = re.match(r"^(ppt素材_\d+_[^（]+\.png)（(左|右|Picture\d+)）?$", part)
                if mm:
                    cur["imgs"].append((mm.group(1), mm.group(2) or "Picture21"))
                else:
                    cur["imgs"].append((part, "Picture21"))
    if cur is not None:
        pages[cur["page"]] = cur
    return pages

# ---------- 2. 克隆 slide ----------
def clone_slide(prs, src_slide):
    dest = prs.slides.add_slide(src_slide.slide_layout)
    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)
    for shp in src_slide.shapes:
        dest.shapes._spTree.append(copy.deepcopy(shp._element))
    rid_map = {}
    for rId, rel in src_slide.part.rels.items():
        if rel.reltype == RT.IMAGE and not rel.is_external:
            try:
                new_rId = dest.part.relate_to(rel.target_part, RT.IMAGE)
                rid_map[rId] = new_rId
            except Exception:
                pass
    if rid_map:
        ns_embed = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
        for el in dest.shapes._spTree.iter():
            tag = el.tag.split('}')[-1]
            if tag == "blip":
                emb = el.get(ns_embed)
                if emb in rid_map:
                    el.set(ns_embed, rid_map[emb])
    return dest

# ---------- 3. 文本工具 ----------
def clear_tf(tf):
    txBody = tf._txBody
    for p in txBody.findall(qn('a:p')):
        txBody.remove(p)

def add_para(tf, text, size=None, bold=None, color=None, align=None):
    txBody = tf._txBody
    p = txBody.makeelement(qn('a:p'), {})
    txBody.append(p)
    if align is not None:
        p.set('algn', align)
    parts = text.split("\x0b")
    for i, seg in enumerate(parts):
        if i > 0:
            br = p.makeelement(qn('a:br'), {})
            p.append(br)
        if seg == "":
            continue
        r = p.makeelement(qn('a:r'), {})
        p.append(r)
        rPr = r.makeelement(qn('a:rPr'), {'lang': 'zh-CN', 'altLang': 'en-US'})
        r.append(rPr)
        if size:
            rPr.set('sz', str(int(size * 100)))
        if bold is not None:
            rPr.set('b', '1' if bold else '0')
        if color:
            sf = rPr.makeelement(qn('a:solidFill'), {})
            clr = rPr.makeelement(qn('a:srgbClr'), {'val': color})
            sf.append(clr)
            rPr.append(sf)
        t = r.makeelement(qn('a:t'), {})
        t.text = seg
        r.append(t)

def set_placeholder(slide, idx, lines):
    for sh in slide.shapes:
        if sh.is_placeholder and sh.placeholder_format.idx == idx:
            tf = sh.text_frame
            clear_tf(tf)
            for line in lines:
                if len(line) == 5:
                    text, size, bold, color, align = line
                else:
                    text, size, bold, color = line
                    align = None
                add_para(tf, text, size, bold, color, align)
            return True
    return False

def set_named_box(slide, name, lines):
    """lines: list of (text,size,bold,color)"""
    for sh in slide.shapes:
        if not sh.is_placeholder and sh.name == name and sh.has_text_frame:
            tf = sh.text_frame
            clear_tf(tf)
            for (text, size, bold, color) in lines:
                add_para(tf, text, size, bold, color)
            return sh
    return None

def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

def add_body_textbox(slide, left_in, top_in, width_in, height_in, lines):
    tb = slide.shapes.add_textbox(Inches(left_in), Inches(top_in),
                                  Inches(width_in), Inches(height_in))
    tf = tb.text_frame
    tf.word_wrap = True
    clear_tf(tf)
    for (text, size, bold, color, align) in lines:
        add_para(tf, text, size, bold, color, align)
    return tb

# ---------- 4. 页型映射 ----------
PAGE_KIND = {
    "封面": ("slide1", "title"), "封底": ("slide1", "title"),
    "章节页": ("slide3", "chapter"), "章节页/模型页": ("slide3", "chapter"),
    "桥页": ("slide4", "question"), "悬念页": ("slide4", "question"),
    "回扣页": ("slide4", "question"),
    "任务卡": ("slide5", "content"), "演示指令卡": ("slide5", "content"),
    "存屏页/任务卡": ("slide5", "content"), "反差页": ("slide5", "content"),
    "命名页": ("slide5", "content"), "公式页": ("slide5", "content"),
    "套用页": ("slide5", "content"), "案例对比页（左）": ("slide5", "content"),
    "案例对比页（右）": ("slide5", "content"), "案例对比页": ("slide5", "content"),
    "拆解页": ("slide5", "content"), "名词页": ("slide5", "content"),
    "链条总图": ("slide5", "content"), "对比/迭代页": ("slide5", "content"),
    "迭代页": ("slide6", "imagepage"),   # P50 有改后版图 → 图文页
    "金句页": ("slide8", "quote"), "锚点页": ("slide8", "quote"),
    "金句＋锚点页": ("slide8", "quote"), "锚点＋行动页": ("slide8", "quote"),
    "金句页＋桥": ("slide8", "quote"),
    "结果证据页": ("slide6", "imagepage"), "结果诊断页": ("slide6", "imagepage"),
    "结果证据页/回扣页": ("slide6", "imagepage"),
}
PREFIX_KIND = [("名词页", ("slide5", "content")), ("案例对比页", ("slide5", "content")),
               ("对比/迭代页", ("slide5", "content"))]
TITLE_LABEL = {
    "金句页": "记住", "锚点页": "锚点", "金句＋锚点页": "记住",
    "锚点＋行动页": "行动", "金句页＋桥": "记住",
    "桥页": "悬念", "悬念页": "想一想", "回扣页": "想一想",
}

# 附加图（slide5 内容页右侧叠加）P17/P18
EXTRA_IMAGES = {
    "P17": ("ppt素材_04_贪吃蛇有分数_v1.png", 8.2, 2.4, 4.4),   # left,top,width(in)
    "P18": ("ppt素材_05_贪吃蛇绿蛇快速度_v1.png", 8.2, 2.4, 4.4),
}

# ---------- 5. 填充 ----------
def fill_slide(slide, page, kind):
    p = page["page"]
    typ = page["type"]
    title = page["title"]
    body = page["body"]
    footnote = page["footnote"]
    notes = "%s | %s | 源段:%s | 时长:%s" % (p, typ, page.get("src", "-"), "-")

    if kind == "title":
        main = title
        sub = page["subtitle"] or (body[0] if body else "")
        full = main + ("\x0b" + sub if sub else "")
        set_placeholder(slide, 0, [(full, 32, True, None)])
        set_notes(slide, notes)
    elif kind == "chapter":
        set_placeholder(slide, 0, [(title, 26, True, None)])
        set_placeholder(slide, 10, [(p[1:], 18, None, None)])
        if body:
            lines = [(b, 18, False, None, "ctr") for b in body]
            add_body_textbox(slide, 2.8, 4.3, 7.7, 1.8, lines)
        set_notes(slide, notes)
    elif kind == "question":
        label = TITLE_LABEL.get(typ, "思考")
        set_placeholder(slide, 0, [(label, None, True, None)])
        lines = [(title, 20, True, None, "ctr")]
        lines += [(b, 24, False, None, "ctr") for b in body]
        if footnote:
            lines.append((footnote, 12, False, "888888", "ctr"))
        set_placeholder(slide, 10, lines)
        set_notes(slide, notes)
    elif kind == "quote":
        label = TITLE_LABEL.get(typ, "金句")
        set_placeholder(slide, 0, [(label, None, True, None)])
        lines = [("\u201c%s\u201d" % title, 30, True, None, "ctr")]
        lines += [(b, 16, False, "666666", "ctr") for b in body]
        if footnote:
            lines.append((footnote, 12, False, "888888", "ctr"))
        set_placeholder(slide, 10, lines)
        set_notes(slide, notes)
    elif kind == "content":
        set_placeholder(slide, 0, [(title, None, None, None)])
        set_named_box(slide, "TextBox 9", [("【%s】" % typ, None, True, None)])
        # 正文要点 → 文本框 5（扩高）
        box5 = set_named_box(slide, "文本框 5",
                             [(b, 14, False, None) for b in body] if body
                             else [("", 14, False, None)])
        if box5 is not None:
            box5.height = Inches(2.1)
        # 底部条 → 文本框 2（无底部条也必须清空母版示例占位文本）
        set_named_box(slide, "文本框 2",
                      [(footnote, 12, True, "C00000")] if footnote else [])
        set_notes(slide, notes)
    elif kind == "imagepage":
        set_placeholder(slide, 0, [(title, None, None, None)])
        set_named_box(slide, "TextBox 9", [("【%s】" % typ, None, True, None)])
        # 图注/正文行 → 文本框 20（扩宽）：body 全部行 + 底部条
        cap_lines = list(body)
        if footnote:
            cap_lines.append(footnote)
        if not cap_lines:
            cap_lines = [""]
        box20 = set_named_box(slide, "文本框 20",
                              [(c, 10, False, "444444") for c in cap_lines])
        if box20 is not None:
            box20.width = Inches(2.2)
            box20.height = Inches(1.3)
        set_notes(slide, notes)

# ---------- 6. 插图 ----------
def clear_pics(slide):
    for sh in list(slide.shapes):
        if sh.shape_type == 13 or "Picture" in sh.name:
            sh._element.getparent().remove(sh._element)

def place_pic(slide, img_path, box_l, box_t, box_w, box_h):
    from PIL import Image
    im = Image.open(img_path)
    iw, ih = im.size
    ar = iw / ih
    w = box_w
    h = w / ar
    if h > box_h:
        h = box_h
        w = h * ar
    left = box_l + (box_w - w) / 2
    top = box_t + (box_h - h) / 2
    return slide.shapes.add_picture(img_path, Inches(left), Inches(top),
                                    Inches(w), Inches(h))

BOX21 = (3.09, 1.88, 5.80, 5.15)
BOX22 = (9.26, 1.88, 3.39, 5.13)

# 主图文页：page -> (主图, 副图)
IMAGE_MAP = {
    "P07": ("ppt素材_01_星空海报_v1.png", None),
    "P08": ("ppt素材_02_简历网页_v1.png", None),
    "P16": ("ppt素材_03_贪吃蛇基础版_v1.png", None),
    "P37": ("ppt素材_02_简历网页_v1.png", None),
    "P48": ("ppt素材_01_星空海报_v1.png", "ppt素材_06_星空海报改后版_v1.png"),
    "P50": ("ppt素材_06_星空海报改后版_v1.png", None),
}

# ---------- 7. 主流程 ----------
def main():
    with open(SCRIPT, encoding="utf-8") as f:
        pages = parse_deep(f.read())
    print("解析页数:", len(pages))
    if len(pages) != 52:
        raise SystemExit("页数不是52！实际%d" % len(pages))

    prs = Presentation(TEMPLATE)
    slides = list(prs.slides)
    src_map = {"slide1": slides[0], "slide3": slides[2], "slide4": slides[3],
               "slide5": slides[4], "slide6": slides[5], "slide8": slides[7]}
    # 注意：先克隆全部新页（add_slide 序号从 slide10 起，避免与母版 slide1-9 冲突），
    # 全部生成后再删除母版原始 9 页的 sldId（避免 zip 同名重复）
    created = []
    for pkey in sorted(pages.keys(), key=lambda x: int(x[1:])):
        page = pages[pkey]
        kind_name = PAGE_KIND.get(page["type"])
        if kind_name is None:
            for prefix, kv in PREFIX_KIND:
                if page["type"].startswith(prefix):
                    kind_name = kv
                    break
        if kind_name is None:
            raise SystemExit("未映射类型: %s (%s)" % (page["type"], page["page"]))
        src_name, kind = kind_name
        new_slide = clone_slide(prs, src_map[src_name])
        fill_slide(new_slide, page, kind)

        # 插图
        if pkey in IMAGE_MAP:
            clear_pics(new_slide)
            main_img, sub_img = IMAGE_MAP[pkey]
            place_pic(new_slide, MAT + "/" + main_img, *BOX21)
            if sub_img:
                place_pic(new_slide, MAT + "/" + sub_img, *BOX22)
        if pkey in EXTRA_IMAGES:
            fname, l, t, w = EXTRA_IMAGES[pkey]
            place_pic(new_slide, MAT + "/" + fname, l, t, w, 3.2)
        created.append(pkey)

    # 删除母版原始 9 页引用（保留克隆出的 52 页）
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst)[:9]:
        sldIdLst.remove(sldId)

    prs.save(OUTPUT)
    print("已生成:", OUTPUT)
    print("页数:", len(created))

if __name__ == "__main__":
    main()
