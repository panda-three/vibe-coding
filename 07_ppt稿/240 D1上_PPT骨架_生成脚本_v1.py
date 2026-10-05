#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Step 4 骨架生成：以 07_ppt稿/课程PPT母版.pptx 的 9 个页型为底，
按 240 D1上_PPT页面脚本表_v1.md 的 52 页生成骨架版 PPTX（只放标题+核心信息）。
"""
import copy
import re
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

BASE = "/Users/panda/Desktop/Vibe Coding 3天营/07_ppt稿"
TEMPLATE = BASE + "/课程PPT母版.pptx"
SCRIPT = BASE + "/240 D1上_PPT页面脚本表_v1.md"
OUTPUT = BASE + "/240 D1上_PPT骨架_v1.pptx"

# ---------- 1. 解析脚本表 ----------
def parse_script(md):
    pages = []
    for line in md.splitlines():
        line = line.strip()
        if not line.startswith("|"):
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        # 表格头与分隔行跳过
        if cells[0] in ("页码", "---") or not re.fullmatch(r"P\d{2}", cells[0]):
            continue
        if len(cells) < 8:
            continue
        pages.append({
            "page": cells[0],
            "type": cells[1].replace("**", "").strip(),
            "title": cells[2],
            "core": cells[3],
            "visual": cells[4],
            "notes_src": cells[5],   # 备注列
            "duration": cells[6],    # 时长
            "src": cells[7],         # 源段
        })
    return pages

# ---------- 2. 克隆 slide（含图片关系重映射） ----------
def clone_slide(prs, src_slide):
    dest = prs.slides.add_slide(src_slide.slide_layout)
    # 移除 dest 自带占位符
    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)
    # 复制 shapes XML
    for shp in src_slide.shapes:
        dest.shapes._spTree.append(copy.deepcopy(shp._element))
    # 复制图片关系并重映射 r:embed
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

# ---------- 3. 文本填充工具 ----------
def clear_tf(tf):
    txBody = tf._txBody
    for p in txBody.findall(qn('a:p')):
        txBody.remove(p)

def add_para(tf, text, size=None, bold=None, color=None, align=None):
    """新增一个段落；text 内 \x0b 转换为 <a:br/> 换行"""
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
    """lines: list of (text, size, bold, color)"""
    for sh in slide.shapes:
        if sh.is_placeholder and sh.placeholder_format.idx == idx:
            tf = sh.text_frame
            clear_tf(tf)
            for (text, size, bold, color) in lines:
                add_para(tf, text, size, bold, color)
            return True
    return False

def set_named_box(slide, name, text, size=None, bold=None, color=None):
    for sh in slide.shapes:
        if not sh.is_placeholder and sh.name == name and sh.has_text_frame:
            tf = sh.text_frame
            clear_tf(tf)
            add_para(tf, text, size, bold, color)
            return True
    return False

def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

# ---------- 4. 页型映射：type -> (源页序号1based, 填充函数) ----------
# 源页: 1封面 2目录 3章节 4思考 5内容页 6内容页+图 7要素页 8金句页 9空白
PAGE_KIND = {
    "封面": ("slide1", "title"),
    "章节页": ("slide3", "chapter"),
    "桥页": ("slide4", "question"),
    "悬念页": ("slide4", "question"),
    "回扣页": ("slide4", "question"),
    "任务卡": ("slide5", "content"),
    "演示指令卡": ("slide5", "content"),
    "存屏页/任务卡": ("slide5", "content"),
    "金句页": ("slide8", "quote"),
    "锚点页": ("slide8", "quote"),
    "金句＋锚点页": ("slide8", "quote"),
    "锚点＋行动页": ("slide8", "quote"),
    "金句页＋桥": ("slide8", "quote"),
    "结果证据页": ("slide6", "imagepage"),
    "结果诊断页": ("slide6", "imagepage"),
    "结果证据页/回扣页": ("slide6", "imagepage"),
    "反差页": ("slide5", "content"),
    "命名页": ("slide5", "content"),
    "公式页": ("slide5", "content"),
    "套用页": ("slide5", "content"),
    "案例对比页（左）": ("slide5", "content"),
    "案例对比页（右）": ("slide5", "content"),
    "案例对比页": ("slide5", "content"),
    "拆解页": ("slide5", "content"),
    "名词页": ("slide5", "content"),
    "链条总图": ("slide5", "content"),
    "对比/迭代页": ("slide5", "content"),
    "迭代页": ("slide5", "content"),
    "封底": ("slide1", "title"),
    "章节页/模型页": ("slide3", "chapter"),
}

TITLE_LABEL = {
    "金句页": "记住", "锚点页": "锚点", "金句＋锚点页": "记住",
    "锚点＋行动页": "行动", "金句页＋桥": "记住",
    "桥页": "悬念", "悬念页": "想一想", "回扣页": "想一想",
}

def fill_slide(slide, page, kind):
    p = page["page"]
    title = page["title"]
    core = page["core"]
    src = page["src"]
    typ = page["type"]
    notes = "%s | %s | 源段:%s | 时长:%s" % (p, typ, src, page["duration"])

    if kind == "title":  # 封面/封底
        # TITLE(0): 主标题 \x0b 副标题
        if page["page"] == "P01":
            subtitle = "你没写一行代码"
        elif "：" in title:
            subtitle = title.split("：", 1)[1]
        else:
            subtitle = ""
        main = title.split("：", 1)[0] if "：" in title else title
        full = main + "\x0b" + subtitle if subtitle else main
        set_placeholder(slide, 0, [(full, 36, True, None)])
        set_notes(slide, notes)
    elif kind == "chapter":  # 章节页
        num = p[1:]  # 06/12/...
        set_placeholder(slide, 0, [(title, 28, False, None)])
        set_placeholder(slide, 10, [(num, None, None, None)])
        set_notes(slide, notes)
    elif kind == "question":  # 思考/悬念/桥页
        label = TITLE_LABEL.get(typ, "思考")
        set_placeholder(slide, 0, [(label, None, True, None)])
        # 正文：问题/悬念句（title）优先；core 若不同则追加为第二行（如选项）
        lines = [(title, None, None, None)]
        if core and core != title:
            lines.append((core, None, None, None))
        set_placeholder(slide, 10, lines)
        set_notes(slide, notes)
    elif kind == "quote":  # 金句/锚点页
        label = TITLE_LABEL.get(typ, "金句")
        set_placeholder(slide, 0, [(label, None, True, None)])
        # 正文＝结论化标题（金句本身）；core 只是说明词，不上正文
        set_placeholder(slide, 10, [("\u201c%s\u201d" % title, 30, True, None)])
        set_notes(slide, notes)
    elif kind == "content":  # 内容页
        set_placeholder(slide, 0, [(title, None, None, None)])
        set_named_box(slide, "TextBox 9", "【%s】" % typ, None, True, None)
        # P43 链条总图：正文放链条文字
        if page["page"] == "P43":
            set_named_box(slide, "文本框 5", CHAIN_TEXT, 14, None, None)
            set_named_box(slide, "文本框 2", "每个都是因为上一个做不到，才有了下一个", 14, None, "FF0000")
        else:
            core_lines = core.split("\n") if core else []
            set_named_box(slide, "文本框 5", "\x0b".join(core_lines) if core_lines else core, 14, None, None)
            set_named_box(slide, "文本框 2", "", 14, None, "FF0000")
        set_notes(slide, notes)
    elif kind == "imagepage":  # 结果证据页（图占位）
        set_placeholder(slide, 0, [(title, None, None, None)])
        set_named_box(slide, "TextBox 9", "【%s】" % typ, None, True, None)
        set_named_box(slide, "文本框 20", "【待放图：%s】" % (core if core else "全屏截图"), 14, None, "FF0000")
        set_notes(slide, notes)

def quote_text(core, title):
    # 金句页 BODY：引用句 + 出处行
    if core and core != title:
        return "\x0b".join(["\u201c%s\u201d" % core])
    return "\u201c%s\u201d" % title


# P43 链条总图正文特判
CHAIN_TEXT = "大模型 → Token → 上下文 → 提示词 → 工具 → 智能体 → 技能包 → Vibe Coding"

# ---------- 5. 主流程 ----------
def main():
    with open(SCRIPT, encoding="utf-8") as f:
        pages = parse_script(f.read())
    print("解析页数:", len(pages))
    if len(pages) != 52:
        raise SystemExit("页数不是52！")

    prs = Presentation(TEMPLATE)
    slides = list(prs.slides)  # 先取 9 个源页引用
    src_map = {"slide1": slides[0], "slide3": slides[2], "slide4": slides[3],
               "slide5": slides[4], "slide6": slides[5], "slide8": slides[7]}
    # 再移除母版原始 9 页（克隆用的是已保存的引用，不受影响）
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        sldIdLst.remove(sldId)

    created = []
    PREFIX_KIND = [("名词页", ("slide5", "content")), ("案例对比页", ("slide5", "content")),
                   ("对比/迭代页", ("slide5", "content"))]
    for page in pages:
        kind_name = PAGE_KIND.get(page["type"])
        if kind_name is None:
            for prefix, kv in PREFIX_KIND:
                if page["type"].startswith(prefix):
                    kind_name = kv
                    break
        if kind_name is None:
            raise SystemExit("未映射类型: %s (%s)" % (page["type"], page["page"]))
        src_name, kind = kind_name
        src = src_map[src_name]
        new_slide = clone_slide(prs, src)
        fill_slide(new_slide, page, kind)
        created.append(page["page"])

    prs.save(OUTPUT)
    print("已生成:", OUTPUT)
    print("页序:", created)

if __name__ == "__main__":
    main()
