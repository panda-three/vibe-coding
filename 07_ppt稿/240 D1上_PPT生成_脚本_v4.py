#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Step 4.6 大屏演示版生成（v6）：
修复①字号过小（大屏后排不可读）②无动画（方法论语阶段5"动画引导注意力"被跳过）。
方案：
- 字号体系放大：正文20pt/任务卡22pt/图注16pt/底部条16pt/金句32pt/question 26pt/chapter 22pt
- 正文逐段独立文本框 + 单击淡入动画（讲一句出一句；标题与标签固定不动画）
依据：096 A4.3"动画只引导注意力不炫技"、任务卡≥28pt判断规则、097"后排能读/指令大字"。
"""
import copy
import re
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches, Emu, Pt

BASE = "/Users/panda/Desktop/Vibe Coding 3天营/07_ppt稿"
TEMPLATE = BASE + "/课程PPT母版.pptx"
SCRIPT = BASE + "/240 D1上_PPT页面内容深化_v2.md"
OUTPUT = BASE + "/240 D1上_PPT骨架_v6.pptx"
MAT = BASE + "/素材"

# ---------- 字号体系（大屏演示版） ----------
SZ_CONTENT_TITLE = 28      # 内容页标题
SZ_LABEL = 16              # 【类型】标签（长标签 2 行也不溢出）
SZ_CONTENT_BODY = 20       # 普通内容页正文（原14）
SZ_TASK_BODY = 22          # 任务卡/指令卡/存屏页正文（原14）
SZ_FOOTNOTE = 16           # 底部条（原12）
SZ_IMAGE_CAP = 16          # 图注（原10）
SZ_QUESTION_TITLE = 24     # 悬念/桥页标题（原20）
SZ_QUESTION_BODY = 26      # 悬念/桥页正文（原24）
SZ_QUOTE_TITLE = 32        # 金句/锚点大字（原30）
SZ_QUOTE_BODY = 20         # 金句页补充（原16）
SZ_CHAPTER_BODY = 22       # 章节页正文（原18）
SZ_CHAPTER_TITLE = 26      # 章节页标题（原26）

# 布局（英寸）
BODY_X, BODY_W = 1.0, 11.2
BODY_Y0 = 1.85            # 正文起始 y
BODY_Y_MAX = 5.55         # 正文最大 y（避开底部条）
LINE_H = 0.34             # 20pt 行高约
LINE_H_TASK = 0.38        # 22pt 行高
PARA_GAP = 0.10           # 段间距
FOOT_X, FOOT_W, FOOT_Y = 1.0, 11.2, 5.72
Q_X, Q_W, Q_Y0 = 5.61, 6.30, 2.70   # question/quote 正文区（母版文本占位符位置）
CAP_X, CAP_W, CAP_Y = 1.0, 11.3, 6.75  # imagepage 图注（底部，避开图片）
CAP_MAX_Y = 7.42                       # imagepage 图注最大 y（页面高 7.5 留边）

# 任务卡/指令卡/存屏页类型 → 用大字号
TASK_TYPES = {"任务卡", "演示指令卡", "存屏页/任务卡"}

# ---------- 1. 解析深化表 ----------
def parse_deep(md):
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

# ---------- 4. 逐段文本框 + 动画 ----------
def _est_lines(text, width_in, size_pt):
    """估算文本显示行数：中文字符宽≈字号pt，留 5% 余量防折行误判"""
    if not text:
        return 1
    chars_per_line = max(1, int(width_in * 72 / size_pt * 0.95))
    n = 0
    for seg in text.split("\x0b"):
        n += max(1, -(-len(seg) // chars_per_line))
    return n

def _line_h(size_pt):
    """行高：字号×1.25/72 英寸"""
    return size_pt * 1.25 / 72.0

def fit_size(text, w, max_sz=32, target_lines=2):
    """按宽度给标题选字号：保证不超过 target_lines 行（优先大字号）"""
    for sz in (max_sz, 28, 24, 20, 18):
        if _est_lines(text, w, sz) <= target_lines:
            return sz
    return 18

def add_body_stack(slide, texts, x, y, w, size, bold=False, color=None,
                   align=None, animate=True, max_y=BODY_Y_MAX, gap=PARA_GAP):
    """把多段文字变成独立文本框纵向排列；空间不足时先紧凑重排（保动画粒度），
    仍不足才把剩余段合并进最后框，绝不压扁或溢出。返回 (shape列表, 结束y)"""
    shapes = []
    cy = y
    lh = _line_h(size)
    remaining = list(texts)
    while remaining:
        t = remaining.pop(0)
        if not t:
            continue
        lines = _est_lines(t, w, size)
        h = lines * lh + gap
        if cy + h > max_y:
            if gap > 0.045:
                # 紧凑重排：先删除已加框，用更小段距重排（保持逐段动画）
                for sh in shapes:
                    sh._element.getparent().remove(sh._element)
                return add_body_stack(slide, texts, x, y, w, size, bold, color,
                                      align, animate, max_y, gap=0.04)
            # 仍放不下：剩余所有文本合并进最后一个框，保证不出界
            rest = "\x0b".join(x for x in remaining if x)
            txt = t + ("\x0b" + rest if rest else "")
            tb = slide.shapes.add_textbox(Inches(x), Inches(cy), Inches(w),
                                          Inches(max(max_y - cy, 0.3)))
            tf = tb.text_frame
            tf.word_wrap = True
            clear_tf(tf)
            add_para(tf, txt, size, bold, color, align)
            shapes.append(tb)
            cy = max_y
            break
        tb = slide.shapes.add_textbox(Inches(x), Inches(cy), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        clear_tf(tf)
        add_para(tf, t, size, bold, color, align)
        shapes.append(tb)
        cy += h
    return shapes, cy

def add_single_textbox(slide, text, x, y, w, size, bold=False, color=None,
                       align=None, animate=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(0.6))
    tf = tb.text_frame
    tf.word_wrap = True
    clear_tf(tf)
    add_para(tf, text, size, bold, color, align)
    return tb

def add_click_fade_animations(slide, shapes, id_start=1000):
    """给一组 shape 加"单击淡入"动画（按传入顺序，每次点击出一个）。
    每个 shape 一个 par：set visible + fade 500ms，click 触发（delay=indefinite）。
    依据 096 A4.3：动画只引导注意力。"""
    if not shapes:
        return id_start
    slide_el = slide._element
    # 移除已有 timing（克隆母版一般没有）
    old = slide_el.find(qn('p:timing'))
    if old is not None:
        slide_el.remove(old)

    def tid():
        nonlocal id_start
        id_start += 1
        return id_start

    root = slide_el.makeelement(qn('p:timing'), {})
    tnLst = slide_el.makeelement(qn('p:tnLst'), {})
    par0 = slide_el.makeelement(qn('p:par'), {})
    cTn0 = slide_el.makeelement(qn('p:cTn'), {'id': str(tid()), 'dur': 'indefinite',
                                              'restart': 'never', 'nodeType': 'tmRoot'})
    child0 = slide_el.makeelement(qn('p:childTnLst'), {})
    seq = slide_el.makeelement(qn('p:seq'), {'concurrent': '1', 'nextAc': 'seek'})
    cTnSeq = slide_el.makeelement(qn('p:cTn'), {'id': str(tid()), 'dur': 'indefinite',
                                                'nodeType': 'mainSeq'})
    childSeq = slide_el.makeelement(qn('p:childTnLst'), {})

    for shp in shapes:
        spid = str(shp.shape_id)
        # 外层 par：click 触发
        p_par = slide_el.makeelement(qn('p:par'), {})
        p_ctn = slide_el.makeelement(qn('p:cTn'), {'id': str(tid()), 'fill': 'hold'})
        stCond = slide_el.makeelement(qn('p:stCondLst'), {})
        cond = slide_el.makeelement(qn('p:cond'), {'delay': 'indefinite'})
        stCond.append(cond)
        p_child = slide_el.makeelement(qn('p:childTnLst'), {})
        # 内层 par：delay 0
        i_par = slide_el.makeelement(qn('p:par'), {})
        i_ctn = slide_el.makeelement(qn('p:cTn'), {'id': str(tid()), 'fill': 'hold'})
        i_stCond = slide_el.makeelement(qn('p:stCondLst'), {})
        i_cond = slide_el.makeelement(qn('p:cond'), {'delay': '0'})
        i_stCond.append(i_cond)
        i_child = slide_el.makeelement(qn('p:childTnLst'), {})
        # set visible
        set_el = slide_el.makeelement(qn('p:set'), {})
        cBhvr = slide_el.makeelement(qn('p:cBhvr'), {})
        b_ctn = slide_el.makeelement(qn('p:cTn'), {'id': str(tid()), 'dur': '1', 'fill': 'hold'})
        b_st = slide_el.makeelement(qn('p:stCondLst'), {})
        b_cd = slide_el.makeelement(qn('p:cond'), {'delay': '0'})
        b_st.append(b_cd)
        b_ctn.append(b_st)
        tgt = slide_el.makeelement(qn('p:tgtEl'), {})
        spTgt = slide_el.makeelement(qn('p:spTgt'), {'spid': spid})
        tgt.append(spTgt)
        attrLst = slide_el.makeelement(qn('p:attrNameLst'), {})
        attr = slide_el.makeelement(qn('p:attrName'), {})
        attr.text = 'style.visibility'
        attrLst.append(attr)
        cBhvr.append(b_ctn)
        cBhvr.append(tgt)
        cBhvr.append(attrLst)
        set_el.append(cBhvr)
        to_el = slide_el.makeelement(qn('p:to'), {})
        strVal = slide_el.makeelement(qn('p:strVal'), {'val': 'visible'})
        to_el.append(strVal)
        set_el.append(to_el)
        i_child.append(set_el)
        # fade effect
        eff = slide_el.makeelement(qn('p:animEffect'), {'transition': 'in', 'filter': 'fade'})
        cBhvr2 = slide_el.makeelement(qn('p:cBhvr'), {})
        b_ctn2 = slide_el.makeelement(qn('p:cTn'), {'id': str(tid()), 'dur': '500', 'fill': 'hold'})
        b_st2 = slide_el.makeelement(qn('p:stCondLst'), {})
        b_cd2 = slide_el.makeelement(qn('p:cond'), {'delay': '0'})
        b_st2.append(b_cd2)
        b_ctn2.append(b_st2)
        tgt2 = slide_el.makeelement(qn('p:tgtEl'), {})
        spTgt2 = slide_el.makeelement(qn('p:spTgt'), {'spid': spid})
        tgt2.append(spTgt2)
        cBhvr2.append(b_ctn2)
        cBhvr2.append(tgt2)
        eff.append(cBhvr2)
        i_child.append(eff)
        # 组装
        i_ctn.append(i_stCond)
        i_ctn.append(i_child)
        i_par.append(i_ctn)
        p_child.append(i_par)
        p_ctn.append(stCond)
        p_ctn.append(p_child)
        p_par.append(p_ctn)
        childSeq.append(p_par)

    cTnSeq.append(childSeq)
    seq.append(cTnSeq)
    child0.append(seq)
    cTn0.append(child0)
    par0.append(cTn0)
    tnLst.append(par0)
    root.append(tnLst)
    slide_el.append(root)
    return id_start

# ---------- 5. 页型映射 ----------
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
    "迭代页": ("slide6", "imagepage"),
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
    "P17": ("ppt素材_04_贪吃蛇有分数_v1.png", 8.2, 2.4, 4.4),
    "P18": ("ppt素材_05_贪吃蛇绿蛇快速度_v1.png", 8.2, 2.4, 4.4),
}

# ---------- 6. 填充 ----------
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
        tsz = 28 if len(main) <= 14 else 24   # 封面主标题超长降档，防占位符溢出
        set_placeholder(slide, 0, [(full, tsz, True, None)])
        set_notes(slide, notes)
    elif kind == "chapter":
        set_placeholder(slide, 0, [(title, SZ_CHAPTER_TITLE, True, None)])
        set_placeholder(slide, 10, [(p[1:], 18, None, None)])
        if body:
            shps, _ = add_body_stack(slide, body, 2.8, 4.3, 7.7, SZ_CHAPTER_BODY,
                                     align="ctr", max_y=6.4)
            add_click_fade_animations(slide, shps)
        set_notes(slide, notes)
    elif kind == "question":
        label = TITLE_LABEL.get(typ, "思考")
        set_placeholder(slide, 0, [(label, None, True, None)])
        # 标题不用母版大占位符（其垂直居中会与正文重叠），自建文本框，高度按行数给足
        set_placeholder(slide, 10, [])
        qsz = fit_size(title, 7.2, max_sz=SZ_QUESTION_TITLE)
        qlines = _est_lines(title, 7.2, qsz)
        qtitle_h = qlines * _line_h(qsz) + 0.12
        qtb = add_single_textbox(slide, title, Q_X, 1.70, 7.2, qsz,
                                 bold=True, align="ctr")
        qtb.height = Inches(qtitle_h)
        body_y = 1.70 + qtitle_h + 0.18
        if body:
            shps, _ = add_body_stack(slide, body, Q_X, body_y, Q_W, SZ_QUESTION_BODY,
                                     align="ctr", max_y=6.25)
            anim = list(shps)
            if footnote:
                fb = add_single_textbox(slide, footnote, Q_X, 6.15, Q_W,
                                        SZ_FOOTNOTE, color="888888", align="ctr")
                anim.append(fb)
            add_click_fade_animations(slide, anim)
        set_notes(slide, notes)
    elif kind == "quote":
        label = TITLE_LABEL.get(typ, "金句")
        set_placeholder(slide, 0, [(label, None, True, None)])
        set_placeholder(slide, 10, [])
        qtxt = "\u201c%s\u201d" % title
        qsz = fit_size(qtxt, 7.2, max_sz=SZ_QUOTE_TITLE)
        qlines = _est_lines(qtxt, 7.2, qsz)
        qtitle_h = qlines * _line_h(qsz) + 0.12
        qtb = add_single_textbox(slide, qtxt, Q_X, 1.70, 7.2, qsz,
                                 bold=True, align="ctr")
        qtb.height = Inches(qtitle_h)
        body_y = 1.70 + qtitle_h + 0.18
        if body:
            shps, _ = add_body_stack(slide, body, Q_X, body_y, Q_W, SZ_QUOTE_BODY,
                                     color="666666", align="ctr", max_y=6.05)
            anim = list(shps)
            if footnote:
                fb = add_single_textbox(slide, footnote, Q_X, 6.15, Q_W,
                                        SZ_FOOTNOTE, color="888888", align="ctr")
                anim.append(fb)
            add_click_fade_animations(slide, anim)
        set_notes(slide, notes)
    elif kind == "content":
        tsz = SZ_CONTENT_TITLE
        if _est_lines(title, 9.59, tsz) > 1:   # 标题超过 1 行自动降档
            tsz = 24
        set_placeholder(slide, 0, [(title, tsz, True, None)])
        set_named_box(slide, "TextBox 9", [("【%s】" % typ, SZ_LABEL, True, None)])
        # 清空母版原正文框（文本框 5）与底部示例条（文本框 2），改用自建逐段文本框
        set_named_box(slide, "文本框 5", [])
        set_named_box(slide, "文本框 2", [])
        is_task = typ in TASK_TYPES
        body_sz = SZ_TASK_BODY if is_task else SZ_CONTENT_BODY
        line_h = LINE_H_TASK if is_task else LINE_H
        bw = 6.8 if p in EXTRA_IMAGES else BODY_W   # P17/P18 右侧有图，正文收窄
        shps, cy = add_body_stack(slide, body, BODY_X, BODY_Y0, bw, body_sz,
                                  max_y=BODY_Y_MAX)
        anim = list(shps)
        if footnote:
            fb = add_single_textbox(slide, footnote, FOOT_X, FOOT_Y, FOOT_W,
                                    SZ_FOOTNOTE, bold=True, color="C00000")
            anim.append(fb)
        add_click_fade_animations(slide, anim)
        set_notes(slide, notes)
    elif kind == "imagepage":
        set_placeholder(slide, 0, [(title, SZ_CONTENT_TITLE, True, None)])
        set_named_box(slide, "TextBox 9", [("【%s】" % typ, SZ_LABEL, True, None)])
        # 清空母版图注框，改用底部大字图注
        set_named_box(slide, "文本框 20", [])
        cap_lines = list(body)
        if footnote:
            cap_lines.append(footnote)
        if cap_lines:
            shps, _ = add_body_stack(slide, cap_lines, CAP_X, CAP_Y, CAP_W,
                                     SZ_IMAGE_CAP, max_y=CAP_MAX_Y)
            add_click_fade_animations(slide, shps)
        set_notes(slide, notes)

# ---------- 7. 插图 ----------
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

BOX21 = (3.09, 1.88, 5.80, 4.45)
BOX22 = (9.26, 1.88, 3.39, 4.45)

IMAGE_MAP = {
    "P07": ("ppt素材_01_星空海报_v1.png", None),
    "P08": ("ppt素材_02_简历网页_v1.png", None),
    "P16": ("ppt素材_03_贪吃蛇基础版_v1.png", None),
    "P37": ("ppt素材_02_简历网页_v1.png", None),
    "P48": ("ppt素材_01_星空海报_v1.png", "ppt素材_06_星空海报改后版_v1.png"),
    "P50": ("ppt素材_06_星空海报改后版_v1.png", None),
}

# ---------- 8. 主流程 ----------
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

    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst)[:9]:
        sldIdLst.remove(sldId)

    prs.save(OUTPUT)
    print("已生成:", OUTPUT)
    print("页数:", len(created))

if __name__ == "__main__":
    main()
