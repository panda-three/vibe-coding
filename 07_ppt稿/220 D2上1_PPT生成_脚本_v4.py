# 220 D2上1 PPT 生成脚本 v4
# 输入：09_修改+备课回填/D2上1_PPT_v2.pptx（作模板，取其母版/版式/皮肤）
# 内容：07_ppt稿/220 D2上1_PPT页面内容深化_v4.md（28页丰满版）
# 输出：07_ppt稿/220 D2上1_PPT骨架_v4.pptx（新文件，不覆盖任何旧版）
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import copy

TEMPLATE = "/Users/panda/Desktop/Vibe Coding 3天营/09_修改+备课回填/D2上1_PPT_v2.pptx"
OUT = "/Users/panda/Desktop/Vibe Coding 3天营/07_ppt稿/220 D2上1_PPT骨架_v6.pptx"

RED = RGBColor(0xC0, 0x00, 0x00)
GRAY = RGBColor(0x66, 0x66, 0x66)

# (页型, 标题/引文, 标签, [正文行], 底部行, 副句)
PAGES = [
    ("cover", "D2上午 · 从一句话到400+", None, [], "朋友一句话 → 400多个付费用户——不是运气，是流程", None),
    ("quote", "“先验证，再投入”", None, [], None, "这四个字，是这个故事的地基"),
    ("content", "出发点：朋友一句话", "【内容】", ["“设计师出图太贵了。”", "就这一句，没头没尾"], "就这一句抱怨，值400多个付费用户", None),
    ("content", "场景：我那朋友是设计师，天天出图", "【内容】", ["AI 出图，一张几毛钱", "一天几百张，天天在掏", "一个月下来，上千块"], "天天掏钱的事，才是真事", None),
    ("content", "我当时什么反应？——“哦，贵就贵呗。”", "【内容】", ["我没当回事", "可那句话，它没走", "回去一想：这句话里，藏着三样东西"], "没走的那句话，里头有货", None),
    ("content", "真需求三特征", "【内容】", ["有人——谁在说这句话？设计师，一群人，不是一个人", "有痛——出图太贵，天天在掏钱", "高频——天天出图，天天痛，不是一年痛一次"], "一句话是抱怨，三样凑齐是需求", None),
    ("content", "边界：什么不是需求", "【内容】", ["“我家门口树太高”——有人、有痛", "但一年痛一次", "那不是需求，是闲事"], "三样凑不齐，别动手", None),
    ("content", "60秒：写一句你听过的抱怨", "【任务卡】", ["工作里、生活里，哪句抱怨你听得最多？", "“太贵了”“太慢了”“太烦了”——写下来，一句就行", "想不起来？你抱怨过别人也行，写下来"], "不是会写代码的人发现需求，是听见抱怨的人发现需求", None),
    ("quote", "“不是会写代码的人发现需求，是听见抱怨的人发现需求”", None, [], None, "你刚才写下的那句，可能就是你未来的“400个付费用户”的起点"),
    ("content", "五行模板", "【内容】", ["第一行：给谁用", "第二行：解决什么", "第三行：现在怎么做的", "第四行：贵在哪", "第五行：他愿不愿意换"], "前四行是事实，第五行是判断", None),
    ("content", "五行填好，长这样", "【内容】", ["给谁用：设计师", "解决什么：出图太贵", "现在怎么做的：自己找图、外包、硬扛", "贵在哪：一张几毛 × 一天几百张 = 一个月上千块", "愿不愿意换：能省一半成本——换"], "事实拷问清楚了，判断才有底气", None),
    ("content", "演示①：提示词原文", "【指令卡】", ["“帮我把‘设计师出图太贵了’这句话，拷问成五行需求模板：给谁用/解决什么/现在怎么做的/贵在哪/愿不愿意换。”"], "讲师机动作：切豆包 → 打这句 → 回车（等待期话术：稍等，它在想）", None),
    ("content", "演示②：它填的，都是有人、有痛、高频", "【内容】", ["看大屏中间，五行一行一行", "有人、有痛、高频——就是刚才那三样"], "先拷问，后动手——顺序不能反", None),
    ("content", "纠偏：直接问，它就给你个随便的", "【内容】", ["你要是直接问“做个生图网站行不行”", "它就给你一个随便的答案", "先拷问，后动手"], "顺序反了，答案就水了", None),
    ("content", "90秒：把你的抱怨逼成五行", "【任务卡】", ["基础版：照模板一行一行填", "进阶版：填完发豆包，让它补第五行", "挑战版：拷问完再问“这个需求值得做吗？”"], "你没画一张图，已经把一个需求想清楚了", None),
    ("content", "初版三件事", "【内容】", ["第一，让它能跑——点一下，出来", "第二，让它能用——丑没关系，功能是真的", "第三，先给人用——没自己藏着，直接给朋友"], "初版的敌人不是丑，是“没跑起来”", None),
    ("content", "我当时那个初版，丑", "【内容】", ["没有设计，没有品牌", "就是一个能出图的东西", "但我没等——直接给朋友：你先用"], "人先用起来，才知道哪里该改；自己憋着改，永远改不完", None),
    ("quote", "“初版的敌人不是丑，是‘没跑起来’”", None, [], None, "有人问“这么丑怎么好意思给人用”——我当时也这么想"),
    ("content", "生图网站本质", "【内容】", ["一个网页 + 一个调用的接口 = 生图网站", "网页好办，难在怎么让网站会出图——调别人的接口", "新问题：接口，怎么接？"], "我没学过接口。我的做法，你也能做", None),
    ("content", "喂文档法第一步：调研", "【内容】", ["不是拿到一家就接——先把几家生图接口的文档都翻一遍", "比什么：接口稳不稳、参数全不全、价格多少、返回啥", "像挑菜，先逛一圈"], "调研不花钱，踩坑才花钱", None),
    ("content", "喂文档法三步", "【内容】", ["① 接口链接——接口地址给它", "② 接口参数——要传什么：提示词、尺寸、数量、钥匙", "③ 返回结果——图片网址藏在哪个字段"], "就这三样喂进去，它基本就能用这个接口了", None),
    ("content", "演示①：提示词原文", "【指令卡】", ["“这是 XXX 生图接口的文档（链接+参数+返回示例），请帮我在我的网站里接入它：用户输入提示词→调这个接口→把返回的图片显示出来。”"], "讲师机动作：切豆包 → 打字 → 回车（等待期话术：它读文档呢）", None),
    ("content", "演示②：看，代码出来了", "【内容】", ["关键在这：我不是让 AI 猜接口", "我是把接口的“说明书”喂给它"], "先让它学会，再让它干活", None),
    ("content", "纠偏：不给文档，它就编一个", "【内容】", ["直接说“接个生图接口”", "它十有八九编一个——那接口根本不存在"], "喂的是真文档，出的才是真代码", None),
    ("quote", "“先让它学会，再让它干活”", None, [], None, "耳熟吗？昨天下午的主锤——今天，用在了接口上"),
    ("content", "上午三件事", "【内容】", ["听见一句话 → 逼成五行 → 搭个能跑的初版"], "没有一件需要“会写代码”——你干的活是听、是问、是判断", None),
    ("quote", "“听、问、判断——你天天都在干”", None, [], None, "400+的第一步，就是这么来的"),
    ("content", "下午：接口自己“翻车”了", "【悬念】", ["用着用着，出事了——接口不稳定", "时而慢、时而挂、时而出错", "怎么接住这一波？下午见"], "只许记住这一句：先验证，再投入", None),
]

def add_text(slide, x, y, w, h, text, size, bold=False, color=None, wrap=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.text = text
    r = tf.paragraphs[0].runs[0]
    r.font.size = Pt(size)
    r.font.bold = bold
    if color is not None:
        r.font.color.rgb = color
    return tb

def est_h(text, size=20):
    # 20pt 中文约 38 字/行（11.2in 宽），每行 0.45in
    import math
    lines = max(1, math.ceil(len(text) / 38))
    return lines * 0.45

def strip_non_title_placeholders(slide):
    # 新建页会从版式继承空占位符，编辑视图里显示"请输入正文内容/此页二级标题-阿里巴巴普惠体18号"等提示文字——全部剥掉，只留标题
    title = slide.shapes.title
    for ph in list(slide.placeholders):
        if title is not None and ph._element is title._element:
            continue
        ph._element.getparent().remove(ph._element)

def main():
    p = Presentation(TEMPLATE)
    cover_layout = p.slides[0].slide_layout
    quote_layout = p.slides[1].slide_layout
    content_layout = p.slides[2].slide_layout

    # 先记旧页数量；新页先加（命名为 slide16+ 避免冲突），最后再删旧页
    n_old = len(p.slides._sldIdLst)

    for kind, title, tag, bullets, bottom, sub in PAGES:
        if kind == "cover":
            s = p.slides.add_slide(cover_layout)
            s.shapes.title.text = title
            s.shapes.title.text_frame.paragraphs[0].runs[0].font.size = Pt(24)
            s.shapes.title.text_frame.paragraphs[0].runs[0].font.bold = True
            strip_non_title_placeholders(s)
            add_text(s, 1.89, 3.20, 10.17, 0.60, bottom, 20)
        elif kind == "quote":
            s = p.slides.add_slide(quote_layout)
            if s.shapes.title is not None:
                s.shapes.title.text = "记住"
            strip_non_title_placeholders(s)
            add_text(s, 5.61, 1.70, 7.20, 1.60, title, 32, bold=True, wrap=True)
            if sub:
                add_text(s, 5.61, 3.45, 7.20, 0.90, sub, 20, color=GRAY, wrap=True)
        else:
            s = p.slides.add_slide(content_layout)
            s.shapes.title.text = title
            s.shapes.title.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
            s.shapes.title.text_frame.paragraphs[0].runs[0].font.bold = True
            strip_non_title_placeholders(s)
            # v6：按用户要求去掉【内容】【任务卡】【指令卡】【悬念】标签框
            y = 1.85
            for b in bullets:
                h = est_h(b)
                add_text(s, 1.00, y, 11.20, h, b, 20, wrap=True)
                y += h + 0.12
            if bottom:
                add_text(s, 1.00, 5.72, 11.20, 0.60, bottom, 16, bold=True, color=RED, wrap=True)

    # 删掉模板自带的前 n_old 页
    xml_slides = p.slides._sldIdLst
    for sld in list(xml_slides)[:n_old]:
        rId = sld.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        p.part.drop_rel(rId)
        xml_slides.remove(sld)

    p.save(OUT)
    print("saved:", OUT, "slides:", len(p.slides._sldIdLst))

if __name__ == "__main__":
    main()
