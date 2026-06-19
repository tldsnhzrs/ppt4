# -*- coding: utf-8 -*-
"""废旧轮胎的回收与资源化利用 —— 中文学术演示 PPT 生成器。

遵循 cli-anything-wps 技能的设计系统：
  · 设计预设：academic（深蓝/橙/绿，视觉优先 65%，每页一主题）
  · 布局模板：cover / toc / overview / quadrant / stats / timeline /
              pipeline / grid_cards / content_image / data_table / closing
  · 演讲类型：conference（学术会议，约 20 页，15 分钟）

坐标系统：模板单位 = 磅(pt)，幻灯片 960×540pt（16:9 宽屏）。
图表由 make_charts.py 用 matplotlib 渲染（中文字体），作为图片嵌入。
"""
import os
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

HERE = os.path.dirname(__file__)
ASSETS = os.path.join(HERE, "assets")

# ====== academic 预设 ======
C = {
    "primary":   RGBColor(0x1A, 0x3C, 0x8B),   # 深蓝
    "secondary": RGBColor(0xE6, 0x77, 0x33),   # 橙
    "accent":    RGBColor(0x18, 0x80, 0x50),   # 绿
    "dark":      RGBColor(0x22, 0x22, 0x22),
    "gray":      RGBColor(0x80, 0x80, 0x80),
    "light":     RGBColor(0xF5, 0xF8, 0xFC),
    "white":     RGBColor(0xFF, 0xFF, 0xFF),
    "accent_lt": RGBColor(0xFF, 0xD9, 0xB8),   # 浅橙（深底标注）
    "light_txt": RGBColor(0xDD, 0xE4, 0xF0),
    "purple":    RGBColor(0x9B, 0x59, 0xB6),
    "red":       RGBColor(0xC8, 0x50, 0x4B),
}
TITLE_FONT = "微软雅黑"   # 用户 Windows/WPS 上的中文标准字体
BODY_FONT = "微软雅黑"

prs = Presentation()
prs.slide_width = Pt(960)
prs.slide_height = Pt(540)
BLANK = prs.slide_layouts[6]


# ---------- 基础工具 ----------
def slide():
    return prs.slides.add_slide(BLANK)


def _set_font(run, size, color, bold, font):
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.name = font
    # 东亚字体
    rpr = run._r.get_or_add_rPr()
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = rpr.makeelement(qn("a:ea"), {})
        rpr.append(ea)
    ea.set("typeface", font)


def text(s, x, y, w, h, content, size=18, color=None, bold=False,
         align=PP_ALIGN.LEFT, font=None, anchor=MSO_ANCHOR.TOP, line_spacing=1.15):
    color = color or C["dark"]
    font = font or BODY_FONT
    tb = s.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(4); tf.margin_right = Pt(4)
    tf.margin_top = Pt(2); tf.margin_bottom = Pt(2)
    lines = content.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        # 支持 [b]粗体段 + 颜色 via tuple markers? keep simple: whole-line
        r = p.add_run()
        r.text = ln
        _set_font(r, size, color, bold, font)
    return tb


def rich(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.2):
    """runs: list of paragraphs; each paragraph is list of (text,size,color,bold)."""
    tb = s.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(6); tf.margin_right = Pt(6)
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(4)
        for (t, sz, col, bd) in para:
            r = p.add_run()
            r.text = t
            _set_font(r, sz, col, bd, BODY_FONT)
    return tb


def rect(s, x, y, w, h, fill=None, line=None, line_w=1.0, shape=MSO_SHAPE.RECTANGLE,
         shadow=False):
    sp = s.shapes.add_shape(shape, Pt(x), Pt(y), Pt(w), Pt(h))
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    sp.shadow.inherit = False
    if shadow:
        el = sp._element.spPr
        ef = el.makeelement(qn("a:effectLst"), {})
        sh = ef.makeelement(qn("a:outerShdw"),
                            {"blurRad": "40000", "dist": "20000", "dir": "5400000",
                             "rotWithShape": "0"})
        clr = sh.makeelement(qn("a:srgbClr"), {"val": "9AA5B5"})
        alpha = clr.makeelement(qn("a:alpha"), {"val": "45000"})
        clr.append(alpha); sh.append(clr); ef.append(sh); el.append(ef)
    return sp


def box_text(s, x, y, w, h, content, size, color, bold=False, align=PP_ALIGN.CENTER,
             fill=None, line=None, anchor=MSO_ANCHOR.MIDDLE, shape=MSO_SHAPE.RECTANGLE,
             shadow=False, line_spacing=1.1):
    sp = rect(s, x, y, w, h, fill=fill, line=line, shape=shape, shadow=shadow)
    tf = sp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(8); tf.margin_right = Pt(8)
    tf.margin_top = Pt(4); tf.margin_bottom = Pt(4)
    for i, ln in enumerate(content.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run(); r.text = ln
        _set_font(r, size, color, bold, BODY_FONT)
    return sp


def pic(s, name, x, y, w=None, h=None):
    kw = {}
    if w: kw["width"] = Pt(w)
    if h: kw["height"] = Pt(h)
    return s.shapes.add_picture(os.path.join(ASSETS, name), Pt(x), Pt(y), **kw)


def pic_h(shape):
    """图片实际高度（pt），用于精确排布其后续元素（如图注）。"""
    return shape.height / 12700.0


def page_header(s, title, idx, accent=None):
    """内容页统一页眉：左竖条 + 标题 + 下划线 + 页码。"""
    accent = accent or C["secondary"]
    rect(s, 0, 0, 14, 540, fill=C["primary"])
    text(s, 50, 30, 740, 52, title, size=30, color=C["primary"], bold=True)
    rect(s, 52, 84, 150, 4, fill=accent)
    text(s, 880, 505, 70, 24, f"{idx:02d}", size=12, color=C["gray"], align=PP_ALIGN.RIGHT)


def footer(s):
    text(s, 50, 508, 500, 22, "废旧轮胎的回收与资源化利用", size=10, color=C["gray"])


# =====================================================================
# 幻灯片
# =====================================================================
N = [0]
def nxt():
    N[0] += 1
    return N[0]


# ---- S1 封面 (cover) ----
def s_cover():
    s = slide()
    rect(s, 0, 0, 960, 540, fill=C["primary"])
    rect(s, 0, 0, 960, 16, fill=C["secondary"])
    # 装饰：右下大圆环呼应轮胎
    rect(s, 700, 280, 320, 320, fill=None, line=RGBColor(0x33,0x55,0xA0), line_w=2,
         shape=MSO_SHAPE.OVAL)
    rect(s, 735, 315, 250, 250, fill=None, line=RGBColor(0x2A,0x4A,0x95), line_w=14,
         shape=MSO_SHAPE.OVAL)
    text(s, 60, 70, 600, 40, "环境工程 · 固体废物资源化", size=20,
         color=C["accent_lt"], bold=True)
    text(s, 60, 150, 840, 170, "废旧轮胎的回收\n与资源化利用", size=66,
         color=C["white"], bold=True, line_spacing=1.05)
    rect(s, 64, 360, 360, 5, fill=C["secondary"])
    text(s, 60, 380, 820, 70,
         "Recycling and Resource Utilization of Waste Tires",
         size=20, color=C["light_txt"])
    text(s, 60, 470, 600, 30, "汇报人：环境工程专业　|　指导教师：×××",
         size=16, color=C["light_txt"])
    text(s, 60, 500, 600, 26, "2026 年 6 月", size=14, color=C["gray"])


# ---- S2 目录 (toc) ----
def s_toc():
    s = slide()
    rect(s, 0, 0, 360, 540, fill=C["primary"])
    text(s, 50, 70, 280, 60, "目  录", size=46, color=C["white"], bold=True)
    text(s, 52, 145, 280, 30, "CONTENTS", size=16, color=C["accent_lt"])
    rect(s, 52, 185, 120, 5, fill=C["secondary"])
    items = [
        ("01", "研究背景与意义", "黑色污染 · 资源属性"),
        ("02", "回收体系与管理", "回收网络 · 政策法规"),
        ("03", "资源化利用途径", "翻新 · 胶粉 · 热解 · 再生胶"),
        ("04", "下游应用领域", "道路 · 场地 · 建材 · 能源"),
        ("05", "效益分析与挑战", "环境 · 经济 · 瓶颈"),
        ("06", "发展趋势与结论", "技术展望 · 总结"),
    ]
    y = 60
    for num, t, d in items:
        box_text(s, 410, y, 56, 56, num, 22, C["white"], bold=True,
                 fill=C["primary"], shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text(s, 485, y + 4, 430, 32, t, size=24, color=C["dark"], bold=True)
        text(s, 485, y + 38, 430, 22, d, size=14, color=C["gray"])
        if num != "06":
            rect(s, 485, y + 66, 430, 1, fill=RGBColor(0xE2,0xE6,0xEC))
        y += 76


# ---- 章节分隔页 ----
def s_section(no, title, sub):
    s = slide()
    rect(s, 0, 0, 960, 540, fill=C["primary"])
    rect(s, 0, 250, 960, 4, fill=RGBColor(0x33,0x55,0xA0))
    text(s, 80, 150, 300, 160, no, size=130, color=RGBColor(0x2E,0x52,0xA0), bold=True)
    rect(s, 300, 200, 5, 120, fill=C["secondary"])
    text(s, 340, 205, 560, 70, title, size=46, color=C["white"], bold=True)
    text(s, 342, 285, 560, 40, sub, size=20, color=C["accent_lt"])
    footer(s)


# ---- S3 概览 (overview) ----
def s_overview():
    s = slide()
    rect(s, 0, 0, 960, 96, fill=C["primary"])
    text(s, 0, 24, 960, 50, "为什么关注废旧轮胎？", size=34, color=C["white"],
         bold=True, align=PP_ALIGN.CENTER)
    cards = [
        ("体量巨大", "我国年产生量已超 2000 万吨，约 3.3 亿条，且以年均 6–8% 速度递增", C["primary"]),
        ("难以降解", "轮胎以橡胶、炭黑、钢丝、纤维复合而成，自然降解需上百年", C["secondary"]),
        ("污染严重", '露天堆存形成"黑色污染"，火灾、蚊媒、土壤与地下水风险并存', C["red"]),
        ("资源富集", "热值高达 32 MJ/kg，富含橡胶烃，是名副其实的'城市矿产'", C["accent"]),
    ]
    x0, y0, cw, ch, gx, gy = 40, 130, 435, 170, 50, 30
    for i, (t, d, col) in enumerate(cards):
        cx = x0 + (i % 2) * (cw + gx)
        cy = y0 + (i // 2) * (ch + gy)
        rect(s, cx, cy, cw, ch, fill=C["light"], shadow=True)
        rect(s, cx, cy, 10, ch, fill=col)
        text(s, cx + 30, cy + 20, cw - 60, 40, t, size=24, color=col, bold=True)
        text(s, cx + 30, cy + 72, cw - 55, 90, d, size=16, color=C["dark"],
             line_spacing=1.3)
    footer(s)


# ---- S4 危害 (content_image with hazard diagram) ----
def s_hazard():
    page_header(s := slide(), '"黑色污染"——废旧轮胎的环境危害', nxt())
    pic(s, "hazard.png", 470, 110, w=470)
    pts = [
        ("占用土地", "全球每年数十亿条退役，露天堆存如山，侵占大量土地资源"),
        ("火灾隐患", "轮胎易燃且燃烧释放 SO₂、多环芳烃等有毒黑烟，火场难以扑灭"),
        ("病媒滋生", "堆体积水成为蚊虫繁殖温床，传播登革热等疾病"),
        ("土壤水体", "降解析出锌、铅及有机助剂，污染土壤与地下水"),
        ("难以降解", "交联橡胶结构稳定，自然界中数百年难以分解"),
    ]
    y = 120
    for t, d in pts:
        box_text(s, 50, y, 14, 14, "", 8, C["white"], fill=C["secondary"],
                 shape=MSO_SHAPE.OVAL)
        text(s, 74, y - 6, 360, 26, t, size=18, color=C["primary"], bold=True)
        text(s, 74, y + 20, 380, 50, d, size=13.5, color=C["dark"], line_spacing=1.2)
        y += 76
    footer(s)


# ---- S5 数字统计 (stats + production chart) ----
def s_stats():
    page_header(s := slide(), "触目惊心的数字", nxt())
    stats = [
        ("20", "百万吨", "我国年产生量"),
        ("3.3", "亿条", "退役轮胎数量"),
        ("64", "%", "综合利用率"),
        ("32", "MJ/kg", "平均热值"),
    ]
    x = 50
    for num, unit, lab in stats:
        rect(s, x, 110, 200, 96, fill=C["primary"], shadow=True,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rich(s, x, 120, 200, 50, [[(num, 40, C["white"], True),
                                   (" " + unit, 16, C["accent_lt"], True)]],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, x, 172, 200, 26, lab, size=14, color=C["light_txt"],
             align=PP_ALIGN.CENTER)
        x += 222
    pic(s, "production_trend.png", 50, 232, w=475)
    rect(s, 545, 235, 365, 250, fill=C["light"])
    rect(s, 545, 235, 365, 44, fill=C["secondary"])
    text(s, 545, 245, 365, 30, "关键判读", size=20, color=C["white"], bold=True,
         align=PP_ALIGN.CENTER)
    text(s, 568, 292, 320, 180,
         "· 产生量持续刚性增长，2024 年突破 2000 万吨\n\n"
         "· 综合利用率稳步提升至 64%，但仍显著低于欧盟、日本（>90%）\n\n"
         "· 提升空间巨大，资源化是必由之路",
         size=14, color=C["dark"], line_spacing=1.25)
    footer(s)


# ---- S6 政策与发展 (timeline) ----
def s_timeline():
    page_header(s := slide(), "政策驱动与产业发展历程", nxt())
    rect(s, 70, 120, 4, 350, fill=C["primary"])
    events = [
        ("2011", "《废轮胎综合利用行业准入条件》发布，行业规范化起步"),
        ("2016", "工信部推行生产者责任延伸（EPR）制度试点"),
        ("2020", "《固体废物污染环境防治法》修订，强化全过程管理"),
        ("2021", '"无废城市"建设深入推进，胶粉改性沥青纳入推广'),
        ("2024", "双碳目标驱动，循环经济与资源化利用上升为国家战略"),
    ]
    y = 122
    for date, ev in events:
        rect(s, 60, y, 24, 24, fill=C["secondary"], shape=MSO_SHAPE.OVAL,
             line=C["white"], line_w=2)
        box_text(s, 100, y - 6, 90, 36, date, 20, C["primary"], bold=True, fill=None,
                 align=PP_ALIGN.LEFT)
        text(s, 200, y - 4, 380, 60, ev, size=16, color=C["dark"], line_spacing=1.2)
        y += 70
    # 侧栏
    rect(s, 620, 120, 300, 350, fill=C["light"])
    rect(s, 620, 120, 300, 46, fill=C["primary"])
    text(s, 620, 130, 300, 30, "政策内核", size=20, color=C["white"], bold=True,
         align=PP_ALIGN.CENTER)
    text(s, 642, 182, 258, 280,
         "EPR  生产者责任延伸\n谁生产、谁回收、谁负责\n\n"
         "减量化 · 资源化 · 无害化\n固废处理三大原则\n\n"
         "双碳 + 无废城市\n循环经济政策双轮驱动\n\n"
         "从'被动处置'转向'主动循环'",
         size=15, color=C["dark"], line_spacing=1.3)
    footer(s)


# ---- S8 回收体系 (pipeline) ----
def s_collection():
    page_header(s := slide(), "废旧轮胎回收体系流程", nxt())
    steps = [
        ("产生源", "汽修厂 · 4S店\n物流车队 · 个人", C["primary"]),
        ("回收网点", "区域回收站\n规范化收集", C["secondary"]),
        ("分拣分类", "按规格/损伤\n分级分流", C["accent"]),
        ("预处理", "去钢丝/纤维\n破碎清洗", C["purple"]),
        ("资源化", "翻新/胶粉/热解\n再生利用", C["primary"]),
        ("产品市场", "改性沥青/制品\n再生材料销售", C["secondary"]),
    ]
    x0, w, gap = 45, 130, 20
    for i, (t, d, col) in enumerate(steps):
        x = x0 + i * (w + gap)
        box_text(s, x, 140, w, 50, t, 19, C["white"], bold=True, fill=col,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(s, x, 200, w, 150, fill=C["light"])
        text(s, x, 220, w, 120, d, size=14, color=C["dark"], align=PP_ALIGN.CENTER,
             line_spacing=1.3)
        if i < len(steps) - 1:
            rect(s, x + w + 2, 158, gap - 4, 16, fill=col, shape=MSO_SHAPE.CHEVRON)
    rect(s, 45, 390, 870, 90, fill=C["primary"])
    text(s, 70, 404, 820, 70,
         "关键瓶颈：回收网络分散、'小散乱'非正规处置占比高、可追溯体系不完善。\n"
         "对策：依托 EPR 构建'互联网+回收'平台，推动规范化、规模化、可追溯的逆向物流。",
         size=16, color=C["white"], line_spacing=1.35, anchor=MSO_ANCHOR.MIDDLE)
    footer(s)


# ---- S10 四大途径概览 (quadrant) ----
def s_routes_overview():
    page_header(s := slide(), "资源化利用的四大途径", nxt())
    pic(s, "routes_pie.png", 560, 105, w=380)
    quad = [
        ("① 轮胎翻新", "在胎体上重贴胎面，恢复使用功能", "最高层级·能耗最低", C["accent"]),
        ("② 胶粉/改性", "常温/低温粉碎制胶粉，改性沥青、橡胶", "规模最大·应用最广", C["primary"]),
        ("③ 热裂解", "无氧高温分解为油、炭黑、钢丝、气", "彻底资源化·高值", C["secondary"]),
        ("④ 再生胶", "脱硫再生使橡胶恢复塑性可再加工", "传统工艺·量大", C["purple"]),
    ]
    y0 = 115
    for i, (t, d, tag, col) in enumerate(quad):
        cy = y0 + i * 98
        rect(s, 50, cy, 480, 86, fill=C["light"], shadow=True)
        rect(s, 50, cy, 10, 86, fill=col)
        text(s, 74, cy + 10, 300, 30, t, size=21, color=col, bold=True)
        box_text(s, 360, cy + 12, 158, 26, tag, 12.5, C["white"], fill=col,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        text(s, 74, cy + 44, 450, 36, d, size=14.5, color=C["dark"])
    text(s, 560, 470, 380, 40,
         "「减量化→资源化」优先序：\n翻新 > 胶粉 > 热解 > 焚烧填埋",
         size=14, color=C["gray"], align=PP_ALIGN.CENTER, line_spacing=1.2)
    footer(s)


# ---- 途径详情：通用图文页 ----
def s_route_detail(idx, title, badge, badge_col, points, img, img_caption,
                   highlight):
    page_header(s := slide(), title, idx, accent=badge_col)
    box_text(s, 800, 36, 150, 32, badge, 13.5, C["white"], bold=True, fill=badge_col,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    # 左侧要点
    y = 120
    for head, body in points:
        rect(s, 50, y + 4, 12, 12, fill=badge_col, shape=MSO_SHAPE.OVAL)
        text(s, 74, y - 4, 380, 28, head, size=19, color=C["primary"], bold=True)
        text(s, 74, y + 24, 400, 56, body, size=14.5, color=C["dark"], line_spacing=1.25)
        y += 84
    # 右侧图
    if img:
        p = pic(s, img, 500, 112, w=430)
        cap_y = 112 + pic_h(p) + 14
        text(s, 500, cap_y, 430, 24, img_caption, size=12, color=C["gray"],
             align=PP_ALIGN.CENTER)
    # 高亮条
    rect(s, 50, 464, 430, 40, fill=C["light"])
    rect(s, 50, 464, 8, 40, fill=badge_col)
    text(s, 70, 464, 410, 40, highlight, size=13.5, color=C["primary"], bold=True,
         anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.1)
    footer(s)


# ---- S16 应用领域 (grid_cards) ----
def s_applications():
    page_header(s := slide(), "再生材料的下游应用领域", nxt())
    pic(s, "applications.png", 510, 115, w=420)
    cards = [
        ("橡胶改性沥青", "掺入胶粉改善沥青高低温性能，路面更耐久、降噪", C["primary"]),
        ("塑胶运动场地", "胶粒铺设跑道、球场，弹性好、环保（需控 VOC）", C["accent"]),
        ("再生橡胶制品", "胶板、密封件、防水卷材、橡胶地砖等", C["secondary"]),
        ("建筑工程材料", "橡胶混凝土、隔震垫、轻质填料、护坡", C["purple"]),
    ]
    y0 = 120
    for i, (t, d, col) in enumerate(cards):
        cy = y0 + (i // 1) * 92 if False else y0 + i * 90
        rect(s, 50, cy, 430, 78, fill=C["light"], shadow=True)
        box_text(s, 62, cy + 14, 50, 50, str(i + 1), 24, C["white"], bold=True,
                 fill=col, shape=MSO_SHAPE.OVAL)
        text(s, 126, cy + 10, 350, 30, t, size=20, color=col, bold=True)
        text(s, 126, cy + 42, 350, 32, d, size=13.5, color=C["dark"], line_spacing=1.15)
    footer(s)


# ---- S17 技术对比 (data_table + radar) ----
def s_compare():
    page_header(s := slide(), "四类技术路线综合对比", nxt())
    headers = ["技术", "资源化程度", "经济性", "环境友好", "成熟度"]
    rows = [
        ["轮胎翻新", "中", "高", "高", "成熟"],
        ["胶粉改性", "高", "较高", "高", "成熟"],
        ["热裂解", "很高", "中", "中", "推广中"],
        ["再生胶", "较高", "中", "中", "成熟"],
    ]
    tx, ty, cw0, cwn, rh = 50, 115, 110, 78, 50
    # 表头
    cols_x = [tx]
    for k in range(len(headers)):
        cols_x.append(cols_x[-1] + (cw0 if k == 0 else cwn))
    for k, hd in enumerate(headers):
        box_text(s, cols_x[k], ty, cols_x[k+1]-cols_x[k], rh, hd, 15, C["white"],
                 bold=True, fill=C["primary"])
    for ri, row in enumerate(rows):
        ry = ty + rh * (ri + 1)
        bg = C["light"] if ri % 2 == 0 else C["white"]
        for k, cell in enumerate(row):
            col = C["primary"] if k == 0 else C["dark"]
            box_text(s, cols_x[k], ry, cols_x[k+1]-cols_x[k], rh, cell,
                     15 if k == 0 else 14, col, bold=(k == 0),
                     fill=bg, line=RGBColor(0xD8,0xDF,0xE8))
    pic(s, "radar.png", 510, 100, w=410)
    text(s, 50, 425, 430, 70,
         "结论：不存在单一最优技术。应按轮胎状况构建'梯级利用'体系——\n"
         "可翻新者优先翻新，其余制胶粉/再生胶高值利用，复杂废胎进入热解彻底回收。",
         size=14.5, color=C["primary"], bold=True, line_spacing=1.3)
    footer(s)


# ---- S19 效益分析 (stats + benefit chart) ----
def s_benefit():
    page_header(s := slide(), "环境效益与经济效益", nxt())
    p1 = pic(s, "benefit.png", 50, 112, w=310)
    pic(s, "country_rate.png", 50, 112 + pic_h(p1) + 15, w=310)
    # 右侧三块效益
    blocks = [
        ("环境效益", C["accent"],
         "· 每利用 1 吨废胎≈节约 0.7 吨石油\n· 减少 CO₂ 排放，缓解黑色污染\n· 替代天然/合成橡胶，降低开采压力"),
        ("经济效益", C["primary"],
         "· 形成胶粉—沥青—制品产业链\n· 热解油、炭黑、钢丝均可变现\n· 带动就业与区域循环经济"),
        ("社会效益", C["secondary"],
         "· 助力'无废城市'与双碳目标\n· 提升固废规范化管理水平\n· 城市矿产战略资源保障"),
    ]
    y = 110
    for t, col, d in blocks:
        rect(s, 460, y, 370, 122, fill=C["light"], shadow=True)
        rect(s, 460, y, 10, 122, fill=col)
        text(s, 487, y + 12, 335, 30, t, size=20, color=col, bold=True)
        text(s, 487, y + 46, 335, 72, d, size=13.5, color=C["dark"], line_spacing=1.25)
        y += 134
    footer(s)


# ---- S20 挑战与对策 (quadrant) ----
def s_challenge():
    page_header(s := slide(), "现存挑战与应对策略", nxt())
    pairs = [
        ("回收体系不健全", "非正规'土法炼油'屡禁不止，污染严重",
         "健全 EPR + 数字化回收平台，强化监管", C["secondary"], C["accent"]),
        ("技术装备待升级", "部分工艺能耗高、二次污染风险",
         "推广清洁热解、连续化智能化装备", C["secondary"], C["accent"]),
        ("产品标准与市场", "再生产品标准不统一、认可度偏低",
         "完善标准体系，政府绿色采购引导", C["secondary"], C["accent"]),
        ("经济性与补贴", "部分路线盈利薄、依赖政策",
         "税收优惠 + 碳交易 + 高值化产品", C["secondary"], C["accent"]),
    ]
    x0, y0, cw, ch, gx, gy = 50, 112, 430, 178, 50, 24
    for i, (prob, pd, sol, pc, sc) in enumerate(pairs):
        cx = x0 + (i % 2) * (cw + gx)
        cy = y0 + (i // 2) * (ch + gy)
        rect(s, cx, cy, cw, ch, fill=C["white"], line=RGBColor(0xDD,0xE3,0xEC), shadow=True)
        box_text(s, cx, cy, cw, 38, "⚠ " + prob, 17, C["white"], bold=True, fill=pc,
                 align=PP_ALIGN.LEFT)
        text(s, cx + 16, cy + 48, cw - 32, 44, pd, size=14, color=C["dark"],
             line_spacing=1.2)
        rect(s, cx + 16, cy + 98, cw - 32, 0.8, fill=RGBColor(0xDD,0xE3,0xEC))
        rich(s, cx + 16, cy + 106, cw - 32, 60,
             [[("对策　", 14, sc, True), (sol, 14, C["dark"], False)]],
             line_spacing=1.2)
    footer(s)


# ---- S21 趋势展望 (pipeline-ish) ----
def s_trend():
    page_header(s := slide(), "发展趋势与未来展望", nxt())
    trends = [
        ("高值化", "从燃料/填充向\n高端再生材料、\n精细化学品升级", C["primary"]),
        ("清洁化", "连续热解、绿色\n脱硫，降低能耗\n与二次污染", C["accent"]),
        ("智能化", "数字化回收平台\n+ AI 分拣，构建\n可追溯逆向物流", C["secondary"]),
        ("低碳化", "纳入碳核算与\n碳交易，服务\n双碳战略", C["purple"]),
        ("循环化", "全生命周期设计\n绿色轮胎，闭环\n材料循环", C["primary"]),
    ]
    x0, w, gap = 48, 158, 16
    for i, (t, d, col) in enumerate(trends):
        x = x0 + i * (w + gap)
        box_text(s, x, 130, w, 56, t, 22, C["white"], bold=True, fill=col,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, shadow=True)
        rect(s, x, 196, w, 170, fill=C["light"])
        text(s, x, 214, w, 150, d, size=15, color=C["dark"], align=PP_ALIGN.CENTER,
             line_spacing=1.35)
    rect(s, 48, 400, 866, 96, fill=C["primary"])
    text(s, 70, 414, 820, 76,
         "总体方向：构建'设计—回收—再生—应用'的全链条闭环体系，\n"
         "推动废旧轮胎从'环境负担'转变为'循环经济与双碳战略下的城市矿产'。",
         size=17, color=C["white"], bold=True, anchor=MSO_ANCHOR.MIDDLE,
         line_spacing=1.35, align=PP_ALIGN.CENTER)
    footer(s)


# ---- S22 结论 (closing) ----
def s_conclusion():
    page_header(s := slide(), "结论", nxt())
    concl = [
        ("现状", "废旧轮胎量大、难降解、污染重，但资源属性突出，是亟待开发的'城市矿产'。"),
        ("路径", "翻新、胶粉、热解、再生胶四大途径各有优劣，应构建梯级利用体系。"),
        ("效益", "资源化兼具显著的环境、经济与社会效益，契合循环经济与双碳目标。"),
        ("展望", "向高值化、清洁化、智能化、低碳化发展，实现全链条闭环循环。"),
    ]
    y = 120
    for i, (t, d) in enumerate(concl):
        col = [C["primary"], C["secondary"], C["accent"], C["purple"]][i]
        box_text(s, 50, y, 110, 64, t, 22, C["white"], bold=True, fill=col,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(s, 175, y, 740, 64, fill=C["light"])
        text(s, 195, y, 700, 64, d, size=17, color=C["dark"],
             anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.2)
        y += 84
    footer(s)


# ---- S23 致谢 (closing) ----
def s_thanks():
    s = slide()
    rect(s, 0, 0, 960, 540, fill=C["primary"])
    rect(s, 0, 0, 960, 16, fill=C["secondary"])
    rect(s, 0, 524, 960, 16, fill=C["secondary"])
    text(s, 0, 175, 960, 110, "感谢聆听", size=80, color=C["white"], bold=True,
         align=PP_ALIGN.CENTER)
    rect(s, 380, 305, 200, 5, fill=C["secondary"])
    text(s, 0, 330, 960, 40, "敬请各位老师批评指正", size=24, color=C["accent_lt"],
         align=PP_ALIGN.CENTER)
    text(s, 0, 430, 960, 30, "废旧轮胎的回收与资源化利用　|　环境工程专业",
         size=16, color=C["light_txt"], align=PP_ALIGN.CENTER)


# =====================================================================
# 组装（conference 演讲序列）
# =====================================================================
s_cover()                                                        # 1
s_toc()                                                          # 2
s_section("01", "研究背景与意义", "黑色污染 · 资源属性 · 政策驱动")   # 3
s_overview()                                                     # 4
s_hazard()                                                       # 5
s_stats()                                                        # 6
s_section("02", "回收体系与管理", "政策法规 · 回收网络")            # 7
s_timeline()                                                     # 8
s_collection()                                                   # 9
s_section("03", "资源化利用途径", "翻新 · 胶粉 · 热解 · 再生胶")     # 10
s_routes_overview()                                              # 11
s_route_detail(nxt(), "途径一：轮胎翻新（Retreading）", "资源化最高层级",
    C["accent"],
    [("原理", "检测合格胎体，打磨后重贴新胎面，硫化复原使用功能"),
     ("优势", "能耗与排放最低，1 条翻新胎≈节约 70% 原材料与能源"),
     ("局限", "仅适用胎体完好的载重/工程胎，乘用胎翻新比例低")],
    None, "",
    "环境友好度最高，应作为优先选项；我国翻新率仍远低于发达国家。")  # 12
s_route_detail(nxt(), "途径二：胶粉与橡胶改性沥青", "应用规模最大",
    C["primary"],
    [("制备", "常温/低温（冷冻）粉碎，去除钢丝纤维制成 20–200 目胶粉"),
     ("改性沥青", "胶粉掺入沥青形成弹性网络，提升高低温性能与抗车辙"),
     ("效益", "路面更耐久、降噪 3–5 dB，大量消纳废胎，技术成熟")],
    "applications.png", "胶粉再生材料下游应用占比",
    "是当前消纳量最大、最具推广价值的资源化方向。")              # 13
s_route_detail(nxt(), "途径三：热裂解（Pyrolysis）", "彻底资源化·高值",
    C["secondary"],
    [("工艺", "破碎后于无氧/缺氧反应釜中 400–600℃ 热裂解"),
     ("产物", "热解油≈45%、炭黑≈33%、钢丝≈12%、可燃气≈10%"),
     ("关键", "需控制二次污染，连续化、清洁化是产业升级方向")],
    "pyrolysis_flow.png", "废旧轮胎热裂解工艺流程与产物分布",
    "可实现近零废弃的彻底资源化，是高值化利用的重要前沿。")        # 14
s_route_detail(nxt(), "途径四：再生橡胶（Reclaimed Rubber）", "传统工艺·量大",
    C["purple"],
    [("原理", "通过热—氧—机械或脱硫剂作用，使硫化橡胶部分脱硫恢复塑性"),
     ("应用", "再生胶可部分替代生胶，用于轮胎、胶管、胶鞋等制品"),
     ("挑战", "传统动态脱硫能耗高、有异味，需向绿色脱硫升级")],
    None, "",
    "工艺成熟、消纳量大，绿色化、连续化脱硫是技术改进重点。")      # 15
s_section("04", "应用领域与对比", "道路 · 场地 · 建材 · 能源")       # 16
s_applications()                                                 # 17
s_compare()                                                      # 18
s_section("05", "效益、挑战与展望", "环境 · 经济 · 趋势")           # 19
s_benefit()                                                      # 20
s_challenge()                                                    # 21
s_trend()                                                        # 22
s_conclusion()                                                   # 23
s_thanks()                                                       # 24

OUT = os.path.join(HERE, "废旧轮胎的回收与资源化利用.pptx")
prs.save(OUT)
print(f"已生成 {len(prs.slides.__iter__.__self__._sldIdLst)} 页 -> {OUT}")
