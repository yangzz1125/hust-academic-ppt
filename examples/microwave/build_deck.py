"""完整示例：10 分钟课程小组汇报《微波炉为什么能加热食物》。

演示整套流程：维基百科条目作材料（source/），由材料参数推出的估算值写进 source/estimates.md；
背景和原理用 fetch_image.py 从 Wikimedia Commons 取图（figures/，来源和许可在同名 .json 里）；
页面用 Grid 排版，构图按页轮换：图为主、上宽下格、便当混合、流程 + 说明、原生示意图、原生图表、关键数字。

取图命令：
    fetch_image.py get "File:Microwave oven (interior).jpg" figures --name oven_interior --width 1600
    fetch_image.py get "File:Magnetron detail.jpg" figures --name magnetron
    fetch_image.py get "File:2021-06-18 12 31 52 The interior of a General Electric Turntable Microwave Oven Model
        JES738WH 001 in the Franklin Farm section of Oak Hill, Fairfax County, Virginia.jpg" figures --name turntable --width 1600

生成和质检：
    uv run --no-project --with python-pptx --with lxml --with xlsxwriter python -X utf8 build_deck.py <输出.pptx>
    uv run --no-project --with python-pptx python -X utf8 ../../scripts/qa_check.py <输出.pptx> --source source/*.txt source/estimates.md
"""
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))
from hust_deck import *  # noqa: E402

FIG = Path(__file__).parent / "figures"
LM = "Times New Roman"  # 公式里的上标、撇号和变量用同一种西文字体
SECTIONS = ["微波从哪里来", "微波怎么加热食物", "为什么会加热不均"]


deck = Deck("hust")
deck.cover("微波炉为什么能加热食物", "课程小组", "2026年10月14日")
deck.toc(SECTIONS)

# ======================================================== 一、微波从哪里来
deck.divider(SECTIONS[0])

# 图为主：炉腔照片 + 窄栏算波长
c = deck.content(SECTIONS[0], "家用微波炉工作在 2.45 GHz，波长 12.2 cm",
                 notes="先看炉腔：一个金属盒子，底部是玻璃转盘。微波炉用 2.45 GHz，用 λ = c/f 算出波长 12.2 cm，"
                       "门上的金属网孔比这小得多，所以微波出不去、光能透过。")
g = Grid()
c.web_figure(FIG / "oven_interior.jpg", g.box(0, 0, 8, 6), "家用微波炉的炉腔：四壁是金属，底部是玻璃转盘")
body = c.card(g.box(8, 0, 4, 3), header="波长怎么算")
c.text(Box(body.l + 0.5, body.t - 0.4, body.w - 1.0, 3.0), [
    P(*var("λ"), " = ", *var("c"), " / ", *var("f")),
    P("= 3 亿 m/s ÷ 24.5 亿 Hz"),
    P(R("≈ 0.122 m = 12.2 cm", color=TONES["accent"].ink, bold=True)),
], line=28)
body = c.card(g.box(8, 3, 4, 3), header="门上的网挡住微波", solid=False)
c.bullets(Box(body.l + 0.5, body.t - 0.4, body.w - 1.0, 3.6),
          ["网孔远小于 12.2 cm", "微波出不去，光能透过"], line=28)

# 图为主（7 + 5）：磁控管剖面 + 部件表
c = deck.content(SECTIONS[0], "磁控管用正交的电场和磁场让电子绕行，产生微波",
                 notes="按图注认部件。阴极发射电子，阳极加高压，永磁体让电子绕着阴极转，"
                       "电子群掠过谐振腔时激发出微波，再由 5、6 号的输出结构送出去。")
c.web_figure(FIG / "magnetron.jpg", g.box(0, 0, 7, 6),
             "磁控管剖面：1 阴极 2 阳极 3 永磁体 4 谐振腔 5 同轴输出线 6 输出针")
body = c.card(g.box(7, 0, 5, 6), header="四个关键部件")
c.table(body.l + 0.5, body.t - 0.4, [0.9, 2.6, body.w - 4.5], [
    ("", "部件", "作用"),
    ("1", "阴极", "加热后发射电子"),
    ("2", "阳极", "铜块，和阴极间加直流高压"),
    ("3", "永磁体", "提供轴向磁场，使电子绕行"),
    ("4", "谐振腔", "电子群掠过时激发微波振荡"),
], row_h=1.5)

# 流程 + 说明合在一张卡里；效率单独成卡，用两个关键数字
c = deck.content(SECTIONS[0], "电能经过高压电源、磁控管、波导三步进入炉腔",
                 notes="顺着箭头讲一遍能量怎么走，每个环节下面是它做的事。磁控管这一步把电能变成微波，用橙色标出。"
                       "最后看效率：普通型号只有一半左右的电能变成微波。")
flow = c.card(g.box(0, 0, 12, 4), header="能量路径")
xs = [flow.l + flow.w * (k + 0.5) / 5 for k in range(5)]
y = flow.cy + 0.9
# 说明按意群手动分行，避免「微波」这类词被自动换行拆开
steps = [("市电", ["交流电输入"], "primary"), ("高压直流电源", ["变压器升压", "电容和二极管倍压"], "primary"),
         ("磁控管", ["把直流电能变成", "2.45 GHz 微波"], "accent"), ("波导", ["一段短金属管", "把微波送进炉腔"], "primary"),
         ("炉腔和食物", ["微波在炉腔里反射", "被食物吸收"], "primary")]
nodes = []
for x, (name, desc, tone) in zip(xs, steps):
    nodes.append(c.node(x, y, name, tone, w=4.6))
    c.text(Box(x - 2.8, y - 1.0, 5.6, 2.0), desc, align="c", line=24,
           color=TONES["accent"].ink if tone == "accent" else TEXT)
for a, b in zip(nodes, nodes[1:]):
    c.link(a, b, color=TONES["primary"].base, lw=1.5)
eff = c.card(g.box(0, 4, 12, 2), header="效率：电能有多少变成了微波", solid=False, tone="accent")
for x, value, label in [(eff.l + eff.w * 0.25, "低至 50%", "普通型号"), (eff.l + eff.w * 0.75, "超过 64%", "节能型号")]:
    c.keynum(x, eff.b + (eff.h - 1.1) / 2, value, label)

# ======================================================== 二、微波怎么加热食物
deck.divider(SECTIONS[1])

# 图为主（原生示意图）：水分子跟着电场翻转
c = deck.content(SECTIONS[1], "水分子一端带正电、一端带负电，跟着电场来回转向",
                 notes="水分子的氧一端带负电，两个氢一端带正电。电场向上时氢一端朝上；半个周期后电场反向，分子跟着翻过来。"
                       "电场每秒振荡 24.5 亿次，分子来回转动，电场的能量就变成了热。")
big = c.card(g.box(0, 0, 8, 6), header="示意：电场反向，水分子跟着翻转")


def molecule(cx, cy, up):
    """水分子示意：up=True 时两个氢原子在上方（正电一端朝上）。"""
    s = 1 if up else -1
    o = (cx, cy - 0.7 * s)
    hs = [(cx - 1.5, cy + 0.6 * s), (cx + 1.5, cy + 0.6 * s)]
    for h in hs:
        c.line(o, h, GRAY, 3.0)
    shp = c.shape("原子", MSO_SHAPE.OVAL, place("c", *o, 1.5, 1.5), TONES["primary"].tint, TONES["primary"].base, 1.5)
    fill_text(shp.text_frame, "O", size=S_TITLE, color=INK, bold=True, align="c", anchor="m", wrap=False)
    c.at("c", o[0], o[1] - 1.25 * s, "δ−", size=S_SMALL, color=TONES["primary"].ink, bold=True)
    for h in hs:
        shp = c.shape("原子", MSO_SHAPE.OVAL, place("c", *h, 1.0, 1.0), TONES["neutral"].tint, GRAY, 1.25)
        fill_text(shp.text_frame, "H", size=S_BODY, color=INK, bold=True, align="c", anchor="m", wrap=False)
        c.at("c", h[0], h[1] + 0.95 * s, "δ+", size=S_SMALL, color=GRAY, bold=True)


def field(x, up):
    y0, y1 = (3.4, 9.9) if up else (9.9, 3.4)
    c.line((x, y0), (x, y1), TONES["accent"].base, 2.0, arrow=True)
    c.at("w", x + 0.25, 9.9 if up else 3.4, "E", size=S_BODY, color=TONES["accent"].ink, bold=True)


for cx, up, label in [(big.l + big.w * 0.27, True, "电场向上"), (big.l + big.w * 0.73, False, "半个周期后：电场向下")]:
    c.title("n", cx, big.t - 1.6, label)
    field(cx - 3.3, up)
    field(cx + 3.0, up)
    molecule(cx, 6.9, up)
c.line((big.cx - 0.7, 6.9), (big.cx + 0.7, 6.9), GRAY, 1.5, arrow=True)
body = c.card(g.box(8, 0, 4, 6), header="为什么会发热", solid=False)
c.bullets(Box(body.l + 0.5, body.t - 0.5, body.w - 1.0, 8.0),
          ["电场每秒振荡 24.5 亿次", "分子一直想对齐电场，跟着来回转动", "转动的能量变成热",
           "脂肪、糖也有极性基团，但吸热比水慢"], line=26)

# 便当混合：公式为主角，两个配角
c = deck.content(SECTIONS[1], "吸热功率与频率、损耗因子和场强平方成正比",
                 notes="这是介质加热的功率密度公式。代入 2.45 GHz 算出角频率；"
                       "两个推论：含水多的地方吸热快，电场强的地方吸热快，后者就是下一章热点的来源。")
big = c.card(g.box(0, 0, 7, 6), header="公式 · 介质加热的功率密度", tone="formula")
eq = Box(big.l + 0.5, big.t - 0.5, big.w - 1.0, 1.8)
c.tint(eq, "formula")
c.text(eq, P(*var("p"), " = ", *var("ω"), " ", *var("ε", "0"), " ", *var("ε", "r"), R("″", latin=LM), " ",
             *var("E"), R("2", latin=LM, sup=True)), size=S_TITLE, color=INK, align="c", anchor="m")
c.table(eq.l, eq.b - 0.5, [2.0, eq.w - 4.6, 2.6], [
    ("符号", "含义", "单位"),
    (P(*var("p")), "单位体积吸热功率", "W/m³"),
    (P(*var("ω")), "角频率", "rad/s"),
    (P(*var("ε", "0")), "真空介电常数", "F/m"),
    (P(*var("ε", "r"), R("″", latin=LM)), "损耗因子（介电常数虚部）", "—"),
    (P(*var("E")), "电场强度", "V/m"),
], row_h=1.2)
body = c.card(g.box(7, 0, 5, 3), header="代入 2.45 GHz", solid=False)
c.text(Box(body.l + 0.5, body.t - 0.5, body.w - 1.0, 2.2), [
    P(*var("ω"), " = 2π", *var("f"), " ≈ 154 亿 rad/s"),
    P("频率越高，吸热越快", before=6),
], line=26)
body = c.card(g.box(7, 3, 5, 3), header="两个推论", solid=False)
c.bullets(Box(body.l + 0.5, body.t - 0.5, body.w - 1.0, 2.4), [
    P("含水多的部分 ", *var("ε", "r"), R("″", latin=LM), " 大，先热"),
    P(*var("E"), " 大的地方吸热快：形成热点"),
], line=26)

# 便当混合：原生剖面示意 + 两个配角
c = deck.content(SECTIONS[1], "微波只进入食物约 1 cm，里面靠热传导",
                 notes="2.45 GHz 的微波只进入食物外层约 1 cm，厚食物的里面是靠外层传热热起来的，所以不是从里往外熟。"
                       "工业炉用 915 MHz，波长更长，能进得更深。")
big = c.card(g.box(0, 0, 7, 6), header="示意：食物剖面")
outer = Box(big.l + 2.2, big.t - 2.2, big.w - 4.4, 7.4)
inner = Box(outer.l + 1.4, outer.t - 1.4, outer.w - 2.8, outer.h - 2.8)
c.shape("示意", MSO_SHAPE.ROUNDED_RECTANGLE, outer, TONES["accent"].tint, TONES["accent"].base, 1.25, radius=0.5)
c.shape("示意", MSO_SHAPE.ROUNDED_RECTANGLE, inner, TONES["neutral"].tint, GRAY, 1.0, radius=0.3)
for x in (outer.l + outer.w * 0.25, outer.cx, outer.l + outer.w * 0.75):
    c.line((x, outer.t + 1.0), (x, outer.t - 0.6), TONES["accent"].base, 2.0, arrow=True)
c.at("e", outer.l + outer.w * 0.25 - 0.3, outer.t + 0.6, "微波", size=S_SMALL, color=TONES["accent"].ink, bold=True)
c.at("c", inner.cx, inner.cy + 0.6, "内部：靠热传导", color=INK)
for (x0, x1) in [(inner.l + 0.3, inner.l + 1.6), (inner.r - 0.3, inner.r - 1.6)]:
    c.line((x0, inner.cy - 0.6), (x1, inner.cy - 0.6), GRAY, 1.5, arrow=True)
c.at("n", outer.cx, outer.b - 0.3, "外层约 1 cm：微波直接加热", size=S_SMALL, color=TONES["accent"].ink, bold=True)
body = c.card(g.box(7, 0, 5, 3), header="频率越低，进得越深", solid=False)
c.text(Box(body.l + 0.5, body.t - 0.5, body.w - 1.0, 2.4), ["工业炉用 915 MHz", "波长 32.8 cm，能进得更深"], line=26)
body = c.card(g.box(7, 3, 5, 3), header="不是从里往外熟", solid=False)
c.text(Box(body.l + 0.5, body.t - 0.5, body.w - 1.0, 2.4), "外层 1 cm 先热，里面靠外层把热传进去", line=26)

# ======================================================== 三、为什么会加热不均
deck.divider(SECTIONS[2])

# 上宽下格：原生折线图（由驻波模型算出）+ 两张卡
c = deck.content(SECTIONS[2], "炉腔里形成驻波，热点间距约半个波长 6.1 cm",
                 notes="微波在金属壁之间来回反射形成驻波。按一维模型，加热强度和 E 的平方成正比，"
                       "峰值每隔半个波长出现一次，约 6.1 cm。用湿的热敏纸放进炉腔，变黑的地方就是热点。")
lam = 12.2
xs = list(range(31))
top = c.card(g.box(0, 0, 12, 4), header="一维简化模型：离炉壁 x 处的相对加热强度（∝ E²），横轴单位 cm")
c.chart("line", Box(top.l + 0.3, top.t - 1.2, top.w - 0.6, top.h - 1.6), [str(x) for x in xs],
        {"相对加热强度": [round(math.sin(2 * math.pi * x / lam) ** 2, 2) for x in xs]},
        unit="相对强度", number_format="0.0", vmax=1.0)
b = c.card(g.box(0, 4, 6, 2), header="估算", solid=False, tone="accent")
c.at("w", b.l + 0.5, b.cy, P("热点间距 = ", *var("λ"), " / 2 = 12.2 cm ÷ 2 = 6.1 cm"))
b = c.card(g.box(6, 4, 6, 2), header="怎么验证", solid=False)
c.at("w", b.l + 0.5, b.cy, "湿热敏纸放进炉腔，变黑的地方就是热点")

# 便当混合：照片为主角 + 两种办法 + 总结条
c = deck.content(SECTIONS[2], "转盘和搅拌器让食物轮流经过热点，中心仍可能不均",
                 notes="两种办法：转盘让食物动，搅拌器让场动。转盘占掉方形炉腔近四分之一的容积，"
                       "而且转盘中心附近的食物位置基本不变，仍可能受热不均。")
gb = Grid(bottombar=True)
c.web_figure(FIG / "turntable.jpg", gb.box(0, 0, 7, 6), "带玻璃转盘的家用微波炉炉腔")
body = c.card(gb.box(7, 0, 5, 3), header="转盘：让食物动", solid=False)
c.text(Box(body.l + 0.5, body.t - 0.5, body.w - 1.0, 2.4), "食物转一圈，轮流经过不同的热点", line=26)
body = c.card(gb.box(7, 3, 5, 3), header="搅拌器：让场动", solid=False)
c.text(Box(body.l + 0.5, body.t - 0.5, body.w - 1.0, 2.4), "顶部金属叶片转动，改变炉腔里的场分布", line=26)
c.bottombar("代价", "转盘占掉方形炉腔近四分之一的容积，中心附近仍可能受热不均")

# ======================================================== 总结
# 三张卡用箭头连成三步，底部数字带把每一步的关键数字对齐在它正下方
c = deck.content("总结", "微波炉加热靠三件事：产生微波、分子吸能、平均热点",
                 notes="三步各回顾一个数字：2.45 GHz 和 12.2 cm，约 1 cm 的穿透深度，6.1 cm 的热点间距。停在这一页答问。")
g3 = Grid(cols=3, gap=1.2)
cards = []
for i, (head, items) in enumerate([
        ("1  产生微波", ["磁控管把直流电变成微波", "高压电源 → 磁控管 → 波导", "门上的网挡住微波"]),
        ("2  分子吸能", ["水分子跟着电场来回转向",
                     P("功率 ∝ ", *var("ω"), " · 损耗因子 · ", *var("E"), R("2", latin=LM, sup=True)),
                     "只进入食物外层约 1 cm"]),
        ("3  平均热点", ["驻波在炉腔里形成热点", "热敏纸能显示热点位置", "转盘、搅拌器让食物轮流经过"])]):
    card = g3.box(i, 0, 1, 4)
    body = c.card(card, header=head, solid=i == 1, highlight=i == 1)
    # 条目纵向居中、拉开间距，不让要点挤在卡片上半部
    c.bullets(Box(body.l + 0.5, body.t - 0.3, body.w - 1.0, body.h - 0.6), items, line=28, anchor="m", gap=14)
    cards.append(card)
for a, b in zip(cards, cards[1:]):
    c.line((a.r + 0.15, a.cy), (b.l - 0.15, a.cy), TONES["primary"].base, 1.5, arrow=True)
band = c.card(g.box(0, 4, 12, 2), header="关键数字", solid=False, tone="accent")
for card, (value, label) in zip(cards, [("12.2 cm", "2.45 GHz 的波长"), ("约 1 cm", "穿透深度"), ("6.1 cm", "热点间距")]):
    c.keynum(card.cx, band.b + (band.h - 1.1) / 2, value, label)

deck.save(sys.argv[1])
print("written", sys.argv[1])
