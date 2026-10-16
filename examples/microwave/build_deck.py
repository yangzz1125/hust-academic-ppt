"""完整示例：10 分钟课程小组汇报《微波炉为什么能加热食物》。

演示整套流程：维基百科条目作材料（source/），由材料参数推出的估算值写进 source/estimates.md；
背景和原理用 fetch_image.py 从 Wikimedia Commons 取图（figures/，来源和许可在同名 .json 里）。

提纲按「现象 → 产生 → 吸收 → 分布」组织：开头用四个日常现象引出问题，每章回答其中一两个；
同一对象（磁控管）、同一机制（分子转向与功率公式）只占一页，定性和定量放在一起讲。
构图按页轮换：照片 + 对照表、照片 + 步骤、流程 + 数字 + 对比表、原生示意图 + 公式、剖面对比、原生图表、对比表、关键数字。

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


def eps_r2():
    """损耗因子 εr″ 的写法：斜体 ε、正体下标 r、Times New Roman 的双撇号。"""
    return [*var("ε", "r"), R("″", latin=LM)]


deck = Deck("hust")
deck.cover("微波炉为什么能加热食物", "课程小组", "2026年10月14日")
deck.toc(SECTIONS)

# ======================================================== 引入：从现象提出问题
# 照片 + 对照表：每个现象对应一个要回答的问题和回答它的章节，后面每章都能回扣到这里
c = deck.content("从现象说起", "微波炉的四个日常现象，都能用电磁波解释",
                 notes="先从大家都见过的四个现象说起。家用炉都用 2.45 GHz，工业炉却常用 915 MHz；厚的食物外面烫了里面还凉；"
                       "隔着门能看到食物，微波却出不来；同一盘菜有冷有热。这四个现象分别对应微波怎么产生、"
                       "怎么变成热、在炉腔里怎么分布，也就是后面三章。")
g = Grid()
c.web_figure(FIG / "oven_interior.jpg", g.box(0, 0, 5, 6), "家用微波炉：金属炉腔、门上的观察窗、底部的玻璃转盘")
body = c.card(g.box(5, 0, 7, 6), header="现象背后的问题")
# 表格上沿和左侧照片上沿明显错开（照片居中放，上沿约 y 12.6）
c.table(body.l + 0.5, body.t - 0.5, [6.0, 8.1, body.w - 15.1], [
    ("现象", "要回答的问题", "章节"),
    ("家用炉都用 2.45 GHz", "怎么产生，为什么是这个频率", "一"),
    ("外面烫了，里面还凉", "微波怎么变成热，能进多深", "二"),
    ("看得见，微波却出不来", "金属炉腔对微波做了什么", "三"),
    ("同一盘菜有冷有热", "热点从哪来，怎样分散", "三"),
], row_h=1.3, first_col_bold=True)
line = Box(body.l + 0.5, body.b + 3.6, body.w - 1.0, 3.2)
c.tint(line, "accent")
c.text(Box(line.l + 0.5, line.t - 0.2, line.w - 1.0, line.h - 0.4), [
    P(R("主线", color=TONES["accent"].ink, bold=True)),
    P("磁控管产生微波 → 食物里的分子吸收微波变成热 → 微波在炉腔里的分布决定哪里热", before=4),
], line=24, anchor="m")

# ======================================================== 一、微波从哪里来
deck.divider(SECTIONS[0])

# 照片 + 步骤：部件不单列一张表，按工作过程逐步点到图上的编号，结构和原理在同一页讲完
c = deck.content(SECTIONS[0], "磁控管让电子绕行掠过谐振腔，像吹哨一样激发微波",
                 notes="对着照片上的编号讲五步。阴极加热发射电子；阳极是一块铜，和阴极之间加直流高压；"
                       "永磁体的轴向磁场让电子走弯曲的路；电子落到阳极的位置不均匀，电流绕着谐振腔流动来抹平，就形成振荡；"
                       "最后由输出线和输出针把微波引出去送进波导。谐振腔两壁像电容、圆孔像电感，频率完全由尺寸决定，"
                       "就像气流吹过哨子开口发出固定的音调。")
gb = Grid(bottombar=True)
c.web_figure(FIG / "magnetron.jpg", gb.box(0, 0, 5, 6), "1 阴极 2 阳极 3 永磁体 4 谐振腔 5 输出线 6 输出针")
card = gb.box(5, 0, 7, 6)
body = c.card(card, header="工作过程（括号里是图上的编号）")
steps = [("发射", "阴极（1）加热后发射电子"),
         ("加速", "阳极（2）是铜块，和阴极间加直流高压"),
         ("绕行", "永磁体（3）的轴向磁场让电子走弯曲的路"),
         ("振荡", "电子落点不均，电流绕谐振腔（4）流动形成振荡"),
         ("输出", "输出线（5）和输出针（6）把微波引进波导")]
for k, (name, desc) in enumerate(steps):
    y = body.t - 0.9 - k * 1.3
    c.num(body.l + 0.9, y, str(k + 1))
    c.at("w", body.l + 1.7, y, P(R(name + "  ", color=INK, bold=True), desc))
lc = Box(body.l + 0.5, body.b + 3.3, body.w - 1.0, 2.9)
c.tint(lc)
c.text(Box(lc.l + 0.5, lc.t - 0.2, lc.w - 1.0, lc.h - 0.4), [
    P(R("谐振腔为什么能定频率", color=TONES["primary"].ink, bold=True)),
    P("槽的两壁像电容、圆孔像电感，组成实心铜的 LC 电路", before=4),
    P("所以频率完全由谐振腔的尺寸决定"),
], line=24, anchor="m")
c.bottombar("类比", "像吹哨：气流掠过哨子开口发出固定音调，电子流掠过谐振腔激发出微波")

# 流程 + 数字 + 对比表：能量怎么走、走掉多少、为什么用这个频率，都是「电变成微波」这一段的事
c = deck.content(SECTIONS[0], "电能约一半变成 2.45 GHz 微波，这个频段全球通用",
                 notes="顺着箭头讲能量怎么走，磁控管这一步把电能变成微波，用橙色标出。左下看效率：普通型号只有一半左右，"
                       "节能型号超过 64%，磁控管 1 kW 输入大约得到 700 W 微波。右下回答为什么用 2.45 GHz："
                       "微波炉必须用 ISM 频段，避免干扰其他无线电业务；915 MHz 只在部分国家能用，5.8 GHz 产生大功率成本太高，"
                       "2.45 GHz 全球都能用。")
flow = c.card(g.box(0, 0, 12, 3), header="能量路径")
xs = [flow.l + flow.w * (k + 0.5) / 5 for k in range(5)]
y = flow.cy + 0.75
# 说明按意群手动分行，避免「微波」这类词被自动换行拆开
steps = [("市电", ["交流电输入"], "primary"), ("高压直流电源", ["变压器升压", "电容和二极管倍压"], "primary"),
         ("磁控管", ["把直流电能变成", "2.45 GHz 微波"], "accent"), ("波导", ["一段短金属管", "把微波送进炉腔"], "primary"),
         ("炉腔和食物", ["微波在炉腔里反射", "被食物吸收"], "primary")]
nodes = []
for x, (name, desc, tone) in zip(xs, steps):
    nodes.append(c.node(x, y, name, tone, w=4.6))
    c.text(Box(x - 2.8, y - 0.9, 5.6, 1.8), desc, align="c", line=22,
           color=TONES["accent"].ink if tone == "accent" else TEXT)
for a, b in zip(nodes, nodes[1:]):
    c.link(a, b, color=TONES["primary"].base, lw=1.5)
eff = c.card(g.box(0, 3, 5, 3), header="效率：电能有多少变成微波", solid=False, tone="accent")
for k, (value, label) in enumerate([("低至 50%", "普通型号"), ("超过 64%", "节能型号"), ("约 700 W", "1 kW 输入时")]):
    c.keynum(eff.l + eff.w * (k + 0.5) / 3, eff.b + (eff.h - 1.1) / 2, value, label)
body = c.card(g.box(5, 3, 7, 3), header="为什么是 2.45 GHz：只能用 ISM 频段")
accent = TONES["accent"].ink
c.table(body.l + 0.5, body.t - 0.25, [2.8, 2.8, 4.6, body.w - 11.2], [
    ("频率", "波长", "用在哪", "取舍"),
    ("915 MHz", "32.8 cm", "工业、商用炉", "只在部分国家可用"),
    (P(R("2.45 GHz", color=accent, bold=True)), P(R("12.2 cm", color=accent, bold=True)),
     P(R("家用微波炉", color=accent, bold=True)), P(R("全球都可用", color=accent, bold=True))),
    ("5.8 GHz", "约 5.2 cm", "不用于烹饪", "大功率成本太高"),
], row_h=1.05)

# ======================================================== 二、微波怎么加热食物
deck.divider(SECTIONS[1])

# 原生示意图 + 公式：分子转向（定性）和功率公式（定量）讲的是同一个机制，放在同一页
c = deck.content(SECTIONS[1], "水分子跟着电场转向发热，功率与频率和场强平方成正比",
                 notes="水分子氧一端带负电、氢一端带正电。电场每秒振荡 24.5 亿次，分子一直想对齐电场，来回转动，"
                       "推挤、碰撞周围的分子，电场的能量就变成了热。右边是定量的写法：单位体积吸热功率等于角频率、"
                       "真空介电常数、损耗因子和场强平方的乘积。损耗因子是介电常数的虚部，衡量材料把场能变成热的能力，"
                       "含水多的部分大，所以先热；场强大的地方吸热快，这就是第三章热点的来源。"
                       "常有人说微波炉靠水分子共振，其实所有极性分子在很宽的频率范围都能这样吸热。")
big = c.card(g.box(0, 0, 5, 6), header="示意：电场反向，水分子跟着翻转")


def molecule(cx, cy, up):
    """水分子示意：up=True 时两个氢原子在上方（正电一端朝上）。"""
    s = 1 if up else -1
    o = (cx, cy - 0.55 * s)
    hs = [(cx - 1.2, cy + 0.5 * s), (cx + 1.2, cy + 0.5 * s)]
    for h in hs:
        c.line(o, h, GRAY, 3.0)
    shp = c.shape("原子", MSO_SHAPE.OVAL, place("c", *o, 1.25, 1.25), TONES["primary"].tint, TONES["primary"].base, 1.5)
    fill_text(shp.text_frame, "O", size=S_BODY, color=INK, bold=True, align="c", anchor="m", wrap=False)
    c.at("c", o[0], o[1] - 1.05 * s, "δ−", size=S_SMALL, color=TONES["primary"].ink, bold=True)
    for h in hs:
        shp = c.shape("原子", MSO_SHAPE.OVAL, place("c", *h, 0.85, 0.85), TONES["neutral"].tint, GRAY, 1.25)
        fill_text(shp.text_frame, "H", size=S_SMALL, color=INK, bold=True, align="c", anchor="m", wrap=False)
        c.at("c", h[0], h[1] + 0.8 * s, "δ+", size=S_SMALL, color=GRAY, bold=True)


def field(x, up, outside):
    """电场箭头；outside 为 "w" 或 "e"，E 标在箭头外侧，不挨着分子的 δ+ 标注。"""
    y0, y1 = (6.3, 10.0) if up else (10.0, 6.3)
    c.line((x, y0), (x, y1), TONES["accent"].base, 2.0, arrow=True)
    dx = -0.2 if outside == "w" else 0.2
    c.at("e" if outside == "w" else "w", x + dx, 10.0 if up else 6.3, "E", size=S_BODY,
         color=TONES["accent"].ink, bold=True)


# 纵向从上到下：标题（11.25 起）→ 电场箭头和分子（10.0–6.3）→ 要点（5.6 起），互不重叠
for cx, up, label in [(big.l + big.w * 0.27, True, "电场向上"), (big.l + big.w * 0.73, False, "半周期后：电场向下")]:
    c.title("n", cx, big.t - 1.45, label)
    field(cx - 2.3, up, "w")
    field(cx + 2.1, up, "e")
    molecule(cx, 8.15, up)
c.line((big.cx - 0.35, 8.15), (big.cx + 0.35, 8.15), GRAY, 1.5, arrow=True)
# 窄卡片一行约 18 个字，每条控制在一行内
c.bullets(Box(big.l + 0.5, 5.6, big.w - 1.0, 3.6),
          ["电场每秒振荡 24.5 亿次，分子来回转", "转动时推挤、碰撞周围分子，变成热", "对液态水最有效，对脂肪和糖弱得多"],
          line=26)
body = c.card(g.box(5, 0, 7, 4), header="定量：介质加热的功率密度", tone="formula")
eq = Box(body.l + 0.5, body.t - 0.4, body.w - 1.0, 1.5)
c.tint(eq, "formula")
c.text(eq, P(*var("p"), " = ", *var("ω"), " ", *var("ε", "0"), " ", *eps_r2(), " ", *var("E"), R("2", latin=LM, sup=True)),
       size=S_TITLE, color=INK, align="c", anchor="m")
c.table(eq.l, eq.b - 0.35, [2.0, eq.w - 2.0], [
    (P(*var("p")), "单位体积吸热功率，W/m³"),
    (P(*var("ω")), P("角频率：2.45 GHz 时 ", *var("ω"), " = 2π", *var("f"), " ≈ 154 亿 rad/s")),
    (P(*eps_r2()), "损耗因子（介电常数虚部）：材料把场能变成热的能力"),
    (P(*var("E")), "电场强度，V/m"),
], row_h=1.05, header=False)
body = c.card(g.box(5, 4, 7, 2), header="推论", solid=False)
c.bullets(Box(body.l + 0.5, body.t - 0.25, body.w - 1.0, body.h - 0.4), [
    P("含水多的部分 ", *eps_r2(), " 大，先热；", *var("E"), " 大的地方形成热点"),
    "不是水分子共振：所有极性分子在很宽的频率范围都能吸热",
], line=24, anchor="m")

# 剖面对比：只画微波炉看不出「1 cm」意味着什么，和烤箱放在一起才看出差别
c = deck.content(SECTIONS[1], "微波直接加热外层约 1 cm，里面靠热传导",
                 notes="左边是烤箱和烧烤：红外线和热风只把热量落在食物表面薄薄一层，里面全靠传导。"
                       "右边是微波炉：2.45 GHz 的微波能直接加热多数食物外层约 1 cm，含水多时可达几厘米，再往里还是靠传导。"
                       "所以微波炉并不是从里往外熟；表层干、下面湿的时候，里面吸热反而多，才会有这种感觉。"
                       "工业炉用 915 MHz，波长 32.8 cm，能进得更深。")
big = c.card(g.box(0, 0, 7, 6), header="示意：热量先落在哪里")
for cx, name, src, inset, caption, src_y in [
        (big.l + big.w * 0.27, "烤箱、烧烤", "红外、热风", 0.3, "热量只落在表面薄薄一层", 10.0),
        (big.l + big.w * 0.73, "微波炉", "微波", 1.2, "直接加热外层约 1 cm", 9.0)]:
    c.title("n", cx, big.t - 0.45, name)
    outer = place("n", cx, 10.0, 6.4, 5.0)
    inner = Box(outer.l + inset, outer.t - inset, outer.w - 2 * inset, outer.h - 2 * inset)
    c.shape("示意", MSO_SHAPE.ROUNDED_RECTANGLE, outer, TONES["accent"].tint, TONES["accent"].base, 1.25, radius=0.4)
    c.shape("示意", MSO_SHAPE.ROUNDED_RECTANGLE, inner, TONES["neutral"].tint, GRAY, 1.0, radius=0.3)
    for x in (outer.l + outer.w * 0.25, outer.cx, outer.l + outer.w * 0.75):
        c.line((x, outer.t + 0.9), (x, src_y), TONES["accent"].base, 2.0, arrow=True)
    c.at("e", outer.l + outer.w * 0.25 - 0.25, outer.t + 0.55, src, size=S_SMALL, color=TONES["accent"].ink, bold=True)
    c.at("c", inner.cx, inner.cy + 0.45, "里面靠热传导", color=INK)
    for x0, x1 in [(inner.l + 0.3, inner.l + 1.3), (inner.r - 0.3, inner.r - 1.3)]:
        c.line((x0, inner.cy - 0.55), (x1, inner.cy - 0.55), GRAY, 1.5, arrow=True)
    c.at("n", cx, outer.b - 0.3, caption, color=TONES["accent"].ink, bold=True)
c.at("s", big.cx, big.b + 0.5, "含水多的食物，微波直接加热的深度可达数厘米", size=S_SMALL, color=GRAY)
body = c.card(g.box(7, 0, 5, 3), header="频率越低，进得越深", solid=False)
c.bullets(Box(body.l + 0.5, body.t - 0.4, body.w - 1.0, body.h - 0.6), [
    "2.45 GHz 只进入多数食物约 1 cm", "工业炉常用 915 MHz，波长 32.8 cm", "频率越低，电场进入非导体越深",
], line=26, gap=4)
body = c.card(g.box(7, 3, 5, 3), header="不是从里往外熟", solid=False)
c.bullets(Box(body.l + 0.5, body.t - 0.4, body.w - 1.0, body.h - 0.6), [
    "外层先热，里面靠热传导", "表层干、下面湿时，里面吸热更多", "解冻用低功率，给热传导留时间",
], line=26, gap=4)

# ======================================================== 三、为什么会加热不均
deck.divider(SECTIONS[2])

# 原生图表 + 屏蔽 + 估算 + 验证：金属炉腔反射微波，一面把微波关在里面，一面形成驻波热点
c = deck.content(SECTIONS[2], "金属腔里微波来回反射成驻波，热点相距约 6.1 cm",
                 notes="炉腔四壁是金属，像一个法拉第笼，微波出不去；门上的观察窗有一层导电金属网，网孔比 12.2 cm 小得多，"
                       "微波也出不去，可见光波长短得多，能透过。微波在金属壁之间来回反射，炉腔尺寸固定，就形成驻波。"
                       "按一维模型，加热强度和 E 的平方成正比，峰值每隔半个波长出现一次，约 6.1 cm。"
                       "用湿的热敏纸放进炉腔，变黑的地方就是热点；超市小票就是热敏纸，在家就能做。")
lam = 12.2
xs = list(range(31))
top = c.card(g.box(0, 0, 8, 4), header="一维模型：离炉壁 x cm 处的相对加热强度（∝ E²）")
c.chart("line", Box(top.l + 0.3, top.t - 0.3, top.w - 0.6, top.h - 0.75), [str(x) for x in xs],
        {"相对加热强度": [round(math.sin(2 * math.pi * x / lam) ** 2, 2) for x in xs]},
        unit="相对强度", number_format="0.0", vmax=1.0)
body = c.card(g.box(8, 0, 4, 4), header="同一面金属壁挡住微波", solid=False)
c.bullets(Box(body.l + 0.5, body.t - 0.4, body.w - 1.0, body.h - 0.6),
          ["炉腔像法拉第笼", "观察窗里有导电金属网", "网孔远小于 12.2 cm", "微波出不去，只能来回反射", "可见光波长短，能透过"],
          line=26, gap=4)
b = c.card(g.box(0, 4, 6, 2), header="估算", solid=False, tone="accent")
c.text(Box(b.l + 0.5, b.t - 0.15, b.w - 1.0, b.h - 0.3), [
    P(*var("λ"), " = ", *var("c"), " / ", *var("f"), " ≈ 12.2 cm；热点间距 ", *var("λ"), " / 2 = 6.1 cm"),
    "热点离炉壁约 3.05、9.15、15.25 cm……",
], line=24, anchor="m")
b = c.card(g.box(6, 4, 6, 2), header="怎么验证", solid=False)
c.text(Box(b.l + 0.5, b.t - 0.15, b.w - 1.0, b.h - 0.3), [
    "湿热敏纸放进炉腔，变黑的地方就是热点", "超市小票就是热敏纸，在家就能做",
], line=24, anchor="m")

# 照片 + 对比表：四种办法同类、条目多，用表格横向比较，不排成几张小卡片
c = deck.content(SECTIONS[2], "四种办法让加热更均匀，转盘最常见但中心仍可能不均",
                 notes="让加热均匀有四种办法。转盘让食物动，是家用最常见的，夏普 1964 到 1966 年间推出，比早期的搅拌器更均匀，"
                       "代价是占掉方形炉腔近四分之一的容积，而且中心附近的食物位置基本不变。搅拌器让场动，叶片藏在顶部的塑料盖板上方，"
                       "商用炉用得多；商用炉也有用几个磁控管的。变频电源是从时间上改善：低档位持续输出低功率，不再整功率开关。"
                       "使用时把食物放得偏离转盘中心，受热更均匀。")
gb = Grid(bottombar=True)
c.web_figure(FIG / "turntable.jpg", gb.box(0, 0, 4, 6), "带玻璃转盘的家用微波炉炉腔")
body = c.card(gb.box(4, 0, 8, 6), header="四种办法比较")
# 说明列一行约 17 个字，按意群分两行写，不让自动换行拆开词语
c.table(body.l + 0.5, body.t - 0.3, [3.4, 2.8, 2.6, body.w - 9.8], [
    ("办法", "改变什么", "常见于", "说明"),
    ("转盘", "食物位置", "家用", ["夏普 1964–1966 年推出", "比搅拌器均匀，但占掉近 1/4 容积"]),
    ("搅拌器", "场分布", "商用较多", ["反射微波的叶片在顶部转动", "藏在透波塑料盖板上方，食物不动"]),
    ("多个磁控管", "微波源", "商用", "用几个磁控管代替一个"),
    ("变频电源", "输出功率", "新型号", ["低档位持续输出低功率", "不再整功率开关，加热更均匀"]),
], row_h=1.85, first_col_bold=True)
c.bottombar("使用建议", "食物偏离转盘中心放，受热更均匀；最便宜的机型没有这些机构，要中途停下搅拌")

# ======================================================== 总结
# 三张卡用箭头连成三步，底部数字带把每一步的关键数字对齐在它正下方
c = deck.content("总结", "微波炉加热靠三件事：产生微波、分子吸能、平均热点",
                 notes="回到开头的四个现象：2.45 GHz 是全球通用的 ISM 频段；外热里凉是因为微波只直接加热外层约 1 cm；"
                       "看得见却出不来是因为网孔远小于波长；冷热不均来自驻波热点，转盘和搅拌器让食物轮流经过热点。停在这一页答问。")
g3 = Grid(cols=3, gap=1.2)
cards = []
for i, (head, items) in enumerate([
        ("1  产生微波", ["磁控管：电子掠过谐振腔", "约一半电能变成微波", "2.45 GHz 全球通用", "门上的网挡住微波"]),
        ("2  分子吸能", ["水分子跟着电场来回转向",
                     P("功率 ∝ ", *var("ω"), " · 损耗因子 · ", *var("E"), R("2", latin=LM, sup=True)),
                     "只直接加热外层约 1 cm", "里面靠热传导"]),
        ("3  平均热点", ["金属腔里形成驻波", "热点相距约 6.1 cm", "转盘、搅拌器分散热点", "变频电源低档更均匀"])]):
    card = g3.box(i, 0, 1, 4)
    body = c.card(card, header=head, solid=i == 1, highlight=i == 1)
    c.bullets(Box(body.l + 0.5, body.t - 0.3, body.w - 1.0, body.h - 0.6), items, line=26, anchor="m", gap=8)
    cards.append(card)
for a, b in zip(cards, cards[1:]):
    c.line((a.r + 0.15, a.cy), (b.l - 0.15, a.cy), TONES["primary"].base, 1.5, arrow=True)
band = c.card(g.box(0, 4, 12, 2), header="关键数字", solid=False, tone="accent")
for card, (value, label) in zip(cards, [("12.2 cm", "2.45 GHz 的波长"), ("约 1 cm", "直接加热的深度"), ("6.1 cm", "热点间距")]):
    c.keynum(card.cx, band.b + (band.h - 1.1) / 2, value, label)

deck.save(sys.argv[1])
print("written", sys.argv[1])
