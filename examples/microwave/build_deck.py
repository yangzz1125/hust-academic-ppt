"""完整示例：10 分钟课程小组汇报《微波炉为什么能加热食物》。

演示整套流程：维基百科条目作材料（source/），由材料参数推出的估算值写进 source/estimates.md；
背景和原理用 fetch_image.py 从 Wikimedia Commons 取图（figures/，来源和许可在同名 .json 里）。

提纲按「现象 → 产生 → 吸收 → 分布」组织：开头用四个日常现象引出问题，每章回答其中一两个；
同一对象（磁控管）、同一机制（分子转向与功率公式）只占一页，定性和定量放在一起讲。
页型按页轮换（命名页型见 references/layouts.md 第 7 节）：照片 + 对照表、时间线、引线标注、大序号步骤、公式拆解、
剖面对比、原生图表、照片 + 对比表、关键数字。器件和历史节点都用实物照片，原生形状只画热量走向和驻波这类拍不出来的东西。

取图命令：
    fetch_image.py get "File:Microwave oven (interior).jpg" figures --name oven_interior --width 1600
    fetch_image.py get "File:Magnetron detail.jpg" figures --name magnetron
    fetch_image.py get "File:2021-06-18 12 31 52 The interior of a General Electric Turntable Microwave Oven Model
        JES738WH 001 in the Franklin Farm section of Oak Hill, Fairfax County, Virginia.jpg" figures --name turntable --width 1600
    fetch_image.py get "File:Randall and Boot cavity magnetron Science museum.JPG" figures --name magnetron_1940 --width 1400
    fetch_image.py get "File:Radarange NS Svannah.jpg" figures --name radarange_early
    fetch_image.py get "File:Amana Radarange Microwave Oven.jpg" figures --name amana_radarange --width 1400
    fetch_image.py get "File:Plug sockets.jpg" figures --name sockets --width 1200
    fetch_image.py get "File:Caso SMG20 - MD Microwave Oven Manufacturing MD-801EMR-1-0190.jpg" figures --name transformer --width 1200
    fetch_image.py get "File:Caso SMG20 - Witol 2M219J-0166.jpg" figures --name magnetron_unit --width 1200
    fetch_image.py get "File:Waveguide collection.jpg" figures --name waveguides   # 再裁取左上部分 (0, 0, 600, 410)

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

# 时间线：横轴贯穿页宽，年份用 S_MARK；节点上方放当年的实物照片，同样大小（fill=True 裁齐）；
# 没有照片的关键节点（1945 年的偶然发现）用强调色底块写一句经过，圆点也放大
c = deck.content(SECTIONS[0], "微波炉源于雷达工厂的一次偶然发现，约 30 年后才进入家庭",
                 notes="磁控管是二战时为雷达研制的，1940 年伯明翰大学造出腔体磁控管原型，输出约 400 W。"
                       "1945 年雷声公司的斯宾塞在雷达旁工作，发现口袋里的糖果棒化了，第一样有意做熟的食物是爆米花。"
                       "1947 年第一台商用微波炉 Radarange 上市，高约 1.8 米、耗电 3 千瓦、要水冷，只能租给饭店。"
                       "之后夏普加了转盘，Amana 做出台式机，到 1970 年代末价格才降到普通家庭买得起，前后约 30 年。")
AY, pitch = 8.2, 29.8 / 6  # 时间轴的 y、节点间距
events = [("1940", "magnetron_1940", "原型腔体磁控管", ("伯明翰大学造出", "腔体磁控管原型"), ("输出约 400 W", "用于二战雷达")),
          ("1945", None, None, ("斯宾塞偶然发现", "微波能加热食物"), ("第一样有意做熟的", "食物是爆米花")),
          ("1947", "radarange_early", "早期 Radarange", ("第一台商用微波炉", "Radarange 上市"), ("高约 1.8 m", "耗电 3 kW，水冷")),
          ("1964–1966", "turntable", "带转盘的炉腔", ("夏普推出", "带转盘的微波炉"), ("比早期的搅拌器", "加热更均匀")),
          ("1967", "amana_radarange", "Amana 台式微波炉", ("Amana 推出", "台式微波炉"), ("体积小到", "能放上台面")),
          ("1970s 末", "oven_interior", "家用微波炉", ("价格降到", "家庭买得起"), ("1980 年代", "价格快速下降"))]
c.line((0.6, AY), (30.4, AY), TONES["primary"].base, 2.5, arrow=True)
pics = {k: c.web_figure(FIG / f"{img}.jpg", Box(0.6 + pitch * (k + 0.5) - 2.2, 13.8, 4.4, 3.6), cap, fill=True)
        for k, (_, img, cap, _, _) in enumerate(events) if img}
for k, (year, img, _, head, detail) in enumerate(events):
    x = 0.6 + pitch * (k + 0.5)
    t = TONES["primary" if img else "accent"]
    if not img:  # 和照片同样大小的底块
        slot = Box(x - 2.2, 13.8, 4.4, pics[0].h)
        c.tint(slot, "accent")
        c.text(slot, ["雷达让他", "口袋里的", "糖果棒化了"], color=t.ink, bold=True, align="c", anchor="m", line=22)
    d = 0.45 if img else 0.7
    c.shape("节点", MSO_SHAPE.OVAL, place("c", x, AY, d, d), t.base)
    c.at("s", x, AY + 0.5, year, size=S_MARK, color=t.ink, bold=True)
    c.text(Box(x - 2.4, AY - 0.55, 4.8, 1.6), list(head), color=INK, bold=True, align="c", line=22)
    c.text(Box(x - 2.4, AY - 2.3, 4.8, 1.3), list(detail), size=S_SMALL, color=GRAY, align="c", line=18)
band = Box(0.6, 4.2, 29.8, 2.6)
c.shape("底色", MSO_SHAPE.ROUNDED_RECTANGLE, band, TONES["primary"].base, radius=0.247)
c.text(Box(band.l + 0.7, band.t - 0.2, 21.0, band.h - 0.4), [
    P(R("从雷达到厨房", bold=True, size=S_TITLE)),
    P("磁控管为二战雷达研制；商用机器又高又重，体积和价格降下来后才进入家庭", before=4),
], color=WHITE, line=24, anchor="m")
c.at("s", 26.6, band.cy - 0.05, "约 30 年", size=S_MARK, color=WHITE, bold=True)
c.at("n", 26.6, band.cy - 0.15, "1947 年商用 → 1970s 末进入家庭", size=S_SMALL, color=WHITE)

# 引线标注：照片居中占半页，部件说明分在两侧浅底块里，引线从块指到图上自带的编号；
# 按编号在图上的高低排块，引线不交叉。部件不单列一张表，结构和工作过程在同一页讲完
c = deck.content(SECTIONS[0], "磁控管让电子绕行掠过谐振腔，像吹哨一样激发微波",
                 notes="对着照片上的编号讲。阴极 1 加热后发射电子；阳极 2 是一块铜，和阴极之间加直流高压；"
                       "永磁体 3 的轴向磁场让电子走弯曲的路；电子落到阳极的位置不均匀，电流绕着谐振腔 4 流动来抹平，就形成振荡；"
                       "最后由输出线 5 和输出针 6 把微波引出去送进波导。谐振腔两壁像电容、圆孔像电感，频率完全由尺寸决定，"
                       "就像气流吹过哨子开口发出固定的音调。")
gb = Grid(bottombar=True)
img = FIG / "magnetron.jpg"
pic = c.web_figure(img, gb.box(3, 0, 6, 6), "拆开的家用磁控管，白色数字为部件编号", align="n")
# 编号在原图上的像素位置（1665×1143，看图量出）
for col, side, parts in [(0, "left", [(4, "谐振腔", ["电流绕着腔流动", "形成振荡"], (650, 178)),
                                      (1, "阴极", ["加热后发射电子", "电子从这里出发"], (600, 360)),
                                      (2, "阳极", ["一块铜，和阴极", "之间加直流高压"], (655, 505))]),
                         (9, "right", [(3, "永磁体", ["轴向磁场让电子", "走弯曲的路"], (868, 138)),
                                       (6, "输出针", ["把微波引出", "送进波导"], (1325, 410)),
                                       (5, "输出线", ["从谐振腔引出", "振荡的能量"], (1345, 870))])]:
    for k, (n, name, desc, (px, py)) in enumerate(parts):
        blk = gb.box(col, 2 * k, 3, 2)
        c.tint(blk)
        c.num(blk.l + 0.9, blk.t - 0.85, str(n))
        c.title("w", blk.l + 1.7, blk.t - 0.85, name)
        c.text(Box(blk.l + 0.6, blk.t - 1.55, blk.w - 1.0, 1.7), desc, line=22)
        p = on_image(pic, img, px, py)
        c.line((blk.r if side == "left" else blk.l, blk.t - 0.85), p, TONES["primary"].base, 1.25)
        c.shape("标注", MSO_SHAPE.OVAL, place("c", *p, 0.75, 0.75), None, TONES["accent"].base, 2.0)
c.bottombar("类比", "像吹哨：电子流掠过谐振腔激发微波；谐振腔两壁像电容、圆孔像电感，频率由尺寸决定")

# 大序号步骤：五步各配一张实物照片（同样大小），36pt 序号用箭头串起来，不画方框流程；
# 关键一步（磁控管）整列加强调色底。下面一条讲频段、一块放效率
c = deck.content(SECTIONS[0], "电能约一半变成 2.45 GHz 微波，这个频段全球通用",
                 notes="顺着序号讲能量怎么走：插座的交流电进来，高压变压器升压，再经电容和二极管倍压成直流高压；"
                       "磁控管这一步把直流电能变成微波，用橙色标出；波导是一段金属管，把微波送进炉腔。"
                       "右下看效率：普通型号只有一半左右，节能型号超过 64%。左下回答为什么用 2.45 GHz："
                       "微波炉必须用 ISM 频段，避免干扰其他无线电业务；915 MHz 只在部分国家能用，5.8 GHz 产生大功率成本太高，"
                       "2.45 GHz 全球都能用。")
g5 = Grid(cols=5)
cols = [g5.box(i, 0, 1, 4) for i in range(5)]
steps = [("市电", "sockets", "墙上的电源插座", ["交流电输入"]),
         ("高压直流电源", "transformer", "微波炉里的高压变压器", ["变压器升压", "电容和二极管倍压"]),
         ("磁控管", "magnetron_unit", "家用磁控管", ["把直流电能变成", "2.45 GHz 微波"]),
         ("波导", "waveguides", "各种矩形波导部件", ["一段金属管", "把微波送进炉腔"]),
         ("炉腔和食物", "oven_interior", "家用微波炉炉腔", ["微波在炉腔里反射", "被食物吸收"])]
accent = TONES["accent"].ink
c.tint(cols[2], "accent")
for i, (col, (name, img, cap, desc)) in enumerate(zip(cols, steps)):
    key = i == 2
    c.at("n", col.cx, col.t - 0.15, f"0{i + 1}", size=S_DISPLAY, color=accent if key else TONES["primary"].ink, bold=True)
    c.web_figure(FIG / f"{img}.jpg", Box(col.l + 0.3, col.t - 1.9, col.w - 0.6, 3.5), cap, fill=True)
    c.at("n", col.cx, col.t - 5.6, name, size=S_TITLE, color=INK, bold=True)
    c.text(Box(col.l + 0.2, col.t - 6.4, col.w - 0.4, 1.5), desc, align="c", line=22, color=accent if key else TEXT)
for a, b in zip(cols, cols[1:]):
    c.line((a.cx + 1.2, a.t - 1.0), (b.cx - 1.2, a.t - 1.0), TONES["primary"].base, 1.5, arrow=True)
ism = g.box(0, 4, 8, 2)
c.tint(ism)
c.at("nw", ism.l + 0.6, ism.t - 0.35, "为什么是 2.45 GHz：微波炉只能用 ISM 频段", size=S_TITLE, color=TONES["primary"].ink, bold=True)
w3 = (ism.w - 1.2) / 3
for k, (f, lam, note, hl) in enumerate([("915 MHz", "波长 32.8 cm", "工业炉，部分国家可用", False),
                                        ("2.45 GHz", "波长 12.2 cm", "家用炉，全球可用", True),
                                        ("5.8 GHz", "波长约 5.2 cm", "成本太高，不用于烹饪", False)]):
    x = ism.l + 0.6 + k * w3
    c.text(Box(x, ism.t - 1.3, w3 - 0.3, 2.1), [P(R(f, size=S_TITLE, bold=True, color=accent if hl else INK)),
                                                P(lam, before=2), note], size=S_SMALL, color=accent if hl else TEXT, line=20)
eff = g.box(8, 4, 4, 2)
c.tint(eff, "accent")
c.keynum(eff.cx, eff.cy - 0.45, "50–64%", "电能变成微波的比例（普通到节能型号）", big=True)

# ======================================================== 二、微波怎么加热食物
deck.divider(SECTIONS[1])

# 公式拆解：公式铺满页宽，每一项向下引出一块解释，关键项（损耗因子）用强调色；
# 分子怎么转、为什么生热写进对应项的解释里，定性和定量在同一页
c = deck.content(SECTIONS[1], "水分子跟着电场转向发热，功率与频率和场强平方成正比",
                 notes="把公式拆开看。左边 p 是单位体积的吸热功率，就是食物每一处升温有多快。"
                       "ω 是角频率：电场每秒振荡 24.5 亿次，水分子氧一端带负电、氢一端带正电，一直想对齐电场，跟着来回转。"
                       "ε0 是常数。损耗因子是介电常数的虚部，分子转动时推挤、碰撞周围分子，把场能变成热；"
                       "含水多的部分损耗因子大，所以先热。E 的平方说明场强大的地方吸热快，这就是第三章热点的来源。"
                       "常有人说微波炉靠水分子共振，其实所有极性分子在很宽的频率范围都能这样吸热。")
g5b = Grid(cols=5, bottombar=True)
blocks = [g5b.box(i, 2, 1, 4) for i in range(5)]
FY = 12.5  # 公式的 y
terms = [P(*var("p")), P(*var("ω")), P(*var("ε", "0")), P(*eps_r2()), P(*var("E"), R("2", latin=LM, sup=True))]
for i, (blk, term) in enumerate(zip(blocks, terms)):
    t = TONES["accent" if i == 3 else "primary"]
    c.at("c", blk.cx, FY, term, size=S_DISPLAY, color=t.ink if i else INK, kind="公式")
    c.line((blk.cx, FY - 1.1), (blk.cx, blk.t + 0.1), t.base, 1.5, arrow=True)
for a, b, s in zip(blocks, blocks[1:], ["=", "·", "·", "·"]):
    c.at("c", (a.r + b.l) / 2, FY, s, size=S_DISPLAY, color=INK, kind="公式")
info = [("吸热功率密度", ["单位 W/m³", "食物每一处", "升温有多快"], "要求的量"),
        ("角频率", [P(*var("ω"), " = 2π", *var("f")), "电场每秒振荡", "24.5 亿次"], "水分子跟着来回转"),
        ("真空介电常数", ["物理常数", "对所有食物", "都一样"], "不随食物变"),
        ("损耗因子", ["介电常数的虚部", "分子转动时互相", "推挤碰撞生热"], "含水多的部分大，先热"),
        ("场强的平方", [P("电场强度 ", *var("E")), "场强翻一倍", "吸热变成 4 倍"], "场强大处成热点")]
for i, (blk, (name, lines, foot)) in enumerate(zip(blocks, info)):
    if i == 0:  # 等号左边是结果，用实色块
        c.shape("底色", MSO_SHAPE.ROUNDED_RECTANGLE, blk, TONES["primary"].base, radius=0.247)
        fg, fg2, rule = WHITE, WHITE, WHITE
    else:
        t = TONES["accent" if i == 3 else "primary"]
        c.tint(blk, "accent" if i == 3 else "primary")
        fg, fg2, rule = INK, t.ink, LINE
    c.at("n", blk.cx, blk.t - 0.5, name, size=S_TITLE, color=fg, bold=True)
    c.text(Box(blk.l + 0.4, blk.t - 1.6, blk.w - 0.8, 2.7), lines, color=WHITE if i == 0 else TEXT, align="c", line=24)
    c.line((blk.l + 0.8, blk.b + 2.0), (blk.r - 0.8, blk.b + 2.0), rule, 0.75)
    c.text(Box(blk.l + 0.4, blk.b + 1.8, blk.w - 0.8, 1.5), foot, size=S_SMALL, color=fg2, bold=True, align="c", anchor="m")
c.bottombar("误区", "不是水分子共振：所有极性分子在很宽的频率范围都能这样吸热，对液态水最有效")

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
