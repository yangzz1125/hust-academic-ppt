"""排法示例：四种气体放电的照片，演示 references/layouts.md 第 6 节不套卡片、以图为主的几种做法。

A 并列多列、B 每行一项（输入图 → 处理 → 结果图）、C 图片表格、E 大图 + 引线标注、F 一大几小，各一页。
照片取自 Wikimedia Commons（figures/，来源和许可在同名 .json 里）；「局部放大」由本脚本按 CROP 从照片裁出，
生成时写到临时目录再嵌进 PPT。页面上只写照片里看得到的东西。

取图命令：
    fetch_image.py get "File:Corona discharge in atmospheric air 1.jpg" figures --name corona
    fetch_image.py get "File:Large Discharge Spheres on Quadruple Bonetti Machine.jpg" figures --name spark
    fetch_image.py get "File:Glow discharge regions.jpg" figures --name glow
    fetch_image.py get "File:Electric arc.jpg" figures --name arc

生成和质检：
    uv run --no-project --with python-pptx --with lxml python -X utf8 build_deck.py <输出.pptx>
    uv run --no-project --with python-pptx python -X utf8 ../../scripts/qa_check.py <输出.pptx>
"""
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "scripts"))
from hust_deck import *  # noqa: E402
from PIL import Image  # noqa: E402

FIG = Path(__file__).parent / "figures"
# 放电区域在照片里的位置：左、上、右、下，按照片宽高的比例
CROP = {"corona": (0.36, 0.18, 0.60, 0.62), "spark": (0.30, 0.28, 0.80, 0.66),
        "glow": (0.50, 0.18, 0.98, 0.56), "arc": (0.36, 0.34, 0.58, 0.64)}


def zoom(name, out):
    """从照片裁出放电区域，返回裁出图的路径。"""
    im = Image.open(FIG / f"{name}.jpg")
    w, h = im.size
    l, t, r, b = CROP[name]
    path = out / f"{name}_zoom.png"
    im.crop((round(l * w), round(t * h), round(r * w), round(b * h))).save(path)
    return path


def head_h(its):
    """标题块高度：标签 0.75cm + 下面每行 0.6cm，取几项中最高的，各项的图才能上沿对齐。"""
    return 0.75 + 0.6 * max(len(it["lines"]) for it in its)


def head(c, box, it, align="c"):
    c.text(box, [P(R(it["tag"], bold=True, color=PRI, size=S_BODY))]
           + [P(R(s, size=S_SMALL, color=GRAY), before=4 if j == 0 else 0) for j, s in enumerate(it["lines"])],
           align=align)


deck = Deck("hust")
PRI = TONES["primary"].base
deck.cover("气体放电照片的几种排法", "排法示例", "2026年10月16日")
B = Grid().box(0, 0, 12, 6)
SEC = "排法示例"
items = {
    "corona": dict(tag="电晕放电", lines=["尖端电极，大气压"], seen="向下散成扇形",
                   res=["只有电极尖端发亮", "光从尖端向下散开", "其余空间不发光"]),
    "spark": dict(tag="火花放电", lines=["两只金属球之间"], seen="一条弯折的亮通道",
                  res=["两球之间被击穿", "形成一条明亮通道", "通道弯折、不走直线"]),
    "glow": dict(tag="辉光放电", lines=["克鲁克斯管，低气压"], seen="右段排成一串亮斑",
                 res=["管内大段同时发光", "左段偏紫、右段偏红", "右段分成一串亮斑"]),
    "arc": dict(tag="电弧放电", lines=["两根导线之间"], seen="两端更亮",
                res=["导线端部之间发光", "亮区小，两端更亮", "导线端部烧得发红"]),
}
order = ["corona", "spark", "glow", "arc"]
its = [items[n] for n in order]

with tempfile.TemporaryDirectory() as tmp:
    D = {n: (FIG / f"{n}.jpg", zoom(n, Path(tmp))) for n in CROP}

    # ---------------------------------------------------------------- A 并列多列
    c = deck.content(SEC, "四种气体放电形态不同，从照片上就能区分",
                     notes="方案 A 并列多列。四种放电各占一列：上面是整张照片，下面是放电区域的局部放大。"
                           "电晕只在尖端发亮，火花是一条弯折的通道，辉光是整段管子发光，电弧的亮区很小、集中在两段导线之间。")
    widths, gap = [6.0, 7.6, 8.8, 5.6], 0.6
    hh, x = head_h(its), B.l
    for k, (w, n) in enumerate(zip(widths, order)):
        head(c, Box(x, B.t, w, hh), items[n])
        c.gallery(Box(x, B.t - hh - 0.3, w, 8.0 - hh), [[D[n][0]], [D[n][1]]], valign="n")
        c.text(Box(x, B.t - 8.6, w, 2.8), items[n]["res"], size=S_BODY, align="c", line=26)
        if k < len(order) - 1:
            c.line((x + w + gap / 2, B.t), (x + w + gap / 2, B.t - 11.4), LINE, 0.75)
        x += w + gap

    # ---------------------------------------------------------------- B 每行一项：输入图 → 处理 → 结果图
    c = deck.content(SEC, "裁出放电区域放大后，每种放电的发光结构都看得清",
                     notes="方案 B 每行一项。每行从左到右是整张照片、裁切和放大后的放电区域。火花是一条弯折的亮通道；"
                           "辉光管右段的光分成一串亮斑；电弧只在两段导线之间发光，两端更亮。"
                           "电晕的放大图又高又窄，放进一行太小，所以这一页不放它。")
    rows = ["spark", "glow", "arc"]
    H, cap_h = 3.1, 0.7
    pitch = (B.h - H - cap_h) / 2
    x_in, x_out = B.l + 12.4, B.l + 15.8
    for i, n in enumerate(rows):
        y = B.t - i * pitch
        c.text(Box(B.l, y, 4.6, H), [P(R(items[n]["tag"], size=S_BODY, bold=True, color=PRI)),
                                     P(items[n]["lines"][0], before=4)], size=S_SMALL, anchor="m")
        c.strip(Box(B.l + 4.9, y, x_in - B.l - 4.9, H + cap_h), [([D[n][0]], "整张照片")], align="e")
        c.strip(Box(x_out, y, B.r - x_out, H + cap_h), [([D[n][1]], items[n]["seen"])], align="w")
        mid = y - H / 2
        c.line((x_in + 0.3, mid), (x_out - 0.3, mid), PRI, 1.5, arrow=True)
        c.at("s", (x_in + x_out) / 2, mid + 0.15, ["裁出", "放电区域"], size=S_SMALL, color=INK)
        if i < len(rows) - 1:
            c.line((B.l, y - (pitch + H + cap_h) / 2), (B.r, y - (pitch + H + cap_h) / 2), LINE, 0.75)

    # ---------------------------------------------------------------- C 图片表格
    c = deck.content(SEC, "同样放大放电区域，四种放电的发光结构各不相同",
                     notes="方案 C 图片表格。列是四种放电，上一行是整张照片，下一行是放大后的放电区域，"
                           "横着看能直接比较四种放电的形状：扇形、弯折的通道、一串亮斑、两段导线间的一小段。")
    lab_w, hh = 1.5, head_h(its)
    widths = [5.2, 7.4, 8.4, B.w - lab_w - 0.3 - 3 * 0.6 - 21.0]
    rh = (B.h - hh - 0.3) / 2
    x = B.l + lab_w + 0.3
    for k, (w, n) in enumerate(zip(widths, order)):
        head(c, Box(x, B.t, w, hh), items[n])
        for r, cap in enumerate([None, items[n]["seen"]]):
            c.strip(Box(x, B.t - hh - 0.3 - r * rh, w, rh - 0.2), [([D[n][r]], cap)], valign="n")
        if k < len(order) - 1:
            c.line((x + w + 0.3, B.t), (x + w + 0.3, B.b), LINE, 0.75)
        x += w + 0.6
    for r, name in enumerate([["整张", "照片"], ["局部", "放大"]]):
        y = B.t - hh - 0.3 - r * rh
        c.text(Box(B.l, y, lab_w, rh), [P(R(t, bold=True, color=INK)) for t in name], anchor="m")
        c.line((B.l, y + 0.15), (B.r, y + 0.15), LINE, 0.75)

    # ---------------------------------------------------------------- E 大图 + 引线标注
    c = deck.content(SEC, "火花在两只金属球之间击穿空气，放电通道弯折明亮",
                     notes="方案 E 大图加引线标注。这一页只放一张照片，要讲的三处用引线指出来：左右两只金属球是放电电极，"
                           "中间弯折的亮线是放电通道，背后的圆盘是起电机产生高压的部分。")
    pic = c.figure(D["spark"][0], Box(B.l, B.t, 20.0, B.h), align="w")
    marks = [  # 标注点在照片里的位置（按宽、高的比例，自上而下）和说明
        ((0.57, 0.22), "起电机圆盘", ["转动起电，", "给两只球加上高压"]),
        ((0.86, 0.45), "放电球", ["左右各一只，", "是放电的两个电极"]),
        ((0.48, 0.58), "放电通道", ["两球之间被击穿，", "通道弯折、不走直线"]),
    ]
    lx, prev = pic.r + 1.6, None
    for (fx, fy), name, desc in marks:
        px, py = pic.l + fx * pic.w, pic.t - fy * pic.h
        ly = py + 0.35 if prev is None else min(py + 0.35, prev - 2.6)  # 说明和标注点同高，放不下再往下错开
        prev = ly
        c.line((lx - 0.2, ly - 0.35), (px, py), PRI, 1.25, arrow=True)
        c.text(Box(lx, ly, B.r - lx, 2.4), [P(R(name, bold=True, color=PRI, size=S_BODY))]
               + [P(R(t, size=S_SMALL), before=4 if j == 0 else 0) for j, t in enumerate(desc)])

    # ---------------------------------------------------------------- F 一大几小
    c = deck.content(SEC, "辉光放电沿管长分段，右段的光排成一串亮斑",
                     notes="方案 F 一大几小。这一页重点讲辉光放电，放大在左边：整张照片上管子左段偏紫、右段偏红，"
                           "放大右段能看到光分成一串亮斑。右边三种放电缩小，只作对照。")
    main = items["glow"]
    hm = head_h([main])
    head(c, Box(B.l, B.t, 14.0, hm), main, align="l")
    c.gallery(Box(B.l, B.t - hm - 0.3, 14.0, 9.2), [[D["glow"][0]], [D["glow"][1]]], valign="n")
    c.text(Box(B.l, B.b + 1.0, 14.0, 0.8), "上：整张照片；下：右段放大，光分成一串亮斑",
           size=S_SMALL, color=GRAY)
    x0 = B.l + 15.2
    c.line((x0 - 0.6, B.t), (x0 - 0.6, B.b), LINE, 0.75)
    w = (B.r - x0 - 2 * 0.6) / 3
    for k, n in enumerate(["corona", "spark", "arc"]):
        x = x0 + k * (w + 0.6)
        head(c, Box(x, B.t, w, hm), items[n])
        # 照片一行、放大图一行，三列逐行上沿对齐；第一行上沿与左边主图对齐
        c.strip(Box(x, B.t - hm - 0.3, w, 2.8), [([D[n][0]], None)], valign="n")
        c.strip(Box(x, B.t - hm - 3.8, w, 3.0), [([D[n][1]], None)], valign="n")
        c.text(Box(x, B.t - 8.5, w, 2.4), items[n]["res"], size=S_SMALL, align="c", line=20)

    deck.save(sys.argv[1])
print("written", sys.argv[1])
