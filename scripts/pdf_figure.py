"""从论文 PDF 里取图：先找图在哪页，再看带网格的预览定范围，最后按 300dpi 裁出子图。

用法：
    python pdf_figure.py list <论文.pdf>
    python pdf_figure.py preview <论文.pdf> <页码> <预览.png> [--dpi 100]
    python pdf_figure.py crop <论文.pdf> <页码> <输出.png> --box x0,y0,x1,y1 [--dpi 300]

页码从 1 开始。坐标是页面比例（0–1），原点在左上角：x 向右、y 向下，和预览图上的网格读数一致。
list：   每页的图注位置（Fig. / Figure / 图 n 开头的文字块）和嵌入位图的位置，用于定位图在哪页、大致在哪。
preview：整页渲染，叠加 10% 一格的网格和刻度（5% 处有细线），Read 看图后读出裁切范围。
crop：   按 --box 裁切并以 --dpi 渲染（矢量图在 300dpi 下放进 PPT 也清楚）；四周自动留 1% 余量。
依赖：pymupdf、pillow（uv run --with pymupdf --with pillow）。
"""
import argparse
import re

import pymupdf
from PIL import Image, ImageDraw, ImageFont

CAPTION = re.compile(r"^\s*(fig\.?|figure|图|表|table)\s*\d+", re.I)
MARGIN = 0.01


def frac(rect, page):
    w, h = page.rect.width, page.rect.height
    return f"{rect.x0 / w:.2f},{rect.y0 / h:.2f},{rect.x1 / w:.2f},{rect.y1 / h:.2f}"


def cmd_list(pdf):
    doc = pymupdf.open(pdf)
    for k, page in enumerate(doc, 1):
        lines = []
        for x0, y0, x1, y1, text, *_ in page.get_text("blocks"):
            if CAPTION.match(text):
                lines.append(f"  图注 {frac(pymupdf.Rect(x0, y0, x1, y1), page)}  {' '.join(text.split())[:70]}")
        for img in page.get_images(full=True):
            for r in page.get_image_rects(img[0]):
                if r.width * r.height > 0.02 * page.rect.width * page.rect.height:  # 忽略图标、公式小图
                    lines.append(f"  位图 {frac(r, page)}  {img[2]}×{img[3]}px")
        if lines:
            print(f"第 {k} 页")
            print("\n".join(lines))


def cmd_preview(pdf, page_no, out, dpi):
    page = pymupdf.open(pdf)[page_no - 1]
    pix = page.get_pixmap(dpi=dpi)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    d = ImageDraw.Draw(img)
    font = ImageFont.truetype("C:/Windows/Fonts/consola.ttf", max(12, pix.width // 60))
    for i in range(1, 20):
        major = i % 2 == 0
        color, width = ((220, 38, 38), 2) if major else ((248, 180, 180), 1)
        x, y = pix.width * i / 20, pix.height * i / 20
        d.line((x, 0, x, pix.height), fill=color, width=width)
        d.line((0, y, pix.width, y), fill=color, width=width)
        if major:
            d.text((x + 3, 3), f"{i / 20:.1f}", font=font, fill=(220, 38, 38))
            d.text((3, y + 3), f"{i / 20:.1f}", font=font, fill=(220, 38, 38))
    img.save(out)
    print("written", out)


def cmd_crop(pdf, page_no, out, box, dpi):
    page = pymupdf.open(pdf)[page_no - 1]
    x0, y0, x1, y1 = (float(v) for v in box.split(","))
    if not (0 <= x0 < x1 <= 1 and 0 <= y0 < y1 <= 1):
        raise ValueError(f"--box {box} 不合法：要求 0 ≤ x0 < x1 ≤ 1、0 ≤ y0 < y1 ≤ 1")
    x0, y0 = max(0.0, x0 - MARGIN), max(0.0, y0 - MARGIN)
    x1, y1 = min(1.0, x1 + MARGIN), min(1.0, y1 + MARGIN)
    w, h = page.rect.width, page.rect.height
    clip = pymupdf.Rect(x0 * w, y0 * h, x1 * w, y1 * h)
    pix = page.get_pixmap(dpi=dpi, clip=clip)
    pix.save(out)
    print(f"written {out}  {pix.width}×{pix.height}px，"
          f"按 150dpi 最大可放 {pix.width / 150 * 2.54:.1f}×{pix.height / 150 * 2.54:.1f}cm")


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("list")
    a.add_argument("pdf")
    p = sub.add_parser("preview")
    p.add_argument("pdf")
    p.add_argument("page", type=int)
    p.add_argument("out")
    p.add_argument("--dpi", type=int, default=100)
    c = sub.add_parser("crop")
    c.add_argument("pdf")
    c.add_argument("page", type=int)
    c.add_argument("out")
    c.add_argument("--box", required=True)
    c.add_argument("--dpi", type=int, default=300)
    args = ap.parse_args()
    if args.cmd == "list":
        cmd_list(args.pdf)
    elif args.cmd == "preview":
        cmd_preview(args.pdf, args.page, args.out, args.dpi)
    else:
        cmd_crop(args.pdf, args.page, args.out, args.box, args.dpi)


if __name__ == "__main__":
    main()
