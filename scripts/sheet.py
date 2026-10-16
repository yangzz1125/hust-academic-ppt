"""把几张渲染图拼成一张带标签的对照图，用来让用户比较同一页的几种排法。

用法：
    python sheet.py <输出.png> <图1.png> <图2.png> ... --labels "A 并列多列,B 每行一项,..." [--cols 2] [--width 800]

标签按逗号分隔，个数必须和图片数相同，写在每张图的上方。图片按 --cols 列从左到右、从上到下排，
每张缩放到 --width 像素宽。依赖：pillow（uv run --with pillow）。
"""
import argparse
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

FONT = "C:/Windows/Fonts/msyh.ttc"  # 微软雅黑，与 PPT 正文字体一致
GAP, LABEL_H, BORDER = 24, 44, "#D7E0EC"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("out")
    ap.add_argument("images", nargs="+")
    ap.add_argument("--labels", required=True, help="逗号分隔，个数与图片数相同")
    ap.add_argument("--cols", type=int, default=2)
    ap.add_argument("--width", type=int, default=800, help="每张图缩放后的宽度（像素）")
    args = ap.parse_args()

    labels = args.labels.split(",")
    if len(labels) != len(args.images):
        raise SystemExit(f"标签 {len(labels)} 个，图片 {len(args.images)} 张，个数要相同")
    tiles = []
    for path in args.images:
        im = Image.open(path).convert("RGB")
        tiles.append(im.resize((args.width, round(im.height * args.width / im.width)), Image.LANCZOS))
    cols = min(args.cols, len(tiles))
    rows = -(-len(tiles) // cols)
    tile_h = max(t.height for t in tiles)
    sheet = Image.new("RGB", (GAP + cols * (args.width + GAP), GAP + rows * (LABEL_H + tile_h + GAP)), "white")
    draw = ImageDraw.Draw(sheet)
    font = ImageFont.truetype(FONT, 26)
    for k, (tile, label) in enumerate(zip(tiles, labels)):
        x = GAP + (k % cols) * (args.width + GAP)
        y = GAP + (k // cols) * (LABEL_H + tile_h + GAP)
        draw.text((x, y + LABEL_H - 12), label.strip(), font=font, fill="#1E293B", anchor="ls")
        sheet.paste(tile, (x, y + LABEL_H))
        draw.rectangle((x - 1, y + LABEL_H - 1, x + tile.width, y + LABEL_H + tile.height), outline=BORDER)
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    sheet.save(args.out)
    print("written", args.out, f"{sheet.width}×{sheet.height}px")


if __name__ == "__main__":
    main()
