"""把代码、JSON 或终端输出渲染成 PNG，用来在 PPT 里展示真实材料（数据片段、报错、测试输出）。

用法：
    python snippet.py <文本文件> <输出.png> --kind json|python|terminal [--title 文件名] [--pt 14]

依赖：pillow、pygments、fonttools（uv run --with pillow --with pygments --with fonttools）。
图片按 300dpi 生成：放进 PPT 时用 Canvas.snippet(png, ...) 按原尺寸放置，字号就是 --pt 指定的大小。
排版参数（标题区高、内边距、行高）写进 PNG 的 snippet 文本块，Canvas.snippet 据此算出每行的位置，
用来把箭头对准某一行。
西文用 Consolas，中文用微软雅黑，两者都没有的符号（如 ∇）用 Cambria；按字形表逐字选择，统一按基线绘制。
terminal：深色底；以 # 开头的行是灰色注释，含「失败 / 错误 / Error」的行标红，其余浅色。
"""
import argparse
import json
from pathlib import Path

from fontTools.ttLib import TTCollection, TTFont
from PIL import Image, ImageDraw, ImageFont
from PIL.PngImagePlugin import PngInfo
from pygments.lexers import get_lexer_by_name
from pygments.token import Token

DPI = 300
FONTS = ["C:/Windows/Fonts/consola.ttf", "C:/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/cambria.ttc"]  # 按优先级

LIGHT = {"bg": "#F6F8FA", "border": "#D7E0EC", "text": "#24292F", "title": "#64748B"}
DARK = {"bg": "#1E293B", "border": "#1E293B", "text": "#E2E8F0", "title": "#94A3B8",
        "comment": "#94A3B8", "error": "#FCA5A5"}
TOKEN_COLORS = [  # 按顺序匹配，取第一个
    (Token.Name.Tag, "#0550AE"),          # JSON 的键
    (Token.Keyword.Constant, "#CF222E"),  # null / true / false
    (Token.Keyword, "#CF222E"),
    (Token.Literal.String, "#0A3069"),
    (Token.Literal.Number, "#0550AE"),
    (Token.Comment, "#6E7781"),
    (Token.Name.Function, "#8250DF"),
]


def color_of(ttype):
    for t, c in TOKEN_COLORS:
        if ttype in t:
            return c
    return LIGHT["text"]


def spans(text, kind):
    """把文本切成 [(行内片段, 颜色)] 的行列表。"""
    lines = []
    if kind == "terminal":
        for line in text.splitlines():
            if line.lstrip().startswith("#"):
                c = DARK["comment"]
            elif any(w in line for w in ("失败", "错误", "Error", "error")):
                c = DARK["error"]
            else:
                c = DARK["text"]
            lines.append([(line, c)])
        return lines
    cur = []
    for ttype, value in get_lexer_by_name(kind).get_tokens(text):
        parts = value.split("\n")
        for k, part in enumerate(parts):
            if k:
                lines.append(cur)
                cur = []
            if part:
                cur.append((part, color_of(ttype)))
    if cur:
        lines.append(cur)
    while lines and not lines[-1]:
        lines.pop()
    return lines


def render(text, out, kind="json", title=None, pt=14):
    px = pt * DPI / 72
    fonts = [ImageFont.truetype(f, round(px * (1 if k == 0 else 0.95))) for k, f in enumerate(FONTS)]
    cmaps = [_cmap(f) for f in FONTS]
    theme = DARK if kind == "terminal" else LIGHT
    lines = spans(text.expandtabs(4), kind)

    def font_for(ch):
        for f, cmap in zip(fonts, cmaps):
            if ord(ch) in cmap:
                return f
        raise ValueError(f"没有字体包含字符 {ch!r}")

    def width(s):
        return sum(font_for(ch).getlength(ch) for ch in s)

    pad, lh = round(px * 0.9), round(px * 1.5)
    head = round(px * 1.9) if title else 0
    w = round(max([width("".join(t for t, _ in ln)) for ln in lines] + [width(title or "")]) + 2 * pad)
    h = head + len(lines) * lh + 2 * pad - round(px * 0.3)
    img = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((0, 0, w - 1, h - 1), radius=round(px * 0.6), fill=theme["bg"], outline=theme["border"],
                        width=3)
    if title:
        x = pad
        for ch in title:
            d.text((x, head * 0.62), ch, font=font_for(ch), fill=theme["title"], anchor="ls")
            x += font_for(ch).getlength(ch)
        d.line((3, head, w - 4, head), fill=theme["border"] if kind != "terminal" else "#334155", width=2)
    y = head + pad
    for ln in lines:
        x = pad
        for seg, col in ln:
            for ch in seg:
                f = font_for(ch)
                d.text((x, y + round(px * 1.05)), ch, font=f, fill=col, anchor="ls")
                x += f.getlength(ch)
        y += lh
    info = PngInfo()
    # 行 k 的字形中心在图片顶端下方 head + pad + k*lh + 0.7*px 像素处
    info.add_text("snippet", json.dumps({"dpi": DPI, "head": head, "pad": pad, "lh": lh, "center": round(px * 0.7)}))
    img.save(out, dpi=(DPI, DPI), pnginfo=info)
    return out


def _cmap(path):
    """字体覆盖的码位集合；.ttc 取第一个字体。"""
    font = TTCollection(path).fonts[0] if path.endswith(".ttc") else TTFont(path)
    return set(font.getBestCmap())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--kind", default="json", choices=["json", "python", "terminal"])
    ap.add_argument("--title")
    ap.add_argument("--pt", type=float, default=14)
    a = ap.parse_args()
    render(Path(a.src).read_text(encoding="utf-8"), a.out, a.kind, a.title, a.pt)
    print("written", a.out)


if __name__ == "__main__":
    main()
