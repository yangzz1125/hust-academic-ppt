"""大纲（outline.json）的检查、便利贴墙和策划稿。

一份汇报的大纲写成一个 JSON 文件，三个子命令都读它：

    python outline.py check outline.json              # 校验格式，列出只读标题序列和提示
    python outline.py wall  outline.json wall.html    # 便利贴墙：一页一张便利贴，带构图缩略图，给用户审大纲
    python outline.py draft outline.json draft.pptx   # 策划稿：真模板 + 每块的位置和内容，图表位置用虚线占位框

运行：uv run --no-project --with python-pptx --with lxml python -X utf8 outline.py <子命令> ...

格式（字段说明见 references/outline.md）：

    {
      "meta": {"title": "...", "presenter": "...", "date": "...", "template": "hust",
               "occasion": "...", "audience": "...", "minutes": 10, "thesis": "..."},
      "parts": [
        {"title": "章节名", "toc": true, "pages": [
          {"title": "结论句标题", "composition": "图为主：左照片，右要点",
           "cols": 12, "bottombar": ["类比", "一句话"],
           "blocks": [{"grid": [0, 0, 8, 6], "kind": "图片", "content": "磁控管剖面照片，标 1–5 编号"},
                      {"grid": [8, 0, 4, 6], "kind": "卡片", "header": "工作过程", "content": ["...", "..."]}],
           "notes": "口播稿", "sources": ["source/xxx.txt"]}]}
      ]
    }
"""
import argparse
import html
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from hust_deck import *  # noqa: E402
from hust_deck import cn_num, pt2cm, tw  # noqa: E402
from pptx.enum.dml import MSO_LINE  # noqa: E402

TEXT_KINDS = {"卡片", "文字"}
FIGURE_KINDS = {"图片", "图表", "表格", "节点图", "公式", "代码"}
ROWS = 6
TITLE_EM = 32  # 华科模板标题框宽 28.9 cm，24pt 一行约放 32 个汉字；其他模板以 draft 渲染和 qa_check 为准


def fail(where, msg):
    raise ValueError(f"{where}：{msg}")


def load(path):
    """读入并校验大纲。大纲由人和模型手写，这里是信任边界：缺字段、类型不对、块越界或重叠都直接报错。"""
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    meta = data.get("meta")
    if not isinstance(meta, dict):
        fail("meta", "缺少 meta 对象")
    for k in ("title", "presenter", "date"):
        if not isinstance(meta.get(k), str) or not meta[k]:
            fail("meta", f"缺少字符串字段 {k}")
    parts = data.get("parts")
    if not isinstance(parts, list) or not parts:
        fail("parts", "缺少章节列表")
    for i, part in enumerate(parts):
        where = f"第 {i + 1} 章"
        if not isinstance(part.get("title"), str) or not part["title"]:
            fail(where, "缺少 title")
        pages = part.get("pages")
        if not isinstance(pages, list) or not pages:
            fail(where, "缺少 pages")
        for j, page in enumerate(pages):
            check_page(page, f"{where}「{part['title']}」第 {j + 1} 页")
    return data


def check_page(page, where):
    if not isinstance(page.get("title"), str) or not page["title"]:
        fail(where, "缺少 title（结论句标题）")
    blocks = page.get("blocks")
    if not isinstance(blocks, list) or not blocks:
        fail(where, "缺少 blocks")
    bar = page.get("bottombar")
    if bar is not None and not (isinstance(bar, list) and len(bar) == 2 and all(isinstance(s, str) for s in bar)):
        fail(where, "bottombar 应为 [标签, 一句话]")
    cells = {}
    for k, b in enumerate(blocks):
        bw = f"{where} 第 {k + 1} 块"
        if b.get("kind") not in TEXT_KINDS | FIGURE_KINDS:
            fail(bw, f"kind 应为 {'、'.join(sorted(TEXT_KINDS | FIGURE_KINDS))} 之一，实际是 {b.get('kind')!r}")
        g = b.get("grid")
        if not (isinstance(g, list) and len(g) == 4 and all(isinstance(v, int) for v in g)):
            fail(bw, "grid 应为 [列, 行, 占几列, 占几行] 四个整数")
        cols = block_cols(page, b)
        col, row, w, h = g
        if col < 0 or row < 0 or w < 1 or h < 1 or col + w > cols or row + h > ROWS:
            fail(bw, f"grid {g} 超出 {cols}×{ROWS} 网格")
        content = b.get("content")
        if not (isinstance(content, str) or (isinstance(content, list) and all(isinstance(s, str) for s in content))):
            fail(bw, "content 应为字符串或字符串列表")
        # 不同列数的网格按页宽比例换算成 60 等分再查重叠
        x0, x1 = col * 60 // cols, (col + w) * 60 // cols
        for x in range(x0, x1):
            for y in range(row, row + h):
                if (x, y) in cells:
                    fail(bw, f"和第 {cells[(x, y)] + 1} 块重叠")
                cells[(x, y)] = k


def block_cols(page, block):
    return block.get("cols", page.get("cols", 12))


def em(s):
    """按汉字宽度折算的字数（西文按 0.55 个）。"""
    return tw(s, 1) / pt2cm(1)


def signature(page):
    """构图签名：各块按页宽比例换算后的位置，连续两页相同说明排法没变。"""
    return tuple(sorted((b["grid"][0] / block_cols(page, b), b["grid"][1],
                         b["grid"][2] / block_cols(page, b), b["grid"][3]) for b in page["blocks"]))


def walk(data):
    """按幻灯片顺序给出 (页码, 章节, 页面)；过渡页只占页码。页码和 draft 生成的 pptx 一致。"""
    n = 2  # 封面、目录
    for part in data["parts"]:
        if part.get("toc", True):
            n += 1
        for page in part["pages"]:
            n += 1
            yield n, part, page


def hints(data):
    """每页的提示：标题过长、构图和上一页相同、没写口播稿。返回 {页码: [提示]}。"""
    out, prev = {}, None
    for n, _, page in walk(data):
        h = []
        if em(page["title"]) > TITLE_EM:
            h.append(f"标题约 {em(page['title']):.0f} 字，一行放不下（约 {TITLE_EM} 字）")
        sig = signature(page)
        if sig == prev:
            h.append("构图和上一页相同")
        prev = sig
        if not page.get("notes"):
            h.append("没写口播稿")
        out[n] = h
    return out


# ---------------------------------------------------------------- check

def cmd_check(args):
    data = load(args.outline)
    hs = hints(data)
    pages = list(walk(data))
    print(f"内容页 {len(pages)} 页", end="")
    if "minutes" in data["meta"]:
        print(f"，时长 {data['meta']['minutes']} 分钟（约每分钟一页）", end="")
    print("\n\n只读标题：")
    for n, part, page in pages:
        print(f"  {n:>2}  [{part['title']}] {page['title']}")
    flagged = [(n, h) for n, h in hs.items() if h]
    if flagged:
        print("\n提示：")
        for n, h in flagged:
            print(f"  ⚠️ 第 {n} 页：{'；'.join(h)}")


# ---------------------------------------------------------------- wall

WALL_CSS = """
:root { --bg:#f4f1ea; --note:#fff8d6; --note-line:#e9dc9c; --ink:#1e293b; --text:#334155; --gray:#64748b;
  --primary:#0259a0; --accent:#e8710a; --cell:#ffffff; --cell-line:#cbd5e1; --fig:#cfe0f5; --warn-bg:#fdecdc; }
@media (prefers-color-scheme: dark) { :root:not([data-theme="light"]) { --bg:#1b1d22; --note:#2b2a22; --note-line:#4a4630;
  --ink:#f1f5f9; --text:#cbd5e1; --gray:#94a3b8; --primary:#6aa8e8; --accent:#f5a25d; --cell:#2f3440; --cell-line:#4b5263; --fig:#233650; --warn-bg:#3d2a1c; } }
:root[data-theme="dark"] { --bg:#1b1d22; --note:#2b2a22; --note-line:#4a4630; --ink:#f1f5f9; --text:#cbd5e1; --gray:#94a3b8;
  --primary:#6aa8e8; --accent:#f5a25d; --cell:#2f3440; --cell-line:#4b5263; --fig:#233650; --warn-bg:#3d2a1c; }
* { box-sizing:border-box; }
body { margin:0; background:var(--bg); color:var(--text); font:15px/1.55 "Microsoft YaHei","PingFang SC",sans-serif; }
main { max-width:1280px; margin:0 auto; padding:24px 16px 48px; }
h1 { color:var(--ink); font-size:24px; margin:0 0 6px; }
.meta { color:var(--gray); margin:0 0 4px; }
.thesis { color:var(--ink); margin:8px 0 20px; }
.thesis b { color:var(--accent); }
h2 { color:var(--primary); font-size:18px; margin:28px 0 10px; }
.wall { display:grid; grid-template-columns:repeat(auto-fill,minmax(280px,1fr)); gap:16px; }
.note { background:var(--note); border:1px solid var(--note-line); border-radius:4px; padding:12px 14px 14px;
  box-shadow:0 2px 6px rgba(0,0,0,.08); display:flex; flex-direction:column; gap:8px; }
.no { color:var(--gray); font-size:13px; }
.title { color:var(--ink); font-weight:700; font-size:16px; }
svg { width:100%; height:auto; display:block; }
svg rect.cell { fill:var(--cell); stroke:var(--cell-line); stroke-width:.4; }
svg rect.fig { fill:var(--fig); stroke:var(--primary); stroke-dasharray:1.5 1; stroke-width:.4; }
svg rect.bar { fill:var(--accent); opacity:.5; }
svg text { fill:var(--ink); font-size:5px; }
.comp { color:var(--gray); font-size:13px; }
ul { margin:0; padding-left:18px; font-size:13px; }
li b { color:var(--ink); }
.notes { font-size:13px; color:var(--gray); border-top:1px dashed var(--note-line); padding-top:6px; }
.warn { font-size:13px; color:var(--accent); background:var(--warn-bg); border-radius:3px; padding:2px 6px; }
"""


def mini_layout(page):
    """构图缩略图：按 31×14.3 的正文区比例画出各块的位置。"""
    W, H, gap = 124.0, 57.2, 1.6
    rows = []
    bottom = 6.0 if page.get("bottombar") else 0
    rh = (H - bottom - gap * (ROWS - 1)) / ROWS
    for b in page["blocks"]:
        cols = block_cols(page, b)
        cw = (W - gap * (cols - 1)) / cols
        col, row, w, h = b["grid"]
        x, y = col * (cw + gap), row * (rh + gap)
        bw, bh = w * cw + (w - 1) * gap, h * rh + (h - 1) * gap
        cls = "fig" if b["kind"] in FIGURE_KINDS else "cell"
        label = html.escape(b.get("header") or b["kind"])
        rows.append(f'<rect class="{cls}" x="{x:.1f}" y="{y:.1f}" width="{bw:.1f}" height="{bh:.1f}" rx="1.5"/>'
                    f'<text x="{x + 2:.1f}" y="{y + 6.5:.1f}">{label}</text>')
    if bottom:
        rows.append(f'<rect class="bar" x="0" y="{H - bottom + 1.5:.1f}" width="{W}" height="{bottom - 1.5:.1f}" rx="1.5"/>')
    return f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="构图">{"".join(rows)}</svg>'


def block_items(b):
    head = html.escape(b.get("header") or b["kind"])
    body = b["content"] if isinstance(b["content"], str) else "；".join(b["content"])
    return f"<li><b>{head}</b>：{html.escape(body)}</li>"


def cmd_wall(args):
    data = load(args.outline)
    meta, hs = data["meta"], hints(data)
    e = html.escape
    info = " · ".join(e(str(meta[k])) for k in ("occasion", "audience") if meta.get(k))
    if meta.get("minutes"):
        info += f" · {meta['minutes']} 分钟"
    out = [f'<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">'
           f'<meta name="viewport" content="width=device-width,initial-scale=1">'
           f'<title>{e(meta["title"])[:20]} 大纲</title><style>{WALL_CSS}</style></head><body><main>',
           f'<h1>{e(meta["title"])}</h1><p class="meta">{e(meta["presenter"])} · {e(meta["date"])}'
           f'{" · " + info if info else ""}</p>']
    if meta.get("thesis"):
        out.append(f'<p class="thesis"><b>要听众接受的结论：</b>{e(meta["thesis"])}</p>')
    current = None
    for n, part, page in walk(data):
        if part is not current:
            if current is not None:
                out.append("</div>")
            k = [p for p in data["parts"] if p.get("toc", True)]
            num = f"{cn_num(k.index(part) + 1)}、" if part in k else ""
            out.append(f'<h2>{num}{e(part["title"])}</h2><div class="wall">')
            current = part
        out.append(f'<article class="note"><div class="no">第 {n} 页</div>'
                   f'<div class="title">{e(page["title"])}</div>{mini_layout(page)}')
        if page.get("composition"):
            out.append(f'<div class="comp">{e(page["composition"])}</div>')
        out.append("<ul>" + "".join(block_items(b) for b in page["blocks"]) + "</ul>")
        if page.get("notes"):
            out.append(f'<div class="notes">口播：{e(page["notes"])}</div>')
        out.extend(f'<div class="warn">{e(h)}</div>' for h in hs[n])
        out.append("</article>")
    out.append("</div></main></body></html>")
    Path(args.out).write_text("".join(out), encoding="utf-8")
    print("written", args.out)


# ---------------------------------------------------------------- draft

def draw_block(c, page, b):
    g = Grid(cols=block_cols(page, b), bottombar=bool(page.get("bottombar")))
    box = g.box(*b["grid"])
    items = [b["content"]] if isinstance(b["content"], str) else b["content"]
    if b["kind"] in TEXT_KINDS:
        if b["kind"] == "卡片" or b.get("header"):
            body = c.card(box, header=b.get("header"), solid=False)
            box = Box(body.l + 0.5, body.t - 0.3, body.w - 1.0, body.h - 0.6)
        c.bullets(box, items)
        return
    if b.get("header"):
        body = c.card(box, header=b["header"], solid=False)
        box = Box(body.l + 0.3, body.t - 0.3, body.w - 0.6, body.h - 0.6)
    shp = c.shape("占位", MSO_SHAPE.RECTANGLE, box, TONES["neutral"].tint, GRAY, 0.8)
    shp.line.dash_style = MSO_LINE.DASH
    fill_text(shp.text_frame, [P(R(f"【{b['kind']}】", color=GRAY, bold=True)), *items],
              size=S_SMALL, color=GRAY, align="c", anchor="m")


def cmd_draft(args):
    data = load(args.outline)
    meta = data["meta"]
    deck = Deck(meta.get("template", "hust"))
    deck.cover(meta["title"], meta["presenter"], meta["date"])
    deck.toc([p["title"] for p in data["parts"] if p.get("toc", True)])
    for part in data["parts"]:
        if part.get("toc", True):
            deck.divider(part["title"])
        for page in part["pages"]:
            c = deck.content(part["title"], page["title"], notes=page.get("notes", ""))
            for b in page["blocks"]:
                draw_block(c, page, b)
            if page.get("bottombar"):
                c.bottombar(*page["bottombar"])
    deck.save(args.out)
    print("written", args.out)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name, fn, extra in (("check", cmd_check, False), ("wall", cmd_wall, True), ("draft", cmd_draft, True)):
        p = sub.add_parser(name)
        p.add_argument("outline")
        if extra:
            p.add_argument("out")
        p.set_defaults(fn=fn)
    args = ap.parse_args()
    args.fn(args)


if __name__ == "__main__":
    main()
