"""讲稿（script.md）、大纲（outline.json）、网页初稿和汇报网页（讲稿 + 便签画布 + 翻页预览）。

    python outline.py check outline.json                    # 校验讲稿、大纲和网页初稿，列出每页的标题、要点句、过渡句和字数
    python outline.py site  outline.json site.html [--design <设计稿 png 目录>]
                                                            # 生成汇报网页：讲稿、便签画布、翻页预览（搜索 / 初稿 / 设计稿）
    python outline.py chrome hust                           # 从模板生成网页初稿用的外观底图（换模板或改模板后跑一次）

运行：uv run --no-project --with python-pptx --with lxml --with pillow python -X utf8 outline.py <子命令> ...
（pillow 用于把初稿里 <img> 引用的图缩小后嵌进网页）
chrome 还要用 render.py（LibreOffice）渲染，结果存在 templates/<模板>/web/，生成后提交进仓库，平时不用再跑。

一份汇报的目录里：script.md 是讲稿（先写，用户确认后再切页），outline.json 是每页的标题、要点句和出处，
draft/<页 id>.html 是网页初稿。讲稿里用单独一行 [页 id] 标出下面几段属于哪一页，这几段就是这一页的口播稿，
第一句是过渡句。设计稿脚本用 page_notes() 取口播稿，不另写。格式和组件见 references/outline.md。
网页的外壳（样式、视图、脚本）在同目录的 site.html，数据替换其中 id="data" 的占位符。
"""
import argparse
import base64
import html
import io
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from hust_deck import COLLOQUIAL, Deck, TONES, cn_num, grid_bottom, pt2cm, tw, use_template  # noqa: E402
from pptx.enum.shapes import PP_PLACEHOLDER  # noqa: E402
from pptx.opc.constants import RELATIONSHIP_TYPE as RT  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402

SCRIPTS = Path(__file__).resolve().parent
TEMPLATES = SCRIPTS.parent / "templates"
ROWS = 6
TITLE_EM = 16  # 页面标题写主题词，16 个汉字以内（content_rules.md 第 3 节）
POINT_EM = 20  # 要点句 20 个汉字以内
CHARS_PER_MIN = 200  # 讲稿按每分钟约 200 字估算时长，只作参考（用户明确说过：流畅比字数要紧）
ID_RE = re.compile(r"^[a-z0-9][a-z0-9_-]*$")
MARK_RE = re.compile(r"^\[([a-z0-9][a-z0-9_-]*)\]$")
GRID_TAG_RE = re.compile(r"<[^>]*\bdata-grid=\"([^\"]*)\"[^>]*>")
COLS_RE = re.compile(r"\bdata-cols=\"(\d+)\"")
VARIANT_RE = re.compile(r"(<section\b[^>]*\bclass=\"variant\"[^>]*>)(.*?)</section>", re.S)
LABEL_RE = re.compile(r"\bdata-label=\"([^\"]*)\"")
# 图单（页面的 figures 字段）至少列几张：背景章每页 3 张以上（用户指出过「图不够丰富」，figures.md 第 1 节），其余内容页至少 1 张
FIG_MIN_BACKGROUND, FIG_MIN = 3, 1
# 图单里写的图片文件（只认图片格式：出处里的论文 PDF 不是要取的图）
FIG_PATH_RE = re.compile(r"[\w./\\-]+\.(?:png|jpe?g|gif|svg|webp)", re.I)
# 材料的角色：核心论文、补充论文、仅作背景（不引用、不取图）
MATERIAL_ROLES = ("核心", "补充", "仅背景")
CITED_RE = re.compile(r"[（(][^（()）]+[）)]\s*$")          # 资料要点末尾括号里的出处
DOI_NEEDED_RE = re.compile(r"\[(?:J|C|D)\]")                # 期刊、会议、学位论文条目要带 DOI（有的话）
WEB_RE = re.compile(r"https?://|Commons|网图|百度|公众号", re.I)  # 严格材料边界下不该出现的来源
IMG_SRC_RE = re.compile(r"(<img\b[^>]*?\bsrc=\")([^\"]+)(\")", re.I)
DRAFT_IMG_W = 900  # 初稿里的图缩到这个宽度（像素）再嵌进网页，整个网页才不会超过发布的大小上限


def fail(where, msg):
    raise ValueError(f"{where}：{msg}")


# ---------------------------------------------------------------- 大纲

def load(path):
    """读入并校验大纲和讲稿。两者都是手写的，这里是信任边界：缺字段、类型不对、id 重复、讲稿里的页 id 不存在都直接报错。"""
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    meta = data.get("meta")
    if not isinstance(meta, dict):
        fail("meta", "缺少 meta 对象")
    for k in ("title", "presenter", "date"):
        if not isinstance(meta.get(k), str) or not meta[k]:
            fail("meta", f"缺少字符串字段 {k}")
    mats = meta.get("materials", [])
    if not isinstance(mats, list):
        fail("meta.materials", "应为列表")
    for k, m in enumerate(mats):
        if not (isinstance(m, dict) and isinstance(m.get("file"), str) and m["file"] and m.get("role") in MATERIAL_ROLES
                and all(isinstance(m.get(f, ""), str) for f in ("read", "use"))):
            fail(f"meta.materials 第 {k + 1} 项",
                 f"应为 {{\"file\": 文件名, \"role\": {' / '.join(MATERIAL_ROLES)}, \"read\": 读哪些页或节, \"use\": 用途}}")
    if not isinstance(meta.get("strict", False), bool):
        fail("meta.strict", "应为 true 或 false")
    parts = data.get("parts")
    if not isinstance(parts, list) or not parts:
        fail("parts", "缺少章节列表")
    ids = []
    for i, part in enumerate(parts):
        where = f"第 {i + 1} 章"
        if not isinstance(part.get("title"), str) or not part["title"]:
            fail(where, "缺少 title")
        pages = part.setdefault("pages", [])  # 讲稿确认、切页之前，章节下还没有页
        if not isinstance(pages, list):
            fail(where, "pages 应为列表")
        for j, page in enumerate(pages):
            pw = f"{where}「{part['title']}」第 {j + 1} 页"
            pid = page.get("id")
            if not isinstance(pid, str) or not ID_RE.match(pid):
                fail(pw, "id 应为小写字母、数字、- 或 _ 组成的字符串（讲稿标记 [id] 和网页初稿 draft/<id>.html 都用它）")
            if pid in ids:
                fail(pw, f"id「{pid}」重复")
            ids.append(pid)
            for k in ("title", "point"):
                if k in page and not (isinstance(page[k], str) and page[k]):
                    fail(pw, f"{k} 应为非空字符串")
            if "title" not in page:
                fail(pw, "缺少 title（主题词标题）")
            for k in ("sources", "research", "refs", "figures", "todo"):
                v = page.get(k, [])
                if not (isinstance(v, list) and all(isinstance(s, str) for s in v)):
                    fail(pw, f"{k} 应为字符串列表")
            bar = page.get("bottombar")
            if bar is not None and not (isinstance(bar, list) and len(bar) == 2 and all(isinstance(s, str) for s in bar)):
                fail(pw, "bottombar 应为 [标签, 一句话]")
    data["dir"] = Path(path).resolve().parent
    data["script"] = load_script(data["dir"] / "script.md", ids)
    return data


def load_script(path, ids):
    """读讲稿：## 开头是章节，单独一行 [页 id] 开始一页，空行分段。返回 [{"title": 章节, "paras": [(页 id 或 None, 段落)]}]。
    讲稿还没切页时没有 [id] 标记，段落的页 id 都是 None。没有 script.md 时返回空列表。"""
    if not path.exists():
        return []
    chapters, page, buf = [], None, []

    def flush():
        if buf:
            if not chapters:
                chapters.append({"title": "", "paras": []})
            chapters[-1]["paras"].append((page, "".join(s.strip() for s in buf)))
            buf.clear()

    seen = []
    for n, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        s = line.strip()
        if s.startswith("# ") and not s.startswith("## "):
            continue
        if s.startswith("## "):
            flush()
            chapters.append({"title": s[3:].strip(), "paras": []})
            page = None
        elif MARK_RE.match(s):
            flush()
            page = MARK_RE.match(s).group(1)
            if page not in ids:
                fail(f"{path.name} 第 {n} 行", f"[{page}] 不是 outline.json 里的页 id")
            if page in seen:
                fail(f"{path.name} 第 {n} 行", f"[{page}] 出现了两次，同一页的段落要连在一起")
            seen.append(page)
        elif not s:
            flush()
        else:
            buf.append(s)
    flush()
    order = [i for i in ids if i in seen]
    if seen != order:
        fail(path.name, f"页的顺序和 outline.json 不一致：讲稿里是 {'、'.join(seen)}")
    return chapters


def page_paras(data):
    """{页 id: [段落]}：讲稿里属于每一页的段落。"""
    out = {}
    for ch in data["script"]:
        for pid, text in ch["paras"]:
            if pid:
                out.setdefault(pid, []).append(text)
    return out


def page_notes(outline_path):
    """{页 id: 口播稿}：讲稿里属于这一页的段落，段间换行。设计稿脚本用它写 deck.content(..., notes=)。"""
    return {k: "\n".join(v) for k, v in page_paras(load(outline_path)).items()}


REF_NUM_RE = re.compile(r"^\[(\d+)\]")


def collect_refs(data):
    """各页 refs 里的条目，按页序去重：[(条目, 第一次出现的页码)]。"""
    seen, out = set(), []
    for n, _, page in walk(data):
        for r in page.get("refs", []):
            if r not in seen:
                seen.add(r)
                out.append((r, n))
    return out


def refs_in_order(outline_path):
    """末尾参考文献附录的条目（deck.references() 用）：各页 refs 按第一次出现的顺序去重。"""
    return [r for r, _ in collect_refs(load(outline_path))]


def ref_hints(data):
    """参考文献编号的提示：编号要按第一次出现的顺序从 [1] 连续排下去，同一个编号不能对应两条不同的条目。"""
    h, nums = [], {}
    for r, n in collect_refs(data):
        m = REF_NUM_RE.match(r)
        if not m:
            h.append(f"第 {n} 页的参考文献没有「[n]」编号：{r[:24]}")
            continue
        k = int(m.group(1))
        if k in nums:
            h.append(f"编号 [{k}] 对应了两条不同的条目（第 {nums[k]} 页和第 {n} 页），同一篇文献各页要写成完全一样的条目")
        nums.setdefault(k, n)
    order = list(nums)
    if order != list(range(1, len(order) + 1)):
        h.append(f"编号按第一次出现的顺序应为 1–{len(order)} 连续，现在是 {order}")
    return h


def first_sentence(text):
    return re.split(r"(?<=[。！？])", text.strip(), maxsplit=1)[0] if text else ""


def walk(data):
    """按幻灯片顺序给出 (页码, 章节, 页面)，页码和设计稿 pptx 一致：封面 1、目录 2，每个进目录的章节前一张过渡页。"""
    n = 2
    for part in data["parts"]:
        if part.get("toc", True):
            n += 1
        for page in part["pages"]:
            n += 1
            yield n, part, page


def toc_parts(data):
    return [p for p in data["parts"] if p.get("toc", True)]


def em(s):
    """按汉字宽度折算的字数（西文按 0.55 个）。"""
    return tw(s, 1) / pt2cm(1)


def chars(text):
    """讲稿字数：汉字和字母数字都算，标点和空格不算。"""
    return len(re.findall(r"[一-鿿A-Za-z0-9]", text))


# ---------------------------------------------------------------- 网页初稿

def variants(data, page):
    """读 draft/<id>.html，返回 [(方案名, HTML)]。片段里有 <section class="variant" data-label="…"> 时每个是一个方案。"""
    f = data["dir"] / "draft" / f"{page['id']}.html"
    if not f.exists():
        return []
    src = f.read_text(encoding="utf-8")
    secs = list(VARIANT_RE.finditer(src))
    if not secs:
        return [("", src)]
    out = []
    for m in secs:
        lab = LABEL_RE.search(m.group(1))
        out.append((lab.group(1) if lab else f"方案 {len(out) + 1}", m.group(2)))
    return out


def check_blocks(src, page, where):
    """每块的 data-grid 在网格内、互不重叠。不同列数的网格按页宽比例换算成 60 等分再查重叠。"""
    cells = {}
    for k, m in enumerate(GRID_TAG_RE.finditer(src)):
        bw = f"{where} 第 {k + 1} 块"
        try:
            col, row, w, h = (int(v) for v in m.group(1).split(","))
        except ValueError:
            fail(bw, f"data-grid=\"{m.group(1)}\" 应为「列,行,占几列,占几行」四个整数")
        c = COLS_RE.search(m.group(0))
        cols = int(c.group(1)) if c else page.get("cols", 12)
        if col < 0 or row < 0 or w < 1 or h < 1 or col + w > cols or row + h > ROWS:
            fail(bw, f"data-grid {col},{row},{w},{h} 超出 {cols}×{ROWS} 网格")
        for x in range(col * 60 // cols, (col + w) * 60 // cols):
            for y in range(row, row + h):
                if (x, y) in cells:
                    fail(bw, f"和第 {cells[(x, y)] + 1} 块重叠")
                cells[(x, y)] = k
    if not cells:
        fail(where, "没有带 data-grid 的块")


NO_FIG_PARTS = ("总结", "结论", "展望", "下一步", "计划", "规划")  # 这些章节的页不查图单


def fig_need(part):
    """这一章每页图单至少几张：章名带「背景」的 FIG_MIN_BACKGROUND，总结、展望、计划类的 0，其余 FIG_MIN。"""
    if "背景" in part["title"]:
        return FIG_MIN_BACKGROUND
    return 0 if any(w in part["title"] for w in NO_FIG_PARTS) else FIG_MIN


def material_hints(data, page):
    """材料边界和出处的提示（references/content_rules.md 第 4 节）：资料要点没写出处、期刊条目没写 DOI、
    引用了「仅背景」材料；严格材料边界下资料或图来自清单以外（网页、Commons 等）。"""
    h = []
    uncited = [r for r in page.get("research", []) if not CITED_RE.search(r)]
    if uncited:
        h.append(f"资料要点没写出处（末尾括号写文件名和页码或节号）：{'；'.join(r[:16] for r in uncited)}")
    no_doi = [r.split("]")[0] + "]" for r in page.get("refs", []) if DOI_NEEDED_RE.search(r) and "doi" not in r.lower()]
    if no_doi:
        h.append(f"参考文献没写 DOI：{'、'.join(no_doi)}")
    mats = data["meta"].get("materials", [])
    stems = {Path(m["file"]).stem for m in mats}
    texts = page.get("research", []) + page.get("figures", []) + page.get("sources", []) + page.get("refs", [])
    for m in mats:
        stem = Path(m["file"]).stem
        if m["role"] == "仅背景" and any(stem in t for t in texts):
            h.append(f"用到了仅作背景的材料「{m['file']}」：它不引用、不取图")
    if data["meta"].get("strict"):
        outside = [t[:20] for t in texts if WEB_RE.search(t) and not any(s in t for s in stems)]
        if outside:
            h.append(f"严格材料边界下只用材料清单里的论文，这几条来自清单以外：{'；'.join(outside)}")
    return h


def hints(data):
    """每页的提示：标题偏长、没写或太长的要点句、讲稿里没有这一页、图单太少或图还没取到、材料边界和出处。
    返回 {页码: [提示]}。待核对事项（todo）不在这里，check 单独列出。"""
    paras = page_paras(data)
    out = {}
    for n, part, page in walk(data):
        h = []
        if em(page["title"]) > TITLE_EM:
            h.append(f"标题约 {em(page['title']):.0f} 字，写主题词（{TITLE_EM} 字以内）")
        if not page.get("point"):
            h.append("没写要点句")
        elif em(page["point"]) > POINT_EM:
            h.append(f"要点句约 {em(page['point']):.0f} 字，压到 {POINT_EM} 字以内")
        if data["script"] and page["id"] not in paras:
            h.append("讲稿里没有这一页（缺 [id] 标记）")
        spoken = sorted({w for w in COLLOQUIAL for t in (page["title"], page.get("point", "")) if w in t})
        if spoken:
            h.append(f"标题、要点句用了口语说法「{'、'.join(spoken)}」，改成书面语（content_rules.md 第 3 节）")
        figs = page.get("figures", [])
        need = fig_need(part)
        if len(figs) < need:
            h.append(f"图单 {len(figs)} 张（{'背景章每页' if need == FIG_MIN_BACKGROUND else '内容页'}至少 {need} 张），"
                     "在 figures 里列出要放的图和来源")
        missing = [p for f in figs for p in FIG_PATH_RE.findall(f)
                   if not (data["dir"] / p).exists() and not (data["dir"] / "draft" / p).exists()]
        if missing:
            h.append(f"图单里的图还没取到：{'、'.join(missing)}")
        out[n] = h + material_hints(data, page)
    return out


def cmd_check(args):
    data = load(args.outline)
    paras = page_paras(data)
    pages = list(walk(data))
    drafted = 0
    for n, part, page in pages:
        vs = variants(data, page)
        drafted += bool(vs)
        for label, src in vs:
            check_blocks(src, page, f"第 {n} 页「{page['title']}」{label}")
    if data["script"]:
        total = sum(chars(t) for ch in data["script"] for _, t in ch["paras"])
        print(f"讲稿 {total} 字，按每分钟 {CHARS_PER_MIN} 字约 {total / CHARS_PER_MIN:.1f} 分钟", end="")
        if data["meta"].get("minutes"):
            print(f"，汇报时长 {data['meta']['minutes']} 分钟", end="")
        print("（字数只作估算参考，不是硬性要求；讲得顺最要紧）")
        for ch in data["script"]:
            c = sum(chars(t) for _, t in ch["paras"])
            print(f"  {ch['title'] or '（未分章）'}：{c} 字，约 {c / CHARS_PER_MIN:.1f} 分钟")
    else:
        print("还没有讲稿 script.md")
    mats = data["meta"].get("materials", [])
    if mats:
        count = "、".join(f"{r} {sum(m['role'] == r for m in mats)} 份" for r in MATERIAL_ROLES)
        print(f"材料清单：{count}；{'严格材料边界（只用清单里的论文）' if data['meta'].get('strict') else '可上网补资料和图'}")
    else:
        print("还没写材料清单（meta.materials）")
    if not pages:
        print("\n还没切页：讲稿确认后在 outline.json 的章节下写 pages，并在讲稿里加 [页 id]")
        return
    print(f"\n内容页 {len(pages)} 页，已写网页初稿 {drafted} 页")
    print("\n页码  标题 ｜ 要点句 ｜ 过渡句（讲稿里这一页的第一句）：")
    for n, part, page in pages:
        bridge = first_sentence(paras.get(page["id"], [""])[0])
        print(f"  {n:>2}  {page['title']} ｜ {page.get('point', '（没写要点句）')} ｜ {bridge or '—'}")
    flagged = [(n, h) for n, h in hints(data).items() if h]
    if flagged:
        print("\n提示：")
        for n, h in flagged:
            print(f"  ⚠️ 第 {n} 页：{'；'.join(h)}")
    refs = collect_refs(data)
    if refs:
        print(f"\n参考文献 {len(refs)} 条，放在末尾附录（deck.references(refs_in_order(...))）")
        for h in ref_hints(data):
            print(f"  ⚠️ {h}")
    todo = [(n, t) for n, _, page in pages for t in page.get("todo", [])]
    if todo:
        print(f"\n待核对（{len(todo)} 项，交给用户确认前列在回复末尾）：")
        for n, t in todo:
            print(f"  ❓ 第 {n} 页：{t}")


# ---------------------------------------------------------------- 模板外观底图

def theme_colors(prs):
    """母版主题的配色方案：{名字: RRGGBB}，bg1/tx1/bg2/tx2 映射到 lt1/dk1/lt2/dk2。"""
    theme = prs.slide_master.part.part_related_by(RT.THEME)
    scheme = etree.fromstring(theme.blob).find(".//" + qn("a:clrScheme"))
    out = {}
    for el in scheme:
        c = el[0]
        out[etree.QName(el).localname] = c.get("val") if c.tag == qn("a:srgbClr") else c.get("lastClr")
    out.update(bg1=out["lt1"], tx1=out["dk1"], bg2=out["lt2"], tx2=out["dk2"])
    return out


WINGDINGS = {"n": "■", "l": "●", "q": "❑", "u": "◆", "Ø": "➢", "ü": "✓", "§": "▪"}


def text_spec(shape, colors):
    """文字框的位置（cm，从页面左上角量）和第一段文字的样式，网页初稿照此叠放文字。"""
    tf = shape.text_frame
    body = tf._txBody.find(qn("a:bodyPr"))
    p = tf.paragraphs[0]._p
    rpr = p.find(".//" + qn("a:rPr"))
    if rpr is None:
        raise ValueError(f"「{shape.name}」第一段没有文字样式，先在模板里给它写上占位词")
    clr = rpr.find(qn("a:solidFill"))[0]
    if len(clr):
        raise ValueError(f"「{shape.name}」的文字颜色带亮度调整，网页初稿不支持，改成纯色")
    ppr = p.find(qn("a:pPr"))
    algn = ppr.get("algn") if ppr is not None else None
    ins = {k: int(body.get(a, d)) / 360000 for k, a, d in
           (("l", "lIns", 91440), ("r", "rIns", 91440), ("t", "tIns", 45720), ("b", "bIns", 45720))}
    bullet = None
    bu = ppr.find(qn("a:buChar")) if ppr is not None else None
    if bu is not None:  # 华科模板的页面标题前有一个 Wingdings 方块，清空文字后它也没了，网页里照样画出来
        font = ppr.find(qn("a:buFont"))
        char = bu.get("char")
        if font is not None and font.get("typeface") == "Wingdings":
            char = WINGDINGS[char]
        bullet = {"char": char, "indent": int(ppr.get("marL", "0")) / 360000}
    return {"x": shape.left / 360000, "y": shape.top / 360000, "w": shape.width / 360000, "h": shape.height / 360000,
            "pt": int(rpr.get("sz", "1800")) / 100, "bold": rpr.get("b") == "1",
            "color": clr.get("val") if clr.tag == qn("a:srgbClr") else colors[clr.get("val")],
            "align": {"ctr": "center", "r": "right"}.get(algn, "left"),
            "anchor": {"ctr": "center", "b": "flex-end"}.get(body.get("anchor"), "flex-start"),
            "wrap": body.get("wrap") != "none", "ins": ins, "bullet": bullet}


def cmd_chrome(args):
    """把过渡页、内容页上要填字的形状清空、去掉页码，渲染成底图；记下这些形状的位置和文字样式。"""
    deck = Deck(args.template)
    colors = theme_colors(deck.prs)
    spec = {"slide": [deck.prs.slide_width / 360000, deck.prs.slide_height / 360000],
            "body": {"x": deck.template.x0 / 360000, "y": deck.template.y0 / 360000,
                     "w": deck.template.w, "h": deck.template.h},
            "pages": {}}
    for key, src, names in (("divider", deck.divider_t, {"章节标题": "title"}),
                            ("content", deck.content_t, {"页面标题": "title", "章节标签": "label"})):
        s = deck._clone(src)
        texts = {}
        for sh in list(s.shapes):
            if sh.name in names:
                texts[names[sh.name]] = text_spec(sh, colors)
                sh.text_frame.text = ""
            elif sh.is_placeholder and sh.placeholder_format.type == PP_PLACEHOLDER.SLIDE_NUMBER:
                texts["num"] = text_spec(sh, colors)
                sh._element.getparent().remove(sh._element)
        missing = set(names.values()) - set(texts)
        if missing:
            raise ValueError(f"模板「{args.template}」的{key}页缺少形状：{missing}")
        spec["pages"][key] = {"bg": f"{key}.png", "texts": texts}
    out = TEMPLATES / args.template / "web"
    out.mkdir(exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        pptx = Path(tmp) / "chrome.pptx"
        deck.save(pptx)
        subprocess.run([sys.executable, "-X", "utf8", str(SCRIPTS / "render.py"), str(pptx), tmp, "--dpi", "150"],
                       check=True)
        for k, key in ((1, "divider"), (2, "content")):
            (Path(tmp) / f"p-{k:02d}.png").replace(out / f"{key}.png")
    (out / "web.json").write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
    print("written", out)


# ---------------------------------------------------------------- 汇报网页

def data_uri(path):
    return "data:image/png;base64," + base64.b64encode(Path(path).read_bytes()).decode()


def thumb_uri(path):
    """初稿里的图：缩到 DRAFT_IMG_W 宽后转 data URI；有透明通道的存 PNG，其余存 JPEG。"""
    from PIL import Image
    im = Image.open(path)
    if im.width > DRAFT_IMG_W:
        im = im.resize((DRAFT_IMG_W, round(im.height * DRAFT_IMG_W / im.width)), Image.LANCZOS)
    buf = io.BytesIO()
    if im.mode in ("RGBA", "LA", "P"):
        im.convert("RGBA").save(buf, "PNG", optimize=True)
        mime = "png"
    else:
        im.convert("RGB").save(buf, "JPEG", quality=82)
        mime = "jpeg"
    return f"data:image/{mime};base64," + base64.b64encode(buf.getvalue()).decode()


def inline_images(src, root, where):
    """把初稿片段里 <img src="相对路径"> 换成嵌入的图。路径先按 draft/ 目录解析（../figures/x.png），再按汇报目录解析。"""
    def sub(m):
        url = m.group(2)
        if url.startswith(("data:", "http://", "https://")):
            return m.group(0)
        path = next((p for p in (root / "draft" / url, root / url) if p.exists()), None)
        if path is None:
            fail(where, f"图片不存在：{url}（相对 draft/ 或汇报目录）")
        return m.group(1) + thumb_uri(path) + m.group(3)
    return IMG_SRC_RE.sub(sub, src)


def site_slides(data, design):
    """网页脚本用的数据：每页一项，按幻灯片顺序。design 为设计稿渲染图目录（p-NN.png），没有时为 None。"""
    meta = data["meta"]
    tp = toc_parts(data)
    hs = hints(data)
    paras = page_paras(data)

    def shot(n):
        f = design / f"p-{n:02d}.png" if design else None
        return data_uri(f) if f and f.exists() else None

    slides = [{"n": 1, "type": "cover", "title": meta["title"], "design": shot(1)},
              {"n": 2, "type": "toc", "title": "目录", "design": shot(2),
               "items": [f"{cn_num(k + 1)}、{p['title']}" for k, p in enumerate(tp)]}]
    n = 2
    for part in data["parts"]:
        num = f"{cn_num(tp.index(part) + 1)}、" if part in tp else ""
        if part in tp:
            n += 1
            slides.append({"n": n, "type": "divider", "title": f"{num}{part['title']}", "design": shot(n)})
        for page in part["pages"]:
            n += 1
            ps = paras.get(page["id"], [])
            slides.append({
                "n": n, "type": "content", "id": page["id"], "part": part["title"], "partNo": num,
                "title": page["title"], "point": page.get("point", ""), "notes": "\n".join(ps),
                "bridge": first_sentence(ps[0]) if ps else "",
                "research": page.get("research", []), "sources": page.get("sources", []),
                "figures": page.get("figures", []), "todo": page.get("todo", []),
                "bar": page.get("bottombar"), "refs": page.get("refs", []), "cols": page.get("cols", 12),
                "bottom": grid_bottom(bool(page.get("bottombar"))),  # 网格下边离正文区底边的距离，和 Grid 同一个算法
                "variants": [{"label": lab, "html": inline_images(src, data["dir"], f"draft/{page['id']}.html")}
                             for lab, src in variants(data, page)],
                "design": shot(n), "hints": hs[n]})
    refs = [r for r, _ in collect_refs(data)]
    if refs:  # 末尾的参考文献附录；设计稿里按长度可能分成几页，网页只示意一页
        slides.append({"n": n + 1, "type": "refs", "title": "参考文献", "items": refs, "design": shot(n + 1)})
    return slides


def cmd_site(args):
    data = load(args.outline)
    meta = data["meta"]
    name = meta.get("template", "hust")
    use_template(name)
    web = TEMPLATES / name / "web"
    if not (web / "web.json").exists():
        raise ValueError(f"模板「{name}」还没有网页底图，先运行：outline.py chrome {name}")
    chrome = json.loads((web / "web.json").read_text(encoding="utf-8"))
    for p in chrome["pages"].values():
        p["bg"] = data_uri(web / p["bg"])
    design = Path(args.design) if args.design else None
    if design and not design.is_dir():
        raise ValueError(f"设计稿目录不存在：{design}")
    script = [{"title": ch["title"], "paras": [{"page": pid, "text": t, "chars": chars(t)} for pid, t in ch["paras"]]}
              for ch in data["script"]]
    payload = {"meta": {k: meta.get(k) for k in ("title", "presenter", "date", "occasion", "audience", "minutes", "thesis")},
               "chrome": chrome, "slides": site_slides(data, design), "script": script, "cpm": CHARS_PER_MIN,
               "tones": {k: {"base": t.base, "tint": t.tint, "ink": t.ink} for k, t in TONES.items()
                         if k in ("primary", "accent", "neutral", "error")}}
    blob = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    shell = (SCRIPTS / "site.html").read_text(encoding="utf-8")
    page = shell.replace("__TITLE__", html.escape(meta["title"][:12] + " 汇报"), 1).replace("__DATA__", blob, 1)
    Path(args.out).write_text(page, encoding="utf-8")
    print("written", args.out)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("check")
    p.add_argument("outline")
    p.set_defaults(fn=cmd_check)
    p = sub.add_parser("site")
    p.add_argument("outline")
    p.add_argument("out")
    p.add_argument("--design", help="设计稿渲染图目录（render.py 的输出，p-NN.png）")
    p.set_defaults(fn=cmd_site)
    p = sub.add_parser("chrome")
    p.add_argument("template")
    p.set_defaults(fn=cmd_chrome)
    args = ap.parse_args()
    args.fn(args)


if __name__ == "__main__":
    main()
