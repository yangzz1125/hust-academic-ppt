# 讲稿 script.md、大纲 outline.json 和网页初稿

一份汇报的目录里有三样东西，`scripts/outline.py` 读它们：

| 文件 | 内容 | 什么时候写 |
|---|---|---|
| `script.md` | 讲稿：要上台讲的话，切页后每段标出属于哪一页 | 第一步，用户确认后再切页 |
| `outline.json` | 每页的 id、主题词标题、要点句、资料要点和出处 | 切页时 |
| `draft/<页 id>.html` | 网页初稿，每页一个 HTML 片段 | 切页后 |

`check` 列出讲稿字数和每页的「标题 ｜ 要点句 ｜ 过渡句」，`site` 生成汇报网页（讲稿、便签画布、翻页预览），
设计稿脚本用 `page_notes("outline.json")` 取每页的口播稿。
完整示例是 `examples/microwave/`：`script.md`、`outline.json`、`draft/` 是讲稿和初稿，同目录的 `build_deck.py` 是照它做的设计稿。

## 1. script.md

```markdown
# 微波炉为什么能加热食物

## 微波从哪里来

[history]
先看第一个现象：为什么都是 2.45 GHz？这得从微波炉的来历说起……

1947 年，第一台商用微波炉 Radarange 上市……

[magnetron]
这 30 年里，微波炉越做越小，核心却一直是磁控管……
```

- `# ` 开头的一行是汇报标题，只给人看，脚本忽略。`## ` 开头是章节，写 `outline.json` 里的章节名。
- 空行分段。单独一行 `[页 id]` 开始一页，它下面直到下一个 `[页 id]` 或章节的段落都属于这一页。
- 还没切页时不写 `[页 id]`，网页的讲稿视图照样能看，每段旁边显示「还没切页」。
- 页的顺序要和 `outline.json` 一致，同一页的段落要连在一起，`check` 会报错并指出行号。
- 这一页的几段就是它的口播稿（段间换行），第一句是过渡句，显示在便签画布的连线上。写法见 `content_rules.md` 第 2 节。

## 2. outline.json

```json
{
  "meta": {"title": "...", "presenter": "...", "date": "...", "template": "hust",
           "occasion": "...", "audience": "...", "minutes": 10, "thesis": "..."},
  "parts": [
    {"title": "章节名", "pages": [
      {"id": "magnetron", "title": "磁控管的结构与原理", "point": "电子绕着谐振腔转，像吹哨一样激发微波",
       "research": ["要点（出处）"], "sources": ["source/xxx.txt"], "cols": 12, "bottombar": ["类比", "一句话"]}]},
    {"title": "总结", "toc": false, "pages": [...]}
  ]
}
```

| 位置 | 字段 | 必填 | 含义 |
|---|---|---|---|
| meta | `title` `presenter` `date` | 是 | 封面标题、汇报人（最多约 8 个汉字）、日期（「2026年10月14日」） |
| meta | `template` | | 模板名，默认 `hust` |
| meta | `occasion` `audience` `minutes` `thesis` | | 场合、听众、时长（分钟）、要听众接受的那一个结论：需求调研的结果 |
| 章节 | `title` | 是 | 章节名，写进目录、过渡页和章节标签，和讲稿的 `## ` 一致 |
| 章节 | `toc` | | 默认 `true`；`false` 时不进目录、不加过渡页，用于开头的引入页和最后的总结页 |
| 章节 | `pages` | | 这一章的页；讲稿确认、切页之前先不写，`check` 和网页的讲稿视图照样能用 |
| 页 | `id` | 是 | 小写字母、数字、`-`、`_`，全份不重复；讲稿里的 `[id]` 和初稿文件名 `draft/<id>.html` 都用它 |
| 页 | `title` | 是 | 主题词标题，16 字以内（`content_rules.md` 第 3 节） |
| 页 | `point` | | 要点句，20 字以内，画在标题下的浅底条里；写初稿前要写上 |
| 页 | `research` | | 资料要点，一条一句，括号里写出处；显示在预览的「搜索」里 |
| 页 | `sources` | | 出处：`source/` 下的文件或网址 |
| 页 | `cols` | | 这一页网格的列数，默认 12；五步并列用 5，三步用 3 |
| 页 | `bottombar` | | 页底总结条 `["标签", "一句话"]`；有它时网格下边上移，和 `Grid(bottombar=True)` 一致 |

页码按 封面 1、目录 2，每个进目录的章节前一张过渡页 数下去，和设计稿 pptx、讲稿视图、便签上的「#n」一致，
用户说「第 7 页」时几处对得上。

## 3. 网页初稿

`draft/<id>.html` 是一个 HTML 片段，只写正文区里的块；顶栏、页面标题、章节标签、页码、要点句条、页底总结条
由网页按模板画出来。每个块是一个带 `data-grid` 的元素：

```html
<div class="tint tight" data-grid="0,0,3,2"><h4><span class="num">1</span>阴极</h4><p>加热后发射电子，电子从这里出发</p></div>
<div class="ph" data-grid="3,0,6,6" data-kind="照片">拆开的家用磁控管，图上自带 1–6 编号；两侧说明块用引线指到编号</div>
```

- `data-grid="列,行,占几列,占几行"`：和 `Grid(lead=True).box(列, 行, 占几列, 占几行)` 完全相同，行从上往下数，共 6 行；
  写了要点句的页，网格从要点句条下面开始。一页上半五列、下半十二列时，块上加 `data-cols="12"` 改这一块的列数。
  块不能越出网格、不能互相重叠，`check` 会报错。
- 网页里的长度和 PPT 一一对应（`--cm`、`--pt` 换算），字号就是 PPT 的磅数。放不下的块网页标红框，改到没有红框。
- 版式照设计稿的做法排（照片当主角、引线标注、命名页型），不要一格一格地把网格填满。
- 几个方案写在同一个文件里，每个包一层 `<section class="variant" data-label="A 卡片版">…</section>`，预览里切换。
- 内联样式只用来微调排列（居中、换成横排），颜色用 `var(--p)`（主色）、`var(--p-ink)`、`var(--a-ink)`（强调色）这些变量，
  不写别的色值。

## 4. 初稿组件和设计稿的对照

| 初稿写法 | 是什么 | 设计稿里用 |
|---|---|---|
| `.card` + `.hd` + `.bd` | 白底细边框卡片、实色标题条、正文区；并排的卡片都是平的 | `card(box, header=...)` |
| `.hd.soft` / `.hd.accent` / `.card.hl` / `.card.lift` | 浅底标题条 / 强调色标题条 / 加粗边框标重点 / 带阴影（只给压在照片上的那一块） | `solid=False` / `tone="accent"` / `highlight=True` / `shadow=True` |
| `.sub` | 标题条下一行灰字副标题 | `card(..., sub=...)` |
| `.tint`（`.tight` 内边距小一些）/ `.tint.accent` / `.tint.neutral` / `.solid` | 主色浅底块 / 强调色浅底 / 灰底 / 主色实色块 | `tint(box)` / `tint(box, "accent")` / `tint(box, "neutral")` / `shape(..., TONES["primary"].base)` |
| `h4` / `.num` | 块里的 18pt 粗体小标题 / 实心圆序号 | `title()`、`heading()` / `num()` |
| `ul` / `ol.steps` | 项目符号列表 / 带圆形序号的步骤 | `bullets()` / `num()` + `text()` |
| `.kpis` (`.kpis.bigv`) | 关键数字行，竖线分隔 | `kpis(box, [(数值, 说明)], big=...)` |
| `.big` / `.mark` | 36pt 主角数字或大序号 / 28pt 年份 | `keynum(..., big=True)`、`at(..., size=S_DISPLAY)` / `S_MARK` |
| `.chips` | 一排关键词胶囊 | `chips(x, y, [...])` |
| `.tl` | 时间线：横线、圆点、年份、名称和说明 | `timeline(box, [(名称, 说明)])`，或按 layouts.md 第 7 节画大时间线 |
| `.vs` / `.cols` | 两三块并排对比 / 等分几栏 | 几个 `tint()` 并排 |
| `table.tbl` | 无网格表格，表头下一条主色线 | `table()` |
| `.ph` + `data-kind` | 虚线占位框：照片、截图、论文图、原生图表、原生示意图、公式图 | `web_figure()` / `figure()` / `gallery()` / `chart()` / 原生形状 / `formula()` / `snippet()` |
| `.em` | 强调色粗体的关键结论 | `R(..., color=TONES["accent"].ink, bold=True)` |

封面和目录在网页里只画示意，设计稿用模板的封面页和目录页。

## 5. 提示

`check`、便签和网页会标出：标题超过 16 字、没写要点句或要点句超过 20 字、讲稿里没有这一页。这些是提示，不阻止生成。
讲稿字数按每分钟约 200 字估算时长，只作参考。

## 6. 换模板

网页初稿的模板外观来自 `templates/<模板>/web/`：过渡页和内容页的底图（要填字的形状清空、去掉页码后渲染），
以及页面标题、章节标签、页码的位置和字体。新增或修改模板后运行一次（要 LibreOffice）并提交生成的文件：

```bash
uv run --no-project --with python-pptx --with lxml python -X utf8 <skill>/scripts/outline.py chrome <模板名>
```
