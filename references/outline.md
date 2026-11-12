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
           "occasion": "...", "audience": "...", "minutes": 10, "thesis": "...",
           "materials": [{"file": "Dias2023.pdf", "role": "核心", "read": "第 3–9 页、图 2–6", "use": "机理和主导反应"},
                         {"file": "report.md", "role": "补充", "read": "第 2 章", "use": "背景和应用"},
                         {"file": "Alves2019.pdf", "role": "仅背景", "read": "", "use": "不引用、不取图"}],
           "strict": false},
  "parts": [
    {"title": "章节名", "pages": [
      {"id": "magnetron", "title": "磁控管的结构与原理", "point": "电子绕着谐振腔转，像吹哨一样激发微波",
       "research": ["电能变成微波的比例约 50%（microwave_oven.txt「Efficiency」段）"], "sources": ["source/xxx.txt"],
       "figures": ["拆开的家用磁控管，图上自带 1–6 编号（figures/magnetron.jpg）",
                   "O₂(a) 损失通道占比（Dias2023.pdf 第 7 页图 4b，保留横轴气压和图例 → figures/dias_fig4b.png）"],
       "todo": ["Kanarik 的「成本降一半」是哪一组对比，原文第 709 页待核对"],
       "cols": 12, "bottombar": ["类比", "一句话"]}]},
    {"title": "总结", "toc": false, "pages": [...]}
  ]
}
```

| 位置 | 字段 | 必填 | 含义 |
|---|---|---|---|
| meta | `title` `presenter` `date` | 是 | 封面标题、汇报人（最多约 8 个汉字）、日期（「2026年10月14日」） |
| meta | `template` | | 模板名，默认 `hust` |
| meta | `occasion` `audience` `minutes` `thesis` | | 场合、听众、时长（分钟）、要听众接受的那一个结论：需求调研的结果 |
| meta | `materials` | | 材料清单：每份材料的 `file`、`role`（核心 / 补充 / 仅背景）、`read`（读哪些页、节、图）、`use`（用来支撑什么）。需求调研时和用户一起定；「仅背景」的材料不引用、不取图，`check` 发现用到会提示 |
| meta | `strict` | | 严格材料边界，默认 `false`。论文精读、文献汇报设为 `true`：只读清单里写明的页和节，资料和图都只从清单里的论文来，不上网补、不画描述性插图；`check` 会提示来自网页、Commons 的条目 |
| 章节 | `title` | 是 | 章节名，写进目录、过渡页和章节标签，和讲稿的 `## ` 一致 |
| 章节 | `toc` | | 默认 `true`；`false` 时不进目录、不加过渡页，用于开头的引入页和最后的总结页 |
| 章节 | `pages` | | 这一章的页；讲稿确认、切页之前先不写，`check` 和网页的讲稿视图照样能用 |
| 页 | `id` | 是 | 小写字母、数字、`-`、`_`，全份不重复；讲稿里的 `[id]` 和初稿文件名 `draft/<id>.html` 都用它 |
| 页 | `title` | 是 | 主题词标题，16 字以内（`content_rules.md` 第 3 节） |
| 页 | `point` | | 要点句，20 字以内，画在标题下的浅底条里；写初稿前要写上 |
| 页 | `research` | | 资料要点，一条一句，末尾括号写出处：文件名加页码或节号（「Dias2023.pdf 第 7 页」「report.md 2.3 节」），有适用条件一起写（「30 mA、10 Torr 下」）。`check` 提示没写出处的条目；显示在预览的「搜索」里 |
| 页 | `sources` | | 出处：`source/` 下的文件或网址 |
| 页 | `figures` | | 图单：这一页要放的图，一条一张，写「放什么（从哪来）」；论文图写到论文、PDF 页码、图号和要保留的坐标轴、图例，取到后写上文件名（「O₂(a) 损失通道占比（Dias2023.pdf 第 7 页图 4b，保留横轴气压和图例 → figures/dias_fig4b.png）」）。切页时就列，章名带「背景」的章每页至少 3 张、其余页至少 1 张（总结、展望、计划类章节不查；`figures.md` 第 1 节）；原生图表、示意图也写进图单，注明「脚本画」；`check` 报出图单太少和写了文件名但还没取到的图，网页「搜索」里能看到 |
| 页 | `todo` | | 待核对：材料缺页、数据找不到原文依据、出处对不上时，写明缺什么，不猜、不自行扩大阅读范围。`check` 单独列出，网页便签和「搜索」里显示；每个检查点交给用户时列在回复末尾，用户答复后删掉对应条目 |
| 页 | `cols` | | 这一页网格的列数，默认 12；五步并列用 5，三步用 3 |
| 页 | `bottombar` | | 页底总结条 `["标签", "一句话"]`；有它时网格下边上移，和 `Grid(bottombar=True)` 一致 |
| 页 | `refs` | | 这一页引用的文献（答辩、开题讲到别人工作的页）：GB/T 7714 完整条目的列表，带「[n]」，期刊、会议、学位论文有 DOI 的写在末尾。页面上只标「[n]」，全文由 `deck.references(refs_in_order(...))` 汇总到末尾参考文献页；同一篇各页写成完全一样的条目。`check` 查编号是否从 [1] 连续、缺 DOI；网页「搜索」显示本页引用，翻页预览最后有参考文献页 |

页码按 封面 1、目录 2，每个进目录的章节前一张过渡页 数下去，和设计稿 pptx、讲稿视图、便签上的「#n」一致，
用户说「第 7 页」时几处对得上。

## 3. 网页初稿

`draft/<id>.html` 是一个 HTML 片段，只写正文区里的块；顶栏、页面标题、章节标签、页码、要点句条、页底总结条
由网页按模板画出来。每个块是一个带 `data-grid` 的元素：

```html
<div class="tint tight" data-grid="0,0,3,2"><h4><span class="num">1</span>阴极</h4><p>加热后发射电子，电子从这里出发</p></div>
<figure class="fig" data-grid="3,0,6,6"><img src="../figures/magnetron.jpg"><figcaption>1 阴极 2 阳极 3 永磁体 4 谐振腔</figcaption></figure>
<div class="ph" data-grid="9,0,3,6" data-kind="论文图">Witman 2019 图 1：射流装置，还没裁</div>
```

多面板页（`layouts.md` 第 7 节）的写法，一个面板里套两张方法卡片：

```html
<div class="panel" data-grid="4,0,8,6"><span class="ptab">工艺优化：贝叶斯优化</span>
  <div class="cols" style="flex:1;min-height:0">
    <div class="card"><div class="hd soft">虚拟刻蚀 [7]</div><div class="bd"><div class="figblock">
      <div class="imgs"><div><img src="../figures/kanarik_fig1_sim.png"><img src="../figures/kanarik_fig2.png"></div>
                        <div><img src="../figures/kanarik_fig3.png"></div></div>
      <div class="cap">刻蚀模拟器 · 工程师进度 · 人机配合的成本</div>
      <div class="ends"><p><b>调节变量：</b>11 个工艺参数</p><p class="concl">人机协同，达标成本降低约 50%</p></div>
    </div></div></div>
    <div class="card">…第二张卡片同样写…</div>
  </div></div>
```

- **图先放真图**：`<img src>` 写相对 `draft/` 的路径（`../figures/x.png`）或相对汇报目录的路径，`site` 会缩到 900 像素宽后嵌进网页；
  图片不存在时报错。还没取到的图才用 `.ph` 虚线占位框写明放什么、从哪来。

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
| `figure.fig` + `img` + `figcaption` / 块里直接一个 `img` | 真图，等比缩进块里；图注 13pt 灰字 | `figure(path, box, caption=...)` / `web_figure()` / 几张图用 `gallery()`、`strip()` |
| `.panel` + `.ptab`（`.panel.side` 竖排标签，内容包在 `.pbody` 里） | 虚线分组框 + 实色标签：多面板页 | `panel(box, 标签, side=...)` |
| `table.ftbl`（`th` 表头和行名，`td` 里放 `img` 或数值） | 图片表格 | `figtable(box, 表头, 行)` |
| `.concl` | 面板里一句强调色粗体结论 | `R(..., color=TONES["accent"].ink, bold=True)` |
| `.fine` | 12pt 注解：贴着图写的参数、数值、图注 | `size=S_NOTE` |
| `.figblock` > `.fhead`（可选）+ `.imgs`（每个子 `div` 一行图）+ `.cap` + `.ends`（几个 `p`，`b` 是对比项名，`.concl` 是结果） | 顶部说明 + 一组文献图 + 图注 + 底部对齐的几行 | `figblock(box, rows, 图注, lines, head=...)` |
| `.ph` + `data-kind` | 虚线占位框：还没取到的照片、截图、论文图，或原生图表、原生示意图、公式图 | `web_figure()` / `figure()` / `gallery()` / `chart()` / 原生形状 / `formula()` / `snippet()` |
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
