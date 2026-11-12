# 设计稿生成脚本：hust_deck API

写 `build_deck.py` 时读。结构照 `examples/microwave/build_deck.py`；初稿组件和这里方法的对照表在 `outline.md` 第 4 节。

## 1. 脚本结构

每份汇报写一个生成脚本，放在项目里（如 `reports/<汇报>/slides/build_deck.py`）：

```python
import sys
sys.path.insert(0, r"<skill>/scripts")
from hust_deck import *
from outline import page_notes

NOTES = page_notes("outline.json")       # {页 id: 讲稿里这一页的几段}，当口播稿，不在脚本里另写
deck = Deck("hust")                      # 模板名，见 templates.md
deck.cover(标题, 汇报人, 日期)          # 标题、汇报人过长时自动缩小字号
deck.toc([章节...])
deck.divider(章节)                       # 每章前一张，编号取目录顺序
c = deck.content(章节, 主题词标题, notes=NOTES[页 id], point=要点句,
                 refs=页["refs"])        # 要点句画在标题下的浅底条里；refs 是答辩、开题的页脚参考文献，可省
g = Grid(lead=True, refs=页["refs"])     # 有要点句的页，网格从要点句条下面开始；有页脚文献时下边让出位置
...                                      # 用 c 画正文
deck.save(输出路径)
```

运行：`uv run --no-project --with python-pptx --with lxml --with xlsxwriter python -X utf8 build_deck.py`（xlsxwriter 用于 `chart()` 内嵌的数据表）

## 2. 正文画布和网格

原点在左下角、y 向上，单位 cm，宽高由模板定（华科模板宽 31cm、高 14.3cm，下文坐标都按它举例）。
`Box(l, t, w, h)` 的 `t` 是上边。主色随模板变，脚本里要用主色时写 `TONES["primary"].base`，不写色值。

**排版用 `Grid` 网格对齐**：`g = Grid(lead=True)` 把要点句条以下的正文区切成 12 列 × 6 行，格间距统一 0.6cm；
`g.box(列, 行, 占几列, 占几行)` 返回一个块（列、行从 0 起，行从上往下数），边都落在网格线上，和初稿的 `data-grid` 用同样的四个数。
带总结条时加 `bottombar=True`；五列阶段卡用 `Grid(cols=5, lead=True)`。
几张图拼在一块区域里时用 `gallery()` / `strip()` 按宽高比算尺寸（`aspect(图)` 取宽高比），区域本身仍落在网格上；
只有 `layouts.md` 第 6 节那类几项并列的图片页，才从正文区 `B = Grid(lead=True).box(0, 0, 12, 6)` 起按图的形状排。
常见构图和例子见 `layouts.md` 第 5、6 节。

## 3. 方法

| 方法 | 用途 |
|---|---|
| `card(box, header=None, tone, solid=True, highlight=False, sub=None, shadow=False)` | 白底细边框卡片；header 给出时画标题条并返回其下的正文区 Box，sub 是标题条下一行灰字副标题。`shadow=True` 只给压在照片上的那一块 |
| `tint(box, tone)` | 浅色底块 |
| `panel(box, 标签, tone="primary", side=False)` | 分组框：虚线圆角边框 + 压在上边框正中的实色标签（`side=True` 为左侧竖排标签），返回放内容的 Box。多面板页用，见 `layouts.md` 第 7 节 |
| `figblock(box, rows, 图注, lines)` | 一组图（rows 同 `gallery()`）→ 一行 12pt 图注 → 底部几行 16pt 居中字（如「调的参数」+ 强调色结果）。底部从 box 底边往上排，几块行数相同就自动对齐。多面板页的面板和卡片里用 |
| `figtable(box, 表头, [(行名, [格子, ...]), ...])` | 图片表格：格子是图片路径、文字或 None；含图的行平分高度，纯文字行（指标值）0.8cm。返回 {(行, 列): 格子 Box} |
| `refs(条目)` | 页脚参考文献，9pt 灰字贴页面底部；由 `deck.content(refs=)` 调用，网格用 `Grid(refs=)` 让位。和 `bottombar()` 不能同页 |
| `heading(card, 标题, n=None)` / `title(anchor, x, y, s)` | 无标题条卡片的左上角标题（给 n 时带序号）/ 18pt 粗体标题 |
| `text(box, paras, size, color, align, anchor, line)` | 文本框；paras 可为 str、`P(...)`、`R(...)` 或列表 |
| `at(anchor, x, y, paras, size, w=None)` | 按锚点放文字；不给 w 时单行不换行 |
| `bullets(box, items, tone, anchor="t", gap=0)` | 项目符号列表（真项目符号）；`anchor="m"` 纵向居中，`gap` 为条目间距（pt） |
| `num(x, y, "01", tone)` / `chip(anchor, x, y, 文字, tone, solid=False)` | 序号圆点 / 小标签 |
| `node(x, y, paras, tone)` | 图中节点：tone 浅底 + 实色边框；知识点用类型色（返回 Box）。只用在知识关系图这类本身就是节点图的地方 |
| `line` / `poly` / `link(a, b, 说明)` / `label` | 直线 / 折线 / 两节点连线（自动裁到边框）/ 连线说明 |
| `legend(x, y, ["law", "formula"])` | 类型图例 + 有方向 / 无方向线型 |
| `table(l, t, widths, rows, row_h, header, first_col_bold)` | 无网格原生表格，表头下一条横线、行间细线；单元格可传列表，按意群分行 |
| `output(card, 产出, tone="accent")` | 卡片底部的「产出」框 |
| `keynum(x, y, 数值, 说明, tone="accent", big=False)` | 关键数字：18pt 粗体数值 + 下方 13pt 说明，以 (x, y) 为中心；`big=True` 时数值 36pt，当整页主角 |
| `chart(kind, box, 类别, {系列: 数值}, tones, unit, number_format, labels, highlight, vmax)` | 原生图表：`bar` / `barh` / `line`，数据可在 PowerPoint 里改；最多 3 个系列，默认本工作主色、基线灰色；`highlight=类别名` 把一根柱标成橙色；柱图的轴从 0 开始 |
| `figure(path, box, align="c", caption=None, fill=False)` / `web_figure(path, box, 图注, align, fill)` | 等比放图，`align`（c / w / e / n / s）定图在 box 里的位置，图注要写成 `caption=`（第三个位置参数是 align）；`fill=True` 裁掉多余部分铺满 box，一排实物照片要同样大小时用 / 放网上取来的图，图注必填 |
| `on_image(pic, 图, px, py)` | 原图像素坐标换算成画布坐标（`pic` 是 `figure()` 的返回值），引线标注的终点用 |
| `gallery(box, [[图, 图], [图]], gap, valign)` | 一块区域里拼多张图：同行等高、占满整行宽度，整组缩放到放得下。一张卡片放 2 张以上图时用它，不要逐张放进固定格子 |
| `strip(box, [([图, 图], 图注), ...], gap, ggap, align, valign)` | 几组图排成等高的一行，每组下面一行图注；`align="e"` 贴右边（如贴住箭头）。返回图片占用的 Box |
| `aspect(图)` | 图片宽高比，按图的形状算版面时用 |
| `formula(png, x, y)` / `snippet(png, anchor, x, y)` | 按 600dpi 原尺寸放公式图 / 按原尺寸放代码或终端图（返回的 Box 带 `line_y(k)`） |
| `bottombar(标签, 一句话)` | 页底总结条：一句设计原则或补充说明，不复述标题 |
| `kpis(box, [(数值, 说明)], big=False)` / `chips(x, y, [词])` / `timeline(box, [(名称, 说明)])` | 关键数字行（竖线分隔）/ 一排关键词胶囊 / 迷你时间线；和网页初稿的 `.kpis` `.chips` `.tl` 对应 |

短公式用 `P(...)` + `var()` 写成可编辑文字；代码截图、复杂公式、网上的图、论文里的图怎么取，见 `figures.md`。

## 4. 陷阱

- **输出文件被 PowerPoint 打开时写不进去**（WinError 32）。先输出到临时目录做质检，请用户关掉文件后再写入，或另存新文件名。
- **`qa_check.py` 把第 1 页当封面**，不查它。单独试做一页时先 `deck.cover(...)`，试做的页放在第 2 页。
- **改 `hust_deck.py` 前先读它的注释**：中文断行、主题样式、表格网格、幻灯片复制、横条图坐标轴、封面汇报人这些坑都在对应代码旁写明了。
- **宽度是估算的**：`tw()` 按汉字 1em、西文 0.55em 算。卡片里的文字留 10% 余量，最终以渲染图为准。
  `qa_check.py` 不估算表格单元格，表格里的词语有没有被拆开只能看渲染图；窄列先算「列宽 ÷ 0.564cm」得出一行能放几个字。
- **Bash heredoc 会吞反斜杠**：含 LaTeX、正则或 `\n` 的脚本用 Write 工具写成文件再运行。
- **复制形状时，形状 ID 和标签部件都不能重复**：同一页两个形状 ID 相同，或者共用同一个 `p:tags` 标签部件，
  PowerPoint 都会拒绝打开整个文件，`qa_check.py` 会报 ❌。换模板时先删掉 `p:custDataLst` 和 tags 部件。
- **过渡页编号取自目录顺序**：先调用 `toc()`，`divider(章节名)` 的章节名必须在目录里。
- **封面汇报人最多约 8 个汉字**：`cover()` 会去掉字间距、缩到最小 14pt，还放不下就报错，改写组名或第一作者。
