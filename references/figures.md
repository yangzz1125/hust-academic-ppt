# 取图与嵌图

页面要放代码、终端输出、复杂公式、网上的图或论文里的图时读这一份。结构图（流程、框图、关系图、表格）一律原生画，不在这里。

## 1. 代码、JSON、终端输出

```bash
uv run --no-project --with pillow --with pygments --with fonttools python -X utf8 <skill>/scripts/snippet.py \
    片段.json 片段.png --kind json|python|terminal --title "文件名（节选）" --pt 14
```

- 字号用 13–14pt（放进 PPT 后就是这个字号）；一行放不下就把片段改窄（节选、换行），不缩小字号。
- terminal 模式里以 # 开头的行是灰色注释，含「失败 / 错误」的行标红。
- 报错示例要真实运行得到（例如改坏一份副本再加载），不手写。原文和渲染命令记在生成脚本里。
- 用 `c.snippet(png, anchor, x, y)` 按原尺寸放入；箭头用返回值的 `line_y(k)` 对准第 k 行，不按像素手算。

## 2. 公式

- 短式子写成可编辑文字：`P(R("∇ × ", latin="Cambria Math"), *var("E"), " = …")`，变量用 `var()`（斜体、可带下标）。
- 分式、积分等复杂公式用 LaTeX standalone 编译后 `pdftoppm -r 600 -png -singlefile` 导出，再用 `c.formula(png, x, y)` 放入。

## 3. 网上的图

网上的图都可以用，按哪张最合适来选，不受来源限制。讲到器件、设备、部件、实验装置、历史事件时先找照片，
找不到再考虑原生示意（用户明确要求过，见 SKILL.md 第 3 节）：

- **Wikimedia Commons**：用 `search` / `get` 直接检索下载，作者和许可自动记下，SVG 可取 PNG 渲染图。
  常用词可用中文（「磁控管」「微波炉」），专业术语用英文更准（「驻波」搜不准，`standing wave` 准）；
  要图上是中文标注的示意图，关键词后加 `intitle:zh`（如 `"electromagnetic spectrum intitle:zh"`）。
- **其他网站**（百度图片、百度百科、公众号、博客、论文、厂商页面等）：百度对脚本直接访问返回安全验证，
  所以先用 WebSearch / WebFetch 或浏览器（Claude in Chrome）找到图片直链和所在网页，再用 `add` 下载登记。

```bash
F="uv run --no-project --with pillow python -X utf8 <skill>/scripts/fetch_image.py"
$F search "magnetron" --kind photo          # --kind diagram 只找 SVG 示意图
$F get "File:Magnetron detail.jpg" <汇报目录>/figures --name magnetron
$F add <图片直链> <汇报目录>/figures --name xxx --title 题名 --source "百度百科" --page <图片所在网页> [--author 作者]
```

- `add` 用浏览器 UA 请求，并把 `--page` 作为 Referer（公众号、百科图床会查防盗链），WebP 自动转成 PNG。
  报「返回的不是图片」时，多半是防盗链或验证页：补上 `--page`，或在浏览器里打开图片拿原图直链。
- **公众号**：文章页 `mp.weixin.qq.com/s/...` 能直接抓，正文图片是 `data-src="https://mmbiz.qpic.cn/..."`（`&amp;` 还原成 `&`）。
  `--page` 填这篇文章的链接，或者不填；填别的网站会拿到「此图片来自微信公众平台 未经允许不可引用」的 140×140 占位图。
  图片地址里的 `wx_fmt=gif` 是动图，PPT 里能正常播放。
- `add` 会打印下载到的像素尺寸，尺寸异常小时先用 Read 看是不是占位图。
- `get` / `add` 在图片旁写同名 `.json`，记下作者、来源、许可（其他网站没写许可的留空，作者不明写「佚名」）；
  `reference` 是 GB/T 7714-2015 的 [EB/OL] 条目，
  需要参考文献页时直接用它。用 `c.web_figure(path, box, 图注)` 放入，缺 `.json` 会报错。
- 一次连续取十几张图后 Commons 会返回 HTTP 429，脚本会按 Retry-After 等待后重试；批量取图时每张之间停几秒。
- 取图后用 Read 看一遍：内容对不对、有没有裁掉的标注、图上的字放进页面后是否太小、有没有水印。有缺陷就换图，
  或者改用原生示意图、`chart()` 重画（`examples/microwave/` 的水分子示意和驻波曲线就是这样来的）。
- **图上自带的编号写进图注**：`get` 会打印 Commons 的描述，里面常有「1) cathode 2) anode」这样的编号含义，
  图注按图上的编号全部列出（如「磁控管剖面：1 阴极 2 阳极 3 永磁体……」），讲解也按这个编号来。
- 图片按放置尺寸至少 150dpi（`qa_check.py` 会查），不够就换大图或放小。示意图优先选 SVG，可以取任意宽度。
- 图里的英文标注（如 E、H、λ）听众能看懂就保留；需要中文时在图旁用原生文字加注，不改原图。

## 4. 论文里的图

从 PDF 按 300dpi 裁到子图，不截整页、不把整张多面板图缩小塞进页面。

```bash
P="uv run --no-project --with pymupdf --with pillow python -X utf8 <skill>/scripts/pdf_figure.py"
$P list 论文.pdf                                   # 各页图注（Fig. n）和位图的位置，坐标是页面比例
$P preview 论文.pdf 5 <scratchpad>/p5.png           # 带 10% 网格的整页预览，Read 看图读出裁切范围
$P crop 论文.pdf 5 <汇报目录>/figures/fig3b.png --box 0.12,0.54,0.82,0.89
```

- 裁切命令记在生成脚本里，换范围时改命令重跑。不裁论文原图注，图注在 PPT 里用 `figure(..., caption=)` 重新写
  （「论文图 3b：……」）。
- **裁切质检**（Read 看裁出的图，每项都要满足，缺一项就是缺陷，不是可以接受的取舍）：
  坐标轴、轴标题、刻度、图例、色标、比例尺都在；讲解要提到的子图字母（a、b、c）在；方法名、分组名可读；
  四周留了余量；按计划尺寸放进页面后最小的字不小于 12pt。
- 修法依次是：扩大裁切范围、拆成两页、整页只放这张图、简单的柱图和表格用 `chart()` / `table()` 按论文里的数值重画。
