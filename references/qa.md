# 质检

每轮生成设计稿后都过一遍。

## 1. 命令

```bash
uv run --no-project --with python-pptx python -X utf8 <skill>/scripts/qa_check.py deck.pptx --source 报告.tex estimates.md --figures figures
uv run --no-project python -X utf8 <skill>/scripts/render.py deck.pptx <scratchpad>/png [--pages 4,18] [--engine powerpoint]
```

- `qa_check.py`：❌ 必须修完；⚠️ 逐条修掉，或在交付说明里写明为何保留。最后列出的「标题 ｜ 要点句 ｜ 过渡句」按顺序读一遍，检查论证和衔接。
  「没有图片也没有图表」「图只占正文区不到两成」「figures 里有图没放进 PPT」这三类先处理：回到 `figures.md` 第 1 节，
  看取了没用的图里有没有能证明要点句的，换掉自己画的方框。
- 再用 pptx skill 的 `scripts/office/validate.py` 做格式校验。它在 `~/.claude/skills/synced/<id>/pptx/` 下，
  `<id>` 随同步变化，用通配符找（Bash）：
  `uv run --no-project --with lxml --with defusedxml python -X utf8 $(ls ~/.claude/skills/synced/*/pptx/scripts/office/validate.py | head -1) deck.pptx`
  找不到这个文件时（pptx skill 没同步或目录改了），在交付说明里写明这一步没做，不要默默跳过；PowerPoint 渲染那一步照常做。
- `render.py` 默认用 LibreOffice（字体是本机微软雅黑，宽度可信），日常目检用它。**交付前必须用 `--engine powerpoint` 渲染一次**：
  LibreOffice 和格式校验都能通过、PowerPoint 却拒绝打开的文件确实出现过，只有这一步能发现。
  有 `chart()` 的页只看 PowerPoint 渲染图，LibreOffice 画图表时图例顺序和字号都不同。

## 2. 逐页目检

用 Read 看图，每一项都答「是」才算过（第 8 项看整份的总览图）：

| # | 问题 |
|---|---|
| 1 | 文字没有溢出、没有压到别的文字、线条或图片（`qa_check.py` 会报文字压图），词语没有被换行拆开？ |
| 2 | 主角是原图、一眼可见？要对着讲的图够大、图上的字看得清（只作展示的文献图不要求，但要多放几张、拼满所在的块）？整页没有大块空白或挤在一角，卡片里没有半张空着？（`qa_check.py` 按估算查卡片空白，以看图为准） |
| 3 | 标题是主题词、要点句一行放下？同一件事没有在页上讲两遍？ |
| 4 | 箭头方向、连线端点正确，说明文字没有压在节点上？ |
| 5 | 公式变量斜体、下标正确？ |
| 6 | 页面有成块的颜色而不是全灰（多面板页看主色标签，`layouts.md` 第 6 节的并列图片页看主色标签和箭头）？颜色含义和全篇一致，类型色有图例，红色只表示错误？橙色只落在要点句点名的那一项，平级并列项没有被单独标橙？ |
| 7 | 页面上的数字都能在材料里找到？ |
| 8 | 整份翻下来有轻重变化，不是每页同一种结构、同一种密度？（把全部页拼成一张总览图看：`render.py` 后在 png 目录里运行 `sheet.py <scratchpad>/sheet.png p-*.png --labels 1,2,...,n --cols 4`，`--labels` 个数和图片数相同） |

## 3. 交付

交付时说明哪些内容需要作者本人复核（数字、要点句措辞、待补的图）。未经要求不写额外的报告或说明文件。
