# 来源说明

本仓库的代码和文档按 MIT 许可发布（见 `LICENSE`），以下内容除外，按各自的来源和许可使用。

## 模板

- `templates/hust/template.pptx`：由 2026-09-22 的「电磁场与波 AI 知识图谱调研汇报」PPT 抽取版式，
  保留华中科技大学校徽、封面图和配色，文字换成占位词，并删掉了原编辑软件留下的形状标签（p:custDataLst / tags）。
  校徽和校名属于华中科技大学，不在 MIT 许可范围内。

## 示例材料

- `examples/microwave/source/*.txt`：英文维基百科条目 Microwave oven、Dielectric heating、Cavity magnetron 的纯文本，
  CC BY-SA 4.0，2026-10-14 通过 MediaWiki API 取得，仅作示例材料。
- `examples/microwave/figures/`：Wikimedia Commons 图片，作者、来源和许可见各图同名 `.json`。

## 参考的项目

`references/content_rules.md` 与 SKILL.md 的单页法则参考了以下项目的思路，文字和代码均为重新编写：

- [academic-pptx-skill](https://github.com/gabberflast/academic-pptx-skill)（MIT License, Copyright (c) 2026 Gabberflast）：
  结论句标题、只读标题测试、一页一个证据、引用规范。
- [p-ppt-generate](https://github.com/ShengLin1001/p-ppt-generate)：按模板克隆版式、按时长定结构、
  数字与材料交叉核对、逐页目检清单。该仓库未声明许可，本项目没有复制其代码。
- [nature-skills](https://github.com/Yuan1z0825/nature-skills) 的 nature-paper2ppt：子图裁切质检、按论文类型定主线、
  结果页图文比例、同构图重复上限、近似错位和套话检查。
- [jnu-academic-ppt-skill](https://github.com/LiGonghao3/jnu-academic-ppt-skill)：组会「进展 → 证据 → 卡点 → 下一步」主线、原生图表。
- [Auto-claude-code-research-in-sleep / paper-slides](https://github.com/wanshuiyin/Auto-claude-code-research-in-sleep)：图占页面主体。
