"""示例：一份 15 分钟阶段汇报的骨架。演示版式页、主题词标题 + 要点句条（point=）、口播稿（第一句接上一页），
以及按含义用色的几种正文排法。讲稿 script.md、outline.json 和网页初稿的完整流程见 examples/microwave/。
每页的证据都铺满正文区：内容少时加大卡片间距或纵向居中，不要让整页挤在上半部分。
有要点句的页，正文从 y 12.3 往下排（要点句条占 12.9–13.9），Grid 用 lead=True。

    uv run --no-project --with python-pptx --with lxml python -X utf8 example_deck.py <输出.pptx>
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
from hust_deck import *  # noqa: E402

deck = Deck("hust")
deck.cover("课程知识图谱阶段汇报", "汇报人", "2026年10月7日")
deck.toc(["本次工作", "方法设计", "后续计划"])

# ---- 路线图：上排是全部阶段和产出，本次阶段用实色标题条 + 粗边框；下排展开本次完成的工作
c = deck.content("本次工作", "本阶段工作与产出", point="本次完成对象设计，写成了校验程序",
                 notes="先看整体路线。路线分五个阶段，这次处于第三阶段对象设计，产出是格式说明和校验程序。"
                       "下面一排是这次具体做的三件事。")
# 每个阶段写清做了哪几件事，不只写一个词：听众要能看出每一步具体干什么
stages = [("需求调研", ["明确研究问题", "调研相关案例", "确定课程范围"], "需求说明"),
          ("架构设计", ["划分系统层次", "明确模块职责", "梳理数据流向"], "架构方案"),
          ("对象设计", ["设计知识对象", "定义字段与关系", "编写校验程序"], "格式说明、校验程序"),
          ("原型实现", ["确定技术选型", "整理样例数据", "实现查询与问答"], "原型系统"),
          ("效果评测", ["扩充教材知识", "检索与问答评测", "开展师生试用"], "评测结果")]
for i, (name, items, out) in enumerate(stages):
    card = Box(0.6 + i * 6.03, 11.7, 5.68, 5.6)
    current = i == 2
    # 已完成用灰色，本次用主色实色，未开始用主色浅底
    tone = "neutral" if i < 2 else "primary"
    body = c.card(card, header=f"{i + 1}  {name}", tone=tone, solid=current, highlight=current)
    c.text(Box(body.l + 0.45, body.t - 0.3, body.w - 0.7, 2.2), items, line=20)
    c.output(card, out, h=1.5)
    if current:
        c.chip("s", card.cx, card.t + 0.15, "本次进度", "accent", solid=True)
done = Box(0.6, 5.6, 29.8, 4.0)
body = c.card(done, header="本次完成的工作", solid=False)
for k, (head, desc) in enumerate([("知识点对象", ["六类知识点共用一套基本字段", "各类再加自己的专有内容"]),
                                  ("九种关系", ["七种有方向，方向由类型决定", "等价于、相关两种无方向"]),
                                  ("数据校验", ["启动时逐项检查 20 条规则", "出错直接报出具体位置"])]):
    x = done.l + k * done.w / 3
    if k:
        c.line((x, body.t - 0.6), (x, body.b + 0.6), LINE, 1.0)
    c.title("w", x + 0.9, body.cy + 0.75, head)
    c.text(Box(x + 0.9, body.cy + 0.2, done.w / 3 - 1.4, 1.7), desc, line=22)

# ---- 节点图：节点用知识点类型色，必须配图例；节点铺开到正文区上下两端
c = deck.content("方法设计", "知识点之间的关系", point="七种关系有方向，方向由类型决定",
                 notes="对象定下来以后，知识点之间怎么连？以电磁学为例，三条定律组成麦克斯韦方程组，再推导出波动方程，"
                       "波动方程描述电磁波。颜色表示知识点类型，箭头表示关系方向，方向由关系类型决定。")
laws = [c.node(5.5, y, name, "law") for y, name in [(11.0, "高斯定律"), (7.4, "法拉第定律"), (3.8, "安培环路定律")]]
mx = c.node(15.5, 7.4, "麦克斯韦方程组", "law")
for a in laws:
    c.link(a, mx, "组成")
wave = c.node(25.5, 7.4, "波动方程", "formula")
c.link(mx, wave, "推导出")
em = c.node(25.5, 3.8, "电磁波", "concept")
c.link(wave, em, "描述")
c.legend(9.5, 1.3, ["law", "formula", "concept"], undirected=False)

# ---- 类型色标题条 + 表格：同类条目多时用表格，不排成一排排小卡片。
# 用 Grid 排版：左边 6 列，右边上下各 3 行，边和间距自动对齐；有总结条时 bottombar=True，有要点句时 lead=True
c = deck.content("方法设计", "公式变量与单位检查", point="单位不一致时，启动检查直接报错",
                 notes="关系之外，公式里的变量也要和物理量对上。左边是法拉第定律，每个变量对应一个物理量和单位；"
                       "程序启动时做四类检查，单位写错就直接报错，右下是一条报错的样子。")
g = Grid(bottombar=True, lead=True)
card = g.box(0, 0, 6, 6)
body = c.card(card, header="公式 · 法拉第定律微分形式", tone="formula")
eq = Box(body.l + 0.5, body.t - 0.5, card.w - 1.0, 2.0)
c.tint(eq, "formula")
c.text(eq, P(R("∇ × ", latin="Cambria Math"), *var("E"), " = −∂", *var("B"), "/∂", *var("t")),
       size=S_TITLE, color=INK, align="c", anchor="m")
c.table(eq.l, eq.b - 0.6, [2.4, 6.6, eq.w - 9.0], [  # 左右边和公式底块对齐
    ("符号", "对应物理量", "单位"),
    (P(*var("E")), "电场强度", "V/m"),
    (P(*var("B")), "磁感应强度", "T"),
    (P(*var("t")), "时间", "s"),
], row_h=1.3)
body = c.card(g.box(6, 0, 6, 3), header="启动时的四类检查", solid=False)
c.bullets(Box(body.l + 0.6, body.t - 0.3, 12.6, body.h - 0.5), [
    P(R("编号唯一：", bold=True), "编号不重复，不重复建同类关系"),
    P(R("引用存在：", bold=True), "所属章节、关系两端必须存在"),
    P(R("内容匹配：", bold=True), "专有内容与知识点类型对应"),
    P(R("取值一致：", bold=True), "变量名称、单位与物理量一致"),
], line=20)
body = c.card(g.box(6, 3, 6, 3), header="报错示例", tone="error", solid=False)
c.text(Box(body.l + 0.6, body.t - 0.3, 12.6, body.h - 0.5), [
    P("变量 ", *var("B"), " 的单位写成 V/m"),
    P(R("报错：单位与物理量「磁感应强度」不一致", color=TONES["error"].base, bold=True), before=4),
    P("报错指出是哪个知识点的第几个变量", before=4),
    P("改正后重新启动，检查通过才提供查询", before=4),
], line=20)
c.bottombar("记录原则", "公式里出现哪些物理量只通过变量记录，不再另建关系")

# ---- 结尾：下一步。要点句说「先选型，再做四项」，版式就按这个结构排：左边一张选型卡，箭头指向四项原型工作
c = deck.content("后续计划", "下一阶段工作", point="先定技术选型，再做四项原型工作",
                 notes="对象和校验已经就绪，下一阶段先定技术选型，再做四项原型工作。先讲选型，再逐项说明四项工作的产出，"
                       "讲完停在这一页回答问题。")
group = Box(9.4, 12.3, 21.0, 10.7)
# 选型卡和右侧一行卡片差不多高，纵向对准原型工作组的中线，避免卡片里留大块空白
first = Box(0.6, group.cy + 3.5, 7.6, 7.0)
body = c.card(first, header="1  技术选型", solid=True, highlight=True)
# 窄卡片一行约 10 个字，条目写短，每条一行
c.bullets(Box(body.l + 0.5, body.t - 0.4, first.w - 1.0, 3.4),
          ["覆盖后端到前端五层", "逐层列出候选技术", "按需求逐项比较", "写明选择理由"], line=24)
c.output(first, "选型说明")
c.tint(group)
c.line((first.r + 0.15, group.cy), (group.l - 0.1, group.cy), TONES["primary"].base, 1.5, arrow=True)
for k, (name, desc, out) in enumerate([("样例数据", ["选一个代表性章节", "按本次结构整理并跑通校验"], "样例章节"),
                                       ("查询接口", ["搜索、知识详情、学习路径", "查询结果由知识库即时生成"], "查询接口"),
                                       ("知识检索", ["按问题找出相关知识点", "连同它们之间的关系一起返回"], "检索模块"),
                                       ("课程问答", ["大模型根据检索结果组织回答", "回答里返回引用的知识点"], "问答接口")]):
    card = Box(group.l + 0.5 + k % 2 * 10.15, group.t - 0.35 - k // 2 * 5.15, 9.85, 4.85)
    body = c.card(card, header=f"{k + 2}  {name}", solid=False)
    c.text(Box(body.l + 0.5, body.t - 0.25, body.w - 1.0, 1.6), desc, line=22)
    c.output(card, out, h=1.5)

deck.save(sys.argv[1])
print("written", sys.argv[1])
