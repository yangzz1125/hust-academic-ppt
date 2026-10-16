"""示例：一份 15 分钟阶段汇报的骨架。演示版式页、结论句标题、口播稿，以及按含义用色的几种正文排法。
每页的证据都铺满正文区：内容少时加大卡片间距或纵向居中，不要让整页挤在上半部分。

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
c = deck.content("本次工作", "本阶段完成对象设计，并写成了校验程序",
                 notes="先说明整体路线，再点出本次处于第三阶段，产出是格式说明和校验程序。")
stages = [("需求调研", "梳理使用场景", "需求说明"), ("架构设计", "划分系统分层", "架构方案"),
          ("对象设计", "定义对象和关系", "格式说明、校验程序"), ("原型实现", "实现查询与问答", "原型系统"),
          ("效果评测", "评测检索效果", "评测结果")]
for i, (name, desc, out) in enumerate(stages):
    card = Box(0.6 + i * 6.03, 13.2, 5.68, 5.4)
    current = i == 2
    # 已完成用灰色，本次用主色实色，未开始用主色浅底
    tone = "neutral" if i < 2 else "primary"
    body = c.card(card, header=f"{i + 1}  {name}", tone=tone, solid=current, highlight=current)
    c.at("nw", body.l + 0.45, body.t - 0.6, desc)
    c.output(card, out)
    if current:
        c.chip("s", card.cx, card.t + 0.15, "本次进度", "accent", solid=True)
done = Box(0.6, 7.2, 29.8, 5.6)
body = c.card(done, header="本次完成的工作", solid=False)
for k, (head, desc) in enumerate([("知识点对象", "六类知识点，字段统一"),
                                  ("九种关系", "七种有方向，两种无方向"),
                                  ("数据校验", "启动时检查，错误定位到字段")]):
    x = done.l + k * done.w / 3
    if k:
        c.line((x, body.t - 0.8), (x, body.b + 0.8), LINE, 1.0)
    c.title("w", x + 0.9, body.cy + 0.55, head)
    c.at("w", x + 0.9, body.cy - 0.55, desc)

# ---- 节点图：节点用知识点类型色，必须配图例；节点铺开到正文区上下两端
c = deck.content("方法设计", "知识点之间用有方向的关系连接，方向由关系类型决定",
                 notes="三条定律组成麦克斯韦方程组，再推导出波动方程。颜色表示知识点类型，箭头表示关系方向。")
laws = [c.node(5.5, y, name, "law") for y, name in [(12.2, "高斯定律"), (8.4, "法拉第定律"), (4.6, "安培环路定律")]]
mx = c.node(15.5, 8.4, "麦克斯韦方程组", "law")
for a in laws:
    c.link(a, mx, "组成")
wave = c.node(25.5, 8.4, "波动方程", "formula")
c.link(mx, wave, "推导出")
em = c.node(25.5, 4.6, "电磁波", "concept")
c.link(wave, em, "描述")
c.legend(9.5, 1.3, ["law", "formula", "concept"], undirected=False)

# ---- 类型色标题条 + 表格：同类条目多时用表格，不排成一排排小卡片。
# 用 Grid 排版：左边 6 列，右边上下各 3 行，边和间距自动对齐；有总结条时 bottombar=True
c = deck.content("方法设计", "公式变量直接关联物理量，单位不一致时自动报错",
                 notes="点出变量表的三列，强调单位检查是程序做的，右下是一条报错的样子。")
g = Grid(bottombar=True)
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
], row_h=1.5)
body = c.card(g.box(6, 0, 6, 3), header="检查规则", solid=False)
c.bullets(Box(body.l + 0.6, body.t - 0.5, 12.6, 3.6), ["变量名称与物理量一致", "单位与物理量一致", "不一致时启动报错"])
body = c.card(g.box(6, 3, 6, 3), header="报错示例", tone="error", solid=False)
c.text(Box(body.l + 0.6, body.t - 0.5, 12.6, 3.4), [
    P("变量 ", *var("B"), " 的单位写成 V/m"),
    P(R("报错：单位与物理量「磁感应强度」不一致", color=TONES["error"].base, bold=True), before=8),
], line=24)
c.bottombar("记录原则", "公式里出现哪些物理量只通过变量记录，不再另建关系")

# ---- 结尾：下一步。标题说「先选型，再做四项」，版式就按这个结构排：左边一张选型卡，箭头指向四项原型工作
c = deck.content("后续计划", "下一阶段先定技术选型，再做四项原型工作",
                 notes="先讲技术选型，再逐项说明四项原型工作的产出，最后停在这一页回答问题。")
group = Box(9.4, 13.8, 21.0, 11.8)
# 选型卡和右侧一行卡片差不多高，纵向对准原型工作组的中线，避免卡片里留大块空白
first = Box(0.6, group.cy + 3.5, 7.6, 7.0)
body = c.card(first, header="1  技术选型", solid=True, highlight=True)
c.bullets(Box(body.l + 0.5, body.t - 0.5, first.w - 1.0, 2.8), ["列出各层候选技术", "按需求逐项比较", "写明选择理由"])
c.output(first, "选型说明")
c.tint(group)
c.line((first.r + 0.15, group.cy), (group.l - 0.1, group.cy), TONES["primary"].base, 1.5, arrow=True)
c.title("w", group.l + 0.5, group.t - 0.6, "原型工作")
for k, (name, desc, out) in enumerate([("样例数据", "选一个代表性章节整理", "样例章节"),
                                       ("查询接口", "搜索、详情、学习路径", "查询接口"),
                                       ("知识检索", "按问题找出相关知识点", "检索模块"),
                                       ("课程问答", "回答并返回引用依据", "问答接口")]):
    card = Box(group.l + 0.5 + k % 2 * 10.15, group.t - 1.3 - k // 2 * 5.1, 9.85, 4.7)
    body = c.card(card, header=f"{k + 2}  {name}", solid=False)
    c.at("nw", body.l + 0.5, body.t - 0.35, desc)
    c.output(card, out)

deck.save(sys.argv[1])
print("written", sys.argv[1])
