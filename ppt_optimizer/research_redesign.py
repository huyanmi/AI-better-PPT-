from __future__ import annotations

import re
from dataclasses import dataclass, replace
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

from .style_reference import DEFAULT_STYLE, ReferenceStyle


A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


SLIDE_W = 1280
SLIDE_H = 720
EMU_PER_PX = 9525


@dataclass(frozen=True)
class TextBox:
    x: int
    y: int
    w: int
    h: int
    text: str
    size: int = 18
    color: str = "1E1E1E"
    bold: bool = False
    align: str = "left"
    fill: str = "none"
    line: str | None = None
    font: str | None = None
    name: str = "Text"


@dataclass(frozen=True)
class ShapeBox:
    x: int
    y: int
    w: int
    h: int
    fill: str = "FFFFFF"
    line: str | None = None
    geometry: str = "rect"
    name: str = "Shape"


@dataclass(frozen=True)
class ConnectorBox:
    x1: int
    y1: int
    x2: int
    y2: int
    color: str
    width: int = 2
    arrow: bool = True
    name: str = "Connector"


class SlideBuilder:
    def __init__(self, style: ReferenceStyle) -> None:
        self.style = style
        self.parts: list[str] = []
        self.shape_id = 2

    def shape(self, box: ShapeBox) -> None:
        self.parts.append(_shape_xml(self._next_id(), box))

    def connector(self, box: ConnectorBox) -> None:
        self.parts.append(_connector_xml(self._next_id(), box))

    def text(self, box: TextBox) -> None:
        if box.font is None:
            box = replace(box, font=self.style.body_font)
        self.parts.append(_text_xml(self._next_id(), box))

    def _next_id(self) -> int:
        value = self.shape_id
        self.shape_id += 1
        return value

    def xml(self) -> bytes:
        body = "\n".join(self.parts)
        xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="{A_NS}" xmlns:r="{R_NS}" xmlns:p="{P_NS}">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
{body}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"""
        return xml.encode("utf-8")


def redesign_slide_xml(
    data: bytes,
    slide_number: int | None = None,
    style: ReferenceStyle | None = None,
) -> bytes:
    style = style or DEFAULT_STYLE
    texts = extract_slide_text(data)
    profile = classify_slide(texts)
    if profile == "architecture":
        if (slide_number or 1) % 2 == 0:
            return _stack_architecture_slide(texts, slide_number, style)
        return _architecture_slide(texts, slide_number, style)
    if profile == "workflow":
        return _workflow_slide(texts, slide_number, style)
    if profile == "results":
        return _results_slide(texts, slide_number, style)
    if profile == "matrix":
        return _matrix_slide(texts, slide_number, style)
    return _summary_slide(texts, slide_number, style)


def extract_slide_text(data: bytes) -> list[str]:
    root = ET.fromstring(data)
    texts: list[str] = []
    for node in root.findall(f".//{{{A_NS}}}t"):
        text = _clean_text(node.text or "")
        if text:
            texts.append(text)
    return _dedupe_texts(texts)


def classify_slide(texts: list[str]) -> str:
    joined = " ".join(texts)
    if re.search(r"结果|实验|准确率|损失|曲线|ROC|消融|误差|分析|性能|精度|召回|F1", joined, re.I):
        return "results"
    if len(texts) >= 24 or re.search(r"局限|方法类别|核心问题|评测|指标|痛点|对比|矩阵|竞品|预算|计划", joined):
        return "matrix"
    if re.search(r"方法|流程|步骤|设计|实现|训练|平台|数据集|特征|分类|映射|Token|Prompt", joined, re.I):
        return "workflow"
    if re.search(r"系统|架构|RAG|工具|调用|接口|模型层|治理层|知识库|向量库", joined, re.I):
        return "architecture"
    if len(texts) >= 18:
        return "matrix"
    return "summary"


def _architecture_slide(texts: list[str], slide_number: int | None, style: ReferenceStyle) -> bytes:
    title_red, title_blue = _title_parts(texts, "系统架构", "RAG、工具调用与策略约束协同")
    builder = _base_builder(title_red, title_blue, slide_number, style)

    lanes = [
        ("接入层", "用户问题\n权限与策略\n日志采集"),
        ("语义层", "意图识别\nEmbedding\n查询改写"),
        ("知识层", "向量库\n文档库\n结构化数据"),
        ("模型层", "LLM 推理\n上下文规划\n答案生成"),
        ("执行层", "工具调用\n代码执行\n业务 API"),
        ("治理层", "安全过滤\n证据链\n审计回放"),
    ]
    for index, (label, body) in enumerate(lanes):
        x = 48 + index * 196
        color = style.accent if index == len(lanes) - 1 else style.secondary
        _section(builder, label, x, 138, 142, color)
        builder.text(TextBox(x, 180, 142, 124, body, 15, align="center", line=color))
        if index < len(lanes) - 1:
            _arrow(builder, x + 144, 230, 44, style.secondary)

    builder.shape(ShapeBox(74, 378, 1070, 138, style.pale, style.border))
    notes = _evidence_notes(texts)
    headings = ["关键设计", "证据优先", "工具受控", "全链路审计"]
    for index, heading in enumerate(headings):
        x = 102 + index * 260
        color = style.accent if index == 3 else style.primary
        builder.text(TextBox(x, 404, 180, 24, heading, 18, color, True))
        if index > 0:
            builder.text(TextBox(x, 438, 220, 52, notes[index - 1], 14))

    builder.text(TextBox(84, 560, 900, 28, _conclusion(texts, "模型输出从文本生成升级为证据驱动的可执行决策。"), style.body_size + 2, style.accent, True))
    return builder.xml()


def _stack_architecture_slide(texts: list[str], slide_number: int | None, style: ReferenceStyle) -> bytes:
    title_red, title_blue = _title_parts(texts, "系统架构", "分层治理与证据链闭环")
    builder = _base_builder(title_red, title_blue, slide_number, style)

    layers = [
        ("输入治理", "任务意图\n权限边界\n数据清洗"),
        ("知识增强", "检索召回\n证据排序\n上下文压缩"),
        ("模型推理", "Prompt 规划\nLLM 生成\n自检修正"),
        ("输出审计", "结果解释\n风险提示\n日志回放"),
    ]
    x0, y0, w, h = 92, 132, 530, 72
    for index, (label, body) in enumerate(layers):
        y = y0 + index * 96
        fill = style.pale if index % 2 == 0 else "FFFFFF"
        line = style.accent if index == len(layers) - 1 else style.secondary
        builder.shape(ShapeBox(x0, y, w, h, fill, line))
        builder.text(TextBox(x0 + 24, y + 18, 118, 28, label, style.body_size + 2, line, True))
        builder.text(TextBox(x0 + 168, y + 14, 292, 34, body, max(13, style.body_size - 2)))
        if index < len(layers) - 1:
            _arrow(builder, x0 + w // 2, y + h + 8, 0, style.secondary, vertical=26)

    builder.shape(ShapeBox(720, 138, 420, 330, "FFFFFF", style.border))
    builder.text(TextBox(748, 166, 220, 28, "科研表达重点", style.title_size - 6, style.primary, True))
    callouts = _evidence_notes(texts)
    labels = ["主旨先行", "证据承载", "风险闭环"]
    for index, label in enumerate(labels):
        y = 218 + index * 78
        color = style.accent if index == 2 else style.primary
        builder.shape(ShapeBox(750, y + 8, 10, 10, color, None, "ellipse"))
        builder.text(TextBox(774, y, 108, 22, label, style.body_size, color, True))
        builder.text(TextBox(892, y - 2, 206, 36, callouts[index], max(12, style.body_size - 3)))

    _conclusion_bar(builder, texts, style, "分层架构页应同时回答“模块是什么、证据在哪里、风险如何闭环”。")
    return builder.xml()


def _workflow_slide(texts: list[str], slide_number: int | None, style: ReferenceStyle) -> bytes:
    title_red, title_blue = _title_parts(texts, "方法流程", "从问题定义到实验验证的研究路径")
    builder = _base_builder(title_red, title_blue, slide_number, style)

    steps = _workflow_steps(texts)
    x_positions = [76, 356, 636, 916]
    y_positions = [176, 286, 176, 286]
    for index, (label, body) in enumerate(steps):
        x, y = x_positions[index], y_positions[index]
        color = style.accent if index == len(steps) - 1 else style.primary
        builder.shape(ShapeBox(x, y, 190, 106, "FFFFFF", color))
        builder.shape(ShapeBox(x + 16, y - 26, 42, 42, color, None, "ellipse"))
        builder.text(TextBox(x + 25, y - 17, 24, 18, f"{index + 1}", 17, "FFFFFF", True, "center", fill=color))
        builder.text(TextBox(x + 22, y + 20, 140, 22, label, style.body_size + 1, color, True))
        builder.text(TextBox(x + 22, y + 52, 140, 34, body, max(12, style.body_size - 3)))
        if index < len(steps) - 1:
            start_x = x + 190
            start_y = y + 53
            end_x = x_positions[index + 1] - 18
            end_y = y_positions[index + 1] + 53
            builder.connector(ConnectorBox(start_x, start_y, end_x, end_y, style.secondary))

    builder.shape(ShapeBox(78, 466, 1030, 58, style.pale, style.border))
    builder.text(TextBox(104, 484, 112, 18, "设计原则", style.body_size, style.primary, True))
    builder.text(TextBox(226, 478, 790, 30, _conclusion(texts, "流程页应呈现清晰的阅读路径：问题、方法、验证与输出逐步收束。"), style.body_size))
    _conclusion_bar(builder, texts, style, "方法流程要避免堆文字，关键是让评审一眼看出变量、步骤和验证关系。", y=558)
    return builder.xml()


def _results_slide(texts: list[str], slide_number: int | None, style: ReferenceStyle) -> bytes:
    title_red, title_blue = _title_parts(texts, "结果分析", "指标、对比与结论的证据板")
    builder = _base_builder(title_red, title_blue, slide_number, style)

    builder.shape(ShapeBox(72, 132, 610, 330, "FFFFFF", style.border))
    builder.text(TextBox(96, 158, 230, 24, "核心证据图", style.body_size + 2, style.primary, True))
    _mini_chart(builder, 116, 226, 500, 160, style)
    builder.text(TextBox(116, 400, 480, 26, _conclusion(texts, "结果页优先展示关键指标变化，再解释原因和适用边界。"), max(13, style.body_size - 2), style.muted))

    findings = _bullets(texts, 3)
    for index, finding in enumerate(findings):
        y = 138 + index * 104
        color = style.accent if index == 0 else style.primary
        builder.shape(ShapeBox(738, y, 390, 76, style.pale if index == 0 else "FFFFFF", style.border))
        builder.text(TextBox(760, y + 16, 70, 20, f"发现 {index + 1}", max(13, style.body_size - 2), color, True))
        builder.text(TextBox(842, y + 12, 246, 34, finding, max(13, style.body_size - 2)))

    _conclusion_bar(builder, texts, style, "结果页应先给结论，再用图表、对比和误差解释支撑判断。")
    return builder.xml()


def _matrix_slide(texts: list[str], slide_number: int | None, style: ReferenceStyle) -> bytes:
    title_red, title_blue = _title_parts(texts, "评测体系", "任务簇、失败模式与科研证据矩阵")
    builder = _base_builder(title_red, title_blue, slide_number, style)
    builder.text(TextBox(48, 92, 920, 24, "将页面信息重排为科研汇报中更常见的表格证据结构，突出问题、证据和改进机制。", max(13, style.body_size - 2), style.muted))

    headers = ["方面", "核心信息", "证据/样本", "风险或不足", "优化方向"]
    widths = [120, 220, 250, 250, 300]
    x0, y0 = 50, 136
    x = x0
    for header, width in zip(headers, widths):
        _table_cell(builder, header, x, y0, width, 38, style.secondary, "FFFFFF", True, max(14, style.body_size - 1), style)
        x += width

    rows = _matrix_rows(texts)
    for row_index, row in enumerate(rows):
        y = y0 + 38 + row_index * 74
        x = x0
        for col_index, (value, width) in enumerate(zip(row, widths)):
            fill = style.pale if col_index == 0 else "FFFFFF"
            color = style.accent if col_index == 3 else style.body
            _table_cell(builder, value, x, y, width, 74, fill, color, col_index in (0, 3), max(12, style.body_size - 4), style)
            x += width

    builder.text(TextBox(58, 558, 1060, 24, _conclusion(texts, "科研型 PPT 应把信息组织成可验证的问题、证据和方案链路。"), style.body_size, style.accent, True))
    return builder.xml()


def _summary_slide(texts: list[str], slide_number: int | None, style: ReferenceStyle) -> bytes:
    title_red, title_blue = _title_parts(texts, "研究结论", "从文字堆叠转向结构化科研叙事")
    builder = _base_builder(title_red, title_blue, slide_number, style)

    builder.text(TextBox(52, 108, 500, 28, "核心判断", style.title_size - 4, style.accent, True))
    bullets = _bullets(texts, 5)
    for index, bullet in enumerate(bullets):
        y = 158 + index * 54
        builder.shape(ShapeBox(62, y + 8, 10, 10, "FFFFFF", style.primary))
        builder.text(TextBox(86, y, 510, 36, bullet, style.body_size))

    loop = [
        ("研究问题", "定义目标"),
        ("技术路线", "形成方法"),
        ("实验验证", "沉淀证据"),
        ("场景落地", "闭环迭代"),
    ]
    for index, (label, body) in enumerate(loop):
        x = 660 + (index % 2) * 250
        y = 150 + (index // 2) * 168
        color = style.accent if index == 3 else style.secondary
        _section(builder, label, x, y, 160, color)
        builder.text(TextBox(x, y + 42, 160, 82, body, 20, color, True, "center", line=color))
    _arrow(builder, 820, 205, 80, style.secondary)
    _arrow(builder, 820, 373, 80, style.secondary)
    builder.shape(ShapeBox(736, 274, 3, 58, style.primary))
    builder.shape(ShapeBox(986, 274, 3, 58, style.accent))

    builder.shape(ShapeBox(72, 526, 1060, 58, _light_accent(style.accent), style.accent))
    builder.text(TextBox(96, 544, 70, 18, "结论", max(14, style.body_size - 1), style.accent, True))
    builder.text(TextBox(176, 540, 830, 26, _conclusion(texts, "该页已重排为红蓝标题、结构化要点和科研汇报结论栏。"), style.body_size))
    return builder.xml()


def _base_builder(title_red: str, title_blue: str, slide_number: int | None, style: ReferenceStyle) -> SlideBuilder:
    builder = SlideBuilder(style)
    builder.shape(ShapeBox(0, 0, SLIDE_W, SLIDE_H, "FFFFFF", None))
    builder.text(TextBox(38, 22, 210, 38, title_red, style.title_size, style.accent, True, font=style.title_font))
    builder.text(TextBox(252, 24, 910, 34, f"—{title_blue}", 22, style.primary, font=style.title_font))
    builder.shape(ShapeBox(38, 82, 1188, 2, style.primary, None))
    builder.shape(ShapeBox(36, 686, 1188, 1, style.border, None))
    builder.text(TextBox(38, 694, 520, 14, "Research-style redesigned by PPT Optimizer", 10, style.muted))
    builder.text(TextBox(1192, 694, 34, 14, str(slide_number or 1).zfill(2), 10, style.muted, align="right"))
    return builder


def _section(builder: SlideBuilder, label: str, x: int, y: int, w: int, color: str) -> None:
    builder.shape(ShapeBox(x, y, w, 26, color, None))
    builder.text(TextBox(x + 6, y + 4, w - 12, 16, label, 13, "FFFFFF", True, "center", fill=color))


def _arrow(
    builder: SlideBuilder,
    x: int,
    y: int,
    w: int,
    color: str = "4F7DB9",
    vertical: int = 0,
) -> None:
    end_x = x + w
    end_y = y + vertical
    builder.connector(ConnectorBox(x, y, end_x, end_y, color, arrow=False))
    if vertical:
        direction = 1 if vertical > 0 else -1
        builder.connector(ConnectorBox(end_x, end_y, end_x - 5, end_y - direction * 8, color, arrow=False))
        builder.connector(ConnectorBox(end_x, end_y, end_x + 5, end_y - direction * 8, color, arrow=False))
        return
    direction = 1 if w >= 0 else -1
    builder.connector(ConnectorBox(end_x, end_y, end_x - direction * 9, end_y - 5, color, arrow=False))
    builder.connector(ConnectorBox(end_x, end_y, end_x - direction * 9, end_y + 5, color, arrow=False))


def _table_cell(
    builder: SlideBuilder,
    text: str,
    x: int,
    y: int,
    w: int,
    h: int,
    fill: str,
    color: str,
    bold: bool,
    size: int,
    style: ReferenceStyle,
) -> None:
    builder.text(TextBox(x, y, w, h, text, size, color, bold, "center", fill, style.border))


def _conclusion_bar(
    builder: SlideBuilder,
    texts: list[str],
    style: ReferenceStyle,
    fallback: str,
    y: int = 540,
) -> None:
    builder.shape(ShapeBox(72, y, 1060, 54, _light_accent(style.accent), style.accent))
    builder.text(TextBox(98, y + 18, 76, 18, "结论", max(14, style.body_size - 1), style.accent, True))
    builder.text(TextBox(184, y + 13, 806, 24, _conclusion(texts, fallback), style.body_size, style.body, True))


def _workflow_steps(texts: list[str]) -> list[tuple[str, str]]:
    labels = ["问题定义", "方法构建", "实验验证", "结果输出"]
    bullets = _bullets(texts, 8)
    while len(bullets) < 8:
        bullets.append("")
    defaults = [
        "明确目标与约束",
        "形成模型或算法路径",
        "用指标检验有效性",
        "沉淀结论与边界",
    ]
    return [
        (label, bullets[index * 2] or defaults[index])
        for index, label in enumerate(labels)
    ]


def _mini_chart(builder: SlideBuilder, x: int, y: int, w: int, h: int, style: ReferenceStyle) -> None:
    builder.shape(ShapeBox(x, y + h, w, 2, style.border, None))
    builder.shape(ShapeBox(x, y, 2, h, style.border, None))
    bars = [0.46, 0.62, 0.78, 0.88]
    labels = ["Baseline", "RAG", "Tool", "Ours"]
    bar_w = 54
    gap = 62
    for index, value in enumerate(bars):
        bx = x + 72 + index * (bar_w + gap)
        bh = int(h * value)
        color = style.accent if index == len(bars) - 1 else style.secondary
        builder.shape(ShapeBox(bx, y + h - bh, bar_w, bh, color, None))
        builder.text(TextBox(bx - 8, y + h + 12, bar_w + 16, 16, labels[index], 10, style.muted, align="center"))


def _title_parts(texts: list[str], fallback_red: str, fallback_blue: str) -> tuple[str, str]:
    title = _best_title(texts)
    if not title:
        return fallback_red, fallback_blue
    for marker in ("-", "—", "–", ":"):
        if marker in title:
            left, right = title.split(marker, 1)
            left = _clean_title(left) or fallback_red
            right = _clean_title(right) or fallback_blue
            return _short(left, 8), _short(right, 24)
    if len(title) <= 12:
        return _short(title, 8), fallback_blue
    return fallback_red, _short(title, 24)


def _best_title(texts: list[str]) -> str:
    candidates = [text for text in texts if 5 <= len(text) <= 42]
    if not candidates:
        return texts[0] if texts else ""
    return max(candidates[:8], key=lambda item: (sum("\u4e00" <= c <= "\u9fff" for c in item), len(item)))


def _evidence_notes(texts: list[str]) -> list[str]:
    defaults = [
        "答案必须绑定可追溯文档片段，降低幻觉风险。",
        "工具调用经过权限、参数和沙箱校验，避免越权执行。",
        "保留 prompt、检索证据、工具返回和最终输出。",
    ]
    long_texts = [text for text in texts if len(text) >= 14]
    notes = [_short(text, 34) for text in long_texts[1:4]]
    return (notes + defaults)[:3]


def _matrix_rows(texts: list[str]) -> list[list[str]]:
    items = _bullets(texts, 16)
    while len(items) < 16:
        items.append("")
    aspects = ["研究问题", "技术路线", "实验评测", "场景落地", "治理审计"]
    rows: list[list[str]] = []
    for index, aspect in enumerate(aspects):
        chunk = items[index * 3 : index * 3 + 3]
        rows.append(
            [
                aspect,
                chunk[0] or "核心信息",
                chunk[1] or "证据样本",
                chunk[2] or "风险不足",
                "形成数据回流与版式重排机制",
            ]
        )
    return rows


def _bullets(texts: list[str], limit: int) -> list[str]:
    values: list[str] = []
    for text in texts:
        if len(text) < 3:
            continue
        if re.fullmatch(r"\d+|[A-Za-z]{1,3}", text):
            continue
        values.append(_short(text, 42))
        if len(values) >= limit:
            break
    if values:
        return values
    return ["提炼核心问题", "重排技术路线", "突出实验依据", "形成应用闭环"][:limit]


def _conclusion(texts: list[str], fallback: str) -> str:
    candidates = [text for text in texts if len(text) >= 18]
    if not candidates:
        return fallback
    return _short(candidates[-1], 58)


def _clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def _clean_title(value: str) -> str:
    return re.sub(r"^[\s:：\-—–]+|[\s:：\-—–]+$", "", value).strip()


def _dedupe_texts(texts: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for text in texts:
        if text in seen:
            continue
        seen.add(text)
        result.append(text)
    return result


def _short(value: str, limit: int) -> str:
    value = _clean_text(value)
    if len(value) <= limit:
        return value
    return f"{value[: limit - 1]}…"


def _text_xml(shape_id: int, box: TextBox) -> str:
    line_xml = _line_xml(box.line)
    fill_xml = _fill_xml(box.fill)
    paragraphs = "\n".join(_paragraph_xml(line, box) for line in box.text.splitlines() or [""])
    return f"""      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="{_xml(box.name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{_emu(box.x)}" y="{_emu(box.y)}"/><a:ext cx="{_emu(box.w)}" cy="{_emu(box.h)}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          {fill_xml}
          {line_xml}
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" anchor="ctr"><a:spAutoFit/></a:bodyPr>
          <a:lstStyle/>
{paragraphs}
        </p:txBody>
      </p:sp>"""


def _paragraph_xml(text: str, box: TextBox) -> str:
    align = {"left": "l", "center": "ctr", "right": "r"}.get(box.align, "l")
    bold = ' b="1"' if box.bold else ""
    font = box.font or DEFAULT_STYLE.body_font
    return f"""          <a:p>
            <a:pPr algn="{align}"/>
            <a:r>
              <a:rPr lang="zh-CN" sz="{box.size * 100}"{bold}>
                <a:solidFill><a:srgbClr val="{box.color}"/></a:solidFill>
                <a:latin typeface="{_xml(font)}"/><a:ea typeface="{_xml(font)}"/><a:cs typeface="{_xml(font)}"/>
              </a:rPr>
              <a:t>{_xml(text)}</a:t>
            </a:r>
          </a:p>"""


def _shape_xml(shape_id: int, box: ShapeBox) -> str:
    return f"""      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="{_xml(box.name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="{_emu(box.x)}" y="{_emu(box.y)}"/><a:ext cx="{_emu(box.w)}" cy="{_emu(box.h)}"/></a:xfrm>
          <a:prstGeom prst="{box.geometry}"><a:avLst/></a:prstGeom>
          {_fill_xml(box.fill)}
          {_line_xml(box.line)}
        </p:spPr>
      </p:sp>"""


def _connector_xml(shape_id: int, box: ConnectorBox) -> str:
    x = min(box.x1, box.x2)
    y = min(box.y1, box.y2)
    w = max(1, abs(box.x2 - box.x1))
    h = max(1, abs(box.y2 - box.y1))
    flip_h = ' flipH="1"' if box.x2 < box.x1 else ""
    flip_v = ' flipV="1"' if box.y2 < box.y1 else ""
    end_xml = '<a:tailEnd type="triangle" w="sm" len="sm"/>' if box.arrow else ""
    return f"""      <p:sp>
        <p:nvSpPr><p:cNvPr id="{shape_id}" name="{_xml(box.name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm{flip_h}{flip_v}><a:off x="{_emu(x)}" y="{_emu(y)}"/><a:ext cx="{_emu(w)}" cy="{_emu(h)}"/></a:xfrm>
          <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
          <a:ln w="{box.width * 6350}">
            <a:solidFill><a:srgbClr val="{box.color}"/></a:solidFill>
            {end_xml}
          </a:ln>
        </p:spPr>
      </p:sp>"""


def _fill_xml(color: str | None) -> str:
    if not color or color.lower() == "none":
        return "<a:noFill/>"
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


def _line_xml(color: str | None) -> str:
    if not color:
        return "<a:ln><a:noFill/></a:ln>"
    return f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>'


def _light_accent(color: str) -> str:
    r = int(color[0:2], 16)
    g = int(color[2:4], 16)
    b = int(color[4:6], 16)
    return "".join(f"{round(channel + (255 - channel) * 0.92):02X}" for channel in (r, g, b))


def _emu(value: int | float) -> int:
    return int(round(value * EMU_PER_PX))


def _xml(value: str) -> str:
    return escape(value, {'"': "&quot;"})
