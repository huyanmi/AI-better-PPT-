from __future__ import annotations

import re
from dataclasses import dataclass
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape


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
    font: str = "Microsoft YaHei"
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


class SlideBuilder:
    def __init__(self) -> None:
        self.parts: list[str] = []
        self.shape_id = 2

    def shape(self, box: ShapeBox) -> None:
        self.parts.append(_shape_xml(self._next_id(), box))

    def text(self, box: TextBox) -> None:
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


def redesign_slide_xml(data: bytes, slide_number: int | None = None) -> bytes:
    texts = extract_slide_text(data)
    profile = classify_slide(texts)
    if profile == "architecture":
        return _architecture_slide(texts, slide_number)
    if profile == "matrix":
        return _matrix_slide(texts, slide_number)
    return _summary_slide(texts, slide_number)


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
    if re.search(r"系统|架构|RAG|工具|调用|接口|LLM|模型层|治理层", joined, re.I):
        return "architecture"
    if len(texts) >= 18 or re.search(r"评测|指标|痛点|对比|矩阵|竞品|预算|计划", joined):
        return "matrix"
    return "summary"


def _architecture_slide(texts: list[str], slide_number: int | None) -> bytes:
    title_red, title_blue = _title_parts(texts, "系统架构", "RAG、工具调用与策略约束协同")
    builder = _base_builder(title_red, title_blue, slide_number)

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
        color = "D60000" if index == len(lanes) - 1 else "4F7DB9"
        _section(builder, label, x, 138, 142, color)
        builder.text(TextBox(x, 180, 142, 124, body, 15, align="center", line=color))
        if index < len(lanes) - 1:
            _arrow(builder, x + 144, 230, 44)

    builder.shape(ShapeBox(74, 378, 1070, 138, "F4F7FD", "B8C6DB"))
    notes = _evidence_notes(texts)
    headings = ["关键设计", "证据优先", "工具受控", "全链路审计"]
    for index, heading in enumerate(headings):
        x = 102 + index * 260
        color = "D60000" if index == 3 else "243A8F"
        builder.text(TextBox(x, 404, 180, 24, heading, 18, color, True))
        if index > 0:
            builder.text(TextBox(x, 438, 220, 52, notes[index - 1], 14))

    builder.text(TextBox(84, 560, 900, 28, _conclusion(texts, "模型输出从文本生成升级为证据驱动的可执行决策。"), 20, "D60000", True))
    return builder.xml()


def _matrix_slide(texts: list[str], slide_number: int | None) -> bytes:
    title_red, title_blue = _title_parts(texts, "评测体系", "任务簇、失败模式与科研证据矩阵")
    builder = _base_builder(title_red, title_blue, slide_number)
    builder.text(TextBox(48, 92, 920, 24, "将页面信息重排为科研汇报中更常见的表格证据结构，突出问题、证据和改进机制。", 15, "6E747A"))

    headers = ["方面", "核心信息", "证据/样本", "风险或不足", "优化方向"]
    widths = [120, 220, 250, 250, 300]
    x0, y0 = 50, 136
    x = x0
    for header, width in zip(headers, widths):
        _table_cell(builder, header, x, y0, width, 38, "4F7DB9", "FFFFFF", True, 16)
        x += width

    rows = _matrix_rows(texts)
    for row_index, row in enumerate(rows):
        y = y0 + 38 + row_index * 74
        x = x0
        for col_index, (value, width) in enumerate(zip(row, widths)):
            fill = "EDF3FC" if col_index == 0 else "FFFFFF"
            color = "D60000" if col_index == 3 else "1E1E1E"
            _table_cell(builder, value, x, y, width, 74, fill, color, col_index in (0, 3), 13)
            x += width

    builder.text(TextBox(58, 558, 1060, 24, _conclusion(texts, "科研型 PPT 应把信息组织成可验证的问题、证据和方案链路。"), 18, "D60000", True))
    return builder.xml()


def _summary_slide(texts: list[str], slide_number: int | None) -> bytes:
    title_red, title_blue = _title_parts(texts, "研究结论", "从文字堆叠转向结构化科研叙事")
    builder = _base_builder(title_red, title_blue, slide_number)

    builder.text(TextBox(52, 108, 500, 28, "核心判断", 22, "D60000", True))
    bullets = _bullets(texts, 5)
    for index, bullet in enumerate(bullets):
        y = 158 + index * 54
        builder.shape(ShapeBox(62, y + 8, 10, 10, "FFFFFF", "243A8F"))
        builder.text(TextBox(86, y, 510, 36, bullet, 17))

    loop = [
        ("研究问题", "定义目标"),
        ("技术路线", "形成方法"),
        ("实验验证", "沉淀证据"),
        ("场景落地", "闭环迭代"),
    ]
    for index, (label, body) in enumerate(loop):
        x = 660 + (index % 2) * 250
        y = 150 + (index // 2) * 168
        color = "D60000" if index == 3 else "4F7DB9"
        _section(builder, label, x, y, 160, color)
        builder.text(TextBox(x, y + 42, 160, 82, body, 20, color, True, "center", line=color))
    _arrow(builder, 820, 205, 80)
    _arrow(builder, 820, 373, 80)
    builder.shape(ShapeBox(736, 274, 3, 58, "243A8F"))
    builder.shape(ShapeBox(986, 274, 3, 58, "D60000"))

    builder.shape(ShapeBox(72, 526, 1060, 58, "FFF7F7", "D60000"))
    builder.text(TextBox(96, 544, 70, 18, "结论", 16, "D60000", True))
    builder.text(TextBox(176, 540, 830, 26, _conclusion(texts, "该页已重排为红蓝标题、结构化要点和科研汇报结论栏。"), 18))
    return builder.xml()


def _base_builder(title_red: str, title_blue: str, slide_number: int | None) -> SlideBuilder:
    builder = SlideBuilder()
    builder.shape(ShapeBox(0, 0, SLIDE_W, SLIDE_H, "FFFFFF", None))
    builder.text(TextBox(38, 22, 210, 38, title_red, 28, "D60000", True, font="SimHei"))
    builder.text(TextBox(252, 24, 910, 42, f"—{title_blue}", 24, "243A8F"))
    builder.shape(ShapeBox(38, 74, 1188, 2, "243A8F", None))
    builder.shape(ShapeBox(36, 686, 1188, 1, "8FA6C7", None))
    builder.text(TextBox(38, 694, 520, 14, "Research-style redesigned by PPT Optimizer", 10, "6E747A"))
    builder.text(TextBox(1192, 694, 34, 14, str(slide_number or 1).zfill(2), 10, "6E747A", align="right"))
    return builder


def _section(builder: SlideBuilder, label: str, x: int, y: int, w: int, color: str) -> None:
    builder.shape(ShapeBox(x, y, w, 26, color, None))
    builder.text(TextBox(x + 6, y + 4, w - 12, 16, label, 13, "FFFFFF", True, "center", fill=color))


def _arrow(builder: SlideBuilder, x: int, y: int, w: int, color: str = "4F7DB9") -> None:
    builder.shape(ShapeBox(x, y + 8, w, 3, color, None))
    builder.shape(ShapeBox(x + w - 2, y, 18, 18, color, None, "rtTriangle"))


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
) -> None:
    builder.text(TextBox(x, y, w, h, text, size, color, bold, "center", fill, "B8C6DB"))


def _title_parts(texts: list[str], fallback_red: str, fallback_blue: str) -> tuple[str, str]:
    title = _best_title(texts)
    if not title:
        return fallback_red, fallback_blue
    for marker in ("-", "—", "–", ":"):
        if marker in title:
            left, right = title.split(marker, 1)
            left = _clean_title(left) or fallback_red
            right = _clean_title(right) or fallback_blue
            return _short(left, 8), _short(right, 28)
    if len(title) <= 12:
        return _short(title, 8), fallback_blue
    return fallback_red, _short(title, 30)


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
    return f"""          <a:p>
            <a:pPr algn="{align}"/>
            <a:r>
              <a:rPr lang="zh-CN" sz="{box.size * 100}"{bold}>
                <a:solidFill><a:srgbClr val="{box.color}"/></a:solidFill>
                <a:latin typeface="{_xml(box.font)}"/><a:ea typeface="{_xml(box.font)}"/><a:cs typeface="{_xml(box.font)}"/>
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


def _fill_xml(color: str | None) -> str:
    if not color or color.lower() == "none":
        return "<a:noFill/>"
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'


def _line_xml(color: str | None) -> str:
    if not color:
        return "<a:ln><a:noFill/></a:ln>"
    return f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>'


def _emu(value: int | float) -> int:
    return int(round(value * EMU_PER_PX))


def _xml(value: str) -> str:
    return escape(value, {'"': "&quot;"})
