from __future__ import annotations

import re
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from statistics import median
from xml.etree import ElementTree as ET


A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
SLIDE_RE = re.compile(r"ppt/slides/slide\d+\.xml$")


@dataclass(frozen=True)
class ReferenceStyle:
    primary: str = "243A8F"
    secondary: str = "4F7DB9"
    accent: str = "D60000"
    pale: str = "F4F7FD"
    border: str = "B8C6DB"
    body: str = "1E1E1E"
    muted: str = "6E747A"
    title_font: str = "SimHei"
    body_font: str = "Microsoft YaHei"
    title_size: int = 26
    body_size: int = 16
    source: str | None = None


DEFAULT_STYLE = ReferenceStyle()


def analyze_reference_style(path: Path) -> ReferenceStyle:
    source = path.expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"reference PPT not found: {source}")
    if source.suffix.lower() != ".pptx" or not zipfile.is_zipfile(source):
        raise ValueError("reference PPT must be a valid .pptx file")

    fonts: Counter[str] = Counter()
    colors: Counter[str] = Counter()
    sizes: list[float] = []

    with zipfile.ZipFile(source, "r") as zf:
        for name in zf.namelist():
            if not SLIDE_RE.match(name):
                continue
            try:
                root = ET.fromstring(zf.read(name))
            except ET.ParseError:
                continue
            _collect_fonts(root, fonts)
            _collect_colors(root, colors)
            _collect_sizes(root, sizes)

    primary = _best_color(colors, _is_blue, DEFAULT_STYLE.primary)
    secondary = _best_secondary_blue(colors, primary)
    accent = _best_color(colors, _is_red, DEFAULT_STYLE.accent)
    pale = _best_color(colors, _is_pale, DEFAULT_STYLE.pale)
    border = _lighten(primary, 0.55)
    title_font, body_font = _choose_fonts(fonts)
    title_size, body_size = _choose_sizes(sizes)

    return ReferenceStyle(
        primary=primary,
        secondary=secondary,
        accent=accent,
        pale=pale,
        border=border,
        title_font=title_font,
        body_font=body_font,
        title_size=title_size,
        body_size=body_size,
        source=str(source),
    )


def summarize_style(style: ReferenceStyle) -> str:
    source = f", source={style.source}" if style.source else ""
    return (
        f"primary=#{style.primary}, accent=#{style.accent}, "
        f"title_font={style.title_font}, body_font={style.body_font}{source}"
    )


def _collect_fonts(root: ET.Element, fonts: Counter[str]) -> None:
    for r_pr in root.findall(f".//{{{A_NS}}}rPr"):
        for tag in ("latin", "ea", "cs"):
            node = r_pr.find(f"{{{A_NS}}}{tag}")
            font = (node.get("typeface") if node is not None else "") or ""
            font = font.strip()
            if font and not font.startswith("+"):
                fonts[font] += 1


def _collect_colors(root: ET.Element, colors: Counter[str]) -> None:
    for node in root.findall(f".//{{{A_NS}}}srgbClr"):
        color = (node.get("val") or "").upper()
        if re.fullmatch(r"[0-9A-F]{6}", color):
            colors[color] += 1


def _collect_sizes(root: ET.Element, sizes: list[float]) -> None:
    for r_pr in root.findall(f".//{{{A_NS}}}rPr"):
        raw = r_pr.get("sz")
        if raw and raw.isdigit():
            size = int(raw) / 100
            if 8 <= size <= 80:
                sizes.append(size)


def _choose_fonts(fonts: Counter[str]) -> tuple[str, str]:
    if not fonts:
        return DEFAULT_STYLE.title_font, DEFAULT_STYLE.body_font

    ranked = [name for name, _count in fonts.most_common()]
    title_font = _first_font(
        ranked,
        ("黑体", "SimHei", "Microsoft YaHei", "微软雅黑", "DengXian", "等线", "Aptos"),
        DEFAULT_STYLE.title_font,
    )
    body_font = _first_font(
        ranked,
        ("Microsoft YaHei", "微软雅黑", "DengXian", "等线", "SimSun", "宋体", title_font),
        DEFAULT_STYLE.body_font,
    )
    return title_font, body_font


def _first_font(ranked: list[str], preferred: tuple[str, ...], fallback: str) -> str:
    lookup = {font.lower(): font for font in ranked}
    for font in preferred:
        if font.lower() in lookup:
            return lookup[font.lower()]
    for font in ranked:
        if not re.search(r"times|impact|arial", font, re.I):
            return font
    return fallback


def _choose_sizes(sizes: list[float]) -> tuple[int, int]:
    if not sizes:
        return DEFAULT_STYLE.title_size, DEFAULT_STYLE.body_size
    body = int(round(median(sizes)))
    body = max(14, min(body, 18))
    title = max(24, min(body + 8, 30))
    return title, body


def _best_color(colors: Counter[str], predicate, fallback: str) -> str:
    candidates = [(color, count) for color, count in colors.items() if predicate(color)]
    if not candidates:
        return fallback
    candidates.sort(key=lambda item: item[1], reverse=True)
    return candidates[0][0]


def _best_secondary_blue(colors: Counter[str], primary: str) -> str:
    candidates = [
        (color, count)
        for color, count in colors.items()
        if color != primary and _is_blue(color) and _distance(color, primary) > 24
    ]
    if candidates:
        candidates.sort(key=lambda item: item[1], reverse=True)
        return candidates[0][0]
    return _lighten(primary, 0.28)


def _rgb(color: str) -> tuple[int, int, int]:
    return int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)


def _is_blue(color: str) -> bool:
    r, g, b = _rgb(color)
    return b > r + 25 and b > g - 5 and b >= 95


def _is_red(color: str) -> bool:
    r, g, b = _rgb(color)
    return r > g + 45 and r > b + 45 and r >= 130


def _is_pale(color: str) -> bool:
    r, g, b = _rgb(color)
    return min(r, g, b) >= 215 and max(r, g, b) - min(r, g, b) <= 45 and color != "FFFFFF"


def _lighten(color: str, amount: float) -> str:
    r, g, b = _rgb(color)
    values = [round(channel + (255 - channel) * amount) for channel in (r, g, b)]
    return "".join(f"{value:02X}" for value in values)


def _distance(left: str, right: str) -> float:
    lr, lg, lb = _rgb(left)
    rr, rg, rb = _rgb(right)
    return ((lr - rr) ** 2 + (lg - rg) ** 2 + (lb - rb) ** 2) ** 0.5
