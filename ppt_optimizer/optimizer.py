from __future__ import annotations

import io
import posixpath
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from PIL import Image
except ImportError:  # pragma: no cover - optional dependency
    Image = None


A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

ET.register_namespace("a", A_NS)
ET.register_namespace("p", P_NS)
ET.register_namespace("r", R_NS)


SLIDE_RE = re.compile(r"ppt/slides/slide\d+\.xml$")
NOTE_SLIDE_RE = re.compile(r"ppt/notesSlides/notesSlide\d+\.xml$")
NOTE_RELATED_RE = re.compile(r"ppt/notesSlides/(?:notesSlide\d+\.xml|_rels/notesSlide\d+\.xml\.rels)$")
COMMENT_RE = re.compile(r"ppt/comments/comment\d+\.xml$")
COMMENT_RELATED_RE = re.compile(r"ppt/comments/(?:comment\d+\.xml|_rels/comment\d+\.xml\.rels)$")
COMMENT_AUTHORS = "ppt/commentAuthors.xml"
MEDIA_RE = re.compile(r"ppt/media/.+\.(?:jpe?g|png)$", re.IGNORECASE)
CONTENT_TYPES = "[Content_Types].xml"


@dataclass(frozen=True)
class OptimizationOptions:
    font_family: str | None = "Microsoft YaHei"
    remove_notes: bool = True
    report_only: bool = False
    image_quality: int = 82
    max_image_width: int | None = 2560
    slide_numbers: tuple[int, ...] | None = None


@dataclass
class DeckReport:
    source: Path
    output: Path | None
    slides: int = 0
    text_runs: int = 0
    font_runs_changed: int = 0
    notes_removed: int = 0
    comments_removed: int = 0
    images_seen: int = 0
    images_optimized: int = 0
    bytes_saved: int = 0
    target_slides: tuple[int, ...] | None = None
    warnings: list[str] = field(default_factory=list)

    def to_text(self) -> str:
        lines = [
            f"Source: {self.source}",
            f"Slides: {self.slides}",
            f"Text runs: {self.text_runs}",
            f"Font runs changed: {self.font_runs_changed}",
            f"Notes removed: {self.notes_removed}",
            f"Comments removed: {self.comments_removed}",
            f"Images seen: {self.images_seen}",
            f"Images optimized: {self.images_optimized}",
            f"Bytes saved: {self.bytes_saved}",
        ]
        if self.target_slides:
            lines.insert(2, f"Target slides: {', '.join(str(n) for n in self.target_slides)}")
        if self.output:
            lines.append(f"Output: {self.output}")
        if self.warnings:
            lines.append("Warnings:")
            lines.extend(f"- {warning}" for warning in self.warnings)
        return "\n".join(lines)


def optimize_pptx(
    source: Path,
    output: Path | None,
    options: OptimizationOptions | None = None,
) -> DeckReport:
    options = options or OptimizationOptions()
    source = source.resolve()
    if not source.exists():
        raise FileNotFoundError(f"input file not found: {source}")
    if source.suffix.lower() != ".pptx":
        raise ValueError("input must be a .pptx file")
    if not zipfile.is_zipfile(source):
        raise ValueError("input is not a valid .pptx zip package")
    if options.report_only:
        output = None
    elif output is None:
        output = source.with_name(f"{source.stem}.optimized{source.suffix}")
    else:
        output = output.resolve()

    report = DeckReport(source=source, output=output)

    with zipfile.ZipFile(source, "r") as zin:
        names = zin.namelist()
        report.slides = sum(1 for name in names if SLIDE_RE.match(name))
        report.target_slides = _normalize_target_slides(options.slide_numbers, report.slides)
        related_parts = _collect_related_parts(zin, report.target_slides)

        if options.report_only:
            _analyze_only(zin, names, report, related_parts)
            return report

        assert output is not None
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            skipped_notes = set()
            for item in zin.infolist():
                name = item.filename
                data = zin.read(name)

                if options.remove_notes and _should_remove_note_part(name, related_parts):
                    if NOTE_SLIDE_RE.match(name):
                        report.notes_removed += 1
                    skipped_notes.add(name)
                    continue
                if options.remove_notes and _should_remove_comment_part(name, related_parts):
                    if COMMENT_RE.match(name) or name == COMMENT_AUTHORS:
                        report.comments_removed += 1
                    continue
                if name == CONTENT_TYPES:
                    data = _remove_content_type_overrides(data, options, related_parts)
                elif SLIDE_RE.match(name) and _is_target_slide_part(name, report.target_slides):
                    data = _optimize_slide_xml(data, options, report)
                elif MEDIA_RE.match(name) and _is_target_media_part(name, related_parts):
                    data = _optimize_image(data, name, options, report)
                elif name.endswith(".rels"):
                    data = _remove_note_relationships(data, name, options, related_parts)

                zout.writestr(_clone_zipinfo(item), data)

            _warn_about_relationships(skipped_notes, report)

    return report


def parse_slide_selection(value: str | None) -> tuple[int, ...] | None:
    if value is None or not value.strip():
        return None
    selected: set[int] = set()
    for chunk in value.split(","):
        part = chunk.strip()
        if not part:
            continue
        if "-" in part:
            start_text, end_text = part.split("-", 1)
            start = int(start_text.strip())
            end = int(end_text.strip())
            if start <= 0 or end <= 0 or end < start:
                raise ValueError(f"invalid slide range: {part}")
            selected.update(range(start, end + 1))
            continue
        number = int(part)
        if number <= 0:
            raise ValueError(f"invalid slide number: {part}")
        selected.add(number)
    return tuple(sorted(selected)) or None


def _analyze_only(
    zin: zipfile.ZipFile,
    names: list[str],
    report: DeckReport,
    related_parts: dict[str, set[str]],
) -> None:
    for name in names:
        if SLIDE_RE.match(name) and _is_target_slide_part(name, report.target_slides):
            data = zin.read(name)
            root = ET.fromstring(data)
            report.text_runs += len(root.findall(f".//{{{A_NS}}}r"))
        elif MEDIA_RE.match(name) and _is_target_media_part(name, related_parts):
            report.images_seen += 1
        elif _should_remove_note_part(name, related_parts) and NOTE_SLIDE_RE.match(name):
            report.notes_removed += 1
        elif _should_remove_comment_part(name, related_parts) and (COMMENT_RE.match(name) or name == COMMENT_AUTHORS):
            report.comments_removed += 1


def _normalize_target_slides(slide_numbers: tuple[int, ...] | None, slide_count: int) -> tuple[int, ...] | None:
    if slide_numbers is None:
        return None
    invalid = [number for number in slide_numbers if number > slide_count]
    if invalid:
        raise ValueError(
            f"slide selection out of range: {', '.join(str(n) for n in invalid)} "
            f"(deck has {slide_count} slides)"
        )
    return slide_numbers


def _slide_number_from_name(name: str) -> int | None:
    match = re.match(r"ppt/slides/slide(\d+)\.xml$", name)
    if not match:
        return None
    return int(match.group(1))


def _is_target_slide_part(name: str, target_slides: tuple[int, ...] | None) -> bool:
    if target_slides is None:
        return True
    slide_number = _slide_number_from_name(name)
    return slide_number in target_slides


def _is_target_media_part(name: str, related_parts: dict[str, set[str]]) -> bool:
    media_parts = related_parts["media"]
    if related_parts["targeted"]:
        return name in media_parts
    return True


def _collect_related_parts(
    zin: zipfile.ZipFile,
    target_slides: tuple[int, ...] | None,
) -> dict[str, set[str]]:
    related = {
        "targeted": set(),
        "notes": set(),
        "comments": set(),
        "media": set(),
        "slide_rels": set(),
        "content_type_removals": set(),
    }
    if target_slides is None:
        return related
    related["targeted"].add("yes")

    for slide_number in target_slides:
        rel_name = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
        if rel_name not in zin.namelist():
            continue
        related["slide_rels"].add(rel_name)
        try:
            root = ET.fromstring(zin.read(rel_name))
        except ET.ParseError:
            continue
        for rel in root:
            rel_type = rel.get("Type", "")
            target = rel.get("Target", "")
            if not target:
                continue
            part = _resolve_relationship_target(f"ppt/slides/slide{slide_number}.xml", target)
            if rel_type.endswith("/image") and MEDIA_RE.match(part):
                related["media"].add(part)
            elif rel_type.endswith("/notesSlide"):
                _add_related_xml_part(related["notes"], part)
                related["content_type_removals"].add(part)
            elif rel_type.endswith("/comments"):
                _add_related_xml_part(related["comments"], part)
                related["content_type_removals"].add(part)
    return related


def _add_related_xml_part(parts: set[str], part: str) -> None:
    parts.add(part)
    directory, filename = posixpath.split(part)
    if filename:
        parts.add(posixpath.join(directory, "_rels", f"{filename}.rels"))


def _resolve_relationship_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def _should_remove_note_part(name: str, related_parts: dict[str, set[str]]) -> bool:
    note_parts = related_parts["notes"]
    if related_parts["targeted"]:
        return name in note_parts
    return NOTE_RELATED_RE.match(name) is not None


def _should_remove_comment_part(name: str, related_parts: dict[str, set[str]]) -> bool:
    comment_parts = related_parts["comments"]
    if related_parts["targeted"]:
        return name in comment_parts
    return COMMENT_RELATED_RE.match(name) is not None or name == COMMENT_AUTHORS


def _optimize_slide_xml(
    data: bytes,
    options: OptimizationOptions,
    report: DeckReport,
) -> bytes:
    root = ET.fromstring(data)
    for run in root.findall(f".//{{{A_NS}}}r"):
        report.text_runs += 1
        if options.font_family:
            if _set_run_font(run, options.font_family):
                report.font_runs_changed += 1
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _set_run_font(run: ET.Element, font_family: str) -> bool:
    changed = False
    r_pr = run.find(f"{{{A_NS}}}rPr")
    if r_pr is None:
        r_pr = ET.Element(f"{{{A_NS}}}rPr")
        run.insert(0, r_pr)
        changed = True

    for tag in ("latin", "ea", "cs"):
        child = r_pr.find(f"{{{A_NS}}}{tag}")
        if child is None:
            child = ET.SubElement(r_pr, f"{{{A_NS}}}{tag}")
            changed = True
        if child.get("typeface") != font_family:
            child.set("typeface", font_family)
            changed = True
    return changed


def _optimize_image(
    data: bytes,
    name: str,
    options: OptimizationOptions,
    report: DeckReport,
) -> bytes:
    report.images_seen += 1
    if Image is None:
        if "Pillow is not installed" not in report.warnings:
            report.warnings.append("Pillow is not installed; images were not recompressed.")
        return data

    try:
        image = Image.open(io.BytesIO(data))
    except Exception:
        report.warnings.append(f"Skipped unreadable image: {name}")
        return data

    original_size = len(data)
    image = _downscale_image(image, options.max_image_width)

    suffix = Path(name).suffix.lower()
    out = io.BytesIO()
    save_kwargs = {"optimize": True}
    if suffix in (".jpg", ".jpeg"):
        if image.mode not in ("RGB", "L"):
            image = image.convert("RGB")
        image.save(out, format="JPEG", quality=options.image_quality, **save_kwargs)
    elif suffix == ".png":
        image.save(out, format="PNG", **save_kwargs)
    else:
        return data

    optimized = out.getvalue()
    if len(optimized) >= original_size:
        return data

    report.images_optimized += 1
    report.bytes_saved += original_size - len(optimized)
    return optimized


def _downscale_image(image: "Image.Image", max_width: int | None) -> "Image.Image":
    if not max_width or image.width <= max_width:
        return image
    ratio = max_width / image.width
    height = max(1, round(image.height * ratio))
    return image.resize((max_width, height), Image.Resampling.LANCZOS)


def _remove_note_relationships(
    data: bytes,
    rel_part_name: str,
    options: OptimizationOptions,
    related_parts: dict[str, set[str]],
) -> bytes:
    if not options.remove_notes:
        return data
    if related_parts["slide_rels"] and rel_part_name not in related_parts["slide_rels"]:
        return data
    root = ET.fromstring(data)
    changed = False
    for rel in list(root):
        rel_type = rel.get("Type", "")
        target = rel.get("Target", "")
        target_part = _resolve_relationship_target(_source_part_from_rels(rel_part_name), target) if target else ""
        if (
            rel_type.endswith("/notesSlide")
            and (not related_parts["notes"] or target_part in related_parts["notes"])
        ) or (
            rel_type.endswith("/comments")
            and (not related_parts["comments"] or target_part in related_parts["comments"])
        ):
            root.remove(rel)
            changed = True
    if not changed:
        return data
    ET.register_namespace("", REL_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _source_part_from_rels(rel_part_name: str) -> str:
    if "/_rels/" not in rel_part_name or not rel_part_name.endswith(".rels"):
        return rel_part_name
    directory, rel_file = rel_part_name.split("/_rels/", 1)
    return posixpath.join(directory, rel_file[:-5])


def _remove_content_type_overrides(
    data: bytes,
    options: OptimizationOptions,
    related_parts: dict[str, set[str]],
) -> bytes:
    if not options.remove_notes:
        return data
    root = ET.fromstring(data)
    changed = False
    for child in list(root):
        part_name = child.get("PartName", "")
        normalized = part_name.lstrip("/")
        should_remove = (
            normalized in related_parts["content_type_removals"]
            if related_parts["targeted"]
            else part_name.startswith("/ppt/notesSlides/")
            or part_name.startswith("/ppt/comments/")
            or part_name == "/ppt/commentAuthors.xml"
        )
        if should_remove:
            root.remove(child)
            changed = True
    if not changed:
        return data
    ET.register_namespace("", CT_NS)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _warn_about_relationships(skipped_notes: set[str], report: DeckReport) -> None:
    if skipped_notes:
        report.warnings.append(
            "Speaker note slide parts were removed and note relationships were cleaned."
        )


def _clone_zipinfo(item: zipfile.ZipInfo) -> zipfile.ZipInfo:
    cloned = zipfile.ZipInfo(item.filename, item.date_time)
    cloned.comment = item.comment
    cloned.compress_type = zipfile.ZIP_DEFLATED
    cloned.extra = item.extra
    cloned.internal_attr = item.internal_attr
    cloned.external_attr = item.external_attr
    cloned.create_system = item.create_system
    return cloned
