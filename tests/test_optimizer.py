from __future__ import annotations

import zipfile
from pathlib import Path

from ppt_optimizer.optimizer import OptimizationOptions, optimize_pptx, parse_slide_selection


CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
</Types>
"""


SLIDE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:r><a:rPr><a:latin typeface="Arial"/></a:rPr><a:t>Hello</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>
"""


RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
</Relationships>
"""


RELS_XML_2 = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide2.xml"/>
</Relationships>
"""


def make_minimal_pptx(path: Path) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        zf.writestr("ppt/slides/slide1.xml", SLIDE_XML)
        zf.writestr("ppt/slides/_rels/slide1.xml.rels", RELS_XML)
        zf.writestr("ppt/notesSlides/notesSlide1.xml", "<notes/>")
        zf.writestr("ppt/notesSlides/_rels/notesSlide1.xml.rels", "<Relationships/>")


def make_two_slide_pptx(path: Path) -> None:
    content_types = CONTENT_TYPES_XML.replace(
        "</Types>",
        """
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/notesSlides/notesSlide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
</Types>
""",
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("ppt/slides/slide1.xml", SLIDE_XML)
        zf.writestr("ppt/slides/slide2.xml", SLIDE_XML)
        zf.writestr("ppt/slides/_rels/slide1.xml.rels", RELS_XML)
        zf.writestr("ppt/slides/_rels/slide2.xml.rels", RELS_XML_2)
        zf.writestr("ppt/notesSlides/notesSlide1.xml", "<notes/>")
        zf.writestr("ppt/notesSlides/notesSlide2.xml", "<notes/>")
        zf.writestr("ppt/notesSlides/_rels/notesSlide1.xml.rels", "<Relationships/>")
        zf.writestr("ppt/notesSlides/_rels/notesSlide2.xml.rels", "<Relationships/>")


def test_optimize_sets_font_and_removes_notes(tmp_path: Path) -> None:
    source = tmp_path / "deck.pptx"
    output = tmp_path / "deck.optimized.pptx"
    make_minimal_pptx(source)

    report = optimize_pptx(
        source,
        output,
        OptimizationOptions(font_family="Aptos", remove_notes=True),
    )

    assert report.slides == 1
    assert report.text_runs == 1
    assert report.font_runs_changed == 1
    assert report.notes_removed == 1
    assert output.exists()

    with zipfile.ZipFile(output) as zf:
        assert "ppt/notesSlides/notesSlide1.xml" not in zf.namelist()
        assert "ppt/notesSlides/_rels/notesSlide1.xml.rels" not in zf.namelist()
        content_types = zf.read("[Content_Types].xml").decode("utf-8")
        slide = zf.read("ppt/slides/slide1.xml").decode("utf-8")
        rels = zf.read("ppt/slides/_rels/slide1.xml.rels").decode("utf-8")

    assert 'typeface="Aptos"' in slide
    assert "notesSlides" not in content_types
    assert "notesSlide" not in rels


def test_report_only_does_not_write(tmp_path: Path) -> None:
    source = tmp_path / "deck.pptx"
    output = tmp_path / "should_not_exist.pptx"
    make_minimal_pptx(source)

    report = optimize_pptx(
        source,
        output,
        OptimizationOptions(report_only=True),
    )

    assert report.output is None
    assert report.slides == 1
    assert report.text_runs == 1
    assert not output.exists()


def test_targeted_slide_optimization_keeps_other_slides(tmp_path: Path) -> None:
    source = tmp_path / "deck.pptx"
    output = tmp_path / "deck.optimized.pptx"
    make_two_slide_pptx(source)

    report = optimize_pptx(
        source,
        output,
        OptimizationOptions(font_family="Aptos", remove_notes=True, slide_numbers=(2,)),
    )

    assert report.slides == 2
    assert report.target_slides == (2,)
    assert report.text_runs == 1
    assert report.font_runs_changed == 1
    assert report.notes_removed == 1

    with zipfile.ZipFile(output) as zf:
        names = zf.namelist()
        slide1 = zf.read("ppt/slides/slide1.xml").decode("utf-8")
        slide2 = zf.read("ppt/slides/slide2.xml").decode("utf-8")
        rels1 = zf.read("ppt/slides/_rels/slide1.xml.rels").decode("utf-8")
        rels2 = zf.read("ppt/slides/_rels/slide2.xml.rels").decode("utf-8")
        content_types = zf.read("[Content_Types].xml").decode("utf-8")

    assert 'typeface="Arial"' in slide1
    assert 'typeface="Aptos"' in slide2
    assert "ppt/notesSlides/notesSlide1.xml" in names
    assert "ppt/notesSlides/notesSlide2.xml" not in names
    assert "notesSlide1.xml" in rels1
    assert "notesSlide2.xml" not in rels2
    assert "/ppt/notesSlides/notesSlide1.xml" in content_types
    assert "/ppt/notesSlides/notesSlide2.xml" not in content_types


def test_targeted_slide_without_notes_does_not_remove_other_notes(tmp_path: Path) -> None:
    source = tmp_path / "deck.pptx"
    output = tmp_path / "deck.optimized.pptx"
    content_types = CONTENT_TYPES_XML.replace(
        "</Types>",
        """
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>
""",
    )
    with zipfile.ZipFile(source, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("ppt/slides/slide1.xml", SLIDE_XML)
        zf.writestr("ppt/slides/slide2.xml", SLIDE_XML)
        zf.writestr("ppt/slides/_rels/slide1.xml.rels", RELS_XML)
        zf.writestr("ppt/notesSlides/notesSlide1.xml", "<notes/>")
        zf.writestr("ppt/notesSlides/_rels/notesSlide1.xml.rels", "<Relationships/>")

    report = optimize_pptx(
        source,
        output,
        OptimizationOptions(font_family="Aptos", remove_notes=True, slide_numbers=(2,)),
    )

    assert report.notes_removed == 0
    with zipfile.ZipFile(output) as zf:
        names = zf.namelist()
        slide1 = zf.read("ppt/slides/slide1.xml").decode("utf-8")
        slide2 = zf.read("ppt/slides/slide2.xml").decode("utf-8")
        content_types = zf.read("[Content_Types].xml").decode("utf-8")

    assert 'typeface="Arial"' in slide1
    assert 'typeface="Aptos"' in slide2
    assert "ppt/notesSlides/notesSlide1.xml" in names
    assert "/ppt/notesSlides/notesSlide1.xml" in content_types


def test_research_style_redesigns_only_selected_slide(tmp_path: Path) -> None:
    source = tmp_path / "deck.pptx"
    output = tmp_path / "deck.research.pptx"
    make_two_slide_pptx(source)

    report = optimize_pptx(
        source,
        output,
        OptimizationOptions(remove_notes=True, slide_numbers=(2,), research_style=True),
    )

    assert report.slides == 2
    assert report.target_slides == (2,)
    assert report.text_runs == 1
    assert report.font_runs_changed == 0
    assert report.slides_redesigned == 1

    with zipfile.ZipFile(output) as zf:
        slide1 = zf.read("ppt/slides/slide1.xml").decode("utf-8")
        slide2 = zf.read("ppt/slides/slide2.xml").decode("utf-8")

    assert 'typeface="Arial"' in slide1
    assert "Research-style redesigned by PPT Optimizer" in slide2
    assert 'val="D60000"' in slide2
    assert 'val="243A8F"' in slide2


def test_parse_slide_selection() -> None:
    assert parse_slide_selection(None) is None
    assert parse_slide_selection("3") == (3,)
    assert parse_slide_selection("2,4-6") == (2, 4, 5, 6)
