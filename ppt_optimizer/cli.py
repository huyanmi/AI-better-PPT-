from __future__ import annotations

import argparse
from pathlib import Path

from .optimizer import OptimizationOptions, optimize_pptx, parse_slide_selection


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ppt_optimizer",
        description="Analyze and optimize PowerPoint .pptx files.",
    )
    parser.add_argument("input", type=Path, help="Path to the source .pptx file.")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Path for the optimized .pptx. Defaults to <name>.optimized.pptx.",
    )
    parser.add_argument(
        "--font",
        default="Microsoft YaHei",
        help="Font family to apply to text runs. Use an empty value to skip.",
    )
    parser.add_argument(
        "--report-only",
        action="store_true",
        help="Analyze the deck without writing an optimized copy.",
    )
    parser.add_argument(
        "--slides",
        help="Only optimize selected 1-based slide numbers, e.g. 3 or 2,5 or 3-6.",
    )
    parser.add_argument(
        "--research-style",
        action="store_true",
        help=(
            "Redesign selected slides into a Chinese research-report style "
            "with title bars, diagrams, evidence tables, and conclusion strips. "
            "Use together with --slides."
        ),
    )
    parser.add_argument(
        "--reference-ppt",
        type=Path,
        help="Optional reference .pptx used to learn fonts and colors for --research-style.",
    )
    parser.add_argument(
        "--keep-notes",
        action="store_true",
        help="Keep speaker notes instead of removing them.",
    )
    parser.add_argument(
        "--image-quality",
        type=int,
        default=82,
        help="JPEG/WebP quality for recompressed images, from 1 to 95.",
    )
    parser.add_argument(
        "--max-image-width",
        type=int,
        default=2560,
        help="Downscale images wider than this many pixels. Use 0 to disable.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    output = args.output
    if output is None and not args.report_only:
        output = args.input.with_name(f"{args.input.stem}.optimized{args.input.suffix}")

    try:
        if args.research_style and not args.slides:
            raise ValueError("--research-style must be used with --slides, e.g. --slides 3")
        options = OptimizationOptions(
            font_family=args.font or None,
            remove_notes=not args.keep_notes,
            report_only=args.report_only,
            image_quality=max(1, min(args.image_quality, 95)),
            max_image_width=args.max_image_width if args.max_image_width > 0 else None,
            slide_numbers=parse_slide_selection(args.slides),
            research_style=args.research_style,
            reference_ppt=args.reference_ppt,
        )
        result = optimize_pptx(args.input, output, options)
    except Exception as exc:
        parser.exit(1, f"error: {exc}\n")

    print(result.to_text())
    return 0
