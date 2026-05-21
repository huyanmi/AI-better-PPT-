# PPT Optimizer

A small local CLI for analyzing and optimizing `.pptx` files.

It works directly with the PowerPoint Open XML package, so it does not require
PowerPoint or `python-pptx`. If Pillow is installed, it can also recompress large
JPEG/PNG images.

## Features

- Generate a slide-by-slide report.
- Normalize fonts across text runs.
- Remove speaker notes and comment authors.
- Recompress oversized images when Pillow is available.
- Write a new optimized copy without changing the original file.

## Usage

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --font "Microsoft YaHei"
```

Analyze without writing a new file:

```powershell
python -m ppt_optimizer input.pptx --report-only
```

More aggressive image compression:

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --image-quality 72 --max-image-width 1920
```

Optimize only selected slides:

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --slides 3
python -m ppt_optimizer input.pptx -o optimized.pptx --slides 2,5-7
```

## Notes

The optimizer is intentionally conservative. It keeps layout, animations, charts,
and slide masters intact while changing only scoped XML/package parts.
