# PPT Optimizer

## 中文说明

这是一个本地运行的 PPT 优化命令行工具，用于分析和优化 `.pptx`
文件。它直接处理 PowerPoint 的 Open XML 包结构，不依赖 PowerPoint
或 `python-pptx`。如果本机安装了 Pillow，还可以对 PPT 内的大图片进行压缩。

### 功能

- 生成 PPT 分析报告。
- 统一指定页面或整份 PPT 的字体。
- 删除演讲者备注和评论相关信息。
- 压缩或缩放 PPT 内过大的图片。
- 输出新的优化副本，不会覆盖原文件。
- 支持只优化指定页，例如第 3 页或第 2、5-7 页。

### 使用方法

先进入项目目录：

```powershell
cd C:\Users\22625\Documents\Playground
```

优化整份 PPT：

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --font "Microsoft YaHei"
```

只分析，不导出新文件：

```powershell
python -m ppt_optimizer input.pptx --report-only
```

只优化某一页：

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --slides 3
```

只优化多页：

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --slides 2,5-7
```

压缩图片：

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --image-quality 72 --max-image-width 1920
```

### 说明

当前优化器偏保守：它会尽量保留原 PPT 的布局、动画、图表和母版结构，只修改字体、
备注、评论、图片和相关 Open XML 包结构。适合先做单页定向优化，再逐步扩展到更复杂的
视觉重排能力。

## English

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
