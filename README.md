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
- 支持指定页“科研风重绘”：抽取原页文字后，重排为红蓝标题、架构图、
  证据矩阵、结论条等科研汇报版式。
- 支持传入参考 PPT 学习配色和字体，让重绘页更贴近目标答辩/科研汇报风格。

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

把某一页重绘为科研汇报风：

```powershell
python -m ppt_optimizer input.pptx -o research-style.pptx --slides 3 --research-style
```

参考另一份 PPT 的配色和字体再重绘：

```powershell
python -m ppt_optimizer input.pptx -o research-style.pptx --slides 3 --research-style --reference-ppt "C:\Users\22625\Desktop\大四\毕业设计\毕设答辩.pptx"
```

只优化多页：

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --slides 2,5-7
```

压缩图片：

```powershell
python -m ppt_optimizer input.pptx -o optimized.pptx --image-quality 72 --max-image-width 1920
```

### 图形界面 / EXE

也可以使用桌面图形界面，适合“上传 PPT -> 处理 -> 下载优化版”的流程：

```powershell
python ppt_optimizer_gui.py
```

打包为 Windows exe：

```powershell
python -m pip install pyinstaller
.\build_exe.ps1
```

打包完成后，程序位于：

```text
dist\PPTOptimizer.exe
```

### 说明

常规优化模式偏保守：它会尽量保留原 PPT 的布局、动画、图表和母版结构，只修改字体、
备注、评论、图片和相关 Open XML 包结构。

如果你想看到明显的科研风视觉变化，请使用 `--research-style --slides 页码`，或在 EXE
里勾选“科研风重绘”。这个模式会重绘指定页，适合先对单页做精确优化，再逐步扩展到整份
PPT。也可以同时选择“参考 PPT”，程序会从参考文件中学习主色、强调色和中文字体。

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
- Redesign selected slides into a Chinese research-report style with title bars,
  diagrams, evidence tables, and conclusion strips.
- Learn colors and fonts from an optional reference PPTX for the redesign mode.

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

Redesign one slide into the research-report style:

```powershell
python -m ppt_optimizer input.pptx -o research-style.pptx --slides 3 --research-style
```

Redesign with a reference deck:

```powershell
python -m ppt_optimizer input.pptx -o research-style.pptx --slides 3 --research-style --reference-ppt reference.pptx
```

Run the desktop UI:

```powershell
python ppt_optimizer_gui.py
```

Build a Windows executable:

```powershell
python -m pip install pyinstaller
.\build_exe.ps1
```

## Notes

The optimizer is intentionally conservative. It keeps layout, animations, charts,
and slide masters intact while changing only scoped XML/package parts.

Use `--research-style` with `--slides` when you want a visible one-slide layout
redesign instead of a conservative cleanup. Add `--reference-ppt` when you want
the redesign to inherit a reference deck's fonts and palette.
