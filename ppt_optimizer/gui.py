from __future__ import annotations

import os
import queue
import threading
from pathlib import Path
from tkinter import END, DISABLED, NORMAL, StringVar, Tk, filedialog, messagebox
from tkinter import ttk

from .optimizer import OptimizationOptions, optimize_pptx, parse_slide_selection


class PPTOptimizerApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("PPT Optimizer")
        self.root.geometry("820x620")
        self.root.minsize(760, 580)

        self.input_path = StringVar()
        self.output_path = StringVar()
        self.reference_path = StringVar()
        self.font_family = StringVar(value="Microsoft YaHei")
        self.slides = StringVar()
        self.image_quality = StringVar(value="82")
        self.max_image_width = StringVar(value="2560")
        self.status = StringVar(value="请选择一个 PPTX 文件开始。")
        self.keep_notes = StringVar(value="0")
        self.research_style = StringVar(value="0")
        self.result_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None

        self._build_ui()
        self.root.after(120, self._poll_worker)

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        title = ttk.Label(
            self.root,
            text="PPT 优化器",
            font=("Microsoft YaHei", 22, "bold"),
        )
        title.grid(row=0, column=0, sticky="w", padx=22, pady=(18, 4))

        subtitle = ttk.Label(
            self.root,
            text="上传 PPTX -> 选择页码 -> 可勾选科研风重绘 -> 处理完成后保存优化版",
            font=("Microsoft YaHei", 10),
        )
        subtitle.grid(row=1, column=0, sticky="w", padx=24, pady=(0, 12))

        main = ttk.Frame(self.root, padding=(22, 4, 22, 10))
        main.grid(row=2, column=0, sticky="nsew")
        main.columnconfigure(1, weight=1)
        main.rowconfigure(9, weight=1)

        self._row(main, 0, "上传 PPTX", self.input_path, self._choose_input, "选择文件")
        self._row(main, 1, "保存为", self.output_path, self._choose_output, "选择位置")
        self._row(main, 2, "参考 PPT", self.reference_path, self._choose_reference, "选择参考")

        ttk.Label(main, text="页码").grid(row=3, column=0, sticky="w", pady=8)
        ttk.Entry(main, textvariable=self.slides).grid(row=3, column=1, sticky="ew", pady=8)
        ttk.Label(main, text="留空表示整份 PPT；示例：3 或 2,5-7").grid(
            row=3,
            column=2,
            sticky="w",
            padx=(10, 0),
            pady=8,
        )

        ttk.Label(main, text="字体").grid(row=4, column=0, sticky="w", pady=8)
        ttk.Entry(main, textvariable=self.font_family).grid(row=4, column=1, sticky="ew", pady=8)
        ttk.Label(main, text="留空则不统一字体").grid(row=4, column=2, sticky="w", padx=(10, 0), pady=8)

        options = ttk.Frame(main)
        options.grid(row=5, column=1, columnspan=2, sticky="ew", pady=8)
        ttk.Label(options, text="图片质量").pack(side="left")
        ttk.Entry(options, textvariable=self.image_quality, width=8).pack(side="left", padx=(8, 18))
        ttk.Label(options, text="最大图片宽度").pack(side="left")
        ttk.Entry(options, textvariable=self.max_image_width, width=10).pack(side="left", padx=(8, 18))
        ttk.Checkbutton(
            options,
            text="保留备注",
            variable=self.keep_notes,
            onvalue="1",
            offvalue="0",
        ).pack(side="left")

        research = ttk.Frame(main)
        research.grid(row=6, column=1, columnspan=2, sticky="ew", pady=(2, 8))
        ttk.Checkbutton(
            research,
            text="科研风重绘（实验，建议只选 1 页）",
            variable=self.research_style,
            onvalue="1",
            offvalue="0",
        ).pack(side="left")
        ttk.Label(
            research,
            text="会重排指定页的标题、流程图、证据矩阵和结论条",
        ).pack(side="left", padx=(12, 0))

        buttons = ttk.Frame(main)
        buttons.grid(row=7, column=1, columnspan=2, sticky="ew", pady=(12, 8))
        self.run_button = ttk.Button(buttons, text="开始优化", command=self._start_optimize)
        self.run_button.pack(side="left")
        ttk.Button(buttons, text="打开输出目录", command=self._open_output_folder).pack(side="left", padx=10)
        ttk.Button(buttons, text="清空日志", command=self._clear_log).pack(side="left")

        ttk.Label(main, textvariable=self.status).grid(row=8, column=0, columnspan=3, sticky="w", pady=(8, 6))

        self.log = self._make_text(main)
        self.log.grid(row=9, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
        self._append_log("准备就绪。")

    def _row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        value: StringVar,
        command,
        button_text: str,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=8)
        ttk.Entry(parent, textvariable=value).grid(row=row, column=1, sticky="ew", pady=8)
        ttk.Button(parent, text=button_text, command=command).grid(
            row=row,
            column=2,
            sticky="ew",
            padx=(10, 0),
            pady=8,
        )

    def _make_text(self, parent: ttk.Frame):
        from tkinter import Text

        text = Text(parent, height=12, wrap="word", font=("Consolas", 10))
        text.configure(state=DISABLED)
        return text

    def _choose_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择要优化的 PPTX 文件",
            filetypes=[("PowerPoint", "*.pptx"), ("All files", "*.*")],
        )
        if not path:
            return
        self.input_path.set(path)
        source = Path(path)
        self.output_path.set(str(source.with_name(f"{source.stem}.optimized.pptx")))

    def _choose_output(self) -> None:
        initial = self.output_path.get() or "optimized.pptx"
        path = filedialog.asksaveasfilename(
            title="保存优化后的 PPTX",
            defaultextension=".pptx",
            initialfile=Path(initial).name,
            filetypes=[("PowerPoint", "*.pptx"), ("All files", "*.*")],
        )
        if path:
            self.output_path.set(path)

    def _choose_reference(self) -> None:
        path = filedialog.askopenfilename(
            title="选择用于学习风格的参考 PPTX",
            filetypes=[("PowerPoint", "*.pptx"), ("All files", "*.*")],
        )
        if path:
            self.reference_path.set(path)

    def _start_optimize(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("正在处理", "当前 PPT 还在处理中，请稍等。")
            return

        try:
            source = Path(self.input_path.get()).expanduser()
            output = Path(self.output_path.get()).expanduser()
            if not source.exists():
                raise ValueError("请选择存在的 PPTX 文件。")
            if source.suffix.lower() != ".pptx":
                raise ValueError("目前只支持 .pptx 文件。")
            if not output.name:
                raise ValueError("请选择保存位置。")
            if self.research_style.get() == "1" and not self.slides.get().strip():
                raise ValueError("科研风重绘需要填写页码，例如 3。这样不会误改整份 PPT。")
            reference = None
            if self.reference_path.get().strip():
                reference = Path(self.reference_path.get()).expanduser()
                if not reference.exists():
                    raise ValueError("参考 PPT 不存在，请重新选择。")
                if reference.suffix.lower() != ".pptx":
                    raise ValueError("参考 PPT 目前只支持 .pptx 文件。")
            quality = int(self.image_quality.get() or "82")
            max_width_raw = self.max_image_width.get().strip()
            max_width = int(max_width_raw) if max_width_raw else 0
            options = OptimizationOptions(
                font_family=self.font_family.get().strip() or None,
                remove_notes=self.keep_notes.get() != "1",
                image_quality=max(1, min(quality, 95)),
                max_image_width=max_width if max_width > 0 else None,
                slide_numbers=parse_slide_selection(self.slides.get()),
                research_style=self.research_style.get() == "1",
                reference_ppt=reference,
            )
        except Exception as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        mode = "科研风重绘" if options.research_style else "常规优化"
        self.status.set("正在处理，请稍等...")
        self.run_button.configure(state=DISABLED)
        self._append_log("")
        self._append_log(f"模式: {mode}")
        self._append_log(f"输入: {source}")
        if options.reference_ppt:
            self._append_log(f"参考: {options.reference_ppt}")
        self._append_log(f"输出: {output}")

        self.worker = threading.Thread(
            target=self._run_optimize,
            args=(source, output, options),
            daemon=True,
        )
        self.worker.start()

    def _run_optimize(self, source: Path, output: Path, options: OptimizationOptions) -> None:
        try:
            result = optimize_pptx(source, output, options)
        except Exception as exc:
            self.result_queue.put(("error", exc))
            return
        self.result_queue.put(("done", result))

    def _poll_worker(self) -> None:
        try:
            kind, payload = self.result_queue.get_nowait()
        except queue.Empty:
            self.root.after(120, self._poll_worker)
            return

        self.run_button.configure(state=NORMAL)
        if kind == "error":
            self.status.set("优化失败。")
            self._append_log(f"错误: {payload}")
            messagebox.showerror("优化失败", str(payload))
        else:
            self.status.set("优化完成。")
            self._append_log(str(payload.to_text()))
            messagebox.showinfo("优化完成", f"已生成优化版:\n{payload.output}")
        self.root.after(120, self._poll_worker)

    def _open_output_folder(self) -> None:
        path_text = self.output_path.get() or self.input_path.get()
        if not path_text:
            messagebox.showinfo("没有路径", "请先选择 PPT 文件或输出位置。")
            return
        folder = Path(path_text).expanduser()
        if folder.is_file() or folder.suffix:
            folder = folder.parent
        if not folder.exists():
            messagebox.showerror("目录不存在", str(folder))
            return
        os.startfile(folder)

    def _append_log(self, line: str) -> None:
        self.log.configure(state=NORMAL)
        self.log.insert(END, f"{line}\n")
        self.log.see(END)
        self.log.configure(state=DISABLED)

    def _clear_log(self) -> None:
        self.log.configure(state=NORMAL)
        self.log.delete("1.0", END)
        self.log.configure(state=DISABLED)


def main() -> None:
    root = Tk()
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    PPTOptimizerApp(root)
    root.mainloop()
