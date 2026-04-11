from __future__ import annotations

from pathlib import Path

from .common import ConversionOptions, ToolState, ensure_output_allowed, output_path_for
from .converter import convert_one_with_options
from .pdf_utils import compress_pdf, merge_pdfs, split_pdf
from .renderers import renderer_candidates


def launch_gui(tools: ToolState) -> int:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Offline Document Converter")
    root.geometry("760x620")

    status_var = tk.StringVar(value="Ready")
    selected_inputs: list[Path] = []
    output_dir_var = tk.StringVar(value=str(Path.cwd()))
    target_var = tk.StringVar(value="pdf")
    mode_var = tk.StringVar(value="auto")
    ocr_var = tk.BooleanVar(value=False)
    ocr_lang_var = tk.StringVar(value="eng")
    sensitive_var = tk.BooleanVar(value=False)
    overwrite_var = tk.BooleanVar(value=False)
    skip_existing_var = tk.BooleanVar(value=True)

    def refresh_input_list() -> None:
        input_box.configure(state="normal")
        input_box.delete("1.0", tk.END)
        if selected_inputs:
            input_box.insert(tk.END, "\n".join(str(path) for path in selected_inputs))
        input_box.configure(state="disabled")

    def choose_files() -> None:
        nonlocal selected_inputs
        files = filedialog.askopenfilenames(title="Choose files to convert")
        if files:
            selected_inputs = [Path(item).resolve() for item in files]
            refresh_input_list()
            status_var.set(f"Selected {len(selected_inputs)} file(s)")

    def clear_files() -> None:
        nonlocal selected_inputs
        selected_inputs = []
        refresh_input_list()
        status_var.set("Selection cleared")

    def choose_output_dir() -> None:
        chosen = filedialog.askdirectory(title="Choose output directory")
        if chosen:
            output_dir_var.set(chosen)

    def run_gui_convert() -> None:
        if not selected_inputs:
            messagebox.showerror("No files selected", "Choose at least one input file.")
            return
        output_dir = Path(output_dir_var.get()).expanduser().resolve()
        target = target_var.get().strip().lower().lstrip(".")
        if not target:
            messagebox.showerror("Missing target format", "Enter a target format like pdf, docx, txt, or pptx.")
            return

        results: list[str] = []
        try:
            for input_path in selected_inputs:
                destination = output_path_for(input_path, output_dir, target)
                allowed = ensure_output_allowed(
                    destination,
                    overwrite=overwrite_var.get(),
                    skip_existing=skip_existing_var.get(),
                )
                if not allowed:
                    results.append(f"Skipped existing: {destination}")
                    continue
                report = convert_one_with_options(
                    input_path,
                    target,
                    destination,
                    tools,
                    ConversionOptions(
                        ocr=ocr_var.get(),
                        ocr_lang=ocr_lang_var.get().strip() or "eng",
                        mode=mode_var.get(),
                        sensitive=sensitive_var.get(),
                    ),
                )
                results.append(f"Converted: {destination}")
                results.append(f"  Engine: {report.engine}")
                results.append(f"  Fidelity: {report.fidelity}")
                results.append(f"  Verified: {'yes' if report.verified else 'no'}")
                if report.checksum:
                    results.append(f"  SHA256: {report.checksum}")
                for note in report.notes:
                    results.append(f"  Note: {note}")
        except Exception as exc:
            messagebox.showerror("Conversion failed", str(exc))
            status_var.set("Conversion failed")
            return

        result_box.configure(state="normal")
        result_box.delete("1.0", tk.END)
        result_box.insert(tk.END, "\n".join(results) if results else "Nothing to do.")
        result_box.configure(state="disabled")
        status_var.set("Finished")

    def gui_merge() -> None:
        files = filedialog.askopenfilenames(title="Choose PDFs to merge", filetypes=[("PDF files", "*.pdf")])
        if len(files) < 2:
            return
        output = filedialog.asksaveasfilename(
            title="Save merged PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not output:
            return
        try:
            ensure_output_allowed(Path(output).resolve(), overwrite=overwrite_var.get(), skip_existing=False)
            merge_pdfs([Path(item).resolve() for item in files], Path(output).resolve())
            messagebox.showinfo("Done", f"Merged PDF saved to:\n{output}")
            status_var.set("Merged PDFs")
        except Exception as exc:
            messagebox.showerror("Merge failed", str(exc))

    def gui_split() -> None:
        source = filedialog.askopenfilename(title="Choose PDF to split", filetypes=[("PDF files", "*.pdf")])
        if not source:
            return
        out_dir = filedialog.askdirectory(title="Choose output folder")
        if not out_dir:
            return
        try:
            split_pdf(Path(source).resolve(), Path(out_dir).resolve())
            messagebox.showinfo("Done", f"Split pages saved under:\n{out_dir}")
            status_var.set("Split PDF")
        except Exception as exc:
            messagebox.showerror("Split failed", str(exc))

    def gui_compress() -> None:
        source = filedialog.askopenfilename(title="Choose PDF to compress", filetypes=[("PDF files", "*.pdf")])
        if not source:
            return
        output = filedialog.asksaveasfilename(
            title="Save compressed PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not output:
            return
        try:
            ensure_output_allowed(Path(output).resolve(), overwrite=overwrite_var.get(), skip_existing=False)
            compress_pdf(Path(source).resolve(), Path(output).resolve())
            messagebox.showinfo("Done", f"Compressed PDF saved to:\n{output}")
            status_var.set("Compressed PDF")
        except Exception as exc:
            messagebox.showerror("Compression failed", str(exc))

    tool_lines = [
        "PyMuPDF: ok",
        "python-docx: ok",
        "python-pptx: ok",
        "reportlab: ok",
        "pypdf: ok",
        f"tesseract: {'ok' if tools.tesseract else 'optional / missing'}",
        f"libreoffice: {'ok' if tools.libreoffice else 'optional / missing'}",
    ]
    if tools.platform == "Darwin":
        tool_lines.append(
            f"Microsoft Word (macOS): {'ok' if tools.osascript and tools.mac_word_app else 'optional / missing'}"
        )
        tool_lines.append(
            f"Microsoft PowerPoint (macOS): {'ok' if tools.osascript and tools.mac_powerpoint_app else 'optional / missing'}"
        )
    elif tools.platform == "Windows":
        tool_lines.append(f"Windows Office automation host: {'ok' if tools.powershell else 'optional / missing'}")

    root.columnconfigure(0, weight=1)
    root.rowconfigure(3, weight=1)

    ttk.Label(root, text="Offline Document Converter", font=("Helvetica", 18, "bold")).grid(
        row=0, column=0, sticky="w", padx=16, pady=(16, 4)
    )
    ttk.Label(
        root,
        text="Cross-platform local conversions for sensitive documents. No cloud upload step.",
    ).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 12))

    top = ttk.Frame(root, padding=16)
    top.grid(row=2, column=0, sticky="nsew")
    top.columnconfigure(1, weight=1)
    top.rowconfigure(1, weight=1)

    ttk.Label(top, text="Input files").grid(row=0, column=0, sticky="w")
    buttons = ttk.Frame(top)
    buttons.grid(row=0, column=1, sticky="e")
    ttk.Button(buttons, text="Choose Files", command=choose_files).pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="Clear", command=clear_files).pack(side="left")

    input_box = tk.Text(top, height=8, wrap="word")
    input_box.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(6, 12))
    input_box.configure(state="disabled")

    settings = ttk.Frame(top)
    settings.grid(row=2, column=0, columnspan=2, sticky="ew")
    settings.columnconfigure(1, weight=1)

    ttk.Label(settings, text="Target format").grid(row=0, column=0, sticky="w")
    ttk.Combobox(
        settings,
        textvariable=target_var,
        values=["pdf", "docx", "txt", "html", "md", "rtf", "pptx"],
        width=12,
    ).grid(row=0, column=1, sticky="w", padx=(8, 20))

    ttk.Label(settings, text="OCR language").grid(row=0, column=2, sticky="w")
    ttk.Entry(settings, textvariable=ocr_lang_var, width=12).grid(row=0, column=3, sticky="w", padx=(8, 0))

    ttk.Label(settings, text="Engine mode").grid(row=1, column=0, sticky="w", pady=(10, 0))
    ttk.Combobox(
        settings,
        textvariable=mode_var,
        values=["auto", "native", "python"],
        width=12,
        state="readonly",
    ).grid(row=1, column=1, sticky="w", padx=(8, 20), pady=(10, 0))

    ttk.Label(settings, text="Output folder").grid(row=2, column=0, sticky="w", pady=(10, 0))
    ttk.Entry(settings, textvariable=output_dir_var).grid(
        row=2, column=1, columnspan=2, sticky="ew", padx=(8, 8), pady=(10, 0)
    )
    ttk.Button(settings, text="Browse", command=choose_output_dir).grid(row=2, column=3, sticky="e", pady=(10, 0))

    ttk.Label(settings, text="Renderer hints").grid(row=1, column=2, sticky="w", pady=(10, 0))
    renderer_hint = tk.StringVar(value="Selected files will show engine info after conversion.")
    ttk.Label(settings, textvariable=renderer_hint, wraplength=220).grid(row=1, column=3, sticky="w", pady=(10, 0))

    toggles = ttk.Frame(top)
    toggles.grid(row=3, column=0, columnspan=2, sticky="w", pady=(12, 12))
    ttk.Checkbutton(toggles, text="Use OCR for PDFs", variable=ocr_var).pack(side="left", padx=(0, 14))
    ttk.Checkbutton(
        toggles,
        text="Sensitive Document Mode",
        variable=sensitive_var,
    ).pack(side="left", padx=(0, 14))
    ttk.Checkbutton(toggles, text="Overwrite existing files", variable=overwrite_var).pack(side="left", padx=(0, 14))
    ttk.Checkbutton(toggles, text="Skip existing files", variable=skip_existing_var).pack(side="left")

    def refresh_renderer_hint(*_args) -> None:
        if not selected_inputs:
            renderer_hint.set("Selected files will show engine info after conversion.")
            return
        exts = sorted({path.suffix.lower().lstrip(".") for path in selected_inputs})
        hints = []
        for ext in exts:
            candidates = renderer_candidates(ext, tools)
            if candidates:
                hints.append(f".{ext}: {', '.join(candidates)}")
            else:
                hints.append(f".{ext}: python fallback")
        renderer_hint.set(" | ".join(hints))

    mode_var.trace_add("write", refresh_renderer_hint)
    refresh_renderer_hint()

    ttk.Button(top, text="Convert Files", command=run_gui_convert).grid(row=4, column=0, sticky="w")

    utilities = ttk.LabelFrame(root, text="PDF Utilities", padding=16)
    utilities.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 12))
    ttk.Button(utilities, text="Merge PDFs", command=gui_merge).pack(side="left", padx=(0, 10))
    ttk.Button(utilities, text="Split PDF", command=gui_split).pack(side="left", padx=(0, 10))
    ttk.Button(utilities, text="Compress PDF", command=gui_compress).pack(side="left")

    bottom = ttk.Frame(root, padding=(16, 0, 16, 16))
    bottom.grid(row=4, column=0, sticky="nsew")
    bottom.columnconfigure(0, weight=1)
    bottom.columnconfigure(1, weight=1)

    tool_frame = ttk.LabelFrame(bottom, text="Tool Status", padding=12)
    tool_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
    ttk.Label(tool_frame, text="\n".join(tool_lines), justify="left").pack(anchor="w")

    result_frame = ttk.LabelFrame(bottom, text="Results", padding=12)
    result_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
    result_box = tk.Text(result_frame, height=10, wrap="word")
    result_box.pack(fill="both", expand=True)
    result_box.configure(state="disabled")

    ttk.Label(root, textvariable=status_var, relief="sunken", anchor="w").grid(row=5, column=0, sticky="ew")
    root.mainloop()
    return 0
