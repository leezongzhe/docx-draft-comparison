from __future__ import annotations

import argparse
import queue
import sys
import threading
import traceback
from pathlib import Path

from comparator import ComparisonError, compare_docx


def run_cli(arguments: argparse.Namespace) -> int:
    try:
        result = compare_docx(
            arguments.v1,
            arguments.v2,
            arguments.output,
            left_label=arguments.label1,
            right_label=arguments.label2,
            backend=arguments.backend,
            progress=lambda message: print(message, flush=True),
        )
    except ComparisonError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    stats = result.stats
    print(
        f"Created {result.output_pdf} using {result.backend}. "
        f"Sheets: {stats.paired_sheets}; added: {stats.additions}; deleted: {stats.deletions}; "
        f"replaced: {stats.replacements}; moved passages: {stats.moved_passages}."
    )
    for warning in result.warnings:
        print(f"Layout note: {warning}")
    return 0


def launch_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    class Application:
        def __init__(self, root: tk.Tk) -> None:
            self.root = root
            self.root.title("DOCX Draft Comparison")
            self.root.minsize(720, 460)
            self.messages: queue.Queue[tuple[str, object]] = queue.Queue()
            self.v1 = tk.StringVar()
            self.v2 = tk.StringVar()
            self.output = tk.StringVar(value=str(Path.cwd() / "comparison.pdf"))
            self.label1 = tk.StringVar(value="Version 1")
            self.label2 = tk.StringVar(value="Version 2")
            self.backend = tk.StringVar(value="auto")
            self.status = tk.StringVar(value="Choose two DOCX files. The source files will only be read.")
            self._build(ttk)
            self.root.after(150, self._poll)

        def _build(self, ttk_module: object) -> None:
            frame = ttk.Frame(self.root, padding=22)
            frame.pack(fill="both", expand=True)
            ttk.Label(frame, text="Compare two DOCX drafts", font=("Segoe UI", 17, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))
            ttk.Label(frame, text="Everything runs on this computer. No document content is uploaded or logged.").grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 18))
            self._file_row(frame, ttk, 2, "Version 1 DOCX", self.v1, self._choose_v1)
            self._file_row(frame, ttk, 3, "Version 2 DOCX", self.v2, self._choose_v2)
            self._file_row(frame, ttk, 4, "Output PDF", self.output, self._choose_output)
            ttk.Label(frame, text="Left label").grid(row=5, column=0, sticky="w", pady=6)
            ttk.Entry(frame, textvariable=self.label1).grid(row=5, column=1, sticky="ew", padx=(12, 8), pady=6)
            ttk.Label(frame, text="Right label").grid(row=6, column=0, sticky="w", pady=6)
            ttk.Entry(frame, textvariable=self.label2).grid(row=6, column=1, sticky="ew", padx=(12, 8), pady=6)
            ttk.Label(frame, text="Converter").grid(row=7, column=0, sticky="w", pady=6)
            ttk.Combobox(frame, textvariable=self.backend, values=("auto", "word", "libreoffice"), state="readonly").grid(row=7, column=1, sticky="w", padx=(12, 8), pady=6)
            self.progress = ttk.Progressbar(frame, mode="indeterminate")
            self.progress.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(18, 8))
            ttk.Label(frame, textvariable=self.status, wraplength=650).grid(row=9, column=0, columnspan=3, sticky="w", pady=(0, 14))
            self.compare_button = ttk.Button(frame, text="Create comparison PDF", command=self._start)
            self.compare_button.grid(row=10, column=0, columnspan=3, sticky="ew", ipady=6)
            ttk.Separator(frame).grid(row=11, column=0, columnspan=3, sticky="ew", pady=(18, 10))
            ttk.Label(
                frame,
                text="Legend: red = deleted, green = added, orange = replaced, purple outline = likely moved. "
                "Floating objects and text embedded in images are preserved visually but may not receive word highlights.",
                wraplength=650,
            ).grid(row=12, column=0, columnspan=3, sticky="w")
            frame.columnconfigure(1, weight=1)

        @staticmethod
        def _file_row(frame: object, ttk_module: object, row: int, label: str, variable: object, command: object) -> None:
            ttk_module.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=6)
            ttk_module.Entry(frame, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=(12, 8), pady=6)
            ttk_module.Button(frame, text="Browse...", command=command).grid(row=row, column=2, pady=6)

        def _choose_v1(self) -> None:
            path = filedialog.askopenfilename(filetypes=[("Word documents", "*.docx")])
            if path:
                self.v1.set(path)

        def _choose_v2(self) -> None:
            path = filedialog.askopenfilename(filetypes=[("Word documents", "*.docx")])
            if path:
                self.v2.set(path)

        def _choose_output(self) -> None:
            path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if path:
                self.output.set(path)

        def _start(self) -> None:
            if not self.v1.get() or not self.v2.get() or not self.output.get():
                messagebox.showerror("Missing file", "Choose both DOCX files and an output PDF.")
                return
            self.compare_button.configure(state="disabled")
            self.progress.start(12)
            self.status.set("Starting local comparison...")
            threading.Thread(target=self._worker, daemon=True).start()

        def _worker(self) -> None:
            try:
                result = compare_docx(
                    self.v1.get(), self.v2.get(), self.output.get(),
                    left_label=self.label1.get().strip() or "Version 1",
                    right_label=self.label2.get().strip() or "Version 2",
                    backend=self.backend.get(),
                    progress=lambda message: self.messages.put(("status", message)),
                )
                self.messages.put(("done", result))
            except Exception as exc:
                self.messages.put(("error", (exc, traceback.format_exc())))

        def _poll(self) -> None:
            try:
                while True:
                    kind, payload = self.messages.get_nowait()
                    if kind == "status":
                        self.status.set(str(payload))
                    elif kind == "done":
                        self.progress.stop()
                        self.compare_button.configure(state="normal")
                        result = payload
                        self.status.set(f"Created {result.output_pdf}")
                        messagebox.showinfo("Comparison complete", f"The PDF was created successfully.\n\n{result.output_pdf}")
                    elif kind == "error":
                        self.progress.stop()
                        self.compare_button.configure(state="normal")
                        exc, _trace = payload
                        self.status.set("Comparison failed.")
                        messagebox.showerror("Comparison failed", str(exc))
            except queue.Empty:
                pass
            self.root.after(150, self._poll)

    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.15)
    except tk.TclError:
        pass
    Application(root)
    root.mainloop()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create a local side-by-side DOCX comparison PDF.")
    parser.add_argument("--cli", action="store_true", help="Run without the graphical interface.")
    parser.add_argument("--v1", help="Version 1 DOCX path.")
    parser.add_argument("--v2", help="Version 2 DOCX path.")
    parser.add_argument("--output", help="Output PDF path.")
    parser.add_argument("--label1", default="Version 1")
    parser.add_argument("--label2", default="Version 2")
    parser.add_argument("--backend", choices=("auto", "word", "libreoffice"), default="auto")
    arguments = parser.parse_args()
    if arguments.cli and not all((arguments.v1, arguments.v2, arguments.output)):
        parser.error("--cli requires --v1, --v2, and --output")
    return arguments


if __name__ == "__main__":
    args = parse_args()
    raise SystemExit(run_cli(args) if args.cli else (launch_gui() or 0))

