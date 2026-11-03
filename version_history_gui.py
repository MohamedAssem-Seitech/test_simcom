"""Four-field desktop interface for comparing YAML files at two commits."""

from __future__ import annotations

import os
import queue
import re
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


def validate_inputs(
    input_path: str, output_path: str, old_id: str, new_id: str,
) -> tuple[Path, Path, str, str]:
    """Validate form values without changing files or repository state."""
    if not all(value.strip() for value in (input_path, output_path, old_id, new_id)):
        raise ValueError("Please fill in all four fields.")
    source = Path(input_path.strip()).expanduser().resolve()
    output = Path(output_path.strip()).expanduser().resolve()
    if source.exists() and not source.is_file():
        raise ValueError("Input path must name a YAML file, not a folder.")
    if output.suffix.lower() != ".xlsx":
        raise ValueError("Output path must end in .xlsx.")
    if source == output:
        raise ValueError("Input and output must be different files.")
    if not output.parent.is_dir():
        raise ValueError("The output folder does not exist. Choose an existing folder.")
    if output.exists() and not output.is_file():
        raise ValueError("Output path must name a file, not a folder.")
    for name, identifier in (("Old", old_id), ("New", new_id)):
        if not re.fullmatch(r"[0-9a-fA-F]{4,40}", identifier.strip()):
            raise ValueError(f"{name} commit ID must contain 4–40 hexadecimal characters.")
    old_id, new_id = old_id.strip().lower(), new_id.strip().lower()
    if old_id == new_id:
        raise ValueError("Select two different commit IDs.")
    return source, output, old_id, new_id


def generate_report(source: Path, output: Path, old_id: str, new_id: str) -> int:
    """Compare commits and atomically save a report; return component count."""
    # Import here so missing dependencies are reported in the GUI.
    from version_history import compare_commits, locate_repository, write_report

    repo, relative_path = locate_repository(source)
    with repo:
        rows = compare_commits(repo, relative_path, old_id, new_id)
    descriptor, temporary = tempfile.mkstemp(
        suffix=".xlsx", prefix=".version_report_", dir=output.parent,
    )
    os.close(descriptor)
    temporary_path = Path(temporary)
    try:
        write_report(
            rows, temporary_path, rows[0]["Old Version"], rows[0]["New Version"],
        )
        temporary_path.replace(output)
    finally:
        temporary_path.unlink(missing_ok=True)
    return sum(len(row["Components"]) for row in rows)


class VersionHistoryApp:
    """Keep the interface responsive while Git and Excel run in a worker."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("Version History — Compare Commits")
        root.minsize(720, 340)
        self.results: queue.Queue[tuple[bool, str]] = queue.Queue()
        self.busy = False
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar(
            value=str(Path.cwd() / "version_changes_report.xlsx"),
        )
        self.old_id = tk.StringVar()
        self.new_id = tk.StringVar()
        self.status = tk.StringVar(value="Select a YAML file and two commit IDs.")

        frame = ttk.Frame(root, padding=24)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)
        ttk.Label(
            frame, text="Compare software components", font=("Segoe UI", 16, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        ttk.Label(
            frame, text="Compare any two local commits, including different branches.",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 18))
        self.controls: list[ttk.Widget] = []
        fields = [
            ("Input path (YAML)", self.input_path, self.browse_input),
            ("Output path (Excel)", self.output_path, self.browse_output),
            ("Old commit ID", self.old_id, None),
            ("New commit ID", self.new_id, None),
        ]
        for index, (label, variable, browse) in enumerate(fields, start=2):
            ttk.Label(frame, text=label).grid(
                row=index, column=0, sticky="w", padx=(0, 12), pady=7,
            )
            entry = ttk.Entry(frame, textvariable=variable, width=55)
            entry.grid(row=index, column=1, sticky="ew", pady=7)
            self.controls.append(entry)
            if browse is not None:
                button = ttk.Button(frame, text="Browse…", command=browse)
                button.grid(row=index, column=2, padx=(10, 0))
                self.controls.append(button)
        self.generate_button = ttk.Button(
            frame, text="Generate Excel report", command=self.start,
        )
        self.generate_button.grid(row=6, column=1, sticky="e", pady=(16, 8))
        self.controls.append(self.generate_button)
        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.grid(row=7, column=0, columnspan=3, sticky="ew", pady=6)
        ttk.Label(frame, textvariable=self.status, wraplength=660).grid(
            row=8, column=0, columnspan=3, sticky="w", pady=(6, 0),
        )
        self.controls[0].focus_set()
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.after(100, self.poll)

    def browse_input(self) -> None:
        """Select the YAML file inside the local Git working repository."""
        path = filedialog.askopenfilename(
            parent=self.root, title="Select YAML file",
            filetypes=[("YAML files", "*.yml *.yaml"), ("All files", "*.*")],
        )
        if path:
            self.input_path.set(path)

    def browse_output(self) -> None:
        """Choose the destination Excel filename."""
        path = filedialog.asksaveasfilename(
            parent=self.root, title="Save Excel report", defaultextension=".xlsx",
            initialfile="version_changes_report.xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)

    def start(self) -> None:
        """Validate the four fields and launch a background comparison."""
        if self.busy:
            return
        try:
            arguments = validate_inputs(
                self.input_path.get(), self.output_path.get(),
                self.old_id.get(), self.new_id.get(),
            )
        except (ValueError, OSError) as exc:
            messagebox.showerror("Check input", str(exc), parent=self.root)
            return
        if arguments[1].exists() and not messagebox.askyesno(
            "Replace report?", f"Replace this file?\n{arguments[1]}", parent=self.root,
        ):
            return
        self.busy = True
        for control in self.controls:
            control.state(["disabled"])
        self.status.set("Comparing commits and creating the Excel report…")
        self.progress.start()
        threading.Thread(target=self.worker, args=arguments, daemon=True).start()

    def worker(self, source: Path, output: Path, old_id: str, new_id: str) -> None:
        """Queue the result without calling Tk from a worker thread."""
        try:
            count = generate_report(source, output, old_id, new_id)
            self.results.put((True, f"Saved {count} component rows.\n{output}"))
        except PermissionError:
            self.results.put((False, "Cannot save the report. Close it in Excel "
                              "and choose a folder you can write to."))
        except ImportError as exc:
            self.results.put((False, f"Missing dependency: {exc}\n"
                              "Run: python -m pip install -r requirements.txt"))
        except Exception as exc:
            self.results.put((False, str(exc) or type(exc).__name__))

    def poll(self) -> None:
        """Display worker results on the Tk main thread."""
        try:
            success, message = self.results.get_nowait()
        except queue.Empty:
            pass
        else:
            self.busy = False
            self.progress.stop()
            for control in self.controls:
                control.state(["!disabled"])
            self.status.set(message)
            if success:
                messagebox.showinfo("Report ready", message, parent=self.root)
            else:
                messagebox.showerror("Could not generate report", message, parent=self.root)
        self.root.after(100, self.poll)

    def close(self) -> None:
        """Avoid interrupting a report while it is being saved."""
        if self.busy:
            messagebox.showinfo(
                "Report in progress", "Please wait for report generation to finish.",
                parent=self.root,
            )
            return
        self.root.destroy()


def main() -> None:
    """Open the desktop form."""
    root = tk.Tk()
    VersionHistoryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
