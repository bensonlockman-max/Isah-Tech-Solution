# Isah-Tech-Solution#!/usr/bin/env python3
"""
Smart Excel Splitter
A standalone Tkinter + ttkbootstrap application that splits Excel files by selected columns.

Features:
- Browse and load .xlsx files using pandas
- Preview first 20 rows in a popup Treeview with scrollbars
- Multi-column splitting: creates folder per column and Excel file per unique value
- Settings saved to JSON (last file, last selected columns)
- Determinate progress bar with background worker and queue updates
- Themed UI using ttkbootstrap (theme: superhero)

Author: ISAH TECH SOLUTIONS
"""

import os
import re
import json
import threading
import queue
from pathlib import Path
from tkinter import Tk, Toplevel, StringVar, IntVar, messagebox, filedialog
from tkinter import scrolledtext as tk_scrolledtext
from tkinter import ttk as tkttk
import ttkbootstrap as tb
from ttkbootstrap import Style
import pandas as pd

# Application branding
APP_NAME = "Smart Excel Splitter"
BRAND_NAME = "ISAH TECH SOLUTIONS"

# Settings file location (same directory as script)
SETTINGS_FILE = Path(__file__).parent / "settings.json"

# Default output folder
DEFAULT_OUTPUT_DIR = Path(__file__).parent / "Output"

# Invalid filename characters to remove
INVALID_FILENAME_CHARS = r'[<>:"/\\|?*]'


def sanitize_filename(name: str) -> str:
    """Remove characters that are invalid in filenames and strip whitespace."""
    if not isinstance(name, str):
        name = str(name)
    # Replace path separators and other invalid characters
    sanitized = re.sub(INVALID_FILENAME_CHARS, "_", name)
    # Trim and collapse spaces
    sanitized = re.sub(r"\s+", " ", sanitized).strip()
    if sanitized == "":
        sanitized = "empty"
    return sanitized


class ExcelSplitterApp:
    def __init__(self, root: Tk):
        # Main window setup with ttkbootstrap style
        self.root = root
        self.style = Style(theme="superhero")
        self.root.title(f"{APP_NAME} — {BRAND_NAME}")
        self.root.geometry("700x500")
        self.root.minsize(650, 450)

        # Data variables
        self.excel_path_var = StringVar()
        self.output_dir_var = StringVar(value=str(DEFAULT_OUTPUT_DIR))
        self.status_var = StringVar(value="Ready")

        # pandas DataFrame
        self.df = None
        self.columns = []

        # Checkbutton variables
        self.column_vars = {}  # column_name -> IntVar

        # Progress management
        self.progress_var = IntVar(value=0)
        self.progress_max = 100
        self.progress_queue = queue.Queue()
        self.worker_thread = None
        self.stop_requested = False

        # Build UI
        self._build_ui()

        # Load settings if available
        self._load_settings()

        # Start queue polling for progress updates
        self._poll_queue()

        # Per-column totals/done tracking (used for UI updates)
        self.col_totals = {}
        self.col_done = {}

    def preview_selected_columns(self):
        """Preview only selected columns in a popup window."""
        if self.df is None:
            self.append_log("Load an Excel file before previewing.")
            messagebox.showwarning("No File Loaded", "Please load an Excel file first.")
            return

        # Get selected columns from checklist
        selected_columns = [col for col, var in self.column_vars.items() if var.get() == 1]

        if not selected_columns:
            messagebox.showwarning("No Columns Selected", "Select at least one column to preview.")
            return

        # Subset dataframe (first 20 rows)
        preview_df = self.df[selected_columns].head(20)

        # --- Create popup window ---
        popup = tb.Toplevel(self.root)
        popup.title("Preview Selected Columns")
        popup.geometry("800x400")

        # Frame with scrollbars
        preview_frame = tb.Frame(popup)
        preview_frame.pack(fill="both", expand=True, padx=10, pady=10)

        tree_scroll_y = tkttk.Scrollbar(preview_frame, orient="vertical")
        tree_scroll_y.pack(side="right", fill="y")

        tree_scroll_x = tkttk.Scrollbar(preview_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")

        tree = tkttk.Treeview(preview_frame,
                              columns=selected_columns,
                              show="headings",
                              yscrollcommand=tree_scroll_y.set,
                              xscrollcommand=tree_scroll_x.set)

        tree.pack(fill="both", expand=True)

        tree_scroll_y.config(command=tree.yview)
        tree_scroll_x.config(command=tree.xview)

        # --- Insert columns ---
        for col in selected_columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor="center")

        # --- Insert rows ---
        for _, row in preview_df.iterrows():
            tree.insert("", "end", values=list(row))

        self.append_log(f"Previewing columns: {', '.join(selected_columns)}")


    def _build_ui(self):
        """Create and layout all widgets."""
        pad = dict(padx=12, pady=8)

        # Create a scrollable main area so the whole window can scroll if needed
        self.main_canvas = tb.Canvas(self.root)
        v_scroll_main = tb.Scrollbar(self.root, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=v_scroll_main.set)
        v_scroll_main.pack(side="right", fill="y")
        self.main_canvas.pack(side="left", fill="both", expand=True)

        # Inner frame which will contain all UI widgets
        self.inner_frame = tb.Frame(self.main_canvas)
        self.inner_window = self.main_canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        # Keep canvas scrollregion updated
        self.inner_frame.bind("<Configure>", lambda e: self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all")))
        # Ensure inner window width tracks canvas width
        self.main_canvas.bind('<Configure>', lambda e: self.main_canvas.itemconfig(self.inner_window, width=e.width))
        # Mousewheel support on Windows
        self.main_canvas.bind_all("<MouseWheel>", lambda e: self.main_canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        header = tb.Frame(self.inner_frame)
        header.pack(fill="x", **pad)

        title_lbl = tb.Label(header, text=APP_NAME, font=("Helvetica", 16, "bold"))
        title_lbl.pack(side="left")

        brand_lbl = tb.Label(header, text=BRAND_NAME, font=("Helvetica", 10))
        brand_lbl.pack(side="right")

        # File selection frame
        file_frame = tb.Labelframe(self.inner_frame, text="Excel File", bootstyle="primary")
        file_frame.pack(fill="x", padx=12, pady=(0, 8))

        self.file_entry = tb.Entry(file_frame, textvariable=self.excel_path_var, width=60)
        self.file_entry.grid(row=0, column=0, sticky="w", padx=(10, 6), pady=10)

        self.browse_btn = tb.Button(file_frame, text="Browse", command=self.browse_file)
        self.browse_btn.grid(row=0, column=1, padx=6)

        # Preview button for selected columns (opens popup)
        self.preview_btn = tb.Button(file_frame, text="Preview Selected Columns", command=self.preview_selected_columns, bootstyle="info")
        self.preview_btn.grid(row=0, column=2, padx=6)
        

        # Output folder selection
        out_frame = tb.Frame(self.inner_frame)
        out_frame.pack(fill="x", padx=12, pady=(0, 8))

        out_label = tb.Label(out_frame, text="Output Folder:")
        out_label.pack(side="left", padx=(0, 8))

        self.out_entry = tb.Entry(out_frame, textvariable=self.output_dir_var, width=52)
        self.out_entry.pack(side="left", padx=(0, 8))

        self.out_browse = tb.Button(out_frame, text="Browse", command=self.browse_output_folder)
        self.out_browse.pack(side="left")

        # Columns checklist frame
        # Add an inline preview section above the checklist
        preview_frame = tb.Labelframe(self.inner_frame, text="Data Preview (first 20 rows)", bootstyle="info")
        preview_frame.pack(fill="both", expand=False, padx=12, pady=(0, 8), ipady=4)

        # Treeview for preview
        self.preview_tree = tkttk.Treeview(preview_frame, show="headings")
        self.preview_vsb = tb.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        self.preview_hsb = tb.Scrollbar(preview_frame, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=self.preview_vsb.set, xscrollcommand=self.preview_hsb.set)
        self.preview_vsb.pack(side="right", fill="y")
        self.preview_hsb.pack(side="bottom", fill="x")
        self.preview_tree.pack(fill="both", expand=True)

        columns_frame = tb.Labelframe(self.inner_frame, text="Select Columns to Split By", bootstyle="info")
        columns_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        # Add a canvas and scrollbar for the dynamic checklist so it can scroll when many columns
        self.check_canvas = tb.Canvas(columns_frame)
        self.check_canvas.pack(side="left", fill="both", expand=True)

        self.check_scrollbar = tb.Scrollbar(columns_frame, orient="vertical", command=self.check_canvas.yview)
        self.check_scrollbar.pack(side="right", fill="y")

        self.check_canvas.configure(yscrollcommand=self.check_scrollbar.set)

        self.check_inner = tb.Frame(self.check_canvas)
        self.check_window = self.check_canvas.create_window((0, 0), window=self.check_inner, anchor="nw")

        # Bind resizing events to update scrollregion
        self.check_inner.bind("<Configure>", lambda e: self.check_canvas.configure(scrollregion=self.check_canvas.bbox("all")))
        self.check_canvas.bind('<Configure>', self._on_canvas_configure)

        # Bottom controls: Start button, Cancel and progress
        bottom = tb.Frame(self.inner_frame)
        bottom.pack(fill="x", padx=12, pady=(0, 12))

        # Start and Cancel controls
        self.start_btn = tb.Button(bottom, text="Start Splitting", bootstyle="success", command=self.start_splitting)
        self.start_btn.pack(side="left")

        self.cancel_btn = tb.Button(bottom, text="Cancel", bootstyle="danger", command=self.cancel_job)
        self.cancel_btn.pack(side="left", padx=(6, 0))
        self.cancel_btn.configure(state="disabled")

        self.progress = tb.Progressbar(bottom, orient="horizontal", mode="determinate", variable=self.progress_var)
        self.progress.pack(side="left", fill="x", expand=True, padx=(12, 12))

        status_lbl = tb.Label(bottom, textvariable=self.status_var)
        status_lbl.pack(side="right")

        # Job frame: per-column progress and log output
        job_frame = tb.Labelframe(self.inner_frame, text="Job Status & Log", bootstyle="secondary")
        job_frame.pack(fill="both", expand=False, padx=12, pady=(0, 12), ipady=6)

        # Left: per-column progress Treeview
        self.col_tree = tkttk.Treeview(job_frame, columns=("total", "done"), show="headings", height=4)
        self.col_tree.heading("total", text="Total")
        self.col_tree.heading("done", text="Done")
        self.col_tree.column("total", width=80, anchor="center")
        self.col_tree.column("done", width=80, anchor="center")
        self.col_tree.pack(side="left", fill="y", padx=(8, 6), pady=6)

        # Right: scrollable job log
        log_frame = tb.Frame(job_frame)
        log_frame.pack(side="left", fill="both", expand=True, padx=(0, 8))

        self.log_text = tk_scrolledtext.ScrolledText(log_frame, height=6, state="disabled", wrap="word")
        self.log_text.pack(fill="both", expand=True)

    def _on_canvas_configure(self, event):
        # Resize inner frame to canvas width
        canvas_width = event.width
        self.check_canvas.itemconfig(self.check_window, width=canvas_width)

    def browse_file(self):
        """Open file dialog to select an Excel file and load it."""
        path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if path:
            self.excel_path_var.set(path)
            self._load_excel(path)
            # Save settings immediately
            self._save_settings()

    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_dir_var.set(folder)

    def _load_excel(self, path: str):
        """Load Excel into pandas DataFrame and populate columns checklist."""
        try:
            # Attempt to read with pandas (use engine openpyxl)
            df = pd.read_excel(path, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file:\n{e}")
            self.df = None
            self.columns = []
            self._populate_checklist()
            return

        self.df = df
        self.columns = list(df.columns)
        self._populate_checklist()
        self.status_var.set(f"Loaded: {os.path.basename(path)}")
        # Automatically populate inline preview when a file is loaded
        try:
            self._populate_preview()
        except Exception:
            pass

    def _populate_checklist(self):
        """Create dynamic checklist of columns using IntVar for each column."""
        # Clear existing
        for widget in self.check_inner.winfo_children():
            widget.destroy()
        self.column_vars.clear()

        if not self.columns:
            lbl = tb.Label(self.check_inner, text="No columns available. Load an Excel file.")
            lbl.pack(anchor="w", pady=6, padx=6)
            return

        # Create a Checkbutton for each column
        for col in self.columns:
            var = IntVar(value=0)
            cb = tb.Checkbutton(self.check_inner, text=col, variable=var, bootstyle="round-toggle")
            cb.pack(anchor="w", pady=2, padx=6)
            self.column_vars[col] = var

        # Try to restore previously selected columns from settings
        settings = self._read_settings()
        if settings:
            last_selected = settings.get("last_selected_columns", [])
            if isinstance(last_selected, list):
                for c in last_selected:
                    if c in self.column_vars:
                        self.column_vars[c].set(1)

        # Reset column progress display (clear previous items)
        try:
            for iid in self.col_tree.get_children():
                self.col_tree.delete(iid)
        except Exception:
            pass

    def preview_data(self):
        """(Deprecated) Previously used to open a separate preview window.
        The application now shows an inline preview automatically. Use
        `_populate_preview()` to refresh the inline preview.
        """
        # Backwards compatibility: refresh inline preview
        self._populate_preview()

    def _populate_preview(self):
        """Populate the inline preview Treeview with the first 20 rows of the loaded DataFrame."""
        # If no DataFrame loaded, clear preview and return
        if self.df is None:
            try:
                # clear existing rows and columns
                for iid in self.preview_tree.get_children():
                    self.preview_tree.delete(iid)
                self.preview_tree.configure(columns=())
            except Exception:
                pass
            return

        # Configure columns
        cols = list(self.df.columns.astype(str))
        # If there's a 'Section' column, move it to the front so it's always visible
        section_idx = None
        for i, c in enumerate(cols):
            if str(c).strip().lower() == "section":
                section_idx = i
                break
        if section_idx is not None and section_idx != 0:
            section_col = cols.pop(section_idx)
            cols.insert(0, section_col)
        try:
            self.preview_tree.configure(columns=cols)
        except Exception:
            pass

        # Setup headings and clear previous rows
        for c in cols:
            try:
                self.preview_tree.heading(c, text=c)
            except Exception:
                pass
        try:
            for iid in self.preview_tree.get_children():
                self.preview_tree.delete(iid)
        except Exception:
            pass

        # Insert first N rows
        N = min(20, len(self.df))
        sample = self.df.head(N).fillna("")

        # Determine column widths
        col_widths = {c: max(len(str(c)), 10) for c in cols}
        for _, row in sample.iterrows():
            for c in cols:
                col_widths[c] = max(col_widths[c], len(str(row[c])))

        for _, row in sample.iterrows():
            values = [row[c] for c in cols]
            try:
                self.preview_tree.insert("", "end", values=values)
            except Exception:
                pass

        # Adjust column widths
        for c in cols:
            try:
                width = min(max(col_widths[c] * 8, 80), 400)
                self.preview_tree.column(c, width=width, anchor="w")
            except Exception:
                pass

    def append_log(self, message: str):
        """Append a line to the job log text widget (safe to call from main thread)."""
        try:
            self.log_text.configure(state="normal")
            self.log_text.insert("end", message + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        except Exception:
            pass

    def start_splitting(self):
        """Validate inputs and start the splitting process in a background thread."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load an Excel file before splitting.")
            return

        # Get selected columns
        selected = [col for col, var in self.column_vars.items() if var.get() == 1]
        if not selected:
            messagebox.showwarning("No Columns Selected", "Please select at least one column to split by.")
            return

        output_dir = Path(self.output_dir_var.get())
        if not output_dir.exists():
            try:
                output_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"Unable to create output directory:\n{e}")
                return

        # Save last selections to settings
        settings = self._read_settings() or {}
        settings["last_selected_columns"] = selected
        settings["last_used_file"] = self.excel_path_var.get()
        self._write_settings(settings)

        # Disable UI controls during processing
        self._set_ui_state(disabled=True)
        self.status_var.set("Preparing split job...")
        self.progress_var.set(0)

        # Count total tasks (sum of unique values across selected columns)
        total_tasks = 0
        unique_values_map = {}
        for col in selected:
            try:
                unique_vals = self.df[col].dropna().unique()
            except Exception:
                unique_vals = self.df[col].astype(str).dropna().unique()
            unique_values_map[col] = unique_vals
            total_tasks += len(unique_vals)

        if total_tasks == 0:
            messagebox.showinfo("Nothing to do", "Selected columns contain no unique values to split.")
            self._set_ui_state(disabled=False)
            return

        # Set progress maximum
        self.progress_max = total_tasks
        self.progress.configure(maximum=self.progress_max)

        # Initialize per-column totals and populate treeview
        self.col_totals = {}
        self.col_done = {}
        # Clear any previous entries
        try:
            for iid in self.col_tree.get_children():
                self.col_tree.delete(iid)
        except Exception:
            pass
        for col in selected:
            total = len(unique_values_map.get(col, []))
            self.col_totals[col] = total
            self.col_done[col] = 0
            try:
                self.col_tree.insert("", "end", iid=col, values=(total, 0))
            except Exception:
                pass

        # Clear job log and enable Cancel
        try:
            self.log_text.configure(state="normal")
            self.log_text.delete("1.0", "end")
            self.log_text.configure(state="disabled")
        except Exception:
            pass
        self.cancel_btn.configure(state="normal")

        # Start worker thread
        self.stop_requested = False
        self.worker_thread = threading.Thread(target=self._worker_split, args=(selected, unique_values_map, output_dir), daemon=True)
        self.worker_thread.start()

    def _set_ui_state(self, disabled: bool):
        """Enable or disable UI elements to prevent interaction during processing."""
        state = "disabled" if disabled else "normal"
        widgets = [getattr(self, 'start_btn', None), getattr(self, 'browse_btn', None), getattr(self, 'preview_btn', None), getattr(self, 'out_browse', None), getattr(self, 'file_entry', None), getattr(self, 'out_entry', None)]
        for w in widgets:
            if w is None:
                continue
            try:
                w.configure(state=state)
            except Exception:
                pass

    def cancel_job(self):
        """Request cancellation of the running worker job."""
        if not self.worker_thread or not self.worker_thread.is_alive():
            return
        self.stop_requested = True
        try:
            self.append_log("Cancellation requested. Waiting for worker to stop...")
        except Exception:
            pass

    def _worker_split(self, selected_columns, unique_values_map, output_dir: Path):
        """Background worker that performs the splitting and writes Excel files.
        Uses self.progress_queue to send progress updates back to the main thread.
        """
        try:
            total_done = 0
            # For each selected column, make a folder
            for col in selected_columns:
                if self.stop_requested:
                    break
                folder_name = sanitize_filename(col)
                col_folder = output_dir / folder_name
                try:
                    col_folder.mkdir(parents=True, exist_ok=True)
                except PermissionError as pe:
                    self.progress_queue.put(("error", f"Permission denied creating folder: {col_folder}\n{pe}"))
                    return
                except Exception as e:
                    self.progress_queue.put(("error", f"Failed to create folder: {col_folder}\n{e}"))
                    return

                unique_vals = unique_values_map.get(col, [])
                # Notify start of column
                self.progress_queue.put(("log", f"Starting column '{col}' ({len(unique_vals)} values)..."))
                self.progress_queue.put(("col_start", col, len(unique_vals)))

                done_for_col = 0
                # Iterate unique values and save corresponding subset
                for val in unique_vals:
                    if self.stop_requested:
                        break
                    # Filter rows where column == val (preserve rows where NaN? We dropped NaN earlier)
                    try:
                        subset = self.df[self.df[col] == val]
                    except Exception:
                        # Fallback: compare as strings
                        subset = self.df[self.df[col].astype(str) == str(val)]

                    # Create filename
                    safe_name = sanitize_filename(val)
                    file_path = col_folder / f"{safe_name}.xlsx"
                    try:
                        # Use pandas ExcelWriter to save
                        subset.to_excel(file_path, index=False, engine="openpyxl")
                    except PermissionError as pe:
                        self.progress_queue.put(("error", f"Permission denied saving file: {file_path}\n{pe}"))
                        return
                    except Exception as e:
                        self.progress_queue.put(("error", f"Failed to save file: {file_path}\n{e}"))
                        return

                    total_done += 1
                    done_for_col += 1
                    # Send progress update overall and per-column
                    self.progress_queue.put(("progress", total_done))
                    self.progress_queue.put(("col_progress", col, done_for_col))
                    self.progress_queue.put(("log", f"Saved: {file_path}"))

            # When done or cancelled
            if self.stop_requested:
                self.progress_queue.put(("log", "Job cancelled by user."))
                self.progress_queue.put(("done", "Cancelled"))
            else:
                self.progress_queue.put(("done", "Splitting completed."))
        except Exception as e:
            self.progress_queue.put(("error", f"Unexpected error during splitting:\n{e}"))

    def _poll_queue(self):
        """Poll the progress queue and update UI accordingly. Runs in main thread via after()."""
        try:
            while True:
                item = self.progress_queue.get_nowait()
                typ = item[0]
                if typ == "progress":
                    _, val = item
                    self.progress_var.set(val)
                    self.status_var.set(f"Processed {val} of {self.progress_max}")
                elif typ == "error":
                    _, payload = item
                    # Re-enable UI and show error
                    self._set_ui_state(disabled=False)
                    self.cancel_btn.configure(state="disabled")
                    self.status_var.set("Error")
                    messagebox.showerror("Error", payload)
                elif typ == "done":
                    _, payload = item
                    if payload == "Cancelled":
                        self.status_var.set("Cancelled")
                    else:
                        self.progress_var.set(self.progress_max)
                        self.status_var.set(str(payload))
                    messagebox.showinfo("Done", str(payload))
                    self._set_ui_state(disabled=False)
                    # disable cancel button when finished
                    self.cancel_btn.configure(state="disabled")
                elif typ == "log":
                    _, payload = item
                    self.append_log(payload)
                elif typ == "col_start":
                    _, col, total = item
                    # ensure treeview has the item
                    try:
                        if col not in self.col_tree.get_children():
                            self.col_tree.insert("", "end", iid=col, values=(total, 0))
                        else:
                            self.col_tree.item(col, values=(total, 0))
                    except Exception:
                        pass
                elif typ == "col_progress":
                    _, col, done = item
                    try:
                        total = self.col_totals.get(col, "")
                        self.col_done[col] = done
                        self.col_tree.item(col, values=(total, done))
                    except Exception:
                        pass
        except queue.Empty:
            pass
        # Continue polling
        self.root.after(200, self._poll_queue)

    def _read_settings(self):
        """Read settings JSON if available."""
        try:
            if SETTINGS_FILE.exists():
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            return None
        return None

    def _write_settings(self, data: dict):
        """Write settings to JSON file with exception handling."""
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            # Non-fatal: just show a warning
            messagebox.showwarning("Settings", f"Failed to save settings:\n{e}")

    def _save_settings(self):
        data = self._read_settings() or {}
        data["last_used_file"] = self.excel_path_var.get()
        data["last_selected_columns"] = [c for c, v in self.column_vars.items() if v.get() == 1]
        self._write_settings(data)

    def _load_settings(self):
        """Load settings on startup and restore last used file and selections."""
        settings = self._read_settings()
        if not settings:
            return
        last_file = settings.get("last_used_file")
        if last_file and Path(last_file).exists():
            self.excel_path_var.set(last_file)
            # Try to load the excel quietly
            try:
                self._load_excel(last_file)
            except Exception:
                pass
        # last_selected_columns will be applied when checklist is populated


def main():
    root = tb.Window(themename="superhero")
    app = ExcelSplitterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
