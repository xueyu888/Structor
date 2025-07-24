# excel_template_merger.py
"""
Graphical helper to merge multiple Excel source workbooks into a single workbook that follows a
pre‑defined template sheet.  The program reproduces the workflow described by the user:

1. 选择模板文件（.xlsx/.xls） – the template provides the target column order.
2. 选择表头映射文件 – a two‑column Excel file whose first row is “模板表头” | “源表头”.
   Each subsequent row contains one mapping pair.
3. 添加源文件 – any number of workbooks that will be appended after their columns are
   renamed according to the mapping file.
4. 开始处理并覆盖模板 – the program concatenates all rows, *in template column order*,
   and writes <template name>_filled.xlsx next to the template file (the original is kept).

Requirements
------------
Python 3.9+, plus the following libraries:
    pip install pandas openpyxl xlrd

For legacy “.xls” files Pandas relies on the *xlrd* engine (< 2.0.0).  If your xlrd ≥ 2.0.0
cannot read .xls please install an older version:
    pip install "xlrd<2.0.0"

The program deals with Chinese file names and headers (utf‑8 on Windows) automatically.

"""
from __future__ import annotations

import pathlib
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List

import pandas as pd

# ---------------------------------------------------------------------------
# Helper functions -----------------------------------------------------------
# ---------------------------------------------------------------------------

def _read_excel(path: pathlib.Path) -> pd.DataFrame:
    """Read xlsx/xls using an appropriate Pandas engine."""
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return pd.read_excel(path, engine="openpyxl")
    elif suffix == ".xls":
        # xlrd < 2 supports .xls, otherwise inform the user
        try:
            return pd.read_excel(path, engine="xlrd")
        except ImportError as exc:
            raise RuntimeError(
                "xlrd is required to read .xls files. Install with:`pip install \"xlrd<2.0.0\"`"
            ) from exc
    else:
        raise ValueError(f"Unsupported file extension: {path}")


def load_mapping(mapping_path: pathlib.Path) -> Dict[str, str]:
    df = _read_excel(mapping_path).fillna("")
    if df.shape[1] < 2:
        raise ValueError("映射文件必须至少包含两列：模板表头 和 源表头")

    template_col, source_col = df.columns[:2]
    mapping = {}
    for tpl, src in zip(df[template_col], df[source_col]):
        tpl, src = str(tpl).strip(), str(src).strip()
        if not tpl or not src:
            continue
        mapping[src] = tpl  # rename *from* source header *to* template header
    if not mapping:
        raise ValueError("映射文件未检测到任何表头映射对")
    return mapping


def merge_files(
    template_path: pathlib.Path,
    mapping_path: pathlib.Path,
    source_paths: List[pathlib.Path],
    log_func=lambda msg: None,
) -> pathlib.Path:
    log = log_func
    log("读取模板…")
    template_df = _read_excel(template_path)
    template_cols = list(template_df.columns)
    log(f"模板列数: {len(template_cols)}")

    log("读取映射文件…")
    mapping = load_mapping(mapping_path)
    log(f"映射对数: {len(mapping)}")

    merged: list[pd.DataFrame] = []
    for p in source_paths:
        log(f"处理 {p.name} …")
        df = _read_excel(p)
        # rename columns according to mapping
        df = df.rename(columns=mapping)
        # keep only template columns; add missing columns if necessary
        for col in template_cols:
            if col not in df.columns:
                df[col] = pd.NA
        df = df[template_cols]
        merged.append(df)
        log(f"  行数: {len(df)}")

    if not merged:
        raise RuntimeError("未选择任何源文件")

    result_df = pd.concat(merged, ignore_index=True)
    log(f"合并完成, 总行数: {len(result_df)}")

    out_path = template_path.with_name(template_path.stem + "_filled.xlsx")
    result_df.to_excel(out_path, index=False, engine="openpyxl")
    log(f"已写出: {out_path}")
    return out_path


# ---------------------------------------------------------------------------
# Tkinter GUI ----------------------------------------------------------------
# ---------------------------------------------------------------------------

class ExcelMergerGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("批量填充模板程序")
        self.geometry("840x660")

        # Paths
        self.template_path: pathlib.Path | None = None
        self.mapping_path: pathlib.Path | None = None
        self.source_paths: list[pathlib.Path] = []

        self._build_widgets()

    # ---------------- UI builder -----------------
    def _build_widgets(self) -> None:
        pad = {"padx": 8, "pady": 4}

        # Buttons on top
        tk.Button(self, text="选择模板文件 (.xlsx/.xls)", command=self.on_select_template).pack(fill="x", **pad)
        tk.Button(self, text="导入表头映射配置", command=self.on_select_mapping).pack(fill="x", **pad)

        # Listbox for sources
        tk.Label(self, text="待处理文件列表:").pack(anchor="w", **pad)
        self.listbox = tk.Listbox(self, height=15, selectmode=tk.EXTENDED)
        self.listbox.pack(fill="both", expand=True, **pad)

        # Buttons under listbox
        frame_mid = tk.Frame(self)
        frame_mid.pack(fill="x", **pad)
        tk.Button(frame_mid, text="添加源文件", command=self.on_add_source).pack(side="left", expand=True, fill="x", padx=4)
        tk.Button(frame_mid, text="清空文件列表", command=self.on_clear_sources).pack(side="left", expand=True, fill="x", padx=4)

        # Start button
        tk.Button(self, text="开始处理并覆盖模板", command=self.on_start).pack(fill="x", **pad)

        # Log box
        tk.Label(self, text="运行日志:").pack(anchor="w", **pad)
        self.text_log = tk.Text(self, height=12, state="disabled", bg="#f5f5f5")
        self.text_log.pack(fill="both", expand=False, **pad)

    # ---------------- Helpers -----------------
    def log(self, msg: str) -> None:
        self.text_log.configure(state="normal")
        self.text_log.insert("end", msg + "\n")
        self.text_log.see("end")
        self.text_log.configure(state="disabled")
        self.update_idletasks()

    def _ask_excel_file(self, title: str, multiple: bool = False):
        filetypes = [("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        if multiple:
            return filedialog.askopenfilenames(title=title, filetypes=filetypes)
        else:
            file = filedialog.askopenfilename(title=title, filetypes=filetypes)
            return [file] if file else []

    # ---------------- Callbacks -----------------
    def on_select_template(self):
        paths = self._ask_excel_file("选择模板文件")
        if paths:
            self.template_path = pathlib.Path(paths[0])
            self.log(f"模板文件: {self.template_path}")

    def on_select_mapping(self):
        paths = self._ask_excel_file("导入表头映射配置")
        if paths:
            self.mapping_path = pathlib.Path(paths[0])
            self.log(f"映射文件: {self.mapping_path}")

    def on_add_source(self):
        paths = self._ask_excel_file("添加源文件", multiple=True)
        for p in paths:
            path = pathlib.Path(p)
            if path not in self.source_paths:
                self.source_paths.append(path)
                self.listbox.insert("end", path.name)
        if paths:
            self.log(f"已添加 {len(paths)} 个源文件")

    def on_clear_sources(self):
        self.source_paths.clear()
        self.listbox.delete(0, "end")
        self.log("已清空文件列表")

    def on_start(self):
        if not self.template_path or not self.mapping_path:
            messagebox.showwarning("提示", "请先选择模板文件和映射文件！")
            return
        if not self.source_paths:
            messagebox.showwarning("提示", "请至少添加一个源文件！")
            return
        try:
            out_path = merge_files(
                template_path=self.template_path,
                mapping_path=self.mapping_path,
                source_paths=self.source_paths,
                log_func=self.log,
            )
            messagebox.showinfo("完成", f"处理完成！输出文件:\n{out_path}")
        except Exception as e:
            self.log(f"错误: {e}")
            messagebox.showerror("错误", str(e))

# ---------------------------------------------------------------------------
# CLI entry ------------------------------------------------------------------
# ---------------------------------------------------------------------------

def main():  # pragma: no cover
    app = ExcelMergerGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
