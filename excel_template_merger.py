# -*- coding: utf-8 -*-
"""
excel_template_merger.py – v3.0
--------------------------------
• 先用映射文件扫整表，找到第一行 ≥1 列命中的行作为真正表头
• fuzzy rename：忽略空白/全角空格/换行，允许“包含关系”
• 未映射列 / 未命中源表头都会打印警告
• 仅保留命中≥1列的文件，其余跳过
Python 3.9+    pip install pandas openpyxl xlrd<2.0.0
"""

from __future__ import annotations

import logging
import pathlib
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Dict, Iterable, List, Tuple

import pandas as pd
import numpy as np  # 新增

# -----------------------------------------------------------------------------
# 日志
# -----------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format="%(levelname)-4s | %(message)s")
# -----------------------------------------------------------------------------#
#  通用工具 + 旧 .et 自动转换                                                  #
# -----------------------------------------------------------------------------#
import logging, pathlib, re, zipfile, os
from typing import Dict, Iterable, List, Tuple

import numpy as np
import pandas as pd

# ---------- optional COM -----------
try:
    import win32com.client as win32  # 仅 Windows 安装 pywin32 时成功
except ImportError:
    win32 = None

_WS_RE = re.compile(r"[\s\u3000\r\n]+")


def clean(s) -> str:
    return _WS_RE.sub("", str(s).strip())


# ---------- WPS COM 转换旧 .et ----------
def _convert_et_via_wps(et_path: pathlib.Path) -> pathlib.Path | None:
    """调用本机 WPS 将旧 .et 另存为 .xlsx；成功返回新路径，否则 None"""
    if os.name != "nt" or win32 is None:
        return None
    try:
        xl_path = et_path.with_suffix(".xlsx")
        app = win32.DispatchEx("ET.Application")
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(str(et_path))
        wb.SaveAs(str(xl_path), FileFormat=51)  # 51 = xlOpenXMLWorkbook
        wb.Close(False)
        app.Quit()
        return xl_path if xl_path.exists() else None
    except Exception as e:
        logging.warning(f"自动转换 {et_path.name} 失败: {e}")
        return None


# ---------- 统一读取 ----------
def read_excel_auto(
    path: pathlib.Path,
    *,
    header: int | None = 0,
    nrows: int | None = None,
    sheet_name: str | int | List | None = 0,
):
    kw = dict(header=header, nrows=nrows, sheet_name=sheet_name)
    ext = path.suffix.lower()

    # .xlsx 或 新版 .et (Zip-XML)
    if ext in {".xlsx", ".et"}:
        try:
            return pd.read_excel(path, engine="openpyxl", **kw)
        except zipfile.BadZipFile:  # 老 .et
            xl_path = _convert_et_via_wps(path)
            if xl_path:
                return pd.read_excel(xl_path, engine="openpyxl", **kw)
            raise RuntimeError(
                f"{path.name} 为旧版二进制 .et，且自动转换失败；"
                "请在 WPS 中另存为 .xlsx 再运行合并"
            )

    # .xls
    if ext == ".xls":
        try:
            return pd.read_excel(path, engine="xlrd", **kw)
        except ImportError as exc:
            raise RuntimeError('读取 .xls 需安装: pip install "xlrd<2.0.0"') from exc

    raise ValueError(f"不支持的文件类型: {path}")


# -----------------------------------------------------------------------------#
#  ExcelTemplateMerger  – backend v3.5                                         #
# -----------------------------------------------------------------------------#
class ExcelTemplateMerger:
    def __init__(self, tail_rows: int = 3, probe_rows: int = 200):
        self.tail_rows = tail_rows
        self.probe_rows = probe_rows
        self.log = logging.info

    # ---------------------------- Public API --------------------------------
    def merge(
        self,
        template_path: pathlib.Path,
        mapping_path: pathlib.Path,
        source_paths: Iterable[pathlib.Path],
        log_func=logging.info,
    ) -> pathlib.Path:
        self.log = log_func

        tpl_cols, fwd, rev = self._load_mapping(mapping_path)
        self.log(f"模板列数: {len(tpl_cols)}")

        frames: list[pd.DataFrame] = []
        for src in source_paths:
            frames.extend(self._process_workbook(src, tpl_cols, fwd, rev))

        if not frames:
            raise RuntimeError("未在任何源文件中匹配到列，无法合并")

        result = pd.concat(frames, ignore_index=True)
        self.log(f"合并完成, 总行数: {len(result)}")

        out = template_path.with_name(template_path.stem + "_filled.xlsx")
        result.to_excel(out, index=False, engine="openpyxl")
        self.log(f"已写出: {out}")
        return out

    # --------------------------- Mapping ------------------------------------
    def _load_mapping(
        self, path: pathlib.Path
    ) -> Tuple[List[str], Dict[str, str], Dict[str, str]]:
        df = read_excel_auto(path, sheet_name=0).fillna("")
        if df.shape[1] < 2:
            raise ValueError("映射文件需至少两列：模板表头 | 源表头")

        col_tpl, col_src = df.columns[:2]
        tpl_cols: list[str] = []
        fwd, rev = {}, {}
        for tpl, src in zip(df[col_tpl], df[col_src]):
            t, s = clean(tpl), clean(src)
            if not t or not s:
                continue
            if t not in tpl_cols:
                tpl_cols.append(t)
            fwd[s] = t
            rev[t] = s
        if not fwd:
            raise ValueError("映射文件未检测到有效映射对")
        self.log(f"读取映射文件, 映射对 {len(fwd)}")
        return tpl_cols, fwd, rev

    # --------------------------- Header Probe -------------------------------
    @staticmethod
    def _probe_header_row(df_probe: pd.DataFrame, keys: set[str]) -> int:
        for i, row in df_probe.iterrows():
            cleaned = [clean(c) for c in row if pd.notna(c)]
            if any(c in keys for c in cleaned):
                return i
        return -1

    # --------------------------- Fuzzy Rename -------------------------------
    @staticmethod
    def _fuzzy_rename(
        df: pd.DataFrame, fwd: Dict[str, str], rev: Dict[str, str]
    ) -> Tuple[pd.DataFrame, int]:
        pool_fwd, pool_rev = set(fwd), set(rev)
        ren_fwd, hit_fwd, ren_rev, hit_rev = {}, 0, {}, 0

        for col in df.columns:
            c = clean(col)
            k = next((k for k in pool_fwd if c == k or k in c or c in k), None)
            if k:
                ren_fwd[col] = fwd[k]; hit_fwd += 1
            k2 = next((k for k in pool_rev if c == k or k in c or c in k), None)
            if k2:
                ren_rev[col] = rev[k2]; hit_rev += 1

        if hit_fwd >= hit_rev and hit_fwd:
            return df.rename(columns=ren_fwd), hit_fwd
        if hit_rev:
            return df.rename(columns=ren_rev), hit_rev
        return df, 0

    # ------------------------- Process Workbook ----------------------------
    def _process_workbook(
        self,
        wb_path: pathlib.Path,
        tpl_cols: List[str],
        fwd: Dict[str, str],
        rev: Dict[str, str],
    ) -> List[pd.DataFrame]:
        self.log(f"处理 {wb_path.name} …")
        frames: list[pd.DataFrame] = []

        sheets = read_excel_auto(wb_path, header=None, sheet_name=None)
        for sheet, probe_df in sheets.items():
            hdr = self._probe_header_row(probe_df.head(self.probe_rows), set(fwd) | set(rev))
            if hdr == -1:
                self.log(f"  【{sheet}】未找到表头，跳过")
                continue

            df = read_excel_auto(wb_path, header=hdr, sheet_name=sheet)
            df.columns = [clean(c) for c in df.columns]
            df, hits = self._fuzzy_rename(df, fwd, rev)
            if hits == 0:
                self.log(f"  【{sheet}】0 列命中，跳过")
                continue

            # 去重列名
            _, first_idx = np.unique(df.columns, return_index=True)
            df = df.iloc[:, sorted(first_idx)]

            # 补缺列并对齐
            for col in tpl_cols:
                if col not in df.columns:
                    df[col] = pd.NA
            frames.append(df[tpl_cols])

            self.log(f"  【{sheet}】命中 {hits}, 缺失 {len(tpl_cols) - hits} 列 | 表头行 {hdr}")
            if not df.empty:
                self.log(f"    示例: {df.iloc[0].to_dict()}")

        return frames


# -----------------------------------------------------------------------------
# Tkinter GUI（未改动，但仍放在同一文件方便运行）
# -----------------------------------------------------------------------------
class ExcelMergerGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("批量填充模板程序")
        self.geometry("860x680")

        self.template_path: pathlib.Path | None = None
        self.mapping_path: pathlib.Path | None = None
        self.source_paths: list[pathlib.Path] = []

        self._backend = ExcelTemplateMerger()
        self._build_widgets()

    # ---------- UI ----------
    def _build_widgets(self):
        pad = {"padx": 8, "pady": 4}

        tk.Button(self, text="选择模板文件 (.xlsx/.xls/.et)", command=self._sel_tpl).pack(fill="x", **pad)
        tk.Button(self, text="导入表头映射配置", command=self._sel_map).pack(fill="x", **pad)

        tk.Label(self, text="待处理文件列表:").pack(anchor="w", **pad)
        self.listbox = tk.Listbox(self, height=15, selectmode=tk.EXTENDED)
        self.listbox.pack(fill="both", expand=True, **pad)

        mid = tk.Frame(self); mid.pack(fill="x", **pad)
        tk.Button(mid, text="添加源文件", command=self._add_src).pack(side="left", expand=True, fill="x", padx=4)
        tk.Button(mid, text="清空文件列表", command=self._clr_src).pack(side="left", expand=True, fill="x", padx=4)

        tk.Button(self, text="开始处理并覆盖模板", command=self._start).pack(fill="x", **pad)

        tk.Label(self, text="运行日志:").pack(anchor="w", **pad)
        self.text_log = tk.Text(self, height=14, state="disabled", bg="#f5f5f5")
        self.text_log.pack(fill="both", expand=False, **pad)

    # ---------- log helper ----------
    def _log(self, msg: str):
        self.text_log.configure(state="normal")
        self.text_log.insert("end", msg + "\n")
        self.text_log.see("end")
        self.text_log.configure(state="disabled")
        self.update_idletasks()
        logging.info(msg)

    # ---------- file dialogs ----------
    def _ask_xls(self, title, multi=False):
        ft = [("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        return (filedialog.askopenfilenames if multi else filedialog.askopenfilename)(
            title=title, filetypes=ft)

    # ---------- callbacks ----------
    def _sel_tpl(self):
        p = self._ask_xls("选择模板文件")
        if p:
            self.template_path = pathlib.Path(p[0] if isinstance(p, tuple) else p)
            self._log(f"模板文件: {self.template_path}")

    def _sel_map(self):
        p = self._ask_xls("导入表头映射配置")
        if p:
            self.mapping_path = pathlib.Path(p[0] if isinstance(p, tuple) else p)
            self._log(f"映射文件: {self.mapping_path}")

    def _add_src(self):
        paths = self._ask_xls("添加源文件", multi=True)
        for p in paths:
            path = pathlib.Path(p)
            if path not in self.source_paths:
                self.source_paths.append(path)
                self.listbox.insert("end", path.name)
        if paths:
            self._log(f"已添加 {len(paths)} 个源文件")

    def _clr_src(self):
        self.source_paths.clear()
        self.listbox.delete(0, "end")
        self._log("已清空文件列表")

    def _start(self):
        if not self.template_path or not self.mapping_path:
            messagebox.showwarning("提示", "请先选择模板文件和映射文件！")
            return
        if not self.source_paths:
            messagebox.showwarning("提示", "请至少添加一个源文件！")
            return
        try:
            out = self._backend.merge(
                template_path=self.template_path,
                mapping_path=self.mapping_path,
                source_paths=self.source_paths,
                log_func=self._log,
            )
            messagebox.showinfo("完成", f"处理完成！输出文件:\n{out}")
        except Exception as e:
            self._log(f"错误: {e}")
            messagebox.showerror("错误", str(e))


# -----------------------------------------------------------------------------
# CLI 入口
# -----------------------------------------------------------------------------
def main():
    ExcelMergerGUI().mainloop()


if __name__ == "__main__":
    main()
