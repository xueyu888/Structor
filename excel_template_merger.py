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

# -----------------------------------------------------------------------------
# 日志
# -----------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format="%(levelname)-4s | %(message)s")

# -----------------------------------------------------------------------------
# 通用工具
# -----------------------------------------------------------------------------
_WS_RE = re.compile(r"[\s\u3000\r\n]+")        # 普通空格、全角空格、回车、换行


def clean(s) -> str:
    """转为 str 并去空白字符"""
    return _WS_RE.sub("", str(s).strip())


def read_excel_auto(path: pathlib.Path, *, header=None, nrows=None) -> pd.DataFrame:
    suf = path.suffix.lower()
    if suf == ".xlsx":
        return pd.read_excel(path, engine="openpyxl", header=header, nrows=nrows)
    if suf == ".xls":
        try:
            return pd.read_excel(path, engine="xlrd", header=header, nrows=nrows)
        except ImportError as exc:
            raise RuntimeError('读取 .xls 需安装:  pip install "xlrd<2.0.0"') from exc
    raise ValueError(f"不支持的文件类型: {path}")


# -----------------------------------------------------------------------------
# 核心合并器（保持与 GUI 调用接口一致）
# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------#
#  核心合并器  ExcelTemplateMerger                                             #
# -----------------------------------------------------------------------------#
class ExcelTemplateMerger:
    """
    工作流程
    =========
    1. 读取并解析映射文件（源→模板、模板→源 双向 dict）
    2. **自动探测表头行**（模板文件 + 每个源文件都用同一算法）：
         • 从上往下遍历，第一行 ≥1 单元格匹配映射键视为表头
    3. fuzzy rename：忽略空白/全角空格/换行，支持互相包含
    4. 对齐模板列、补缺列
    5. 合并所有命中≥1列的 DataFrame，写出
    """

    def __init__(self, tail_rows: int = 3, probe_rows: int = 200) -> None:
        self.tail_rows = tail_rows
        self.probe_rows = probe_rows
        self.log = logging.info  # 默认为 console，可被 GUI 注入

    # --------------------------- public API ---------------------------------
    def merge(
        self,
        template_path: pathlib.Path,
        mapping_path: pathlib.Path,
        source_paths: Iterable[pathlib.Path],
        log_func=logging.info,
    ) -> pathlib.Path:
        self.log = log_func

        fwd, rev = self._load_mapping(mapping_path)
        # 1) 先解析模板，拿到标准列顺序
        tpl_cols = self._load_template(template_path, fwd, rev)

        frames: list[pd.DataFrame] = []
        for src in source_paths:
            df = self._process_source(src, tpl_cols, fwd, rev)
            if df is not None:
                frames.append(df)

        if not frames:
            raise RuntimeError("所有源文件均未成功匹配表头，无法合并")

        result = pd.concat(frames, ignore_index=True)
        self.log(f"合并完成, 总行数: {len(result)}")

        out = template_path.with_name(template_path.stem + "_filled.xlsx")
        result.to_excel(out, index=False, engine="openpyxl")
        self.log(f"已写出: {out}")
        return out

    # --------------------------- mapping ------------------------------------
    def _load_mapping(self, path: pathlib.Path) -> Tuple[Dict[str, str], Dict[str, str]]:
        df = read_excel_auto(path).fillna("")
        if df.shape[1] < 2:
            raise ValueError("映射文件至少包含两列：模板表头 | 源表头")

        col_tpl, col_src = df.columns[:2]
        fwd, rev = {}, {}
        for tpl, src in zip(df[col_tpl], df[col_src]):
            t, s = clean(tpl), clean(src)
            if t and s:
                fwd[s] = t
                rev[t] = s
        if not fwd:
            raise ValueError("映射文件未检测到有效映射对")

        self.log(f"读取映射文件, 共 {len(fwd)} 对")
        return fwd, rev

    # --------------------------- header detect ------------------------------
    def _probe_header_row(self, df_probe: pd.DataFrame, keys: set[str]) -> Tuple[int, List[str]]:
        """
        按行扫描 df_probe，第一行出现 ≥1 keys 即视为表头行。
        返回 (idx, matched_list)。若未找到 → (-1, [])。
        """
        for i, row in df_probe.iterrows():
            cleaned = [clean(c) for c in row if pd.notna(c)]
            matched = [c for c in cleaned if c in keys]
            if matched:
                return i, matched
        return -1, []

    # --------------------------- fuzzy rename -------------------------------
    @staticmethod
    def _fuzzy_rename(df: pd.DataFrame, fwd: Dict[str, str], rev: Dict[str, str]) -> Tuple[pd.DataFrame, int]:
        """
        返回 (重命名后的 df, 命中列数)
        逻辑：源→模板 & 模板→源 各算一遍，取命中多的一侧
        """
        pool_fwd, pool_rev = set(fwd), set(rev)
        ren_fwd, hit_fwd = {}, 0
        ren_rev, hit_rev = {}, 0

        for col in df.columns:
            c = clean(col)
            k = next((k for k in pool_fwd if c == k or k in c or c in k), None)
            if k:
                ren_fwd[col] = fwd[k]
                hit_fwd += 1
            k2 = next((k for k in pool_rev if c == k or k in c or c in k), None)
            if k2:
                ren_rev[col] = rev[k2]
                hit_rev += 1

        if hit_fwd >= hit_rev and hit_fwd:
            return df.rename(columns=ren_fwd), hit_fwd
        if hit_rev:
            return df.rename(columns=ren_rev), hit_rev
        return df, 0

    # --------------------------- template load ------------------------------
    def _load_template(self, tpl_path: pathlib.Path, fwd: Dict[str, str], rev: Dict[str, str]) -> List[str]:
        """模板也用探测逻辑，避免标题行/空行导致列名变 0…31"""
        probe = read_excel_auto(tpl_path, header=None, nrows=self.probe_rows)
        hdr, matched = self._probe_header_row(probe, set(fwd) | set(rev))
        if hdr == -1:
            raise RuntimeError("模板文件未找到任何映射列，无法确定表头行")

        df_tpl = read_excel_auto(tpl_path, header=hdr)
        df_tpl.columns = [clean(c) for c in df_tpl.columns]
        # 如果模板列未必就是“模板表头”，也用 fuzzy_rename 把它们转成模板名
        df_tpl, _ = self._fuzzy_rename(df_tpl, rev, fwd)  # 注意方向反过来
        cols = [clean(c) for c in df_tpl.columns]
        self.log(f"读取模板, 表头行 {hdr}, 列 {cols}")
        return cols

    # --------------------------- single source ------------------------------
    def _process_source(
        self,
        src_path: pathlib.Path,
        tpl_cols: List[str],
        fwd: Dict[str, str],
        rev: Dict[str, str],
    ) -> pd.DataFrame | None:
        self.log(f"处理 {src_path.name} …")

        probe = read_excel_auto(src_path, header=None, nrows=self.probe_rows)
        hdr_row, matched = self._probe_header_row(probe, set(fwd) | set(rev))
        if hdr_row == -1:
            self.log("  ⚠️ 未找到任何映射表头，跳过")
            return None
        self.log(f"  表头行: 第 {hdr_row} 行, 初步命中 {matched}")

        df = read_excel_auto(src_path, header=hdr_row)
        df.columns = [clean(c) for c in df.columns]
        df, hits = self._fuzzy_rename(df, fwd, rev)

        if hits == 0:
            self.log("  ⚠️ 0 列命中，跳过")
            return None

        miss = [c for c in tpl_cols if c not in df.columns]
        self.log(f"  命中 {hits} 列, 缺失 {len(miss)} 列")
        if miss:
            self.log(f"  ⚠️ 缺失模板列: {miss}")

        # 打印示例行（若有数据）
        if not df.empty:
            sample = df.iloc[0].to_dict()
            self.log(f"  表头示例: {sample}")

        # 补列并对齐
        for col in tpl_cols:
            if col not in df.columns:
                df[col] = pd.NA
        return df[tpl_cols]


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

        tk.Button(self, text="选择模板文件 (.xlsx/.xls)", command=self._sel_tpl).pack(fill="x", **pad)
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
