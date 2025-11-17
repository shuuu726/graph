"""Optimized graphing utility.

This module keeps the behaviour of ``graph.py`` but restructures the code to
reduce duplicated work during plotting and to simplify the user interface
layout.  ``graph.py`` is left untouched so existing workflows are not broken,
while this file exposes a lighter weight alternative implementation.
"""

from __future__ import annotations

import math
import os
import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import tkinter as tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from tkinter import filedialog, messagebox, simpledialog, ttk


# ---------------------------------------------------------------------------
# Utility dialogs
# ---------------------------------------------------------------------------


class DiameterInputDialog(simpledialog.Dialog):
    """Prompt the user for three diameter values.

    The original implementation relied on a custom ``Toplevel`` window.  Here we
    use :class:`tkinter.simpledialog.Dialog` to obtain a modal dialog while
    keeping the code compact.
    """

    def body(self, master: tk.Misc) -> Optional[tk.Widget]:
        tk.Label(master, text="各パターンの直径 (nm) を入力してください。", anchor="w").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10)
        )
        tk.Label(master, text="100nm パターン (mod 1):").grid(row=1, column=0, sticky="w")
        tk.Label(master, text="2µm パターン (mod 2):").grid(row=2, column=0, sticky="w")
        tk.Label(master, text="500nm パターン (mod 0):").grid(row=3, column=0, sticky="w")

        self.mod1_var = tk.StringVar(value="100.0")
        self.mod2_var = tk.StringVar(value="2000.0")
        self.mod0_var = tk.StringVar(value="500.0")

        tk.Entry(master, textvariable=self.mod1_var).grid(row=1, column=1)
        tk.Entry(master, textvariable=self.mod2_var).grid(row=2, column=1)
        tk.Entry(master, textvariable=self.mod0_var).grid(row=3, column=1)
        return master

    def validate(self) -> bool:
        try:
            self._result = (
                float(self.mod1_var.get()),
                float(self.mod2_var.get()),
                float(self.mod0_var.get()),
            )
        except ValueError:
            messagebox.showerror("入力エラー", "有効な数値を入力してください。", parent=self)
            return False
        return True

    def apply(self) -> None:  # pragma: no cover - handled by tkinter
        self.result = self._result


# ---------------------------------------------------------------------------
# Data handling helpers
# ---------------------------------------------------------------------------


def get_ordinal_label(n: int) -> str:
    if 10 <= n % 100 <= 13:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def parse_display_range(range_str: str) -> Sequence[int]:
    if not range_str.strip():
        return []
    result: List[int] = []
    for part in range_str.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start, end = part.split("-", 1)
            try:
                s_val = int(start)
                e_val = int(end)
            except ValueError as exc:  # pragma: no cover - defensive
                raise ValueError(f"範囲指定の形式が正しくありません: '{part}'") from exc
            if s_val > e_val:
                s_val, e_val = e_val, s_val
            result.extend(range(s_val, e_val + 1))
        else:
            try:
                result.append(int(part))
            except ValueError as exc:  # pragma: no cover - defensive
                raise ValueError(f"数値の形式が正しくありません: '{part}'") from exc
    return sorted(set(result))


def robust_excel_file(file_path: str) -> pd.ExcelFile:
    try:
        return pd.ExcelFile(file_path, engine="openpyxl")
    except Exception as first_error:  # pragma: no cover - I/O heavy
        try:
            return pd.ExcelFile(file_path, engine="xlrd")
        except Exception as second_error:
            raise Exception(
                "ファイルを開けませんでした。\n"
                f"openpyxl (xlsx) エラー: {first_error}\n"
                f"xlrd (xls) エラー: {second_error}\n"
                "ファイルが破損しているか、Excelファイルではない可能性があります。"
            ) from second_error


def robust_read_excel(file_path: str, sheet_name: str) -> pd.DataFrame:
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine="openpyxl")
    except Exception as first_error:  # pragma: no cover - I/O heavy
        try:
            return pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine="xlrd")
        except Exception as second_error:
            raise Exception(
                f"シート '{sheet_name}' を読み込めませんでした。\n"
                f"openpyxl (xlsx) エラー: {first_error}\n"
                f"xlrd (xls) エラー: {second_error}\n"
                "ファイルが破損しているか、Excelファイルではない可能性があります。"
            ) from second_error


@dataclass
class SheetData:
    dataframe: pd.DataFrame
    numeric_matrix: np.ndarray


class SheetLoader:
    """Load Excel sheets either lazily or eagerly and cache numeric matrices."""

    def __init__(self, file_path: str, load_mode: str) -> None:
        self.file_path = file_path
        self.load_mode = load_mode
        self._cache: Dict[str, SheetData] = {}
        self._sheet_names = robust_excel_file(file_path).sheet_names
        if not self._sheet_names:
            raise ValueError("選択されたファイルにシートがありません。")
        if load_mode == "bulk":
            self._preload_all_sheets()

    # ------------------------------------------------------------------
    def _preload_all_sheets(self) -> None:
        progress = ProgressDialog("データ読み込み中...", len(self._sheet_names))
        try:
            for name in self._sheet_names:
                progress.update_message(f"読み込み中: {name}")
                self._cache[name] = self._create_sheet_data(name)
                progress.step()
        finally:
            progress.close()

    # ------------------------------------------------------------------
    def _create_sheet_data(self, sheet_name: str) -> SheetData:
        df = robust_read_excel(self.file_path, sheet_name)
        numeric_df = df.apply(pd.to_numeric, errors="coerce")
        return SheetData(df, numeric_df.to_numpy())

    # ------------------------------------------------------------------
    @property
    def sheet_names(self) -> Sequence[str]:
        return self._sheet_names

    # ------------------------------------------------------------------
    def get_sheet(self, sheet_name: str) -> SheetData:
        if sheet_name not in self._cache:
            self._cache[sheet_name] = self._create_sheet_data(sheet_name)
        return self._cache[sheet_name]


# ---------------------------------------------------------------------------
# Progress feedback
# ---------------------------------------------------------------------------


class ProgressDialog:
    def __init__(self, message: str, maximum: int) -> None:
        self._root = tk.Toplevel()
        self._root.title(message)
        self._root.geometry("320x120")
        self._root.attributes("-topmost", True)
        self._root.resizable(False, False)

        self._message_var = tk.StringVar(value=message)
        ttk.Label(self._root, textvariable=self._message_var).pack(pady=(20, 5))
        self._progress = ttk.Progressbar(self._root, maximum=max(1, maximum), mode="determinate")
        self._progress.pack(fill="x", padx=20)
        self._count_var = tk.StringVar(value=f"0 / {maximum}")
        ttk.Label(self._root, textvariable=self._count_var).pack(pady=(5, 10))
        self._current = 0
        self._maximum = max(1, maximum)
        self._root.update_idletasks()

    def update_message(self, text: str) -> None:
        self._message_var.set(text)
        self._root.update_idletasks()

    def step(self) -> None:
        self._current += 1
        self._progress["value"] = self._current
        self._count_var.set(f"{self._current} / {self._maximum}")
        self._root.update_idletasks()

    def close(self) -> None:
        self._root.destroy()


# ---------------------------------------------------------------------------
# Plotting helpers
# ---------------------------------------------------------------------------


def calc_area_cm2(diameter_nm: float) -> float:
    radius_nm = diameter_nm / 2.0
    return math.pi * (radius_nm ** 2) * 1e-14


def infer_diameter(sheet_name: str, mod_values: Tuple[float, float, float]) -> float:
    match = re.search(r"(\d+)", sheet_name)
    if not match:
        return mod_values[0]
    num = int(match.group(1))
    mod = num % 3
    if mod == 1:
        return mod_values[0]
    if mod == 2:
        return mod_values[1]
    return mod_values[2]


def apply_voltage_filter(x: np.ndarray, y: np.ndarray, vmin: Optional[float], vmax: Optional[float]) -> Tuple[np.ndarray, np.ndarray]:
    mask = ~np.isnan(x) & ~np.isnan(y)
    if vmin is not None:
        mask &= x >= vmin
    if vmax is not None:
        mask &= x <= vmax
    filtered_x = x[mask]
    filtered_y = y[mask]
    return filtered_x, filtered_y


# ---------------------------------------------------------------------------
# Classification
# ---------------------------------------------------------------------------


CATEGORIES: Sequence[str] = (
    "1回目に破壊",
    "2回目に破壊",
    "貫通",
    "複数回NDR",
    "複数回電流減少",
    "特性のばらつき",
)


@dataclass
class ClassificationEntry:
    sheet_name: str
    category: str
    diameter_nm: float


# ---------------------------------------------------------------------------
# Main controller
# ---------------------------------------------------------------------------


class GraphController:
    def __init__(self, root: tk.Tk, file_path: str, loader: SheetLoader, diameters: Tuple[float, float, float]):
        self.root = root
        self.file_path = file_path
        self.loader = loader
        self.diameters = diameters

        self.classified: List[ClassificationEntry] = []

        self.root.title(f"グラフ表示 - {os.path.basename(file_path)} (mode: {loader.load_mode})")
        self.root.geometry("1200x720")

        self.sheet_var = tk.StringVar(value=self.loader.sheet_names[0])
        self.display_counts_var = tk.StringVar()
        self.vmin_var = tk.StringVar()
        self.vmax_var = tk.StringVar()
        self.legend_visible = tk.BooleanVar(value=True)
        self.legend_loc = tk.StringVar(value="best")
        self.line_width = tk.DoubleVar(value=1.5)

        self.title_font = tk.IntVar(value=18)
        self.label_font = tk.IntVar(value=16)
        self.tick_font = tk.IntVar(value=12)
        self.legend_font = tk.IntVar(value=11)

        self.status_var = tk.StringVar(value="")
        self.active_area_var = tk.StringVar()

        self._build_layout()
        self._register_events()
        self._init_plot()
        self.update_plot()

    # ------------------------------------------------------------------
    def _build_layout(self) -> None:
        paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        plot_frame = ttk.Frame(paned)
        control_frame = ttk.Frame(paned, width=360)
        paned.add(plot_frame, weight=4)
        paned.add(control_frame, weight=1)

        self.figure, self.ax = plt.subplots(figsize=(8, 5))
        self.canvas = FigureCanvasTkAgg(self.figure, master=plot_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        notebook = ttk.Notebook(control_frame)
        notebook.pack(fill=tk.BOTH, expand=True)

        self.basic_tab = ttk.Frame(notebook)
        self.style_tab = ttk.Frame(notebook)
        self.classification_tab = ttk.Frame(notebook)
        notebook.add(self.basic_tab, text="基本設定")
        notebook.add(self.style_tab, text="表示調整")
        notebook.add(self.classification_tab, text="分類")

        self._build_basic_tab()
        self._build_style_tab()
        self._build_classification_tab()

        status_frame = ttk.Frame(control_frame)
        status_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(status_frame, textvariable=self.status_var, foreground="navy").pack(anchor="w")
        ttk.Label(status_frame, textvariable=self.active_area_var).pack(anchor="w")

        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill=tk.X, pady=(5, 10))
        ttk.Button(button_frame, text="グラフを保存", command=self.save_graph).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="再読み込み", command=self.update_plot).pack(side=tk.LEFT, padx=5)

    # ------------------------------------------------------------------
    def _build_basic_tab(self) -> None:
        ttk.Label(self.basic_tab, text="シート:").grid(row=0, column=0, sticky="w")
        self.sheet_combo = ttk.Combobox(
            self.basic_tab,
            textvariable=self.sheet_var,
            values=list(self.loader.sheet_names),
            state="readonly",
        )
        self.sheet_combo.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=2)

        ttk.Label(self.basic_tab, text="表示する回数 (例: 1-3,5):").grid(row=1, column=0, sticky="w")
        ttk.Entry(self.basic_tab, textvariable=self.display_counts_var).grid(
            row=1, column=1, sticky="ew", padx=(0, 5), pady=2
        )

        ttk.Label(self.basic_tab, text="電圧下限 [V]:").grid(row=2, column=0, sticky="w")
        ttk.Entry(self.basic_tab, textvariable=self.vmin_var).grid(row=2, column=1, sticky="ew", padx=(0, 5), pady=2)
        ttk.Label(self.basic_tab, text="電圧上限 [V]:").grid(row=3, column=0, sticky="w")
        ttk.Entry(self.basic_tab, textvariable=self.vmax_var).grid(row=3, column=1, sticky="ew", padx=(0, 5), pady=2)

        ttk.Separator(self.basic_tab).grid(row=4, column=0, columnspan=2, sticky="ew", pady=6)

        extract_frame = ttk.LabelFrame(self.basic_tab, text="データ抽出")
        extract_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=4)
        ttk.Label(extract_frame, text="抽出電圧 [V]:").grid(row=0, column=0, sticky="w")
        self.extract_voltage = tk.StringVar(value="1.0")
        ttk.Entry(extract_frame, textvariable=self.extract_voltage, width=8).grid(row=0, column=1, sticky="w")
        ttk.Button(extract_frame, text="抽出して保存", command=self.extract_and_save).grid(row=0, column=2, padx=5)

        self.basic_tab.columnconfigure(1, weight=1)

    # ------------------------------------------------------------------
    def _build_style_tab(self) -> None:
        ttk.Label(self.style_tab, text="凡例の位置:").grid(row=0, column=0, sticky="w")
        legend_combo = ttk.Combobox(
            self.style_tab,
            textvariable=self.legend_loc,
            values=("best", "upper right", "upper left", "lower left", "lower right", "right", "center"),
            state="readonly",
        )
        legend_combo.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=2)
        ttk.Checkbutton(self.style_tab, text="凡例を表示", variable=self.legend_visible).grid(row=0, column=2, padx=5)

        ttk.Label(self.style_tab, text="線の太さ").grid(row=1, column=0, sticky="w")
        ttk.Spinbox(self.style_tab, from_=0.5, to=5.0, increment=0.1, textvariable=self.line_width).grid(
            row=1, column=1, sticky="w"
        )

        font_frame = ttk.LabelFrame(self.style_tab, text="フォントサイズ")
        font_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=6)
        ttk.Label(font_frame, text="タイトル").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(font_frame, from_=8, to=32, textvariable=self.title_font).grid(row=0, column=1, sticky="w")
        ttk.Label(font_frame, text="軸ラベル").grid(row=1, column=0, sticky="w")
        ttk.Spinbox(font_frame, from_=8, to=32, textvariable=self.label_font).grid(row=1, column=1, sticky="w")
        ttk.Label(font_frame, text="目盛り").grid(row=2, column=0, sticky="w")
        ttk.Spinbox(font_frame, from_=8, to=32, textvariable=self.tick_font).grid(row=2, column=1, sticky="w")
        ttk.Label(font_frame, text="凡例").grid(row=3, column=0, sticky="w")
        ttk.Spinbox(font_frame, from_=8, to=24, textvariable=self.legend_font).grid(row=3, column=1, sticky="w")

        self.style_tab.columnconfigure(1, weight=1)

    # ------------------------------------------------------------------
    def _build_classification_tab(self) -> None:
        control_frame = ttk.Frame(self.classification_tab)
        control_frame.pack(fill=tk.X, pady=5)
        ttk.Label(control_frame, text="分類:").pack(side=tk.LEFT)
        self.category_var = tk.StringVar(value=CATEGORIES[0])
        ttk.Combobox(control_frame, textvariable=self.category_var, values=list(CATEGORIES), state="readonly").pack(
            side=tk.LEFT, padx=5
        )
        ttk.Button(control_frame, text="現在のシートを追加", command=self.add_classification).pack(side=tk.LEFT, padx=5)

        tree_frame = ttk.Frame(self.classification_tab)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        columns = ("sheet", "category", "diameter")
        self.class_tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        self.class_tree.heading("sheet", text="シート")
        self.class_tree.heading("category", text="カテゴリ")
        self.class_tree.heading("diameter", text="直径 [nm]")
        self.class_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        summary_frame = ttk.LabelFrame(self.classification_tab, text="サマリー")
        summary_frame.pack(fill=tk.X, padx=5, pady=5)
        self.summary_labels: List[List[tk.StringVar]] = [
            [tk.StringVar(value="0") for _ in range(3)] for _ in CATEGORIES
        ]
        header = ttk.Frame(summary_frame)
        header.grid(row=0, column=0, columnspan=4, pady=(0, 5))
        ttk.Label(header, text="カテゴリ", width=20).grid(row=0, column=0)
        ttk.Label(header, text="100nm").grid(row=0, column=1)
        ttk.Label(header, text="2µm").grid(row=0, column=2)
        ttk.Label(header, text="500nm").grid(row=0, column=3)

        for idx, category in enumerate(CATEGORIES):
            ttk.Label(summary_frame, text=category, width=20).grid(row=idx + 1, column=0, sticky="w")
            for col in range(3):
                ttk.Label(summary_frame, textvariable=self.summary_labels[idx][col]).grid(row=idx + 1, column=col + 1)

        ttk.Button(self.classification_tab, text="結果を保存", command=self.save_classification_summary).pack(pady=5)

    # ------------------------------------------------------------------
    def _register_events(self) -> None:
        self.sheet_var.trace_add("write", lambda *_: self.update_plot())
        self.display_counts_var.trace_add("write", lambda *_: self.update_plot())
        self.vmin_var.trace_add("write", lambda *_: self.update_plot())
        self.vmax_var.trace_add("write", lambda *_: self.update_plot())
        self.legend_visible.trace_add("write", lambda *_: self._apply_axis_settings())
        self.legend_loc.trace_add("write", lambda *_: self._apply_axis_settings())
        self.line_width.trace_add("write", lambda *_: self.update_plot())
        self.title_font.trace_add("write", lambda *_: self._apply_axis_settings())
        self.label_font.trace_add("write", lambda *_: self._apply_axis_settings())
        self.tick_font.trace_add("write", lambda *_: self._apply_axis_settings())
        self.legend_font.trace_add("write", lambda *_: self._apply_axis_settings())

    # ------------------------------------------------------------------
    def _init_plot(self) -> None:
        self.annot = self.ax.annotate(
            "",
            xy=(0, 0),
            xytext=(10, 10),
            textcoords="offset points",
            bbox=dict(boxstyle="round", fc="w"),
            arrowprops=dict(arrowstyle="->"),
        )
        self.annot.set_visible(False)
        self.canvas.mpl_connect("motion_notify_event", self._on_hover)
        self.canvas.mpl_connect("button_press_event", self._on_click)

    # ------------------------------------------------------------------
    def _on_hover(self, event):  # pragma: no cover - matplotlib interaction
        if event.inaxes != self.ax:
            if self.annot.get_visible():
                self.annot.set_visible(False)
                self.canvas.draw_idle()
            return
        for line in self.ax.lines:
            contains, info = line.contains(event)
            if not contains:
                continue
            idx = info["ind"][0]
            x_data, y_data = line.get_data()
            pos = (x_data[idx], y_data[idx])
            self.annot.xy = pos
            area_cm2 = calc_area_cm2(self._current_diameter)
            current = pos[1] * 1e6 * area_cm2
            text = (
                f"{line.get_label()}\n"
                f"Voltage: {pos[0]:.3f} V\n"
                f"Current: {current:.3e} A\n"
                f"Current Density: {pos[1]:.4g} MA/cm²"
            )
            self.annot.set_text(text)
            self.annot.set_visible(True)
            self.canvas.draw_idle()
            break
        else:
            if self.annot.get_visible():
                self.annot.set_visible(False)
                self.canvas.draw_idle()

    # ------------------------------------------------------------------
    def _on_click(self, event):  # pragma: no cover - matplotlib interaction
        if event.inaxes != self.ax or event.button != 1:
            return
        for line in self.ax.lines:
            contains, info = line.contains(event)
            if not contains:
                continue
            idx = info["ind"][0]
            value = line.get_ydata()[idx]
            self.root.clipboard_clear()
            self.root.clipboard_append(f"{value:.6g}")
            self.status_var.set(f"コピーしました: {value:.6g} MA/cm²")
            self.root.after(2000, lambda: self.status_var.set(""))
            break

    # ------------------------------------------------------------------
    def update_plot(self) -> None:
        try:
            sheet_name = self.sheet_var.get()
            sheet_data = self.loader.get_sheet(sheet_name)
            counts = parse_display_range(self.display_counts_var.get())
        except ValueError as exc:
            messagebox.showerror("入力エラー", str(exc), parent=self.root)
            return

        vmin = self._safe_float(self.vmin_var.get())
        vmax = self._safe_float(self.vmax_var.get())

        self.ax.clear()
        matrix = sheet_data.numeric_matrix
        if matrix.shape[1] < 2:
            self.ax.text(0.5, 0.5, "グラフを描画する列が不足しています。", ha="center", va="center", transform=self.ax.transAxes)
            self.canvas.draw_idle()
            return

        self._current_diameter = infer_diameter(sheet_name, self.diameters)
        area_cm2 = calc_area_cm2(self._current_diameter)
        self.active_area_var.set(f"有効面積: {area_cm2:.4g} cm² ({self._current_diameter:.1f} nm)")

        selected_indices = counts or None
        plotted = 0
        for column in range(0, matrix.shape[1], 2):
            if column + 1 >= matrix.shape[1]:
                break
            iteration = column // 2 + 1
            if selected_indices and iteration not in selected_indices:
                continue
            x_raw = matrix[:, column]
            y_raw = matrix[:, column + 1]
            x, y = apply_voltage_filter(x_raw, y_raw, vmin, vmax)
            if not len(x):
                continue
            y_density = (y / area_cm2) * 1e-6
            self.ax.plot(x, y_density, linewidth=self.line_width.get(), label=get_ordinal_label(iteration))
            plotted += 1

        if not plotted:
            self.ax.text(0.5, 0.5, "表示可能なデータが見つかりませんでした。", ha="center", va="center", transform=self.ax.transAxes)
        else:
            self.ax.set_xlabel("Voltage [V]")
            self.ax.set_ylabel("Current Density [MA/cm²]")
            self.ax.grid(True, which="both", linestyle="--", linewidth=0.5)

        self._apply_axis_settings()
        self.canvas.draw_idle()

    # ------------------------------------------------------------------
    def _apply_axis_settings(self) -> None:
        self.ax.set_title(self.sheet_var.get(), fontsize=self.title_font.get())
        self.ax.xaxis.label.set_size(self.label_font.get())
        self.ax.yaxis.label.set_size(self.label_font.get())
        for tick in self.ax.xaxis.get_ticklabels():
            tick.set_fontsize(self.tick_font.get())
        for tick in self.ax.yaxis.get_ticklabels():
            tick.set_fontsize(self.tick_font.get())
        if self.legend_visible.get() and self.ax.lines:
            self.ax.legend(fontsize=self.legend_font.get(), loc=self.legend_loc.get())
        else:
            legend = self.ax.get_legend()
            if legend:
                legend.remove()
        self.canvas.draw_idle()

    # ------------------------------------------------------------------
    def save_graph(self) -> None:
        filepath = filedialog.asksaveasfilename(
            title="グラフを保存",
            defaultextension=".png",
            filetypes=(
                ("PNG", "*.png"),
                ("JPEG", "*.jpeg"),
                ("SVG", "*.svg"),
                ("PDF", "*.pdf"),
                ("すべてのファイル", "*.*"),
            ),
        )
        if not filepath:
            return
        try:
            self.figure.tight_layout()
            self.figure.savefig(filepath, bbox_inches="tight")
            messagebox.showinfo("完了", f"{filepath} に保存しました。", parent=self.root)
        except Exception as exc:  # pragma: no cover - filesystem errors
            messagebox.showerror("エラー", f"ファイルの保存中にエラーが発生しました:\n{exc}", parent=self.root)

    # ------------------------------------------------------------------
    def extract_and_save(self) -> None:
        try:
            target_voltage = float(self.extract_voltage.get())
        except ValueError:
            messagebox.showerror("エラー", "抽出電圧には有効な数値を入力してください。", parent=self.root)
            return

        sheet_name = self.sheet_var.get()
        sheet_data = self.loader.get_sheet(sheet_name)
        area_cm2 = calc_area_cm2(infer_diameter(sheet_name, self.diameters))

        matrix = sheet_data.numeric_matrix
        if matrix.shape[1] < 2:
            messagebox.showwarning("警告", "抽出できるデータがありません。", parent=self.root)
            return

        rows: List[Tuple[float, float, float]] = []
        for column in range(0, matrix.shape[1], 2):
            if column + 1 >= matrix.shape[1]:
                break
            iteration = column // 2 + 1
            x = matrix[:, column]
            y = matrix[:, column + 1]
            if np.isnan(x).all() or np.isnan(y).all():
                continue
            idx = np.nanargmin(np.abs(x - target_voltage))
            voltage = x[idx]
            current = y[idx]
            density = (current / area_cm2) * 1e-6
            rows.append((iteration, voltage, density))

        if not rows:
            messagebox.showinfo("情報", "指定した電圧に近いデータがありませんでした。", parent=self.root)
            return

        filepath = filedialog.asksaveasfilename(
            title="抽出データを保存",
            defaultextension=".xlsx",
            filetypes=(("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")),
        )
        if not filepath:
            return

        df = pd.DataFrame(rows, columns=["Count", "Voltage [V]", "Current Density [MA/cm²]"])
        try:
            df.to_excel(filepath, index=False)
            messagebox.showinfo("完了", f"抽出データを {filepath} に保存しました。", parent=self.root)
        except Exception as exc:  # pragma: no cover - filesystem errors
            messagebox.showerror("保存エラー", str(exc), parent=self.root)

    # ------------------------------------------------------------------
    def add_classification(self) -> None:
        sheet_name = self.sheet_var.get()
        category = self.category_var.get()
        diameter_nm = infer_diameter(sheet_name, self.diameters)

        for entry in self.classified:
            if entry.sheet_name == sheet_name:
                entry.category = category
                entry.diameter_nm = diameter_nm
                break
        else:
            self.classified.append(ClassificationEntry(sheet_name, category, diameter_nm))

        self._refresh_classification_view()
        self.status_var.set(f"'{sheet_name}' を '{category}' に分類しました。")
        self.root.after(3000, lambda: self.status_var.set(""))

    # ------------------------------------------------------------------
    def _refresh_classification_view(self) -> None:
        for item in self.class_tree.get_children():
            self.class_tree.delete(item)
        for entry in self.classified:
            self.class_tree.insert("", tk.END, values=(entry.sheet_name, entry.category, f"{entry.diameter_nm:.1f}"))
        counts = [[0, 0, 0] for _ in CATEGORIES]
        diameter_order = {
            self.diameters[0]: 0,
            self.diameters[1]: 1,
            self.diameters[2]: 2,
        }
        for entry in self.classified:
            if entry.category in CATEGORIES:
                row = CATEGORIES.index(entry.category)
                column = diameter_order.get(entry.diameter_nm, 0)
                counts[row][column] += 1
        for row_idx, vars_row in enumerate(self.summary_labels):
            for col_idx, var in enumerate(vars_row):
                var.set(str(counts[row_idx][col_idx]))

    # ------------------------------------------------------------------
    def save_classification_summary(self) -> None:
        if not self.classified:
            messagebox.showinfo("情報", "分類リストにデータがありません。", parent=self.root)
            return
        directory = filedialog.askdirectory(title="保存先のフォルダを選択してください")
        if not directory:
            return
        grouped: Dict[str, List[ClassificationEntry]] = {}
        for entry in self.classified:
            grouped.setdefault(entry.category, []).append(entry)

        for category, entries in grouped.items():
            file_path = os.path.join(directory, f"{category}.xlsx")
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                used_names: Dict[str, int] = {}
                for entry in entries:
                    sheet_data = self.loader.get_sheet(entry.sheet_name)
                    base_name = entry.sheet_name
                    if base_name not in used_names:
                        used_names[base_name] = 1
                        final_name = base_name
                    else:
                        used_names[base_name] += 1
                        suffix = f"({used_names[base_name]})"
                        final_name = f"{base_name[: 31 - len(suffix)]}{suffix}"
                    sheet_data.dataframe.to_excel(writer, sheet_name=final_name, index=False, header=False)
        messagebox.showinfo("完了", "分類結果を保存しました。", parent=self.root)

    # ------------------------------------------------------------------
    @staticmethod
    def _safe_float(value: str) -> Optional[float]:
        value = value.strip()
        if not value:
            return None
        try:
            return float(value)
        except ValueError:
            return None


# ---------------------------------------------------------------------------
# Application entry points
# ---------------------------------------------------------------------------


def select_excel_file(root: tk.Tk) -> Optional[str]:
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(
        title="グラフを作成するExcelファイルを選択してください",
        filetypes=(("Excelファイル", "*.xlsx *.xls *.xlsm *.xlsb"), ("すべてのファイル", "*.*")),
    )
    root.attributes("-topmost", False)
    return file_path or None


def request_load_mode(root: tk.Tk) -> str:
    result = messagebox.askyesno(
        "読み込みモード",
        "全てのシートをあらかじめ読み込みますか？\n(はい: 一括読み込み / いいえ: 必要時読み込み)",
        parent=root,
    )
    return "bulk" if result else "on_demand"


def start_graph_tool() -> None:
    root = tk.Tk()
    root.withdraw()

    file_path = select_excel_file(root)
    if not file_path:
        root.destroy()
        return

    dialog = DiameterInputDialog(root, title="直径を入力")
    diameters = dialog.result
    if not diameters:
        messagebox.showinfo("キャンセル", "直径が入力されなかったため、処理を中断します。", parent=root)
        root.destroy()
        return

    load_mode = request_load_mode(root)
    try:
        loader = SheetLoader(file_path, load_mode)
    except Exception as exc:
        messagebox.showerror("エラー", str(exc), parent=root)
        root.destroy()
        return

    root.deiconify()
    GraphController(root, file_path, loader, diameters)
    root.mainloop()


if __name__ == "__main__":  # pragma: no cover - interactive entry point
    start_graph_tool()

