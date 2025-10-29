import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import sys
import math
import re
import numpy as np
import os

# ---- 直径入力用のカスタムダイアログ ----
class DiameterInputDialog(tk.Toplevel):
    """ファイル選択直後に直径を入力させるためのモーダルダイアログ"""
    def __init__(self, parent):
        super().__init__(parent)
        self.title("直径を入力")
        self.parent = parent
        self.result = None

        self.lift()
        self.attributes('-topmost', True)

        self.transient(parent)
        self.grab_set()

        # ウィジェットの作成
        tk.Label(self, text="各パターンの直径 (nm) を入力してください。").grid(
            row=0, column=0, columnspan=2, padx=10, pady=10)

        tk.Label(self, text="100nm パターン (mod 1):").grid(row=1, column=0, sticky='w', padx=10, pady=5)
        self.mod1_entry = tk.Entry(self)
        self.mod1_entry.grid(row=1, column=1, padx=10, pady=5)
        self.mod1_entry.insert(0, "100.0")

        tk.Label(self, text="2µm パターン (mod 2):").grid(row=2, column=0, sticky='w', padx=10, pady=5)
        self.mod2_entry = tk.Entry(self)
        self.mod2_entry.grid(row=2, column=1, padx=10, pady=5)
        self.mod2_entry.insert(0, "2000.0")

        tk.Label(self, text="500nm パターン (mod 0):").grid(row=3, column=0, sticky='w', padx=10, pady=5)
        self.mod0_entry = tk.Entry(self)
        self.mod0_entry.grid(row=3, column=1, padx=10, pady=5)
        self.mod0_entry.insert(0, "500.0")

        ok_button = tk.Button(self, text="OK", command=self.on_ok, width=10)
        ok_button.grid(row=4, column=0, columnspan=2, pady=(10, 15))

        self.bind("<Return>", lambda event: self.on_ok())
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.mod1_entry.focus_set()

    def on_ok(self):
        """OKボタンが押されたときの処理"""
        try:
            mod1 = self.mod1_entry.get()
            mod2 = self.mod2_entry.get()
            mod0 = self.mod0_entry.get()
            float(mod1); float(mod2); float(mod0)

            self.result = (mod1, mod2, mod0)
            self.destroy()
        except ValueError:
            messagebox.showerror("入力エラー", "有効な数値を入力してください。", parent=self)

    def on_cancel(self):
        """キャンセルまたはウィンドウが閉じられたときの処理"""
        self.result = None
        self.destroy()

    def show(self):
        """ダイアログを表示し、ユーザーの入力を待つ"""
        self.deiconify()
        self.parent.wait_window(self)
        return self.result

# ---- 共通関数 ----
def get_ordinal_label(n):
    """数値を序数（1st, 2nd, 3rd...）の文字列に変換する"""
    if 10 <= n % 100 <= 13:
        suffix = 'th'
    else:
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"

def parse_display_range(range_str):
    """表示範囲の文字列（例: "1-5,7"）を数値のセットに変換する"""
    result_set = set()
    parts = range_str.split(',')
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                if start > end:
                    start, end = end, start
                result_set.update(range(start, end + 1))
            except ValueError:
                raise ValueError(f"範囲指定の形式が正しくありません: '{part}'")
        else:
            try:
                result_set.add(int(part))
            except ValueError:
                raise ValueError(f"数値の形式が正しくありません: '{part}'")
    return result_set

def robust_excel_file(file_path):
    """
    拡張子と内容が異なるExcelファイルに対応するため、
    openpyxl と xlrd の両方を試す。
    """
    try:
        return pd.ExcelFile(file_path, engine='openpyxl')
    except Exception as e1:
        try:
            return pd.ExcelFile(file_path, engine='xlrd')
        except Exception as e2:
            raise Exception(
                f"ファイルを開けませんでした。\n"
                f"openpyxl (xlsx) エラー: {e1}\n"
                f"xlrd (xls) エラー: {e2}\n"
                "ファイルが破損しているか、Excelファイルではない可能性があります。"
            )

def robust_read_excel(file_path, sheet_name, header):
    """
    拡張子と内容が異なるExcelファイルに対応するため、
    openpyxl と xlrd の両方を試す。
    """
    try:
        return pd.read_excel(file_path, sheet_name=sheet_name, header=header, engine='openpyxl')
    except Exception as e1:
        try:
            return pd.read_excel(file_path, sheet_name=sheet_name, header=header, engine='xlrd')
        except Exception as e2:
            raise Exception(
                f"シート '{sheet_name}' を読み込めませんでした。\n"
                f"openpyxl (xlsx) エラー: {e1}\n"
                f"xlrd (xls) エラー: {e2}\n"
                "ファイルが破損しているか、Excelファイルではない可能性があります。"
            )

def load_all_sheets_with_progress(root, file_path, sheet_names):
    """全シートを読み込み、プログレスバーを表示する"""
    progress_window = tk.Toplevel(root)
    progress_window.title("データ読み込み中...")
    progress_window.geometry("350x130+400+400")
    progress_window.resizable(False, False)
    progress_window.attributes('-topmost', True) 
    
    label = tk.Label(progress_window, text="Excelファイルから全シートを読み込んでいます...", padx=10, pady=10)
    label.pack()
    
    count_label = tk.Label(progress_window, text=f"0 / {len(sheet_names)}", padx=10, pady=5)
    count_label.pack()

    progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate", maximum=len(sheet_names))
    progress_bar.pack(pady=10, padx=10)
    
    root.update()

    all_sheets_data = {}
    try:
        for i, sheet_name in enumerate(sheet_names):
            label.config(text=f"読み込み中: {sheet_name}")
            count_label.config(text=f"{i + 1} / {len(sheet_names)}")
            progress_bar['value'] = i + 1
            root.update_idletasks() 

            try:
                df = robust_read_excel(file_path, sheet_name=sheet_name, header=None)
                all_sheets_data[sheet_name] = df
            except Exception as e:
                print(f"警告: シート '{sheet_name}' の読み込みに失敗しました: {e}")
        
        progress_window.destroy()
        return all_sheets_data
    
    except Exception as e:
        progress_window.destroy()
        messagebox.showerror("読み込みエラー", f"全シートの読み込み中にエラーが発生しました:\n{e}", parent=root)
        return None


# ---- グラフ化ツール ----

def select_excel_file(root):
    """Excelファイルを選択し、ファイルパスとシート名のリストを返す"""
    try:
        root.attributes('-topmost', True)
        file_path = filedialog.askopenfilename(
            title="グラフを作成するExcelファイルを選択してください",
            filetypes=[("Excelファイル", "*.xlsx *.xls *.xlsm *.xlsb"), ("すべてのファイル", "*.*")]
        )
        root.attributes('-topmost', False)
        if not file_path:
            return None, None

        xls = robust_excel_file(file_path)

        sheet_names = xls.sheet_names
        if not sheet_names:
            messagebox.showerror("エラー", "選択されたファイルにシートがありません。")
            return None, None
        return file_path, sheet_names
    except Exception as e:
        messagebox.showerror("エラー", f"ファイル選択中にエラーが発生しました: {e}")
        return None, None

def plot_graph(df, area_cm2, display_counts_str, ax, voltage_min=None, voltage_max=None):
    """DataFrameからグラフをプロットする"""
    try:
        display_all = not display_counts_str.strip()
        display_counts = set()
        if not display_all:
            display_counts = parse_display_range(display_counts_str)
    except ValueError as e:
        messagebox.showerror("エラー", f"表示回数の入力が正しくありません。({e})")
        return None, 0

    num_columns = len(df.columns)
    plot_count = 0
    collected_x_data = []
    for i in range(0, num_columns, 2):
        if i + 1 < num_columns:
            count = (i // 2) + 1
            if display_all or count in display_counts:
                x_data = pd.to_numeric(df.iloc[:, i], errors='coerce').dropna()
                y_data_raw = pd.to_numeric(df.iloc[:, i + 1], errors='coerce').dropna()
                common_index = x_data.index.intersection(y_data_raw.index)
                x, y = x_data.loc[common_index], y_data_raw.loc[common_index]
                if voltage_min is not None:
                    mask = x >= voltage_min
                    x, y = x[mask], y[mask]
                if voltage_max is not None:
                    mask = x <= voltage_max
                    x, y = x[mask], y[mask]

                y_amp_per_cm2 = y / area_cm2
                y_final = y_amp_per_cm2 * 1e-6

                if not x.empty:
                    ax.plot(x, y_final, linestyle='-', label=get_ordinal_label(count), picker=5)
                    plot_count += 1
                    collected_x_data.append(x)

    if not plot_count and display_counts_str:
        messagebox.showinfo("情報", "グラフを描画するデータがありませんでした。指定した回数や電圧範囲を確認してください。")

    return collected_x_data, plot_count

def save_graph(fig):
    """表示されているグラフを画像ファイルとして保存する"""
    filetypes = [("PNG", "*.png"), ("JPEG", "*.jpeg"), ("SVG", "*.svg"), ("PDF", "*.pdf"), ("すべてのファイル", "*.*")]
    filepath = filedialog.asksaveasfilename(title="グラフを保存", defaultextension=".png", filetypes=filetypes)
    if filepath:
        try:
            fig.tight_layout()
            fig.savefig(filepath, bbox_inches='tight')
            messagebox.showinfo("完了", f"グラフを:\n{filepath}\nに保存しました。")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの保存中にエラーが発生しました: {e}")


def run_graphing_tool(root, load_mode):
    """アプリケーションのメインロジック（グラフ化ツール）"""
    file_path, sheet_names = select_excel_file(root)
    if not file_path:
        root.destroy()
        return

    dialog = DiameterInputDialog(root)
    initial_diams = dialog.show()

    if initial_diams is None:
        messagebox.showinfo("キャンセル", "直径が入力されなかったため、処理を中断します。")
        root.destroy()
        return

    all_sheets_data = None
    sheet_names_list = []

    if load_mode == "bulk":
        all_sheets_data = load_all_sheets_with_progress(root, file_path, sheet_names)
        if all_sheets_data is None:
            messagebox.showerror("エラー", "データ読み込みに失敗したため、ツールを終了します。")
            root.destroy()
            return
        sheet_names_list = list(all_sheets_data.keys())
    else: # load_mode == "on_demand"
        all_sheets_data = None # メモリには保持しない
        sheet_names_list = sheet_names # ファイルから読み取った生のシート名リスト
    
    if not sheet_names_list:
         messagebox.showerror("エラー", "処理対象のシートが見つかりません。")
         root.destroy()
         return

    current_area_cm2 = 0.0
    classified_sheets = []
    _is_programmatic_change = False

    graph_window = tk.Toplevel(root)
    graph_window.title(f"グラフ表示 - {file_path.split('/')[-1]} (モード: {load_mode})")
    graph_window.geometry("960x540+50+50")

    control_window = tk.Toplevel(root)
    control_window.title("コントロールパネル")

    fig, ax = plt.subplots(figsize=(8, 4.5))
    canvas = FigureCanvasTkAgg(fig, master=graph_window)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    annot = None

    def update_annot(line, ind):
        x_data, y_density_data = line.get_data()
        pos = (x_data[ind["ind"][0]], y_density_data[ind["ind"][0]])
        annot.xy = pos
        current_density_MA_per_cm2 = pos[1]
        current_density_A_per_cm2 = current_density_MA_per_cm2 * 1e6
        raw_current_value = current_density_A_per_cm2 * current_area_cm2
        data_count_label = line.get_label()
        text = (f"Data: {data_count_label}\n"
                f"Voltage: {pos[0]:.3f} V\n"
                f"Current: {raw_current_value:.3e} A\n"
                f"Current Density: {pos[1]:.4g} MA/cm²")
        annot.set_text(text)
        xlim = ax.get_xlim(); ylim = ax.get_ylim()
        offset = 25
        is_right = pos[0] > (xlim[0] + xlim[1]) / 2
        is_top = pos[1] > (ylim[0] + ylim[1]) / 2
        x_offset = -offset if is_right else offset
        y_offset = -offset if is_top else offset
        annot.xyann = (x_offset, y_offset)
        annot.set_horizontalalignment('right' if is_right else 'left')
        annot.set_verticalalignment('bottom' if is_top else 'top')

    def hover(event):
        if not annot: return
        vis = annot.get_visible()
        if event.inaxes == ax:
            for line in ax.lines:
                contains, ind = line.contains(event)
                if contains:
                    update_annot(line, ind)
                    annot.set_visible(True)
                    fig.canvas.draw_idle()
                    return
        if vis:
            annot.set_visible(False)
            fig.canvas.draw_idle()

    def on_click(event):
        if event.inaxes == ax and event.button == 1:
            for line in ax.lines:
                contains, ind = line.contains(event)
                if contains:
                    x_data, y_data = line.get_data()
                    clicked_index = ind["ind"][0]
                    current_density_value = y_data[clicked_index]
                    graph_window.clipboard_clear()
                    graph_window.clipboard_append(f"{current_density_value:.4g}")
                    status_label.config(text=f"コピーしました: {current_density_value:.4g} MA/cm²")
                    graph_window.after(2000, lambda: status_label.config(text=""))
                    return

    fig.canvas.mpl_connect("motion_notify_event", hover)
    fig.canvas.mpl_connect('button_press_event', on_click)

    control_canvas = tk.Canvas(control_window)
    scrollbar = tk.Scrollbar(control_window, orient="vertical", command=control_canvas.yview)
    control_frame = tk.Frame(control_canvas, padx=10, pady=10)
    control_canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill="y")
    control_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    control_canvas.create_window((0, 0), window=control_frame, anchor="nw")
    def on_frame_configure(event):
        control_canvas.configure(scrollregion=control_canvas.bbox("all"))
    control_frame.bind("<Configure>", on_frame_configure)

    top_bar = tk.Frame(control_frame)
    top_bar.pack(fill='x', pady=(0, 5))
    content_area = tk.Frame(control_frame)
    content_area.pack(fill='both', expand=True)
    left_column = tk.Frame(content_area)
    right_column = tk.Frame(content_area)
    data_source_frame = tk.LabelFrame(content_area, text="データソース")
    graph_control_frame = tk.Frame(content_area)
    axis_frame = tk.LabelFrame(content_area, text="軸の範囲と表示限界")
    sf_frame = tk.LabelFrame(content_area, text="軸の有効数字")
    legend_frame = tk.LabelFrame(content_area, text="凡例の設定")
    font_frame = tk.LabelFrame(content_area, text="フォントと線の設定")
    action_frame = tk.Frame(content_area)
    extract_frame = tk.LabelFrame(content_area, text="データ抽出機能")
    classification_frame = tk.LabelFrame(content_area, text="分類機能")
    col1_widgets = [data_source_frame, graph_control_frame, axis_frame, sf_frame]
    col2_widgets = [classification_frame, legend_frame, font_frame, extract_frame, action_frame]

    def toggle_layout(*args):
        left_column.pack_forget()
        right_column.pack_forget()
        for widget in col1_widgets + col2_widgets: widget.pack_forget()
        if is_two_column_var.get():
            control_window.geometry("840x680+910+50") 
            left_column.pack(side='left', fill='y', anchor='n', padx=(0, 5))
            right_column.pack(side='right', fill='y', anchor='n', padx=(5, 0))
            for widget in col1_widgets: widget.pack(in_=left_column, fill='x', pady=2)
            for widget in col2_widgets: widget.pack(in_=right_column, fill='x', pady=2)
        else:
            control_window.geometry("450x800+910+50") 
            for widget in col1_widgets + col2_widgets: widget.pack(in_=content_area, fill='x', pady=2)

    def extract_and_save_data():
        try:
            target_v_str = extract_voltage_var.get()
            if not target_v_str:
                messagebox.showwarning("入力エラー", "抽出する電圧を入力してください。")
                return
            target_v = float(target_v_str)
        except ValueError:
            messagebox.showerror("エラー", "電圧には有効な数値を入力してください。")
            return

        filepath = filedialog.asksaveasfilename(
            title="抽出データを保存", defaultextension=".xlsx",
            filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")]
        )
        if not filepath: return

        progress_window = tk.Toplevel(control_window)
        progress_window.title("データ抽出中...")
        progress_window.geometry("350x130+450+450")
        progress_window.resizable(False, False)
        progress_window.attributes('-topmost', True)
        label = tk.Label(progress_window, text="指定電圧のデータを抽出しています...", padx=10, pady=10)
        label.pack()
        
        # sheet_names_list はモードに関わらず設定済み
        total_sheets = len(sheet_names_list)
        count_label = tk.Label(progress_window, text=f"0 / {total_sheets}", padx=10, pady=5)
        count_label.pack()
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate", maximum=total_sheets)
        progress_bar.pack(pady=10, padx=10)
        control_window.update()

        try:
            diameters = {
                1: float(diameter_mod1_var.get()), 
                2: float(diameter_mod2_var.get()), 
                0: float(diameter_mod0_var.get())
            }
            
            results_mod1 = [None] * total_sheets
            results_mod2 = [None] * total_sheets
            results_mod0 = [None] * total_sheets

            for idx, sheet_name in enumerate(sheet_names_list):
                
                # --- プログレスバー更新 ---
                label.config(text=f"処理中: {sheet_name}")
                count_label.config(text=f"{idx + 1} / {total_sheets}")
                progress_bar['value'] = idx + 1
                if idx % 10 == 0 or load_mode == "on_demand": # 都度読み込みは毎回UI更新
                    control_window.update_idletasks()
                # ------------------------

                df_sheet = None
                try:
                    if load_mode == "bulk":
                        df_sheet = all_sheets_data.get(sheet_name) # 読み込み失敗シートはNone
                    else: # load_mode == "on_demand"
                        df_sheet = robust_read_excel(file_path, sheet_name=sheet_name, header=None)
                
                except Exception as e:
                    print(f"警告: シート {sheet_name} の読み込み/処理に失敗: {e}")
                    continue # 次のシートへ

                if df_sheet is None or df_sheet.shape[1] < 2: continue
                
                x_data = pd.to_numeric(df_sheet.iloc[:, 0], errors='coerce').dropna()
                y_data_raw = pd.to_numeric(df_sheet.iloc[:, 1], errors='coerce').dropna()
                common_index = x_data.index.intersection(y_data_raw.index)
                x, y = x_data.loc[common_index], y_data_raw.loc[common_index]
                
                if x.empty or len(x) < 2: continue
                
                match = re.search(r'(\d+)', sheet_name)
                mod_result = (int(match.group(1)) % 3) if match else 1
                diameter_nm = diameters.get(mod_result, diameters[1])
                area_cm2 = (math.pi * ((diameter_nm / 2) ** 2)) * 1e-14
                y_final = (y / area_cm2) * 1e-6
                
                if target_v >= x.min() and target_v <= x.max():
                    interpolated_y = np.interp(target_v, x, y_final)
                    if mod_result == 1: results_mod1[idx] = interpolated_y
                    elif mod_result == 2: results_mod2[idx] = interpolated_y
                    else: results_mod0[idx] = interpolated_y

            control_window.update_idletasks()
            progress_window.destroy()

            df_results = pd.DataFrame({
                'Sheet': sheet_names_list,
                f'直径 {diameters[1]} nm (MA/cm²)': results_mod1,
                f'直径 {diameters[2]} nm (MA/cm²)': results_mod2,
                f'直径 {diameters[0]} nm (MA/cm²)': results_mod0
            })
            df_results.to_excel(filepath, index=False, sheet_name=f'Data_at_{target_v}V', engine='openpyxl')
            messagebox.showinfo("完了", f"データ抽出が完了しました。\n{filepath}\nに保存しました。")

        except Exception as e:
            progress_window.destroy()
            messagebox.showerror("エラー", f"データ抽出中にエラーが発生しました:\n{e}")

    is_two_column_var = tk.BooleanVar(value=True)
    tk.Checkbutton(top_bar, text="2列レイアウトで表示", variable=is_two_column_var, command=toggle_layout).pack(side='left')
    tk.Label(data_source_frame, text="シート:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
    sheet_var = tk.StringVar()
    
    def next_sheet():
        current_index = sheet_names_list.index(sheet_var.get())
        sheet_var.set(sheet_names_list[(current_index + 1) % len(sheet_names_list)])
    def prev_sheet():
        current_index = sheet_names_list.index(sheet_var.get())
        sheet_var.set(sheet_names_list[(current_index - 1 + len(sheet_names_list)) % len(sheet_names_list)])
        
    prev_button = tk.Button(data_source_frame, text="<", command=prev_sheet); prev_button.grid(row=0, column=1, padx=(5,0), pady=2)
    
    sheet_menu = ttk.Combobox(data_source_frame, textvariable=sheet_var, values=sheet_names_list, state="readonly", width=15); sheet_menu.grid(row=0, column=2, sticky='ew', pady=2)
    sheet_menu.set(sheet_names_list[0])
    
    next_button = tk.Button(data_source_frame, text=">", command=next_sheet); next_button.grid(row=0, column=3, padx=(0,5), pady=2)
    tk.Label(data_source_frame, text="100nm の直径(nm):").grid(row=1, column=0, sticky='w', padx=5, pady=2)
    diameter_mod1_var = tk.StringVar(value=initial_diams[0]); tk.Entry(data_source_frame, textvariable=diameter_mod1_var, width=10).grid(row=1, column=1, sticky='w', padx=5, pady=2, columnspan=3)
    tk.Label(data_source_frame, text="2µm の直径(nm):").grid(row=2, column=0, sticky='w', padx=5, pady=2)
    diameter_mod2_var = tk.StringVar(value=initial_diams[1]); tk.Entry(data_source_frame, textvariable=diameter_mod2_var, width=10).grid(row=2, column=1, sticky='w', padx=5, pady=2, columnspan=3)
    tk.Label(data_source_frame, text="500nm の直径(nm):").grid(row=3, column=0, sticky='w', padx=5, pady=2)
    diameter_mod0_var = tk.StringVar(value=initial_diams[2]); tk.Entry(data_source_frame, textvariable=diameter_mod0_var, width=10).grid(row=3, column=1, sticky='w', padx=5, pady=2, columnspan=3)
    active_area_var = tk.StringVar()
    tk.Label(data_source_frame, text="適用中の面積:").grid(row=4, column=0, sticky='w', padx=5, pady=2)
    tk.Label(data_source_frame, textvariable=active_area_var, font=("TkDefaultFont", 10, "bold")).grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=2)
    data_source_frame.columnconfigure(2, weight=1)
    
    title_var = tk.StringVar(value=sheet_names_list[0])
    
    tk.Label(graph_control_frame, text="表示回数 (例: 1-5,7):").grid(row=0, column=0, sticky='w', padx=5, pady=2)
    counts_entry = tk.Entry(graph_control_frame, width=30)
    counts_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
    tk.Label(axis_frame, text="電圧範囲(Xデータ):").grid(row=0, column=0, sticky='w'); voltage_min_entry = tk.Entry(axis_frame, width=10); voltage_min_entry.grid(row=0, column=1); tk.Label(axis_frame, text="～").grid(row=0, column=2); voltage_max_entry = tk.Entry(axis_frame, width=10); voltage_max_entry.grid(row=0, column=3)
    tk.Label(axis_frame, text="X軸 表示最大値:").grid(row=1, column=0, sticky='w'); xlim_max_var = tk.StringVar(); xlim_max_entry = tk.Entry(axis_frame, width=10, textvariable=xlim_max_var); xlim_max_entry.grid(row=1, column=1)
    tk.Label(axis_frame, text="Y軸 表示最大値:").grid(row=2, column=0, sticky='w'); ylim_max_var = tk.StringVar(); ylim_max_entry = tk.Entry(axis_frame, width=10, textvariable=ylim_max_var); ylim_max_entry.grid(row=2, column=1)
    x_sf_var = tk.IntVar(value=2); y_sf_var = tk.IntVar(value=2)
    tk.Label(sf_frame, text="X軸 (小数点以下):").grid(row=0, column=0, sticky='w'); tk.Button(sf_frame, text="-", command=lambda: x_sf_var.set(max(0, x_sf_var.get() - 1))).grid(row=0, column=1); tk.Scale(sf_frame, from_=0, to=5, orient=tk.HORIZONTAL, showvalue=0, variable=x_sf_var).grid(row=0, column=2, sticky='ew', padx=5); tk.Button(sf_frame, text="+", command=lambda: x_sf_var.set(min(5, x_sf_var.get() + 1))).grid(row=0, column=3)
    tk.Label(sf_frame, text="Y軸 (小数点以下):").grid(row=1, column=0, sticky='w'); tk.Button(sf_frame, text="-", command=lambda: y_sf_var.set(max(0, y_sf_var.get() - 1))).grid(row=1, column=1); tk.Scale(sf_frame, from_=0, to=5, orient=tk.HORIZONTAL, showvalue=0, variable=y_sf_var).grid(row=1, column=2, sticky='ew', padx=5); tk.Button(sf_frame, text="+", command=lambda: y_sf_var.set(min(5, y_sf_var.get() + 1))).grid(row=1, column=3)
    sf_frame.columnconfigure(2, weight=1)
    legend_visible_var = tk.BooleanVar(value=True); legend_toggle = tk.Checkbutton(legend_frame, text="凡例を表示", variable=legend_visible_var); legend_toggle.grid(row=0, column=2, columnspan=2, sticky='e', padx=5)
    tk.Label(legend_frame, text="位置:").grid(row=0, column=0, sticky='w'); legend_loc_options = ['best', 'upper right', 'upper left', 'lower left', 'lower right', 'right', 'center', 'Free']
    legend_loc_var = tk.StringVar(control_window); legend_loc_var.set(legend_loc_options[0]); legend_menu = ttk.Combobox(legend_frame, textvariable=legend_loc_var, values=legend_loc_options, width=12, state="readonly"); legend_menu.grid(row=0, column=1, padx=5, sticky='w')
    legend_x_var = tk.DoubleVar(value=0.5); legend_y_var = tk.DoubleVar(value=0.5)
    def toggle_free_legend_sliders(*args):
        state = 'normal' if legend_loc_var.get() == 'Free' else 'disabled'
        for w in [legend_x_scale, legend_y_scale, btn_x_minus, btn_x_plus, btn_y_minus, btn_y_plus]: w.config(state=state)
    tk.Label(legend_frame, text="X:").grid(row=1, column=0, sticky='w'); btn_x_minus = tk.Button(legend_frame, text="-", command=lambda: legend_x_var.set(max(0.0, round(legend_x_var.get() - 0.01, 2)))); btn_x_minus.grid(row=1, column=1); legend_x_scale = tk.Scale(legend_frame, from_=0, to=1, resolution=0.01, orient=tk.HORIZONTAL, showvalue=0, variable=legend_x_var); legend_x_scale.grid(row=1, column=2, sticky='ew', padx=5); btn_x_plus = tk.Button(legend_frame, text="+", command=lambda: legend_x_var.set(min(1.0, round(legend_x_var.get() + 0.01, 2)))); btn_x_plus.grid(row=1, column=3)
    tk.Label(legend_frame, text="Y:").grid(row=2, column=0, sticky='w'); btn_y_minus = tk.Button(legend_frame, text="-", command=lambda: legend_y_var.set(max(0.0, round(legend_y_var.get() - 0.01, 2)))); btn_y_minus.grid(row=2, column=1); legend_y_scale = tk.Scale(legend_frame, from_=0, to=1, resolution=0.01, orient=tk.HORIZONTAL, showvalue=0, variable=legend_y_var); legend_y_scale.grid(row=2, column=2, sticky='ew', padx=5); btn_y_plus = tk.Button(legend_frame, text="+", command=lambda: legend_y_var.set(min(1.0, round(legend_y_var.get() + 0.01, 2)))); btn_y_plus.grid(row=2, column=3)
    tk.Label(legend_frame, text="列数:").grid(row=3, column=0, sticky='w'); legend_ncol_var = tk.IntVar(value=1); tk.Button(legend_frame, text="-", command=lambda: legend_ncol_var.set(max(1, legend_ncol_var.get() - 1))).grid(row=3, column=1); legend_ncol_scale = tk.Scale(legend_frame, from_=1, to=5, orient=tk.HORIZONTAL, showvalue=0, variable=legend_ncol_var); legend_ncol_scale.grid(row=3, column=2, padx=5, sticky='ew'); tk.Button(legend_frame, text="+", command=lambda: legend_ncol_var.set(min(5, legend_ncol_var.get() + 1))).grid(row=3, column=3)
    legend_frame.columnconfigure(2, weight=1)
    title_font_var = tk.IntVar(value=18); label_font_var = tk.IntVar(value=18); tick_font_var = tk.IntVar(value=16); legend_font_var = tk.IntVar(value=11); linewidth_var = tk.DoubleVar(value=1.5)
    tk.Label(font_frame, text="タイトル").grid(row=0, column=0, sticky='w'); tk.Button(font_frame, text="-", command=lambda: title_font_var.set(max(8, title_font_var.get() - 1))).grid(row=0, column=1); tk.Scale(font_frame, from_=8, to=30, orient=tk.HORIZONTAL, showvalue=0, variable=title_font_var).grid(row=0, column=2, sticky='ew', padx=5); tk.Button(font_frame, text="+", command=lambda: title_font_var.set(min(30, title_font_var.get() + 1))).grid(row=0, column=3)
    tk.Label(font_frame, text="軸ラベル").grid(row=1, column=0, sticky='w'); tk.Button(font_frame, text="-", command=lambda: label_font_var.set(max(8, label_font_var.get() - 1))).grid(row=1, column=1); tk.Scale(font_frame, from_=8, to=30, orient=tk.HORIZONTAL, showvalue=0, variable=label_font_var).grid(row=1, column=2, sticky='ew', padx=5); tk.Button(font_frame, text="+", command=lambda: label_font_var.set(min(30, label_font_var.get() + 1))).grid(row=1, column=3)
    tk.Label(font_frame, text="目盛り").grid(row=2, column=0, sticky='w'); tk.Button(font_frame, text="-", command=lambda: tick_font_var.set(max(8, tick_font_var.get() - 1))).grid(row=2, column=1); tk.Scale(font_frame, from_=8, to=30, orient=tk.HORIZONTAL, showvalue=0, variable=tick_font_var).grid(row=2, column=2, sticky='ew', padx=5); tk.Button(font_frame, text="+", command=lambda: tick_font_var.set(min(30, tick_font_var.get() + 1))).grid(row=2, column=3)
    tk.Label(font_frame, text="凡例").grid(row=3, column=0, sticky='w'); tk.Button(font_frame, text="-", command=lambda: legend_font_var.set(max(8, legend_font_var.get() - 1))).grid(row=3, column=1); tk.Scale(font_frame, from_=8, to=30, orient=tk.HORIZONTAL, showvalue=0, variable=legend_font_var).grid(row=3, column=2, sticky='ew', padx=5); tk.Button(font_frame, text="+", command=lambda: legend_font_var.set(min(30, legend_font_var.get() + 1))).grid(row=3, column=3)
    tk.Label(font_frame, text="線の太さ:").grid(row=4, column=0, sticky='w'); tk.Button(font_frame, text="-", command=lambda: linewidth_var.set(max(0.5, round(linewidth_var.get() - 0.1, 1)))).grid(row=4, column=1); tk.Scale(font_frame, from_=0.5, to=5.0, resolution=0.1, orient=tk.HORIZONTAL, showvalue=0, variable=linewidth_var).grid(row=4, column=2, sticky='ew', padx=5); tk.Button(font_frame, text="+", command=lambda: linewidth_var.set(min(5.0, round(linewidth_var.get() + 0.1, 1)))).grid(row=4, column=3)
    font_frame.columnconfigure(2, weight=1)
    
    tk.Label(extract_frame, text="抽出する電圧 [V]:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
    extract_voltage_var = tk.StringVar(value="1.0")
    tk.Entry(extract_frame, textvariable=extract_voltage_var, width=10).grid(row=0, column=1, sticky='w', padx=5, pady=5)
    tk.Button(extract_frame, text="データを抽出して保存", command=extract_and_save_data).grid(row=0, column=2, sticky='ew', padx=5, pady=5)
    extract_frame.columnconfigure(2, weight=1)

    CATEGORIES = [
        "1回目に破壊", "2回目に破壊", "貫通", "複数回NDR",
        "複数回電流減少", "特性のばらつき"
    ]
    classification_var = tk.StringVar(value=None)
    radio_frame = tk.Frame(classification_frame)
    radio_frame.pack(pady=2, padx=5, fill='x')
    for i, category in enumerate(CATEGORIES):
        row, col = divmod(i, 3) # 2行x3列
        ttk.Radiobutton(
            radio_frame,
            text=category,
            variable=classification_var,
            value=category
        ).grid(row=row, column=col, sticky='w', padx=5)
    summary_table_frame = tk.LabelFrame(classification_frame, text="分類サマリー")
    summary_table_frame.pack(pady=5, padx=5, fill='x')
    tk.Label(summary_table_frame, text="").grid(row=0, column=0, padx=5, sticky='w') 
    tk.Label(summary_table_frame, text="100nm", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=1, padx=5)
    tk.Label(summary_table_frame, text="500nm", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=2, padx=5)
    tk.Label(summary_table_frame, text="2µm", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=3, padx=5)
    count_vars = [[tk.StringVar(value="0") for _ in range(3)] for _ in range(len(CATEGORIES))]
    for r_idx, category in enumerate(CATEGORIES):
        tk.Label(summary_table_frame, text=category, font=("TkDefaultFont", 9)).grid(row=r_idx + 1, column=0, sticky='w', padx=5)
        for c_idx in range(3):
            tk.Label(summary_table_frame, textvariable=count_vars[r_idx][c_idx], font=("TkDefaultFont", 10, "bold")).grid(row=r_idx + 1, column=c_idx + 1)
    summary_table_frame.columnconfigure(1, weight=1)
    summary_table_frame.columnconfigure(2, weight=1)
    summary_table_frame.columnconfigure(3, weight=1)
    
    def update_summary_table():
        try:
            d_100 = float(diameter_mod1_var.get())
            d_500 = float(diameter_mod0_var.get())
            d_2um = float(diameter_mod2_var.get())
        except ValueError:
            return
        diam_map = { d_100: 0, d_500: 1, d_2um: 2 }
        cat_map = { category: i for i, category in enumerate(CATEGORIES) }
        counts = [[0 for _ in range(3)] for _ in range(len(CATEGORIES))]
        for entry in classified_sheets:
            category = entry.get('category')
            diameter = entry.get('diameter')
            row = cat_map.get(category)
            col = diam_map.get(diameter)
            if row is not None and col is not None:
                counts[row][col] += 1
        for r_idx in range(len(CATEGORIES)):
            for c_idx in range(3):
                count_vars[r_idx][c_idx].set(str(counts[r_idx][c_idx]))

    save_summary_button = tk.Button(classification_frame, text="結果をまとめて保存", command=lambda: summarize_and_save_results())
    save_summary_button.pack(pady=(5,0), padx=5, fill='x')
    classified_count_var = tk.StringVar(value="分類済みリスト: 0 件")
    classified_count_label = tk.Label(classification_frame, textvariable=classified_count_var, fg="navy")
    classified_count_label.pack(pady=(5,0))

    def on_classification_change(*args):
        nonlocal _is_programmatic_change
        if _is_programmatic_change:
            return
        selected_category = classification_var.get()
        if not selected_category or selected_category not in CATEGORIES:
            return
        
        current_sheet_name = sheet_var.get()
             
        try:
            match = re.search(r'(\d+)', current_sheet_name)
            diameter_str = ""
            if match:
                num = int(match.group(1))
                mod_result = num % 3
                if mod_result == 1: diameter_str = diameter_mod1_var.get()
                elif mod_result == 2: diameter_str = diameter_mod2_var.get()
                else: diameter_str = diameter_mod0_var.get()
            else:
                diameter_str = diameter_mod1_var.get()
            diameter_nm = float(diameter_str)
        except ValueError:
            messagebox.showerror("エラー", "直径の数値が無効です。")
            _is_programmatic_change = True
            classification_var.set(None)
            _is_programmatic_change = False
            return
        
        found_index = -1
        existing_data = None
        for i, entry in enumerate(classified_sheets):
            if entry['sheet_name'] == current_sheet_name:
                found_index = i
                existing_data = entry['data'] 
                break
        
        if existing_data is None:
            if load_mode == "bulk":
                if current_sheet_name not in all_sheets_data:
                    messagebox.showerror("エラー", f"シート '{current_sheet_name}' がメモリに見つかりません。")
                    _is_programmatic_change = True
                    classification_var.set(None)
                    _is_programmatic_change = False
                    return
                existing_data = all_sheets_data[current_sheet_name]
            else: # load_mode == "on_demand"
                try:
                    existing_data = robust_read_excel(file_path, sheet_name=current_sheet_name, header=None)
                except Exception as e:
                    messagebox.showerror("読み込みエラー", f"シート '{current_sheet_name}' の読み込みに失敗しました:\n{e}")
                    _is_programmatic_change = True
                    classification_var.set(None)
                    _is_programmatic_change = False
                    return

        new_entry = {
            'sheet_name': current_sheet_name,
            'category': selected_category,
            'diameter': diameter_nm,
            'data': existing_data 
        }
        if found_index != -1:
            classified_sheets[found_index] = new_entry
            status_label.config(text=f"'{current_sheet_name}' を '{selected_category}' に更新。")
        else:
            classified_sheets.append(new_entry)
            status_label.config(text=f"'{current_sheet_name}' を '{selected_category}' に分類。")
        
        classified_count_var.set(f"分類済みリスト: {len(classified_sheets)} 件")
        update_summary_table()
        graph_window.after(3000, lambda: status_label.config(text=""))

    def summarize_and_save_results():
        if not classified_sheets:
            messagebox.showinfo("情報", "分類リストにシートがありません。")
            return
        folder_path = filedialog.askdirectory(title="保存先のフォルダを選択してください")
        if not folder_path:
            return
        try:
            diameter_sort_order = {
                float(diameter_mod1_var.get()): 0,
                float(diameter_mod0_var.get()): 1,
                float(diameter_mod2_var.get()): 2
            }
        except ValueError:
            messagebox.showerror("エラー", "直径の数値が無効なため、ソートできません。")
            return
        
        for category in CATEGORIES:
            sheets_in_category = [s for s in classified_sheets if s['category'] == category]
            if not sheets_in_category:
                continue
            sheets_in_category.sort(key=lambda s: diameter_sort_order.get(s['diameter'], 99))
            save_filename = os.path.join(folder_path, f"{category}.xlsx")
            try:
                with pd.ExcelWriter(save_filename, engine='openpyxl') as writer:
                    written_sheet_names = {}
                    for item in sheets_in_category:
                        base_sheet_name = item['sheet_name']
                        df = item['data'] 
                        if base_sheet_name not in written_sheet_names:
                            written_sheet_names[base_sheet_name] = 1
                            final_sheet_name = base_sheet_name
                        else:
                            written_sheet_names[base_sheet_name] += 1
                            count = written_sheet_names[base_sheet_name]
                            suffix = f"({count})"
                            max_len = 31
                            if len(base_sheet_name) + len(suffix) > max_len:
                                truncate_at = max_len - len(suffix)
                                base_name_truncated = base_sheet_name[:truncate_at]
                                final_sheet_name = f"{base_name_truncated}{suffix}"
                            else:
                                final_sheet_name = f"{base_sheet_name}{suffix}"
                        df.to_excel(writer, sheet_name=final_sheet_name, index=False, header=False)
            except Exception as e:
                messagebox.showerror("保存エラー", f"{save_filename} の保存中にエラーが発生しました:\n{e}")
                return

        messagebox.showinfo("完了", f"結果の保存が完了しました。\n保存先: {folder_path}")
        classified_sheets.clear()
        classified_count_var.set("分類済みリスト: 0 件")
        update_summary_table()
        classification_var.set(None)

    update_button = tk.Button(action_frame, text="グラフを更新", command=lambda: on_update_button())
    update_button.pack(side=tk.LEFT, expand=True, fill='x', padx=2)
    save_button = tk.Button(action_frame, text="グラフを保存", command=lambda: save_graph(fig))
    save_button.pack(side=tk.LEFT, expand=True, fill='x', padx=2)
    reselect_button = tk.Button(action_frame, text="別のファイルを選択", command=lambda: reselect_data())
    reselect_button.pack(side=tk.LEFT, expand=True, fill='x', padx=2)
    status_label = tk.Label(action_frame, text="", fg="blue")
    status_label.pack(side=tk.BOTTOM, fill='x', pady=(5,0))

    collected_x_data = []
    def update_display_settings():
        nonlocal collected_x_data
        try:
            plt.rcParams['font.family'] = 'serif'; plt.rcParams['font.serif'] = ['Times New Roman', 'MS Mincho']
            plt.rcParams['axes.unicode_minus'] = False; plt.rcParams['font.weight'] = 'bold'
            plt.rcParams['axes.labelweight'] = 'bold'; plt.rcParams['axes.titleweight'] = 'bold'
        except: pass

        ax.set_title(title_var.get(), fontsize=title_font_var.get())
        ax.set_xlabel('Voltage [V]', fontsize=label_font_var.get())
        ax.set_ylabel(r'Current density [MA/cm$^2$]', fontsize=label_font_var.get())
        ax.tick_params(axis='x', labelsize=tick_font_var.get()); ax.tick_params(axis='y', labelsize=tick_font_var.get())
        if ax.legend_ is not None: ax.legend_.remove()
        if legend_visible_var.get() and ax.lines:
            loc_val = legend_loc_var.get()
            ncol = legend_ncol_var.get()
            if loc_val == 'Free':
                ax.legend(bbox_to_anchor=(legend_x_var.get(), legend_y_var.get()), fontsize=legend_font_var.get(), ncol=ncol, labelspacing=0.1, columnspacing=0.8)
            else:
                ax.legend(loc=loc_val, fontsize=legend_font_var.get(), ncol=ncol, labelspacing=0.1, columnspacing=0.8)
        for line in ax.lines: line.set_linewidth(linewidth_var.get())
        x_formatter = ticker.FuncFormatter(lambda x, pos: '0' if x == 0 else f'{x:.{x_sf_var.get()}f}')
        y_formatter = ticker.FuncFormatter(lambda x, pos: '0' if x == 0 else f'{x:.{y_sf_var.get()}f}')
        ax.xaxis.set_major_formatter(x_formatter); ax.yaxis.set_major_formatter(y_formatter)
        xmin, xmax = ax.get_xlim(); ymin, ymax = ax.get_ylim()
        if collected_x_data:
            full_x_series = pd.concat(collected_x_data)
            if not full_x_series.empty and full_x_series.min() >= 0:
                xmin = 0; ymin = 0
        try:
            if xlim_max_var.get(): xmax = float(xlim_max_var.get())
            if ylim_max_var.get(): ymax = float(ylim_max_var.get())
        except ValueError: pass
        ax.set_xlim(xmin, xmax); ax.set_ylim(ymin, ymax)
        fig.tight_layout()
        canvas.draw_idle()

    def on_update_button(*args):
        nonlocal collected_x_data, annot, current_area_cm2
        nonlocal _is_programmatic_change
        try:
            ax.cla()
            annot = ax.annotate("", xy=(0,0), xytext=(20,20), textcoords="offset points",
                                bbox=dict(boxstyle="round", fc="wheat", alpha=0.8),
                                arrowprops=dict(arrowstyle="->"))
            annot.set_visible(False)
            current_sheet = sheet_var.get()
            
            # --- シート切り替え時に分類状態を読み込む ---
            _is_programmatic_change = True
            found_category = None
            for entry in classified_sheets:
                if entry['sheet_name'] == current_sheet:
                    found_category = entry['category']
                    break
            classification_var.set(found_category)
            _is_programmatic_change = False
            # ------------------------------------------

            title_var.set(current_sheet)
            match = re.search(r'(\d+)', current_sheet)
            diameter_str = ""
            if match:
                num = int(match.group(1))
                mod_result = num % 3
                if mod_result == 1: diameter_str = diameter_mod1_var.get()
                elif mod_result == 2: diameter_str = diameter_mod2_var.get()
                else: diameter_str = diameter_mod0_var.get()
            else:
                diameter_str = diameter_mod1_var.get()
            diameter_nm = float(diameter_str)
            area_cm2 = (math.pi * ((diameter_nm / 2) ** 2)) * 1e-14
            current_area_cm2 = area_cm2
            active_area_var.set(f"{area_cm2:.4g} cm² ({diameter_nm} nm)")

            # --- データ読み込み分岐 ---
            df = None
            if load_mode == "bulk":
                df = all_sheets_data.get(current_sheet)
                if df is None:
                    ax.text(0.5, 0.5, f"シート '{current_sheet}' の\nデータが読み込まれていません。", 
                            ha='center', va='center', transform=ax.transAxes, color='red')
                    canvas.draw_idle()
                    return
            else: # load_mode == "on_demand"
                try:
                    df = robust_read_excel(file_path, sheet_name=current_sheet, header=None)
                except Exception as e:
                    ax.text(0.5, 0.5, f"シート '{current_sheet}' の\n読み込みに失敗しました:\n{e}", 
                            ha='center', va='center', transform=ax.transAxes, color='red')
                    canvas.draw_idle()
                    return
            # ------------------------

            vmin = float(v_str) if (v_str := voltage_min_entry.get()) else None
            vmax = float(v_str) if (v_str := voltage_max_entry.get()) else None
            collected_x_data, _ = plot_graph(df, area_cm2, counts_entry.get(), ax, vmin, vmax)
            update_display_settings()
        except Exception as e:
            messagebox.showerror("エラー", f"グラフの更新中にエラーが発生しました:\n{e}"); return

    def reselect_data():
        graph_window.destroy()
        control_window.destroy()
        root.destroy()
        main()

    for var in [sheet_var, diameter_mod1_var, diameter_mod2_var, diameter_mod0_var,
                legend_visible_var, xlim_max_var, ylim_max_var,
                title_font_var, label_font_var, tick_font_var, legend_font_var,
                legend_loc_var, legend_x_var, legend_y_var, legend_ncol_var,
                linewidth_var, x_sf_var, y_sf_var]:
        var.trace('w', lambda *args: on_update_button())

    legend_loc_var.trace('w', lambda *args: [toggle_free_legend_sliders(), update_display_settings()])
    classification_var.trace('w', on_classification_change)

    toggle_layout()
    on_update_button() 
    graph_window.bind("<Configure>", lambda event: update_display_settings())
    def on_close():
        root.destroy()
    graph_window.protocol("WM_DELETE_WINDOW", on_close)
    control_window.protocol("WM_DELETE_WINDOW", on_close)

    pass


def run_peak_tool(root, load_mode):
    """アプリケーションのメインロジック（Vpeak/Jpeak ばらつき比較ツール）"""
    file_path, sheet_names = select_excel_file(root)
    if not file_path:
        root.destroy()
        return

    dialog = DiameterInputDialog(root)
    initial_diams = dialog.show()

    if initial_diams is None:
        messagebox.showinfo("キャンセル", "直径が入力されなかったため、処理を中断します。")
        root.destroy()
        return

    all_sheets_data = None
    sheet_names_list = []

    if load_mode == "bulk":
        all_sheets_data = load_all_sheets_with_progress(root, file_path, sheet_names)
        if all_sheets_data is None:
            messagebox.showerror("エラー", "データ読み込みに失敗したため、ツールを終了します。")
            root.destroy()
            return
        sheet_names_list = list(all_sheets_data.keys())
    else: # load_mode == "on_demand"
        all_sheets_data = None # メモリには保持しない
        sheet_names_list = sheet_names # ファイルから読み取った生のシート名リスト
    
    if not sheet_names_list:
         messagebox.showerror("エラー", "処理対象のシートが見つかりません。")
         root.destroy()
         return

    current_area_cm2 = 0.0
    peak_data_store = {}

    graph_window = tk.Toplevel(root)
    graph_window.title(f"Vpeak/Jpeak ばらつき比較 - {file_path.split('/')[-1]} (モード: {load_mode})")
    graph_window.geometry("960x540+50+50")

    control_window = tk.Toplevel(root)
    control_window.title("コントロールパネル")

    fig, ax = plt.subplots(figsize=(8, 4.5))
    canvas = FigureCanvasTkAgg(fig, master=graph_window)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    annot = None

    def update_annot(line, ind):
        x_data, y_density_data = line.get_data()
        pos = (x_data[ind["ind"][0]], y_density_data[ind["ind"][0]])
        annot.xy = pos
        current_density_MA_per_cm2 = pos[1]
        current_density_A_per_cm2 = current_density_MA_per_cm2 * 1e6
        raw_current_value = current_density_A_per_cm2 * current_area_cm2
        data_count_label = line.get_label()
        text = (f"Data: {data_count_label}\n"
                f"Vpeak: {pos[0]:.3f} V\n"
                f"Jpeak: {pos[1]:.4g} MA/cm²\n"
                f"(Raw Current: {raw_current_value:.3e} A)")
        annot.set_text(text)
        xlim = ax.get_xlim(); ylim = ax.get_ylim()
        offset = 25
        is_right = pos[0] > (xlim[0] + xlim[1]) / 2
        is_top = pos[1] > (ylim[0] + ylim[1]) / 2
        x_offset = -offset if is_right else offset
        y_offset = -offset if is_top else offset
        annot.xyann = (x_offset, y_offset)
        annot.set_horizontalalignment('right' if is_right else 'left')
        annot.set_verticalalignment('bottom' if is_top else 'top')

    def hover(event):
        if not annot: return
        vis = annot.get_visible()
        if event.inaxes == ax:
            for line in ax.lines:
                contains, ind = line.contains(event)
                if contains:
                    update_annot(line, ind)
                    annot.set_visible(True)
                    fig.canvas.draw_idle()
                    return
        if vis:
            annot.set_visible(False)
            fig.canvas.draw_idle()

    def get_current_diameter():
        current_sheet = sheet_var.get()
        match = re.search(r'(\d+)', current_sheet)
        diameter_str = ""
        if match:
            num = int(match.group(1))
            mod_result = num % 3
            if mod_result == 1: diameter_str = diameter_mod1_var.get()
            elif mod_result == 2: diameter_str = diameter_mod2_var.get()
            else: diameter_str = diameter_mod0_var.get()
        else:
            diameter_str = diameter_mod1_var.get()
        try:
            return float(diameter_str)
        except ValueError:
            raise ValueError("直径の入力値が有効な数値ではありません。")

    def on_click(event):
        if event.inaxes == ax and event.button == 1:
            for line in ax.lines:
                contains, ind = line.contains(event)
                if contains:
                    x_data, y_data = line.get_data()
                    clicked_index = ind["ind"][0]
                    v_peak = x_data[clicked_index]
                    j_peak = y_data[clicked_index]
                    current_sheet_name = sheet_var.get()
                    try:
                        current_diameter_nm = get_current_diameter()
                        peak_data_store[current_sheet_name] = {
                            "vpeak": v_peak, 
                            "jpeak": j_peak, 
                            "diameter": current_diameter_nm
                        }
                        peak_count_var.set(f"記録済み: {len(peak_data_store)} 件")
                        status_label.config(text=f"{current_sheet_name}:\nVp={v_peak:.3f} V, Jp={j_peak:.4g} MA/cm² を記録")
                        graph_window.after(3000, lambda: status_label.config(text=""))
                    except ValueError as e:
                        messagebox.showerror("エラー", str(e))
                    return

    fig.canvas.mpl_connect("motion_notify_event", hover)
    fig.canvas.mpl_connect('button_press_event', on_click)

    control_canvas = tk.Canvas(control_window)
    scrollbar = tk.Scrollbar(control_window, orient="vertical", command=control_canvas.yview)
    control_frame = tk.Frame(control_canvas, padx=10, pady=10)
    control_canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill="y")
    control_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    control_canvas.create_window((0, 0), window=control_frame, anchor="nw")
    def on_frame_configure(event):
        control_canvas.configure(scrollregion=control_canvas.bbox("all"))
    control_frame.bind("<Configure>", on_frame_configure)

    top_bar = tk.Frame(control_frame)
    top_bar.pack(fill='x', pady=(0, 5))
    content_area = tk.Frame(control_frame)
    content_area.pack(fill='both', expand=True)
    left_column = tk.Frame(content_area)
    right_column = tk.Frame(content_area)
    data_source_frame = tk.LabelFrame(content_area, text="データソース")
    graph_control_frame = tk.Frame(content_area)
    axis_frame = tk.LabelFrame(content_area, text="軸の範囲と表示限界")
    sf_frame = tk.LabelFrame(content_area, text="軸の有効数字")
    legend_frame = tk.LabelFrame(content_area, text="凡例の設定")
    font_frame = tk.LabelFrame(content_area, text="フォントと線の設定")
    action_frame = tk.Frame(content_area)
    peak_tool_frame = tk.LabelFrame(content_area, text="Vpeak/Jpeak 取得機能")
    col1_widgets = [data_source_frame, graph_control_frame, axis_frame, sf_frame]
    col2_widgets = [legend_frame, font_frame, peak_tool_frame, action_frame]

    def toggle_layout(*args):
        left_column.pack_forget()
        right_column.pack_forget()
        for widget in col1_widgets + col2_widgets: widget.pack_forget()
        if is_two_column_var.get():
            control_window.geometry("840x680+910+50") 
            left_column.pack(side='left', fill='y', anchor='n', padx=(0, 5))
            right_column.pack(side='right', fill='y', anchor='n', padx=(5, 0))
            for widget in col1_widgets: widget.pack(in_=left_column, fill='x', pady=2)
            for widget in col2_widgets: widget.pack(in_=right_column, fill='x', pady=2)
        else:
            control_window.geometry("450x800+910+50") 
            for widget in col1_widgets + col2_widgets: widget.pack(in_=content_area, fill='x', pady=2)

    def run_peak_analysis():
        if not peak_data_store:
            messagebox.showinfo("情報", "記録されたピークデータがありません。\nグラフ上のピークをクリックしてデータを記録してください。")
            return
        filepath = filedialog.asksaveasfilename(
            title="Vpeak/Jpeak データを保存",
            defaultextension=".xlsx",
            filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")]
        )
        if not filepath:
            return
        data_list = []
        for sheet, values in peak_data_store.items():
            data_list.append({
                "Sheet": sheet,
                "Diameter (nm)": values['diameter'],
                "Vpeak (V)": values['vpeak'],
                "Jpeak (MA/cm²)": values['jpeak']
            })
        if not data_list:
            messagebox.showinfo("情報", "有効なデータがありません。")
            return
        df = pd.DataFrame(data_list)
        try:
            df.to_excel(filepath, index=False, sheet_name="PeakData")
        except Exception as e:
            messagebox.showerror("保存エラー", f"Excelファイルの保存に失敗しました:\n{e}")
            return
        try:
            diameters = sorted(df['Diameter (nm)'].unique())
            num_diams = len(diameters)
            if num_diams == 0:
                messagebox.showinfo("完了", f"データを {filepath} に保存しました。\n(ヒストグラム用のデータはありません)")
                return
            all_vpeak_data = df['Vpeak (V)'].dropna()
            if all_vpeak_data.empty:
                 messagebox.showinfo("完了", f"データを {filepath} に保存しました。\n(Vpeakデータがなくヒストグラムは作成できません)")
                 return
            v_min_global = all_vpeak_data.min()
            v_max_global = all_vpeak_data.max()
            v_margin = (v_max_global - v_min_global) * 0.05
            v_margin = max(v_margin, 0.01) 
            x_lim_min = v_min_global - v_margin
            x_lim_max = v_max_global + v_margin
            bins_global = np.linspace(v_min_global, v_max_global, 11) 
            y_max_global = 0
            hist_data_map = {} 
            for diam in diameters:
                vpeak_data = df[df['Diameter (nm)'] == diam]['Vpeak (V)'].dropna()
                if not vpeak_data.empty:
                    counts, _ = np.histogram(vpeak_data, bins=bins_global)
                    hist_data_map[diam] = vpeak_data
                    y_max_global = max(y_max_global, counts.max())
                else:
                    hist_data_map[diam] = None
            y_lim_max = y_max_global * 1.15 
            y_lim_max = max(y_lim_max, 5) 
            plt.rcParams['font.family'] = 'serif'; plt.rcParams['font.serif'] = ['Times New Roman', 'MS Mincho']
            plt.rcParams['axes.unicode_minus'] = False; plt.rcParams['font.weight'] = 'bold'
            plt.rcParams['axes.labelweight'] = 'bold'; plt.rcParams['axes.titleweight'] = 'bold'
            fig_hist, axes = plt.subplots(1, num_diams, figsize=(6 * num_diams, 5), squeeze=False)
            for i, diam in enumerate(diameters):
                ax_hist = axes[0, i]
                vpeak_data = hist_data_map.get(diam)
                ax_hist.set_title(f"Vpeak Distribution ({diam} nm)", fontsize=16)
                if vpeak_data is not None and not vpeak_data.empty:
                    ax_hist.hist(vpeak_data, bins=bins_global, edgecolor='black', alpha=0.7)
                    std_dev = vpeak_data.std()
                    mean_val = vpeak_data.mean()
                    count_val = len(vpeak_data)
                    text_str = (f"$\mu = {mean_val:.3f}$ V\n"
                                f"$\sigma = {std_dev:.4f}$ V\n"
                                f"n = {count_val}")
                    ax_hist.text(0.95, 0.95, text_str, 
                                 transform=ax_hist.transAxes, 
                                 ha='right', va='top', fontsize=12,
                                 bbox=dict(boxstyle='round', facecolor='white', alpha=0.8))
                else:
                    ax_hist.text(0.5, 0.5, "No Data", ha='center', va='center', transform=ax_hist.transAxes)
                ax_hist.set_xlabel("Vpeak (V)", fontsize=14)
                ax_hist.set_ylabel("Count", fontsize=14)
                ax_hist.tick_params(axis='both', which='major', labelsize=12)
                ax_hist.set_xlim(x_lim_min, x_lim_max)
                ax_hist.set_ylim(0, y_lim_max)
            fig_hist.tight_layout()
            plt.show(block=False) 
            messagebox.showinfo("完了", f"データを {filepath} に保存し、\nVpeak のヒストグラムを表示しました。")
        except Exception as e:
            messagebox.showerror("ヒストグラムエラー", f"ヒストグラムの作成中にエラーが発生しました:\n{e}")

    is_two_column_var = tk.BooleanVar(value=True)
    tk.Checkbutton(top_bar, text="2列レイアウトで表示", variable=is_two_column_var, command=toggle_layout).pack(side='left')
    tk.Label(data_source_frame, text="シート:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
    sheet_var = tk.StringVar()
    
    def next_sheet():
        current_index = sheet_names_list.index(sheet_var.get())
        sheet_var.set(sheet_names_list[(current_index + 1) % len(sheet_names_list)])
    def prev_sheet():
        current_index = sheet_names_list.index(sheet_var.get())
        sheet_var.set(sheet_names_list[(current_index - 1 + len(sheet_names_list)) % len(sheet_names_list)])
        
    prev_button = tk.Button(data_source_frame, text="<", command=prev_sheet); prev_button.grid(row=0, column=1, padx=(5,0), pady=2)
    
    sheet_menu = ttk.Combobox(data_source_frame, textvariable=sheet_var, values=sheet_names_list, state="readonly", width=15); sheet_menu.grid(row=0, column=2, sticky='ew', pady=2)
    sheet_menu.set(sheet_names_list[0])
    
    next_button = tk.Button(data_source_frame, text=">", command=next_sheet); next_button.grid(row=0, column=3, padx=(0,5), pady=2)
    tk.Label(data_source_frame, text="100nm の直径(nm):").grid(row=1, column=0, sticky='w', padx=5, pady=2)
    diameter_mod1_var = tk.StringVar(value=initial_diams[0]); tk.Entry(data_source_frame, textvariable=diameter_mod1_var, width=10).grid(row=1, column=1, sticky='w', padx=5, pady=2, columnspan=3)
    tk.Label(data_source_frame, text="2µm の直径(nm):").grid(row=2, column=0, sticky='w', padx=5, pady=2)
    diameter_mod2_var = tk.StringVar(value=initial_diams[1]); tk.Entry(data_source_frame, textvariable=diameter_mod2_var, width=10).grid(row=2, column=1, sticky='w', padx=5, pady=2, columnspan=3)
    tk.Label(data_source_frame, text="500nm の直径(nm):").grid(row=3, column=0, sticky='w', padx=5, pady=2)
    diameter_mod0_var = tk.StringVar(value=initial_diams[2]); tk.Entry(data_source_frame, textvariable=diameter_mod0_var, width=10).grid(row=3, column=1, sticky='w', padx=5, pady=2, columnspan=3)
    active_area_var = tk.StringVar()
    tk.Label(data_source_frame, text="適用中の面積:").grid(row=4, column=0, sticky='w', padx=5, pady=2)
    tk.Label(data_source_frame, textvariable=active_area_var, font=("TkDefaultFont", 10, "bold")).grid(row=4, column=1, columnspan=3, sticky='w', padx=5, pady=2)
    data_source_frame.columnconfigure(2, weight=1)
    
    title_var = tk.StringVar(value=sheet_names_list[0])
    
    tk.Label(graph_control_frame, text="表示回数 (例: 1-5,7):").grid(row=0, column=0, sticky='w', padx=5, pady=2)
    counts_entry = tk.Entry(graph_control_frame, width=30)
    counts_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=2)
    tk.Label(axis_frame, text="電圧範囲(Xデータ):").grid(row=0, column=0, sticky='w'); voltage_min_entry = tk.Entry(axis_frame, width=10); voltage_min_entry.grid(row=0, column=1); tk.Label(axis_frame, text="～").grid(row=0, column=2); voltage_max_entry = tk.Entry(axis_frame, width=10); voltage_max_entry.grid(row=0, column=3)
    tk.Label(axis_frame, text="X軸 表示最大値:").grid(row=1, column=0, sticky='w'); xlim_max_var = tk.StringVar(); xlim_max_entry = tk.Entry(axis_frame, width=10, textvariable=xlim_max_var); xlim_max_entry.grid(row=1, column=1)
    tk.Label(axis_frame, text="Y軸 表示最大値:").grid(row=2, column=0, sticky='w'); ylim_max_var = tk.StringVar(); ylim_max_entry = tk.Entry(axis_frame, width=10, textvariable=ylim_max_var); ylim_max_entry.grid(row=2, column=1)
    x_sf_var = tk.IntVar(value=2); y_sf_var = tk.IntVar(value=2)
    tk.Label(sf_frame, text="X軸 (小数点以下):").grid(row=0, column=0, sticky='w'); tk.Button(sf_frame, text="-", command=lambda: x_sf_var.set(max(0, x_sf_var.get() - 1))).grid(row=0, column=1); tk.Scale(sf_frame, from_=0, to=5, orient=tk.HORIZONTAL, showvalue=0, variable=x_sf_var).grid(row=0, column=2, sticky='ew', padx=5); tk.Button(sf_frame, text="+", command=lambda: x_sf_var.set(min(5, x_sf_var.get() + 1))).grid(row=0, column=3)
    tk.Label(sf_frame, text="Y軸 (小数点以下):").grid(row=1, column=0, sticky='w'); tk.Button(sf_frame, text="-", command=lambda: y_sf_var.set(max(0, y_sf_var.get() - 1))).grid(row=1, column=1); tk.Scale(sf_frame, from_=0, to=5, orient=tk.HORIZONTAL, showvalue=0, variable=y_sf_var).grid(row=1, column=2, sticky='ew', padx=5); tk.Button(sf_frame, text="+", command=lambda: y_sf_var.set(min(5, y_sf_var.get() + 1))).grid(row=1, column=3)
    sf_frame.columnconfigure(2, weight=1)
    legend_visible_var = tk.BooleanVar(value=True); legend_toggle = tk.Checkbutton(legend_frame, text="凡例を表示", variable=legend_visible_var); legend_toggle.grid(row=0, column=2, columnspan=2, sticky='e', padx=5)
    tk.Label(legend_frame, text="位置:").grid(row=0, column=0, sticky='w'); legend_loc_options = ['best', 'upper right', 'upper left', 'lower left', 'lower right', 'right', 'center', 'Free']
    legend_loc_var = tk.StringVar(control_window); legend_loc_var.set(legend_loc_options[0]); legend_menu = ttk.Combobox(legend_frame, textvariable=legend_loc_var, values=legend_loc_options, width=12, state="readonly"); legend_menu.grid(row=0, column=1, padx=5, sticky='w')
    legend_x_var = tk.DoubleVar(value=0.5); legend_y_var = tk.DoubleVar(value=0.5)
    def toggle_free_legend_sliders(*args):
        state = 'normal' if legend_loc_var.get() == 'Free' else 'disabled'
        for w in [legend_x_scale, legend_y_scale, btn_x_minus, btn_x_plus, btn_y_minus, btn_y_plus]: w.config(state=state)
    tk.Label(legend_frame, text="X:").grid(row=1, column=0, sticky='w'); btn_x_minus = tk.Button(legend_frame, text="-", command=lambda: legend_x_var.set(max(0.0, round(legend_x_var.get() - 0.01, 2)))); btn_x_minus.grid(row=1, column=1); legend_x_scale = tk.Scale(legend_frame, from_=0, to=1, resolution=0.01, orient=tk.HORIZONTAL, showvalue=0, variable=legend_x_var); legend_x_scale.grid(row=1, column=2, sticky='ew', padx=5); btn_x_plus = tk.Button(legend_frame, text="+", command=lambda: legend_x_var.set(min(1.0, round(legend_x_var.get() + 0.01, 2)))); btn_x_plus.grid(row=1, column=3)
    tk.Label(legend_frame, text="Y:").grid(row=2, column=0, sticky='w'); btn_y_minus = tk.Button(legend_frame, text="-", command=lambda: legend_y_var.set(max(0.0, round(legend_y_var.get() - 0.01, 2)))); btn_y_minus.grid(row=2, column=1); legend_y_scale = tk.Scale(legend_frame, from_=0, to=1, resolution=0.01, orient=tk.HORIZONTAL, showvalue=0, variable=legend_y_var); legend_y_scale.grid(row=2, column=2, sticky='ew', padx=5); btn_y_plus = tk.Button(legend_frame, text="+", command=lambda: legend_y_var.set(min(1.0, round(legend_y_var.get() + 0.01, 2)))); btn_y_plus.grid(row=2, column=3)
    tk.Label(legend_frame, text="列数:").grid(row=3, column=0, sticky='w'); legend_ncol_var = tk.IntVar(value=1); tk.Button(legend_frame, text="-", command=lambda: legend_ncol_var.set(max(1, legend_ncol_var.get() - 1))).grid(row=3, column=1); legend_ncol_scale = tk.Scale(legend_frame, from_=1, to=5, orient=tk.HORIZONTAL, showvalue=0, variable=legend_ncol_var); legend_ncol_scale.grid(row=3, column=2, padx=5, sticky='ew'); tk.Button(legend_frame, text="+", command=lambda: legend_ncol_var.set(min(5, legend_ncol_var.get() + 1))).grid(row=3, column=3)
    legend_frame.columnconfigure(2, weight=1)
    title_font_var = tk.IntVar(value=18); label_font_var = tk.IntVar(value=18); tick_font_var = tk.IntVar(value=16); legend_font_var = tk.IntVar(value=11); linewidth_var = tk.DoubleVar(value=1.5)
    tk.Label(font_frame, text="タイトル").grid(row=0, column=0, sticky='w'); tk.Button(font_frame, text="-", command=lambda: title_font_var.set(max(8, title_font_var.get() - 1))).grid(row=0, column=1); tk.Scale(font_frame, from_=8, to=30, orient=tk.HORIZONTAL, showvalue=0, variable=title_font_var).grid(row=0, column=2, sticky='ew', padx=5); tk.Button(font_frame, text="+", command=lambda: title_font_var.set(min(30, title_font_var.get() + 1))).grid(row=0, column=3)
    tk.Label(font_frame, text="軸ラベル").grid(row=1, column=0, sticky='w'); tk.Button(font_frame, text="-", command=lambda: label_font_var.set(max(8, label_font_var.get() - 1))).grid(row=1, column=1); tk.Scale(font_frame, from_=8, to=30, orient=tk.HORIZONTAL, showvalue=0, variable=label_font_var).grid(row=1, column=2, sticky='ew', padx=5); tk.Button(font_frame, text="+", command=lambda: label_font_var.set(min(30, label_font_var.get() + 1))).grid(row=1, column=3)
    tk.Label(font_frame, text="目盛り").grid(row=2, column=0, sticky='w'); tk.Button(font_frame, text="-", command=lambda: tick_font_var.set(max(8, tick_font_var.get() - 1))).grid(row=2, column=1); tk.Scale(font_frame, from_=8, to=30, orient=tk.HORIZONTAL, showvalue=0, variable=tick_font_var).grid(row=2, column=2, sticky='ew', padx=5); tk.Button(font_frame, text="+", command=lambda: tick_font_var.set(min(30, tick_font_var.get() + 1))).grid(row=2, column=3)
    tk.Label(font_frame, text="凡例").grid(row=3, column=0, sticky='w'); tk.Button(font_frame, text="-", command=lambda: legend_font_var.set(max(8, legend_font_var.get() - 1))).grid(row=3, column=1); tk.Scale(font_frame, from_=8, to=30, orient=tk.HORIZONTAL, showvalue=0, variable=legend_font_var).grid(row=3, column=2, sticky='ew', padx=5); tk.Button(font_frame, text="+", command=lambda: legend_font_var.set(min(30, legend_font_var.get() + 1))).grid(row=3, column=3)
    tk.Label(font_frame, text="線の太さ:").grid(row=4, column=0, sticky='w'); tk.Button(font_frame, text="-", command=lambda: linewidth_var.set(max(0.5, round(linewidth_var.get() - 0.1, 1)))).grid(row=4, column=1); tk.Scale(font_frame, from_=0.5, to=5.0, resolution=0.1, orient=tk.HORIZONTAL, showvalue=0, variable=linewidth_var).grid(row=4, column=2, sticky='ew', padx=5); tk.Button(font_frame, text="+", command=lambda: linewidth_var.set(min(5.0, round(linewidth_var.get() + 0.1, 1)))).grid(row=4, column=3)
    font_frame.columnconfigure(2, weight=1)
    
    peak_count_var = tk.StringVar(value=f"記録済み: {len(peak_data_store)} 件")
    tk.Label(peak_tool_frame, textvariable=peak_count_var, fg="navy").pack(pady=5)
    tk.Label(peak_tool_frame, text="グラフ上のピーク点を左クリックして記録").pack(pady=(0,5))
    tk.Button(peak_tool_frame, text="データを取得し保存 (Excel + ﾋｽﾄｸﾞﾗﾑ)", command=run_peak_analysis).pack(fill='x', padx=5, pady=5)

    update_button = tk.Button(action_frame, text="グラフを更新", command=lambda: on_update_button())
    update_button.pack(side=tk.LEFT, expand=True, fill='x', padx=2)
    save_button = tk.Button(action_frame, text="グラフを保存", command=lambda: save_graph(fig))
    save_button.pack(side=tk.LEFT, expand=True, fill='x', padx=2)
    reselect_button = tk.Button(action_frame, text="別のファイルを選択", command=lambda: reselect_data())
    reselect_button.pack(side=tk.LEFT, expand=True, fill='x', padx=2)
    status_label = tk.Label(action_frame, text="", fg="blue", wraplength=380, justify="left")
    status_label.pack(side=tk.BOTTOM, fill='x', pady=(5,0))

    collected_x_data = []
    def update_display_settings():
        nonlocal collected_x_data
        try:
            plt.rcParams['font.family'] = 'serif'; plt.rcParams['font.serif'] = ['Times New Roman', 'MS Mincho']
            plt.rcParams['axes.unicode_minus'] = False; plt.rcParams['font.weight'] = 'bold'
            plt.rcParams['axes.labelweight'] = 'bold'; plt.rcParams['axes.titleweight'] = 'bold'
        except: pass

        ax.set_title(title_var.get(), fontsize=title_font_var.get())
        ax.set_xlabel('Voltage [V]', fontsize=label_font_var.get())
        ax.set_ylabel(r'Current density [MA/cm$^2$]', fontsize=label_font_var.get())
        ax.tick_params(axis='x', labelsize=tick_font_var.get()); ax.tick_params(axis='y', labelsize=tick_font_var.get())
        if ax.legend_ is not None: ax.legend_.remove()
        if legend_visible_var.get() and ax.lines:
            loc_val = legend_loc_var.get()
            ncol = legend_ncol_var.get()
            if loc_val == 'Free':
                ax.legend(bbox_to_anchor=(legend_x_var.get(), legend_y_var.get()), fontsize=legend_font_var.get(), ncol=ncol, labelspacing=0.1, columnspacing=0.8)
            else:
                ax.legend(loc=loc_val, fontsize=legend_font_var.get(), ncol=ncol, labelspacing=0.1, columnspacing=0.8)
        for line in ax.lines: line.set_linewidth(linewidth_var.get())
        x_formatter = ticker.FuncFormatter(lambda x, pos: '0' if x == 0 else f'{x:.{x_sf_var.get()}f}')
        y_formatter = ticker.FuncFormatter(lambda x, pos: '0' if x == 0 else f'{x:.{y_sf_var.get()}f}')
        ax.xaxis.set_major_formatter(x_formatter); ax.yaxis.set_major_formatter(y_formatter)
        xmin, xmax = ax.get_xlim(); ymin, ymax = ax.get_ylim()
        if collected_x_data:
            full_x_series = pd.concat(collected_x_data)
            if not full_x_series.empty and full_x_series.min() >= 0:
                xmin = 0; ymin = 0
        try:
            if xlim_max_var.get(): xmax = float(xlim_max_var.get())
            if ylim_max_var.get(): ymax = float(ylim_max_var.get())
        except ValueError: pass
        ax.set_xlim(xmin, xmax); ax.set_ylim(ymin, ymax)
        fig.tight_layout()
        canvas.draw_idle()

    def on_update_button(*args):
        nonlocal collected_x_data, annot, current_area_cm2
        try:
            ax.cla()
            annot = ax.annotate("", xy=(0,0), xytext=(20,20), textcoords="offset points",
                                bbox=dict(boxstyle="round", fc="wheat", alpha=0.8),
                                arrowprops=dict(arrowstyle="->"))
            annot.set_visible(False)
            current_sheet = sheet_var.get()

            if current_sheet in peak_data_store:
                stored = peak_data_store[current_sheet]
                status_label.config(text=f"記録済み: Vp={stored['vpeak']:.3f} V, Jp={stored['jpeak']:.4g} MA/cm²")
            else:
                status_label.config(text="") 

            title_var.set(current_sheet)
            
            diameter_nm = get_current_diameter() 
            area_cm2 = (math.pi * ((diameter_nm / 2) ** 2)) * 1e-14
            current_area_cm2 = area_cm2
            active_area_var.set(f"{area_cm2:.4g} cm² ({diameter_nm} nm)")

            # --- データ読み込み分岐 ---
            df = None
            if load_mode == "bulk":
                df = all_sheets_data.get(current_sheet)
                if df is None:
                    ax.text(0.5, 0.5, f"シート '{current_sheet}' の\nデータが読み込まれていません。", 
                            ha='center', va='center', transform=ax.transAxes, color='red')
                    canvas.draw_idle()
                    return
            else: # load_mode == "on_demand"
                try:
                    df = robust_read_excel(file_path, sheet_name=current_sheet, header=None)
                except Exception as e:
                    ax.text(0.5, 0.5, f"シート '{current_sheet}' の\n読み込みに失敗しました:\n{e}", 
                            ha='center', va='center', transform=ax.transAxes, color='red')
                    canvas.draw_idle()
                    return
            # ------------------------

            vmin = float(v_str) if (v_str := voltage_min_entry.get()) else None
            vmax = float(v_str) if (v_str := voltage_max_entry.get()) else None
            collected_x_data, _ = plot_graph(df, area_cm2, counts_entry.get(), ax, vmin, vmax)
            update_display_settings()
        except Exception as e:
            messagebox.showerror("エラー", f"グラフの更新中にエラーが発生しました:\n{e}"); return

    def reselect_data():
        graph_window.destroy()
        control_window.destroy()
        root.destroy()
        main() 

    for var in [sheet_var, diameter_mod1_var, diameter_mod2_var, diameter_mod0_var,
                legend_visible_var, xlim_max_var, ylim_max_var,
                title_font_var, label_font_var, tick_font_var, legend_font_var,
                legend_loc_var, legend_x_var, legend_y_var, legend_ncol_var,
                linewidth_var, x_sf_var, y_sf_var]:
        var.trace('w', lambda *args: on_update_button())

    legend_loc_var.trace('w', lambda *args: [toggle_free_legend_sliders(), update_display_settings()])

    toggle_layout()
    on_update_button() 
    graph_window.bind("<Configure>", lambda event: update_display_settings())
    def on_close():
        root.destroy()
    graph_window.protocol("WM_DELETE_WINDOW", on_close)
    control_window.protocol("WM_DELETE_WINDOW", on_close)

    pass
# ---- ▲▲▲ Vpeak/Jpeak ばらつき比較ツール 終了 ▲▲▲ ----


# ---- 統合ツール ----
def extract_xy_from_sheetname(sheet_name):
    match = re.match(r'x(\d+)y(\d+)', sheet_name, re.IGNORECASE)
    if match:
        try:
            x_val = int(match.group(1))
            y_val = int(match.group(2))
            return (x_val, y_val)
        except ValueError:
            return None
    return None

def run_merge_tool(root):
    try:
        root.attributes('-topmost', True)
        file_paths = filedialog.askopenfilenames(
            title="統合するExcelファイルを選択してください（複数選択可）",
            filetypes=[("Excelファイル", "*.xlsx *.xls *.xlsm *.xlsb"), ("すべてのファイル", "*.*")],
            parent=root
        )
        root.attributes('-topmost', False)
        if not file_paths:
            messagebox.showinfo("キャンセル", "ファイルが選択されませんでした。", parent=root)
            return

        all_sheets_data = []

        progress_read_window = tk.Toplevel(root)
        progress_read_window.title("読み込み中...")
        progress_read_window.geometry("350x100+400+400")
        progress_read_window.resizable(False, False)
        progress_read_label = tk.Label(progress_read_window, text="ファイルからシートを読み込んでいます...", padx=10, pady=10)
        progress_read_label.pack()
        progress_read_bar = ttk.Progressbar(progress_read_window, orient="horizontal", length=300, mode="indeterminate")
        progress_read_bar.pack(pady=10)
        progress_read_bar.start(10)
        root.update()

        try:
            for file_path in file_paths:
                try:
                    progress_read_label.config(text=f"処理中: {file_path.split('/')[-1]}")
                    root.update_idletasks()

                    xls = robust_excel_file(file_path)

                    for sheet_name in xls.sheet_names:
                        xy_vals = extract_xy_from_sheetname(sheet_name)
                        if xy_vals:
                            x, y = xy_vals
                            df = robust_read_excel(file_path, sheet_name=sheet_name, header=None)
                            all_sheets_data.append({
                                'x': x,
                                'y': y,
                                'sheet_name': sheet_name,
                                'data': df
                            })
                        else:
                            print(f"スキップ: シート名 '{sheet_name}' (ファイル: {file_path}) は 'x〇y〇' の形式ではありません。")
                except Exception as e:
                    messagebox.showwarning("読み取りエラー", f"ファイル {file_path} の読み取り中にエラーが発生しました: {e}", parent=root)
        finally:
            progress_read_bar.stop()
            progress_read_window.destroy()

        if not all_sheets_data:
            messagebox.showerror("エラー", "処理対象となる 'x〇y〇' 形式のシートが見つかりませんでした。", parent=root)
            return

        all_sheets_data.sort(key=lambda item: (item['x'], item['y']))

        root.attributes('-topmost', True)
        save_path = filedialog.asksaveasfilename(
            title="統合後のExcelファイルの保存先を選択",
            defaultextension=".xlsx",
            filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")],
            parent=root
        )
        root.attributes('-topmost', False)
        if not save_path:
            messagebox.showinfo("キャンセル", "保存がキャンセルされました。", parent=root)
            return

        progress_window = tk.Toplevel(root)
        progress_window.title("書き込み中...")
        progress_window.geometry("350x100+400+400")
        progress_window.resizable(False, False)
        progress_label = tk.Label(progress_window, text="統合ファイルに書き込んでいます...", padx=10, pady=10)
        progress_label.pack()
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate", maximum=len(all_sheets_data))
        progress_bar.pack(pady=10)
        root.update()

        try:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                processed_count = 0
                written_sheet_names = {}

                for idx, item in enumerate(all_sheets_data):
                    base_sheet_name = item['sheet_name']
                    df = item['data']

                    try:
                        if base_sheet_name not in written_sheet_names:
                            written_sheet_names[base_sheet_name] = 1
                            final_sheet_name = base_sheet_name
                        else:
                            written_sheet_names[base_sheet_name] += 1
                            count = written_sheet_names[base_sheet_name]
                            suffix = f"({count})"

                            max_len = 31
                            if len(base_sheet_name) + len(suffix) > max_len:
                                truncate_at = max_len - len(suffix)
                                base_name_truncated = base_sheet_name[:truncate_at]
                                final_sheet_name = f"{base_name_truncated}{suffix}"
                                print(f"警告: シート名 '{base_sheet_name}' が長すぎるため '{final_sheet_name}' に短縮しました。")
                            else:
                                final_sheet_name = f"{base_sheet_name}{suffix}"

                        df.to_excel(writer, sheet_name=final_sheet_name, index=False, header=False)
                        processed_count += 1

                        progress_bar['value'] = idx + 1
                        progress_label.config(text=f"書き込み中: {final_sheet_name} ({idx+1}/{len(all_sheets_data)})")
                        root.update_idletasks()

                    except Exception as e:
                        print(f"警告: シート '{base_sheet_name}' の処理中にエラー: {e}")

            if processed_count == 0:
                messagebox.showerror("エラー", "すべてのシートの処理に失敗しました。", parent=root)
            else:
                messagebox.showinfo("完了", f"{processed_count} 枚のシートの統合が完了しました。\n保存先: {save_path}", parent=root)

        except Exception as e:
            messagebox.showerror("保存エラー", f"ファイルの保存中にエラーが発生しました: {e}", parent=root)
        finally:
            progress_window.destroy()

    except Exception as e:
        messagebox.showerror("重大なエラー", f"統合ツール実行中に予期せぬエラーが発生しました: {e}", parent=root)
    finally:
        root.destroy()

class ToolSelectionWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ツール選択")
        self.geometry("400x260+300+300") # 高さを広げる
        self.resizable(False, False)

        tk.Label(self, text="起動するツールを選択してください。", font=("TkDefaultFont", 12)).pack(pady=(20, 10))

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=5, fill='x', expand=True, anchor='center')

        graph_btn = tk.Button(btn_frame, text="グラフ化ツール\n(分類機能あり)", command=self.start_graph_tool, width=18, height=3)
        graph_btn.pack(side=tk.LEFT, padx=(15, 5), fill='x', expand=True)

        merge_btn = tk.Button(btn_frame, text="Excel統合ツール", command=self.start_merge_tool, width=18, height=3)
        merge_btn.pack(side=tk.LEFT, padx=5, fill='x', expand=True)

        peak_tool_btn = tk.Button(btn_frame, text="孔径別Vpeak, Jpeak\nばらつき比較ツール", command=self.start_peak_tool, width=18, height=3)
        peak_tool_btn.pack(side=tk.LEFT, padx=(5, 15), fill='x', expand=True)
        
        btn_frame.pack_configure(padx=10)

        # --- 読み込みモード選択 ---
        mode_frame = tk.LabelFrame(self, text="読み込みモード (グラフ/Vpeakツール用)", padx=10, pady=5)
        mode_frame.pack(pady=(10, 15), padx=15, fill='x')
        
        self.load_mode_var = tk.StringVar(value="bulk") # デフォルトは「一括」
        
        ttk.Radiobutton(
            mode_frame, 
            text="一括読み込み (高速動作 / 起動時ロード)", 
            variable=self.load_mode_var, 
            value="bulk"
        ).pack(anchor='w', padx=10)
        
        ttk.Radiobutton(
            mode_frame, 
            text="都度読み込み (低速動作 / 高速起動)", 
            variable=self.load_mode_var, 
            value="on_demand"
        ).pack(anchor='w', padx=10)
        # ------------------------

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def start_graph_tool(self):
        load_mode = self.load_mode_var.get()
        self.destroy()
        root = tk.Tk()
        root.withdraw()
        run_graphing_tool(root, load_mode) # load_mode を渡す

    def start_merge_tool(self):
        # 統合ツールはモード選択に関係ない
        self.destroy()
        root = tk.Tk()
        root.withdraw()
        run_merge_tool(root)

    def start_peak_tool(self):
        load_mode = self.load_mode_var.get()
        self.destroy()
        root = tk.Tk()
        root.withdraw()
        run_peak_tool(root, load_mode) # load_mode を渡す

    def on_close(self):
        self.destroy()

# ---- main 関数 ----
def main():
    app = ToolSelectionWindow()
    app.mainloop()

if __name__ == '__main__':
    main()