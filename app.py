import pandas as pd
from tkinter import Tk, Label, Button, filedialog, StringVar, ttk
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import os

class ExcelCsvApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel/CSV Viewer & Editor")

        # メインフレーム
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill="both", expand=True)

        # 左側: データ追加とフィルタリング
        self.left_frame = ttk.Frame(self.main_frame)
        self.left_frame.pack(side="left", fill="y", padx=10, pady=10)

        # 右側: データ表示
        self.right_frame = ttk.Frame(self.main_frame)
        self.right_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        # ファイルロードボタン
        self.load_button = Button(self.left_frame, text="ファイルを開く（.xlsx, .csv）", command=self.load_file)
        self.load_button.pack(pady=5)

        # フィルタリング用
        self.filter_frame = ttk.LabelFrame(self.left_frame, text="フィルタリング")
        self.filter_frame.pack(fill="x", pady=10)
        Label(self.filter_frame, text="カラム:").pack()
        self.filter_col = StringVar()
        self.filter_dropdown = ttk.Combobox(self.filter_frame, textvariable=self.filter_col)
        self.filter_dropdown.pack()
        self.filter_dropdown.bind("<<ComboboxSelected>>", self.update_filter_values)  # イベントバインド
        Label(self.filter_frame, text="値:").pack()
        self.filter_value = StringVar()
        self.filter_value_dropdown = ttk.Combobox(self.filter_frame, textvariable=self.filter_value)
        self.filter_value_dropdown.pack()
        self.apply_filter_button = Button(self.filter_frame, text="フィルタ適用", command=self.apply_filter)
        self.apply_filter_button.pack(pady=5)
        self.clear_filter_button = Button(self.filter_frame, text="フィルタ解除", command=self.clear_filter)
        self.clear_filter_button.pack()

        # 散布図プロット用
        self.plot_frame = ttk.LabelFrame(self.left_frame, text="散布図プロット")
        self.plot_frame.pack(fill="x", pady=10)
        Label(self.plot_frame, text="X軸:").pack()
        self.x_col = StringVar()
        self.x_dropdown = ttk.Combobox(self.plot_frame, textvariable=self.x_col)
        self.x_dropdown.pack()
        Label(self.plot_frame, text="Y軸:").pack()
        self.y_col = StringVar()
        self.y_dropdown = ttk.Combobox(self.plot_frame, textvariable=self.y_col)
        self.y_dropdown.pack()
        self.plot_button = Button(self.plot_frame, text="プロット", command=self.plot_scatter)
        self.plot_button.pack(pady=5)

        # データ表示エリア
        self.tree = ttk.Treeview(self.right_frame)
        self.tree.pack(expand=True, fill="both")

        # データ管理
        self.file_path = ""
        self.data = None
        self.filtered_data = None
        self.file_type = ""

    def load_file(self):
        """ExcelまたはCSVファイルを読み込み、データを表示"""
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel or CSV files", "*.xlsx *.csv")])
        if self.file_path:
            ext = os.path.splitext(self.file_path)[-1].lower()
            if ext == ".xlsx":
                self.data = pd.read_excel(self.file_path)
                self.file_type = "xlsx"
            elif ext == ".csv":
                self.data = pd.read_csv(self.file_path)
                self.file_type = "csv"
            else:
                print("対応していないファイル形式です。")
                return
            self.filtered_data = self.data.copy()
            self.update_treeview()
            self.populate_columns()

    def update_treeview(self):
        """Treeviewを更新"""
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(self.filtered_data.columns)
        self.tree["show"] = "headings"
        for col in self.filtered_data.columns:
            self.tree.heading(col, text=col)
        for _, row in self.filtered_data.iterrows():
            self.tree.insert("", "end", values=list(row))

    def populate_columns(self):
        """カラム名をドロップダウンに設定"""
        columns = list(self.data.columns)
        self.filter_dropdown["values"] = columns
        self.x_dropdown["values"] = columns
        self.y_dropdown["values"] = columns

    def update_filter_values(self, event):
        """選択したカラムの値を取得してドロップダウンを更新"""
        selected_col = self.filter_col.get()
        if selected_col and selected_col in self.data.columns:
            unique_values = self.data[selected_col].dropna().unique()
            self.filter_value_dropdown["values"] = sorted(unique_values)  # 値をソートして設定
        else:
            self.filter_value_dropdown["values"] = []  # 値が無効な場合は空にする

    def apply_filter(self):
        """フィルタを適用"""
        col = self.filter_col.get()
        value = self.filter_value.get()
        if col and value:
            try:
                if self.data[col].dtype != object:
                    value = self._convert_value(value, self.data[col].dtype)
                self.filtered_data = self.data[self.data[col] == value]
                self.update_treeview()
            except Exception as e:
                print(f"フィルタ適用エラー: {e}")

    def clear_filter(self):
        """フィルタを解除"""
        self.filtered_data = self.data.copy()
        self.update_treeview()

    def _convert_value(self, value, dtype):
        """値を適切なデータ型に変換"""
        if dtype == int:
            return int(value)
        elif dtype == float:
            return float(value)
        else:
            return value

    def plot_scatter(self):
        """散布図をプロット"""
        x_col = self.x_col.get()
        y_col = self.y_col.get()
        if x_col and y_col and x_col in self.filtered_data.columns and y_col in self.filtered_data.columns:
            plt.figure(figsize=(8, 6))
            plt.scatter(self.filtered_data[x_col], self.filtered_data[y_col], alpha=0.7)
            plt.xlabel(x_col)
            plt.ylabel(y_col)
            plt.title(f"Scatter Plot: {x_col} vs {y_col}")
            plt.grid(True)
            plt.show()

if __name__ == "__main__":
    root = Tk()
    app = ExcelCsvApp(root)
    root.mainloop()
