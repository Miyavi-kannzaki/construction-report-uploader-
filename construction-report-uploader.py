import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.simpledialog as sd
import os
import win32com.client
import pythoncom
import shutil
import glob
import datetime
from datetime import datetime as dt

def get_named_value(wb, name):
    try:
        return wb.Names(name).RefersToRange.Value
    except:
        raise ValueError(f"名前定義 '{name}' が見つかりません。Excelの名前定義を確認してください。")


def execute_excel_macro(excel_path):
    """Excelのマクロを実行してファイル名のベースを取得"""
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(excel_path)

    try:
        sheet = wb.Sheets(1)
        building_no = get_named_value(wb, "HOUSES.BUILDING_NO")
        tentative_name = get_named_value(wb, "HOUSES.TENTATIVE_NAME")

        # SK除去後のフォルダ名生成用 building_name を作成
        building_name = tentative_name.replace("【SK】", "").strip()

        filename_base = f"{int(building_no)}_{tentative_name}_工事完了報告書"
        return filename_base, wb, excel, building_name
    except Exception as e:
        wb.Close(SaveChanges=False)
        excel.Quit()
        raise e


def save_as_xlsx(wb, excel_path, filename_base):
    """Excelを保存し、WB と Excel アプリを解放"""
    folder = os.path.join(os.path.expanduser("~"), "Downloads", "CLIENT処理ファイル")
    os.makedirs(folder, exist_ok=True)
    xlsx_path = os.path.join(folder, f"{filename_base}.xlsx")

    wb.SaveAs(xlsx_path, FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)
    wb.Close(SaveChanges=False)
    return xlsx_path, folder


def find_target_folder(year, month, building_name):
    """Box 上の対象フォルダを探す"""
    folder_name = f"{year}年{month}月切替分_{building_name}"
    base_path = os.path.join(
        os.path.expanduser("~"),
        "Box", "SHARED_WORK", "DEPT1", "GROUP",
        "common", "existing", "【CLIENT案件】", "工事完了報告書"
    )
    target_folder = os.path.join(base_path, folder_name)
    os.makedirs(target_folder, exist_ok=True)
    return target_folder


def move_and_copy_files(xlsx_path, folder, building_no, building_name, year, month):
    """Box へファイルコピーし、最終格納フォルダにバックアップ"""
    target_folder = find_target_folder(year, month, building_name)
    shutil.copy2(xlsx_path, target_folder)

    # 最終格納先フォルダ
    final_folder_name = f"{building_no}_{building_name}"
    final_folder = os.path.join(
        os.path.expanduser("~"),
        "Box", "SHARED_WORK", "PROJECT_MGMT", "工事管理", "LEGACY",
        final_folder_name
    )
    os.makedirs(final_folder, exist_ok=True)
    if os.path.exists(xlsx_path):
        shutil.copy2(xlsx_path, final_folder)
    else:
        raise FileNotFoundError(f"{xlsx_path} が見つかりません。")


def cleanup_temp(folder):
    """一時フォルダを削除"""
    if os.path.exists(folder):
        shutil.rmtree(folder)


def run_main(excel_path, year, month):
    try:
        filename_base, wb, excel, building_name = execute_excel_macro(excel_path)
        xlsx_path, folder = save_as_xlsx(wb, excel_path, filename_base)
        building_no = filename_base.split("_")[0]
        move_and_copy_files(xlsx_path, folder, building_no, building_name, year, month)
    finally:
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
        try:
            excel.Quit()
        except:
            pass
        cleanup_temp(folder)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("工事完了報告書 自動処理ツール (匿名化版)")
        self.geometry("450x220")
        self.resizable(False, False)

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text="Excelファイル:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.file_entry = tk.Entry(self, width=40)
        self.file_entry.grid(row=0, column=1, padx=5, pady=10)
        tk.Button(self, text="参照", command=self.select_file).grid(row=0, column=2, padx=5, pady=10)

        tk.Label(self, text="切替年月 (YYYY/MM):").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.date_entry = tk.Entry(self, width=20)
        self.date_entry.grid(row=1, column=1, padx=5, pady=10, sticky="w")

        tk.Button(self, text="実行", command=self.execute).grid(row=2, column=1, pady=20)

    def select_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx;*.xlsm")])
        if filename:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filename)

    def execute(self):
        excel_path = self.file_entry.get()
        date_str = self.date_entry.get()

        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror("エラー", "Excelファイルを選択してください。")
            return

        try:
            year, month = map(int, date_str.split("/"))
        except ValueError:
            messagebox.showerror("エラー", "切替年月を YYYY/MM 形式で入力してください。")
            return

        try:
            run_main(excel_path, year, month)
            messagebox.showinfo("完了", "処理が完了しました。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))


if __name__ == "__main__":
    App().mainloop()
