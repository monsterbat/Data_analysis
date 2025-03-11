import os
import chardet
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        raw_data = f.read(10000)  # 只讀前一部分來偵測就好
    result = chardet.detect(raw_data)
    encoding = result['encoding']
    return encoding

def read_line_from_txt(file_path, line_number):
    try:
        encoding = detect_encoding(file_path)
        print(f"檔案 {file_path} 檢測到編碼: {encoding}")

        with open(file_path, 'r', encoding=encoding) as file:
            lines = file.readlines()
            if line_number <= len(lines):
                return lines[line_number - 1].strip()
            else:
                return f"檔案只有 {len(lines)} 行"
    except Exception as e:
        return f"讀取錯誤: {str(e)}"

def select_files_and_generate_excel():
    root = tk.Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(
        title="請選取多個TXT檔案",
        filetypes=[("Text Files", "*.txt")]
    )

    if not file_paths:
        print("未選取任何檔案！")
        return

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "TXT檔案內容"

    sheet['A1'] = "檔名"
    sheet['B1'] = "第372行內容"

    for idx, file_path in enumerate(file_paths, start=2):
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        line_content = read_line_from_txt(file_path, 372)

        sheet[f'A{idx}'] = file_name
        sheet[f'B{idx}'] = line_content

        print(f"✅ 已處理: {file_name}")

    output_filename = 'TXT結果彙總.xlsx'
    workbook.save(output_filename)
    print(f"\n✅ 完成！結果已儲存為 {output_filename}")

if __name__ == "__main__":
    select_files_and_generate_excel()