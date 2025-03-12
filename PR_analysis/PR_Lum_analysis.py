import os
import chardet
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
from datetime import datetime

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        raw_data = f.read(10000)  # 只讀前一部分來偵測就好
    result = chardet.detect(raw_data)
    encoding = result['encoding']
    return encoding

def read_line_substr_scale(file_path, line_number, start_idx, end_idx, factor=1):
    """
    讀取第 line_number 行，取第 start_idx~end_idx 字元，
    轉成 float 後再乘以 factor 預設是1
    """
    try:
        encoding = detect_encoding(file_path)
        with open(file_path, 'r', encoding=encoding) as file:
            lines = file.readlines()
            if line_number <= len(lines):
                line_str = lines[line_number - 1].strip()
                # 取指定範圍
                line_str = line_str[start_idx:end_idx]
                if factor == "words":
                    return line_str
                else:
                    # 轉成 float
                    num_value = float(line_str)
                    # 乘以 factor
                    num_value *= factor
                    return num_value
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
    sheet.title = "PR data"
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")

    sheet['A1'] = "File name"           # 保留
    sheet['B1'] = "Expl. Time"  # 第6行
    sheet['C1'] = "Data/Time"     # 第7行
    sheet['D1'] = "Aperture"      # 第8行
    sheet['E1'] = "X"             # 第371行
    sheet['F1'] = "Y"             # 第372行
    sheet['G1'] = "Z"             # 第373行
    sheet['H1'] = "x"             # 第377行
    sheet['I1'] = "y"             # 第378行
    sheet['J1'] = "u"             # 第379行
    sheet['K1'] = "v"             # 第380行
    sheet['L1'] = "u'"            # 第381行
    sheet['M1'] = "v'"            # 第382行

    for idx, file_path in enumerate(file_paths, start=2):
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        sheet[f'A{idx}'] = file_name  # 檔名

        # 讀取不同行號並寫入相對應欄位
        # 第 6 行 → B 欄
        sheet[f'B{idx}'] = read_line_substr_scale(file_path, 6, 11, 19, "words")
        # 第 7 行 → C 欄
        sheet[f'C{idx}'] = read_line_substr_scale(file_path, 7, 15, 42, "words")
        # 第 8 行 → D 欄
        sheet[f'D{idx}'] = read_line_substr_scale(file_path, 8, 10, 18, "words")
        # 第 371 行 → E 欄
        sheet[f'E{idx}'] = read_line_substr_scale(file_path, 371, 3, 9, 10000)
        # 第 372 行 → F 欄
        sheet[f'F{idx}'] = read_line_substr_scale(file_path, 372, 3, 9, 10000)
        # 第 373 行 → G 欄
        sheet[f'G{idx}'] = read_line_substr_scale(file_path, 373, 3, 9, 10000)
        # 第 377 行 → H 欄
        sheet[f'H{idx}'] = read_line_substr_scale(file_path, 377, 4, 10, 1)
        # 第 378 行 → I 欄
        sheet[f'I{idx}'] = read_line_substr_scale(file_path, 378, 4, 10, 1)
        # 第 379 行 → J 欄
        sheet[f'J{idx}'] = read_line_substr_scale(file_path, 379, 4, 10, 1)
        # 第 380 行 → K 欄
        sheet[f'K{idx}'] = read_line_substr_scale(file_path, 380, 4, 10, 1)
        # 第 381 行 → L 欄
        sheet[f'L{idx}'] = read_line_substr_scale(file_path, 381, 6, 11, 1)
        # 第 382 行 → M 欄
        sheet[f'M{idx}'] = read_line_substr_scale(file_path, 382, 6, 11, 1)

        print(f"✅ 已處理: {file_name}")

        
    output_filename = f"PR_result_{timestamp}.xlsx"
    workbook.save(output_filename)
    print(f"\n✅ 完成！結果已儲存為 {output_filename}")

if __name__ == "__main__":
    select_files_and_generate_excel()