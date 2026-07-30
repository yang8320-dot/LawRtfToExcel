import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from striprtf.striprtf import rtf_to_text
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

def chinese_to_arabic(chn_str: str) -> int:
    """中文數字轉阿拉伯數字"""
    s = chn_str.replace("、", "").replace("(", "").replace(")", "")
    s = s.replace("（", "").replace("）", "").strip()
    
    if s.isdigit():
        return int(s)
        
    num_dict = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, 
                "六": 6, "七": 7, "八": 8, "九": 9, "十": 10,
                "百": 100, "千": 1000, "零": 0}
    
    total = 0
    current = 0
    
    for char in s:
        if char in num_dict:
            v = num_dict[char]
            if v in [10, 100, 1000]:
                if current == 0:
                    current = 1
                total += current * v
                current = 0
            else:
                current = v
    total += current
    return total

def format_tiao(article_str: str) -> str:
    """格式化法條字串 (例如將 '第 39-1 條' 轉為 '039-1')"""
    pure = article_str.replace("第", "").replace("條", "").replace("之", "-").strip()
    parts = pure.split("-")
    
    if len(parts) > 0:
        main_num = chinese_to_arabic(parts[0])
        if len(parts) > 1:
            sub_num = chinese_to_arabic(parts[1])
            return f"{main_num:03d}-{sub_num}"
        else:
            return f"{main_num:03d}"
    return article_str

def get_rtf_text(file_path: str) -> str:
    """直接解析 RTF 內容，不需依賴 MS Word"""
    try:
        # 嘗試以不同的編碼讀取 (台灣法規網通常為 cp950/Big5，或 utf-8)
        try:
            with open(file_path, 'r', encoding='cp950') as file:
                rtf_content = file.read()
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                rtf_content = file.read()
                
        return rtf_to_text(rtf_content)
    except Exception as e:
        messagebox.showerror("錯誤", f"讀取 RTF 檔案失敗：\n{str(e)}")
        return ""

def process_law_data(rtf_path: str, save_path: str):
    plain_text = get_rtf_text(rtf_path)
    if not plain_text:
        return

    lines = plain_text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    
    law_name = ""
    law_date = ""
    current_article = ""
    current_xiang_num = 0
    current_item_lv3 = ""
    current_item_lv4 = ""
    current_content = ""
    
    # 正規表示式編譯
    reg_date = re.compile(r"(修正|發布)日期：?\s*民國\s*(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日")
    reg_article = re.compile(r"^第\s*([一二三四五六七八九十百千\d]+)\s*(?:[-之]\s*([一二三四五六七八九十百千\d]+)\s*)?條(?:[-之]\s*([一二三四五六七八九十百千\d]+))?")
    reg_kuan = re.compile(r"^[一二三四五六七八九十百千]+、")
    reg_mu = re.compile(r"^[（\(][一二三四五六極七八九十百千]+[）\)]")
    reg_trim_digits = re.compile(r"^\d+\s+")
    
    records = []
    
    def add_record():
        if current_article and current_content.strip():
            records.append([
                law_date, law_name, current_article, f"{current_xiang_num:02d}", 
                current_item_lv3, current_item_lv4, current_content.strip(), 
                "", "", "", "", "", datetime.now().strftime("%Y-%m-%d"), ""
            ])

    for line in lines:
        line = line.strip()
        
        # 移除行首數字與空白
        line = reg_trim_digits.sub("", line)
        if not line:
            continue
            
        # 解析法規名稱
        if not law_name:
            if line.startswith("法規名稱："):
                law_name = line[5:].strip()
                continue
            elif not line.startswith("第") and "日期" not in line and len(line) < 50:
                law_name = line
                continue
                
        # 解析發布日期
        if not law_date:
            match = reg_date.search(line)
            if match:
                tw_year = int(match.group(2))
                western_year = tw_year + 1911
                law_date = f"{western_year}-{int(match.group(3)):02d}-{int(match.group(4)):02d}"
                continue
                
        # 解析法條 (條)
        match_article = reg_article.search(line)
        if match_article:
            add_record()
            
            match_val = match_article.group(0)
            current_article = format_tiao(match_val)
            current_xiang_num = 1
            current_item_lv3 = ""
            current_item_lv4 = ""
            
            remainder = line[len(match_val):].strip()
            current_content = remainder
            continue
            
        # 處理款、目與內文
        if current_article:
            match_kuan = reg_kuan.search(line)
            match_mu = reg_mu.search(line)
            
            if match_kuan:
                add_record()
                current_item_lv3 = f"{chinese_to_arabic(match_kuan.group(0)):02d}"
                current_item_lv4 = ""
                current_content = line[len(match_kuan.group(0)):].strip()
            elif match_mu:
                add_record()
                current_item_lv4 = f"{chinese_to_arabic(match_mu.group(0)):02d}"
                current_content = line[len(match_mu.group(0)):].strip()
            else:
                if not current_item_lv3 and not current_item_lv4:
                    curr_txt = current_content.strip()
                    if curr_txt and curr_txt[-1] in ["。", "：", ":", "."]:
                        add_record()
                        current_content = ""
                        current_xiang_num += 1
                        
                if current_content:
                    current_content += "\n" + line
                else:
                    current_content = line

    add_record()  # 加入最後一筆資料
    write_to_excel(records, save_path)

def write_to_excel(records, save_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "法規轉換資料"
    
    headers = ["日期", "法令名稱", "條", "項", "款", "目", "內容", "重點摘要", 
               "適用性", "符合度", "有提升績效機會", "有潛在不符合風險", "鑑別日期", "備註"]
    
    # 寫入標題與格式
    header_fill = PatternFill(start_color="FFEE99", end_color="FFEE99", fill_type="solid")
    header_font = Font(name="微軟正黑體", bold=True)
    
    ws.append(headers)
    for col_idx in range(1, 15):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 寫入資料
    for row_data in records:
        ws.append(row_data)
        
    # 整體格式設定
    max_row = len(records) + 1
    font_default = Font(name="微軟正黑體")
    
    # 對齊方式設定
    align_center = Alignment(horizontal="center", vertical="center")
    align_left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # 欄寬設定
    col_widths = {'A': 12, 'B': 20, 'C': 8, 'D': 8, 'E': 8, 'F': 8, 'G': 55, 
                  'H': 40, 'I': 10, 'J': 10, 'K': 15, 'L': 15, 'M': 12, 'N': 40}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    for r in range(2, max_row + 1):
        for c in range(1, 15):
            cell = ws.cell(row=r, column=c)
            cell.font = font_default
            
            # 文字格式(避免 0 開頭數字消失)
            if c in [3, 4, 5, 6]: 
                cell.number_format = '@'
                cell.alignment = align_center
            elif c in [2, 7, 8, 14]: # 需要自動換行的欄位 (法令名稱, 內容, 重點摘要, 備註)
                cell.alignment = align_left_wrap
            elif c in [1, 9, 10, 11, 12, 13]: # 置中的欄位
                cell.alignment = align_center
                
    # 凍結第一列
    ws.freeze_panes = "A2"
    
    # 資料驗證 (下拉選單)
    if max_row >= 2:
        dv_i = DataValidation(type="list", formula1='"適用,不適用,參考,確認中"', showDropDown=True)
        dv_j = DataValidation(type="list", formula1='"合法,不合法"', showDropDown=True)
        dv_kl = DataValidation(type="list", formula1='" ,v"', showDropDown=True) # K 和 L
        
        ws.add_data_validation(dv_i)
        ws.add_data_validation(dv_j)
        ws.add_data_validation(dv_kl)
        
        dv_i.add(f"I2:I{max_row}")
        dv_j.add(f"J2:J{max_row}")
        dv_kl.add(f"K2:K{max_row}")
        dv_kl.add(f"L2:L{max_row}")

    try:
        wb.save(save_path)
        messagebox.showinfo("完成", f"法規轉換完成！\n資料已儲存至：\n{save_path}")
    except Exception as e:
        messagebox.showerror("儲存錯誤", f"儲存 Excel 失敗，請確認檔案是否被開啟中。\n{str(e)}")

def main():
    root = tk.Tk()
    root.withdraw() # 隱藏主視窗
    
    rtf_path = filedialog.askopenfilename(
        title="請選擇法規 RTF 檔案",
        filetypes=[("RTF Files", "*.rtf")]
    )
    
    if not rtf_path:
        return
        
    # 預設儲存檔名
    default_save_name = os.path.basename(rtf_path).replace(".rtf", "_轉換結果.xlsx")
    save_path = filedialog.asksaveasfilename(
        title="儲存轉換後的 Excel 檔案",
        defaultextension=".xlsx",
        initialfile=default_save_name,
        filetypes=[("Excel Files", "*.xlsx")]
    )
    
    if not save_path:
        return
        
    process_law_data(rtf_path, save_path)

if __name__ == "__main__":
    main()
