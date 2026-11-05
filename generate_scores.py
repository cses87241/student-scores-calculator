#!/usr/bin/env python3
"""
學生成績自動計算表產生器
執行此腳本即可產生 student_scores.xlsx
需要安裝：pip install openpyxl
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = Workbook()
ws = wb.active
ws.title = "學生成績"

headers = ["座號", "國語", "數學", "社會", "自然", "英語", "平均", "級距"]
for col, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = Font(bold=True, size=12, name="Microsoft JhengHei", color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="4472C4")
    cell.alignment = Alignment(horizontal="center", vertical="center")

for i in range(1, 30):
    row = i + 1
    ws.cell(row=row, column=1, value=i).alignment = Alignment(horizontal="center")
    for col in range(2, 7):
        cell = ws.cell(row=row, column=col, value=None)
        cell.font = Font(color="0000FF", name="Microsoft JhengHei")
        cell.alignment = Alignment(horizontal="center")
    
    avg_cell = ws.cell(row=row, column=7)
    avg_cell.value = f"=IF(COUNT(B{row}:F{row})=0,\"\",AVERAGE(B{row}:F{row}))"
    avg_cell.font = Font(name="Microsoft JhengHei")
    avg_cell.alignment = Alignment(horizontal="center")
    avg_cell.number_format = "0.00"
    
    grade_cell = ws.cell(row=row, column=8)
    grade_cell.value = f'=IF(G{row}="","",IF(G{row}=100,"100",IF(G{row}>=90,"90-99",IF(G{row}>=80,"80-89",IF(G{row}>=70,"70-79",IF(G{row}>=60,"60-69","未滿60"))))))'
    grade_cell.font = Font(name="Microsoft JhengHei")
    grade_cell.alignment = Alignment(horizontal="center")

summary_row = 31
ws.cell(row=summary_row, column=1, value="全班平均").font = Font(bold=True, name="Microsoft JhengHei")
for col in range(2, 8):
    col_letter = get_column_letter(col)
    cell = ws.cell(row=summary_row, column=col)
    if col < 7:
        cell.value = f"=IF(COUNT({col_letter}2:{col_letter}30)=0,\"\",AVERAGE({col_letter}2:{col_letter}30))"
    else:
        cell.value = f"=IF(COUNT(G2:G30)=0,\"\",AVERAGE(G2:G30))"
    cell.font = Font(bold=True, name="Microsoft JhengHei")
    cell.alignment = Alignment(horizontal="center")
    cell.number_format = "0.00"
    cell.fill = PatternFill("solid", fgColor="FFF2CC")

ws.cell(row=33, column=1, value="級距統計").font = Font(bold=True, name="Microsoft JhengHei", size=12)
grades = ["100", "90-99", "80-89", "70-79", "60-69", "未滿60"]
for idx, grade in enumerate(grades):
    r = 34 + idx
    ws.cell(row=r, column=1, value=grade).font = Font(name="Microsoft JhengHei")
    ws.cell(row=r, column=2, value=f'=COUNTIF(H2:H30,"{grade}")').font = Font(name="Microsoft JhengHei")
    ws.cell(row=r, column=3, value="人").font = Font(name="Microsoft JhengHei")

widths = [10, 12, 12, 12, 12, 12, 12, 12]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.row_dimensions[1].height = 25

thin = Border(left=Side(style='thin'), right=Side(style='thin'),
              top=Side(style='thin'), bottom=Side(style='thin'))
for row in ws.iter_rows(min_row=1, max_row=30, min_col=1, max_col=8):
    for cell in row:
        cell.border = thin

ws2 = wb.create_sheet("使用說明")
ws2["A1"] = "學生成績自動計算表使用說明"
ws2["A1"].font = Font(bold=True, size=14, name="Microsoft JhengHei")
instructions = [
    "1. 在「學生成績」工作表中，於 B2:F30 輸入各科分數（藍色字體欄位）。",
    "2. 平均分數與級距會自動計算。",
    "3. 級距標準（六個級距）：",
    "   - 100分：100",
    "   - 90-99分：90-99",
    "   - 80-89分：80-89",
    "   - 70-79分：70-79",
    "   - 60-69分：60-69",
    "   - 未滿60分：未滿60",
    "4. 底部有全班各科平均與級距人數統計。",
    "5. 可直接修改分數，所有公式會自動更新。",
    "6. 適用於任何安裝 Microsoft Excel 或相容軟體（如 LibreOffice Calc）的電腦。"
]
for i, text in enumerate(instructions, 3):
    ws2[f"A{i}"] = text
    ws2[f"A{i}"].font = Font(name="Microsoft JhengHei", size=11)
ws2.column_dimensions["A"].width = 85

wb.save("student_scores.xlsx")
print("已產生 student_scores.xlsx")
