from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

wb = load_workbook('excel.xlsx')
ws = wb.active

# ls = [1, 2, 3, 4, 5]
# ws.append(ls)   # 在游標位置(新檔案在A1)下一行 append內容

# 寫值進去
for row in range(1,6):
    for col in range(1,6):
        char = get_column_letter(col)
        ws[char + str(row)] = char + str(row)

# 讀值出來
for row in range(1,6):
    for col in range(1,6):
        # A1 B1 C1 ... A2 B2 C2 ...
        char = get_column_letter(col)
        print(ws[char + str(row)].value)

# 移動資料 (rows 正上負下; cols 正右負左)
# ws.move_range('A1:E5', rows=2, cols=2)
# ws.move_range('C3:G7', rows=-2, cols=-2)


# 插入直行橫列 
# ws.insert_cols(3)   # 在 C 行插入
# ws.insert_rows(3)   # 在第三列插入

# 合併儲存格
# ws.merge_cells('A1:E1')
# ws.unmerge_cells('A1:E1')

# 變粗體字 (import Font)
# for col in range(1,6):
#     char = get_column_letter(col)
#     ws[char + '1'].font = Font(bold=True)

wb.save('excel.xlsx')