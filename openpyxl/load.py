from openpyxl import Workbook, load_workbook

# wb = Workbook() # 自動創建excel檔案(obj)
# wb.save('new.xlsx')    # 儲存成'檔案名稱'

wb = load_workbook('excel.xlsx')
ws = wb.active  # 打開的時候預設的工作表
wb.create_sheet('新工作表') # 創建新工作表
# ws = wb['工作表2']  # 指定工作表
ws.title = 'MyTitle'
# print(ws)

B5 = ws['B5'].value # 取值
# print(B5)

ws['B5'].value = 0
wb.save('excel.xlsx')   # 修改要先save (先把excel關掉)