from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

wb = Workbook()
ws = wb.active
ws.title = 'MySheet'

data = [
    {
        'name' : '小白', 
        'height' : 180, 
        'weight' : 74, 
        'age' : 23
    },
    {
        'name' : '小黃', 
        'height' : 177, 
        'weight' : 90, 
        'age' : 28
    },
    {
        'name' : '小綠', 
        'height' : 160, 
        'weight' : 60, 
        'age' : 30
    },
    {
        'name' : '小灰', 
        'height' : 155, 
        'weight' : 50, 
        'age' : 50
    },
    {
        'name' : '小黑', 
        'height' : 170, 
        'weight' : 60, 
        'age' : 46
    },
]

title = ['姓名', '身高', '體重', '年紀']
ws.append(title)

for person in data: # data是一個list
    # print(person.values())  # dict_values(['小白', 180, 74, 23])
    ws.append(list(person.values()))    # append 只能接list['小白', 180, 74, 23]
    
# excel 函式計算 e.g. 某一格[=AVERAGE(E2:E6)]
for col in range(2,5):
    char = get_column_letter(col)
    ws[char + '7'] = f'=AVERAGE({char + "2"}:{char + "6"})'
    
wb.save('practice.xlsx')