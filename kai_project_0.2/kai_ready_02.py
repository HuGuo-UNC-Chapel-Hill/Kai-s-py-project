###################
# Author: Hu Guo
# This program is used to automatically schedule work for members of Champaign Chinese Christian Church on Sunday.
# v0.1002
###################
# 安裝Python的Excel插件
# 在命令行下輸入： pip3 install openpyxl
import openpyxl
from openpyxl import Workbook
from datetime import datetime
from openpyxl.styles import PatternFill, Alignment

# 讀取Excel表格初始可用人員
wb = openpyxl.load_workbook("./kai_Excel_2.xlsx", data_only=True)
# print(type(wb))

sheets = wb.sheetnames
# print(sheets)
# print(wb.active.title)

sh1 = wb['Sheet1']
# print(type(sh1))

####################################################
# 獲取上月已排班兩次人員名單
last_month_assigned_twice = []

month = sh1['F1']
print(month)


# 生成上月值班表格文件名
str_p = str(sh1['F1'].value - 1) + "月擔班情況.xlsx"

try:
    wp = openpyxl.load_workbook(str_p)
    shp_1 = wp['sheet1']
    # print(shp_1.cell(1, 1).value)
    for i in range(1, 25):
        name = shp_1.cell(2, i).value
        ct = shp_1.cell(3, i).value
        if ct == 2:
            # print(name)
            last_month_assigned_twice.append(name)
except:
    print("沒有發現" + str_p)

print("上月擔班了兩次的人員名單: ", last_month_assigned_twice)
print()

# 生成周日具體日期list
sundays = []
for i in range(2, 7):
    date = sh1.cell(4, i).value
    if date != None:
        date2 = str(date)
        dt = datetime.strptime(date2, '%Y-%m-%d %H:%M:%S')
        str_date = dt.strftime('%m/%d')
        if date != None:
            sundays.append(str_date)
print("本月所有週日日期： ", sundays)
print()

# 本月所有週日天數
days = len(sundays)
# print(days)

# 待安排項目數量
tasks = 4

# 建立空白初始人員2D list
lists = [[], [], [], [], []]

# 建立空白的Dict來存儲已安排了項目的人員與值班的天數
assigned = dict()

# 從初始表格讀取待分配人員並加入到待分配人員2D list
cols = 22
k = 0
while (k < days):
    for j in range(17, cols):
        for i in range(5, 26):
            name = sh1.cell(i, j).value
            if name == None:
                continue
            lists[k].append(name)
            assigned[name] = 0
        k += 1

# 測試打印初始人員名單
for i in range(0, days):
    print(sundays[i], "週日可安排人員如下：")
    print(lists[i])
    print()

# 建立空白項目分配2D list
arranged_schedule = [[], [], [], [], []]

# 參照上月擔班概要降低上月已經擔班兩次的人員的優先級別
for x in assigned:
    if x in last_month_assigned_twice:
        assigned[x] = 1
print("本月所有可安排人員如下， 共", len(assigned), "人, 如果上月已經擔班兩次，那麼初始優先級會通過 \"+1\" 值班日會降低：")
print(assigned)
print()
#############################################################################
# 進行排班運算並生成2D list
for d in range(0, days):
    temp = [x for x in lists[d] if assigned.get(x) == 0]
    #    print(d, temp)
    if len(temp) >= tasks:
        for i in range(0, tasks):
            arranged_schedule[d].append(temp[i])
            assigned[temp[i]] = assigned.get(temp[i]) + 1

    elif len(temp) < tasks:
        for i in range(0, len(temp)):
            arranged_schedule[d].append(temp[i])
            assigned[temp[i]] = assigned.get(temp[i]) + 1

        temp_1 = [x for x in lists[d] if assigned.get(x) == 1]
        #    print(d, temp_0)
        if len(temp_1) >= tasks:
            for i in range(0, tasks - len(temp)):
                arranged_schedule[d].append(temp_1[i])
                assigned[temp_1[i]] = assigned.get(temp_1[i]) + 1

        elif len(temp_1) < tasks:
            for i in range(0, tasks - len(temp) - len(temp_1)):
                arranged_schedule[d].append("缺少人員")

print("本月排班結果： 共", len(sundays), "個週日。")
for day in range(0, days):
    print(sundays[day], arranged_schedule[day])
print()

print("根據上月擔班概要和本月值班概要, 增加只擔班一次人員的下月擔班優先級.")
print("如果上月已經擔班兩次，本月在最開始運行程序時已經手動 \"+1\" 降低過擔班優先級。")
print("所以上月擔班兩次的人員在本月實際擔班一次的情況下會顯示值班2次。那麼現在會通過 \"-1\" 增加下月排班的優先級，下月能夠擔班兩次：")
print()

# 根據上月擔班概要和本月值班概要增加只擔班一次人員的下月擔班優先級
for x in assigned:
    if x in last_month_assigned_twice:
        assigned[x] = assigned.get(x) - 1
print("優化完畢")
print()

one_assigned = [x for x in assigned if assigned.get(x) == 1]
print("本月實際擔班一次的人員有", len(one_assigned), "人。 下月可以擔班兩次。")
print(one_assigned)
print()

twice_assigned = [x for x in assigned if assigned.get(x) == 2]
print("本月實際擔班兩次的人員有", len(twice_assigned), "人。 下月盡量只安排一次擔班。")
print(twice_assigned)

####################################################
# 保存排班概況與詳細安排到新xlsx文件
wv = Workbook()
wv['Sheet'].title = "sheet1"
shv = wv.active
shv['A1'].value = sh1['E1'].value
shv['B1'].value = "月擔班情況"
count = 1
for x in assigned:
    shv.cell(2, count).value = x
    currentCell = shv.cell(2, count)  # or currentCell = ws['A1']
    currentCell.alignment = Alignment(horizontal='center')
    shv.cell(3, count).value = assigned[x]
    shv.cell(4, count).value = "次"
    currentCell = shv.cell(4, count)  # or currentCell = ws['A1']
    currentCell.alignment = Alignment(horizontal='right')
    count += 1

b = 7
for x in sundays:
    shv.cell(b, 1).value = x
    b += 1

shv['B6'].value = "影視主控"
shv['B6'].fill = PatternFill("solid", fgColor="FFC300")
currentCell = shv['B6']
currentCell.alignment = Alignment(horizontal='center')
shv['C6'].value = "影視副控"
shv['C6'].fill = PatternFill("solid", fgColor="FFC300")
currentCell = shv['C6']
currentCell.alignment = Alignment(horizontal='center')
shv['D6'].value = "門口招待"
shv['D6'].fill = PatternFill("solid", fgColor="FFC300")
currentCell = shv['D6']
currentCell.alignment = Alignment(horizontal='center')
shv['E6'].value = "堂內招待"
shv['E6'].fill = PatternFill("solid", fgColor="FFC300")
currentCell = shv['E6']
currentCell.alignment = Alignment(horizontal='center')

for i in range(7, 7 + days):
    for j in range(2, 6):
        shv.cell(i, j).value = arranged_schedule[i - 7][j - 2]
        currentCell = shv.cell(i, j)
        currentCell.alignment = Alignment(horizontal='center')
        # if currentCell.value == "藍凱威":
        if currentCell.value == "缺少人員":
            shv.cell(i, j + 5).value = "空缺建議："
            list_suggest = [x for x in lists[i - 7] if x not in arranged_schedule[i - 7]]
            if len(list_suggest) == 0:
                shv.cell(i, j + 6).value = "無建議"
            else:
                for s in range(0, len(list_suggest)):
                    shv.cell(i, j + 6 + s).value = list_suggest[s]

str_1 = str(sh1['E1'].value) + "月擔班情況" + ".xlsx"
wv.save(str_1)
