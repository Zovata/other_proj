from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl import Workbook
from xjzbaobei_dic import *

############################################################
dengdaocheliang = [jilili]
############################################################

# 获取当前日期和时间
current_date = datetime.now()

# 计算明天的日期
tomorrow = current_date + timedelta(days=1)
tomorrow_month = tomorrow.month
tomorrow_day = tomorrow.day

# 读取文件
file_path = 'D:/GS/xjz/新济州报备/'
print(file_path + '12月13日 南京大学登岛车辆报备.xlsx')

# 加载已有的 Excel 文件
wb_car = load_workbook(file_path + '12月13日 南京大学登岛车辆报备.xlsx')
# 获取工作表
sheet_car = wb_car['Sheet1']

# 修改表格的内容
sheet_car['A2'] = f'填报单位：南京大学                    填报日期:   2024 年 {tomorrow_month}月 {tomorrow_day}日'

for i in dengdaocheliang:
    no = dengdaocheliang.index(i)
    sheet_car[f'B{7+no}'] = no + 1
    sheet_car[f'C{7+no}'] = i['单位名称']
    sheet_car[f'D{7+no}'] = i['车牌号']
    sheet_car[f'E{7+no}'] = i['驾驶员姓名']
    sheet_car[f'F{7+no}'] = i['联系号码']
    sheet_car[f'G{7+no}'] = i['车辆型号']

# 保存工作簿
wb_car.save(file_path + f'{tomorrow_month}月{tomorrow_day}日 南京大学登岛车辆报备.xlsx')