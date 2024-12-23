from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl import Workbook
from xjzbaobei_ren_dic import *

############################################################
dengdaorenyuan = [jilili]
############################################################

# 获取当前日期和时间
current_date = datetime.now()

# 计算明天的日期
tomorrow = current_date + timedelta(days=1)
tomorrow_month = tomorrow.month
tomorrow_day = tomorrow.day

# 读取文件
file_path = 'D:/GS/xjz/新济州报备/'

# 加载已有的 Excel 文件
wb_h = load_workbook(file_path + '12月18日 南京大学上岛人员报备表.xlsx')
# 获取工作表
sheet_h = wb_h['Sheet1']

# 修改表格的内容
sheet_h['A2'] = f'填报单位：南京大学                    填报日期:   2024 年 {tomorrow_month}月 {tomorrow_day}日'

for i in dengdaorenyuan:
    no = dengdaorenyuan.index(i)
    sheet_h[f'B{7+no}'] = no + 1
    sheet_h[f'C{7+no}'] = i['姓名']
    sheet_h[f'D{7+no}'] = i['年龄']
    sheet_h[f'E{7+no}'] = i['现居地']
    sheet_h[f'F{7+no}'] = i['身份证号码']
    sheet_h[f'G{7+no}'] = i['联系电话']

# 保存工作簿
wb_h.save(file_path + f'{tomorrow_month}月{tomorrow_day}日 南京大学上岛人员报备表.xlsx')