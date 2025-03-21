import os
import xlwings as xw
import time
from utils import get_project_path
#调用utils模块获取文件相对路径
#############################################################################
path=get_project_path()
print(path)
#对获取的路径用r进行\和/的转换
excel_list = os.listdir(r'%s/data'%(path))
#对文件进行排序
list.sort(excel_list)
print(excel_list)
#读取excel数据
#############################################################################
app = xw.App(visible=False,add_book=False)
wb = app.books.open(r'%s/data'%(path)+'/'+excel_list[0])
app.display_alerts = False
# 工作簿屏幕更新,不更新
app.screen_updating = False
#遍历所有sheet，并生成待序号的list
sheet_name_name=[]
sheet_name_name_number=[]
for sheet_name_name_counting in range(0,len(wb.sheets)):
        if sheet_name_name_counting >=0:
                sheet= wb.sheets[sheet_name_name_counting]
                sheet_name_name_number.append(sheet_name_name_counting)
                sheet_name_name.append(sheet.name)
                sheet_name=list(zip(sheet_name_name_number,sheet_name_name))
                sheet_name_name_counting+=1
        else:
                pass
wb.close()
app.quit()
for k in range(0,len(sheet_name)):
        print(sheet_name[k])
        time.sleep(0.05)
j=int(input('请输入想生成的表格序号：'))
while int(j) not in sheet_name_name_number:
        j=int(input('输入有误，请重新输入:'))
print('正在生成...请稍后...')
#主程序
#############################################################################
app = xw.App(visible=False,add_book=False)
wb2 = app.books.add()
data_matrix=[]
#############################################################################
for excel_count in range(0,len(excel_list)):
        print('当前进度%sin%s'%(excel_count+1,len(excel_list)+1))
        wb1 = app.books.open(r'%s/data'%(path)+'/'+excel_list[excel_count])
        # Excel工作簿显示警告,不显示
        app.display_alerts = False
        # 工作簿屏幕更新,不更新
        app.screen_updating = False

        # 获取活动的工作表
        sheet_list=[]
        for i in range (0,len(sheet_name_name)-1):
                sheet_list.append(wb1.sheets[i])
        #print(sheet_list)

        sheet1 = wb1.sheets[sheet_list[j]]
        sheet2 = wb2.sheets[0]
        
        last_cell = sheet1.used_range.last_cell
        # 最大行数
        last_row = last_cell.row
        # 最大列数
        last_col = last_cell.column
        last_cell = sheet1.used_range.last_cell

        sheet1.range((1, 1), (last_row, last_col)).copy()
        # 最大行数
        last_cell = sheet2.used_range.last_cell
        last_row_up= last_cell.row
        last_col_up= last_cell.column
        # 写入二维列表,追加模式
        sheet2.range((last_row_up + 2, 1)).paste()
        wb1.app.api.CutCopyMode=False
        wb2.app.api.CutCopyMode=False

        last_cell = sheet2.used_range.last_cell
        # 最大行数
        last_row_down= last_cell.row
        # 最大列数
        last_col_down= last_cell.column
        #在最右边那一列加上公司文件名称
        sheet2.range((last_row_up+2,9),(last_row_down+2,9)).value=('%s'%(excel_list[excel_count]))
        sheet2.range((1, 1), (last_row, last_col)).columns.autofit()
        wb1.close

wb2.save('%s.xlsx'%(sheet1.name))
wb2.close
app.quit()