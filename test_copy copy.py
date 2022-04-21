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
print('请输入想生成的表格序号：')
j=input()
print('正在生成...请稍后...')

