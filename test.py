import os
import xlwings as xw


excel_list = os.listdir('C:/Users/andyy/Desktop/project1/data')

list.sort(excel_list)

#print(excel_list)

#读取excel数据
#############################################################################
app = xw.App(visible=False,add_book=False)
app.quit()
app = xw.App(visible=False,add_book=False)
wb2 = app.books.add()
data_matrix=[]
#复制的sheet序号，计算从0开始
j=7
#############################################################################
for excel_count in range(0,len(excel_list)):
        wb1 = app.books.open('C:/Users/andyy/Desktop/project1/data'+'/'+excel_list[excel_count])
        # Excel工作簿显示警告,不显示
        app.display_alerts = False
        # 工作簿屏幕更新,不更新
        app.screen_updating = False

        # 获取活动的工作表
        sheet_list=[]
        for i in range (0,26):
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
        print(last_row_up)
        print(last_col_up)
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
app.quit()