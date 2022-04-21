import os
import xlwings as xw

excel_list = os.listdir('C:/Users/andyy/Desktop/project1/data')

list.sort(excel_list)

#print(excel_list)

#读取excel数据
#############################################################################
app = xw.App(visible=False,add_book=False)
data_matrix=[]
#print(excel_list[0])
for excel_count in range(0,len(excel_list)):
        wb1 = app.books.open('C:/Users/andyy/Desktop/project1/data'+'/'+excel_list[excel_count])
        wb2 = app.books.add()

# 获取活动的工作表
        sheet_list=[]
        for i in range (0,26):
                sheet_list.append(wb1.sheets[i])
        #print(sheet_list)

        sheet1 = wb1.sheets[sheet_list[3]]
        sheet2 = wb2.sheets[0]
        

        last_cell = sheet1.used_range.last_cell
        # 最大行数
        last_row = last_cell.row
        # 最大列数
        last_col = last_cell.column

        data = sheet1.range((1, 1), (last_row, last_col)).api.Copy()
        
        wb1.app.api.CutCopyMode=False
        wb1.close()

print(data_matrix)
#print(data)
app.quit()



#写入excel数据
#############################################################################
app = xw.App(visible=False, add_book=False)
# Excel工作簿显示警告,不显示
app.display_alerts = False
# 工作簿屏幕更新,不更新
app.screen_updating = False
# 创建工作簿
wb1 = app.books.add()

sheet1 = wb1.sheets[0]

for data_count in range(0,72):

        last_cell = sheet1.used_range.last_cell
        # 最大行数
        last_row_up= last_cell.row
        last_col_up= last_cell.column

        # 写入二维列表,追加模式
        sheet1.range((last_row_up + 2, 1)).options(expand='table').api.Select = data_matrix[data_count]
            
        last_cell = sheet1.used_range.last_cell
                # 最大行数
        last_row_down= last_cell.row
                # 最大列数
        last_col_down= last_cell.column
                # 在range中,cell的大小自适应
        sheet1.range((1, 1), (last_row, last_col)).columns.autofit()

        #for wrinte_count in (last_row_up,last_row_down):
        #print(wrinte_count)
        sheet1.range((last_row_up+2,9),(last_row_down+2,9)).value=('%s'%excel_list[data_count])


# 保存文件
wb1.save()
# 关闭工作簿
wb1.close()
# 退出Excel
app.quit()


