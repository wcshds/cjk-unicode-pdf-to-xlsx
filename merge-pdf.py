import xlwings as xw


# 僅可在 Windows 或 Mac OS 上運行

path1 = './result/basic-1.xlsx'
path2 = './result/basic-2.xlsx'

app = xw.App(visible=False)

wb1 = app.books.open(path1)
wb2 = app.books.open(path2)

# 第二個xlsx
wb2_ws4 = wb2.sheets(1)
wb2_ws5 = wb2.sheets(2)
wb2_ws6 = wb2.sheets(3)

wb2_ws4.api.Copy(After=wb1.sheets(3).api)
wb2_ws5.api.Copy(After=wb1.sheets(4).api)
wb2_ws6.api.Copy(After=wb1.sheets(5).api)

wb1.sheets(4).name = "Sheet4"
wb1.sheets(5).name = "Sheet5"
wb1.sheets(6).name = "Sheet6"

wb1.save("./result/basic-merge.xlsx")
app.quit()
