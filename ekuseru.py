import openpyxl

wb = openpyxl.Workbook()

ws = wb.active


wb.save("app_data.xlsx")

print("作成完了")