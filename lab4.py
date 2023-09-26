import openpyxl as op
wb= op.load_workbook("Lab4Data.xlsx", read_only=True, data_only=True)

ws=wb["Table 9"]

for row in ws[]