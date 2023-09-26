import openpyxl as op
import csv
wb = op.load_workbook("Lab4Data.xlsx", read_only=True, data_only=True)

ws = wb.active

categoryNames = ["Child Labour Total", "Child Labour Male", "Child Labour Female", "Child marriage <15", "Child marriage <18", "Birth Registration Total", "FGM Prevalence Women", "FGM Prevalence Girls", "FGM Support", "Wife Beating Justification Male", "Wife Beating Justification Female", "Violent Discipline Total", "Violent Discipline Male", "Violent Discipline Female"]

outputList = []

for row in ws["B15:AE211"]:
    outputList.append(row[1].value+categoryNames[1] + str(row[3].value))

print(outputList)