import openpyxl as op
import csv
wb = op.load_workbook("Lab4Data.xlsx", read_only=True, data_only=True)

ws = wb.active

categoryNames = ["Child Labour Total", "Child Labour Male", "Child Labour Female", "Child marriage <15", "Child marriage <18", "Birth Registration Total", "FGM Prevalence Women", "FGM Prevalence Girls", "FGM Support", "Wife Beating Justification Male", "Wife Beating Justification Female", "Violent Discipline Total", "Violent Discipline Male", "Violent Discipline Female"]

outputList = []

for row in ws["B15:AE211"]:
    for catInd in range(len(categoryNames)):
        if row[3+2*catInd].value != "–" and row[3+2*catInd].value != None:
            outputList.append([row[0].value, categoryNames[catInd], row[3+2*catInd].value])

fptr = open("group10Lab4.csv", "w", newline="")
writer = csv.writer(fptr)
writer.writerows(outputList)
fptr.close()

fptr2 = open("group10Lab4.csv", "r")
print(sum(1 for row in fptr2))
fptr2.close()

