# Lab4-DSC200
# Nathan Reed Abhishek Shrestha
# Write a program that reads data from an Excel file a list of child abuse events for several countries
# and write the data into a csv file. Output the length of the resulting csv file.


# import libraries for opening excel and csv files
import openpyxl as op
import csv
# load the workbook from the Excel file
wb = op.load_workbook("Lab4Data.xlsx", read_only=True, data_only=True)

# Select the active sheet of the workbook.
ws = wb.active

# Define a list of category names that will be used later.
categoryNames = ["Child Labour Total", "Child Labour Male", "Child Labour Female", "Child marriage <15",
                 "Child Marriage <18", "Birth Registration Total", "FGM Prevalence Women",
                 "FGM Prevalence Girls", "FGM Support", "Wife Beating Justification Male",
                 "Wife Beating Justification Female", "Violent Discipline Total", "Violent Discipline Male",
                 "Violent Discipline Female"]

# Create an empty list to store the extracted data.
outputList = [["CountryName", "CategoryName", "CategoryTotal"]]


# Iterate through rows in the Excel worksheet from B15 to AE211.
for row in ws["B15:AE211"]:
    # Iterate through category indices in the categoryNames list.
    for catInd in range(len(categoryNames)):
        # Check if the cell value is not "–" (en dash), not None, and not 0
        if row[3+2*catInd].value != "–" and row[3+2*catInd].value is not None and row[3+2*catInd].value != 0:
            # Append data to the outputList as a list containing country name, category name, and cell value.
            outputList.append([row[0].value, categoryNames[catInd], row[3+2*catInd].value])


# Open a CSV file named "group10Lab4.csv" for writing
fptr = open("group10Lab4.csv", "w", newline="")
writer = csv.writer(fptr)
# Write the data from outputList to the CSV file.
writer.writerows(outputList)
fptr.close()

# Open the CSV file "group10Lab4.csv" for reading.
fptr2 = open("group10Lab4.csv", "r")
# Print the number of rows in the CSV file (counting the number of lines).
print(sum(1 for row in fptr2))
# Close the CSV file.
fptr2.close()
