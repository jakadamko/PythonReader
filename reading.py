#!/usr/bin/env python

#import xlrd module
import xlrd
import csv

#open spreadsheet workbook
workbook = xlrd.open_workbook("UsersBook.xlsx")

#open sheet
worksheet = workbook.sheet_by_name("Users")

#extract value of specific data cell
print("the value at row 4 and column 2 is : {0}".format(worksheet.cell(4,2).value))


#total number of rows and columns in the sheet
total_rows = worksheet.nrows
total_cols = worksheet.ncols
print("number of rows : {0}, and number of columns : {1}".format(total_rows,total_cols))


#loop through every record in the worksheet and store them in a list then display the final list
table = list()
record = list()

#Read all data in worksheet
for x in range(total_rows):
    for y in range(total_cols):
        record.append(worksheet.cell(x,y).value)
    table.append(record)
    record = []
    x +=1
print(table)


#export CSV file
with open('a_file.csv', 'w', newline="") as f:
    c = csv.writer(f)
    for r in range(worksheet.nrows):
        c.writerow(worksheet.row_values(r))

#Confirm to user file created
print("CSV File successfuly exported")
        
