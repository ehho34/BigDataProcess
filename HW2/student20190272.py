#!/usr/bin/python3

from openpyxl import load_workbook
wb = openpyxl.load_workbook('student.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')

for i in range (len(sheet['a'])-1):
  midterm = sheet.cell(row=2+i, column=3).value
  final = sheet.cell(row=2+i, column=4).value
  homework = sheet.cell(row=2+i, column=5).value
  attendance = sheet.cell(row=2+i, column=6).value
  
  sheet.cell(row=2+i, column=7).value = midterm*0.2 + final*0.4 + homework*0.39 + attendance*0.1

  total = sheet.cell(row=2+i, column=7).value
  
Ap = int(0.15 * (len(sheet['a'])-1))
Az = int(0.30 * (len(sheet['a'])-1))
Bp = int(0.40 * (len(sheet['a'])-1))
Bz = int(0.50 * (len(sheet['a'])-1))
Cp = int(0.75 * (len(sheet['a'])-1))
Cz = int(1.00 * (len(sheet['a'])-1))

grade = []
result = []
for i in range (len(sheet['a'])-1):
  grade.append(sheet.cell(row=2+i, column=7).value)
  
  
for i in range(0, len(grade)):
    r = 1
    for j in range(0, len(grade)):
        if grade[i] < grade[j]: r += 1
    result.append(r)
    
for i in range(len(result)):
  if result[i] <= Ap:
    sheet.cell(row=2+i, column=8).value = "A+"
  elif result[i] > Ap and result[i] <= Az:
    sheet.cell(row=2+i, column=8).value = "A0"
  elif result[i] > Az and result[i] <= Bp:
    sheet.cell(row=2+i, column=8).value = "B+"
  elif result[i] > Bp and result[i] <= Bz:
    sheet.cell(row=2+i, column=8).value = "B0"
  elif result[i] > Bz and result[i] <= Cp:
    sheet.cell(row=2+i, column=8).value = "C+"
  elif result[i] > Cp and result[i] <= Cz:
    sheet.cell(row=2+i, column=8).value = "C0"
    
wb.save('student.xlsx')
