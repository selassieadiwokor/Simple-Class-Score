##################################################################################
#   Class Score computation in python using sample data from Excel.
#    XLRD IS USED TO IMPORT & XLWT IS USED TO EXPORT TO EXCEL
#
#   Copyright 2018, Selassie Adiwokor, selassieadiwokor@gmail.com
###################################################################################
import xlrd
import xlwt
loc = "/Users/macos/Desktop/Simple Class Score/exams_score.xlsx"  #navigate to the file path

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

workbook = xlwt.Workbook()
result_sheet = workbook.add_sheet("Grade_Result")

# Specifying style
style = xlwt.easyxf('font: bold 1, color red;')
style2 = xlwt.easyxf('font: bold 1')
style3 = xlwt.easyxf('font: bold 1, color green;')

result_sheet.write(0, 0, 'Name', style)
result_sheet.write(0, 1, '100 Percentage', style)
result_sheet.write(0, 2, "Remarks", style)

for i in range(sheet.nrows):
    in_rows = sheet.row_values(i)
    if(i==0):
        continue
    result = round(in_rows[1]+in_rows[2],2)
    if(result>=80):
        result_sheet.write(i, 0, in_rows[0],style2)
        result_sheet.write(i, 1, result)
        result_sheet.write(i, 2, "Excellent",style3)
    elif(result>=70):
        result_sheet.write(i, 0, in_rows[0],style2)
        result_sheet.write(i, 1, result)
        result_sheet.write(i, 2, "Very Good",style2)
    elif(result>=60):
        result_sheet.write(i, 0, in_rows[0],style2)
        result_sheet.write(i, 1, result)
        result_sheet.write(i, 2, "Good",style2)
    elif(result>=50):
        result_sheet.write(i, 0, in_rows[0],style2)
        result_sheet.write(i, 1, result)
        result_sheet.write(i, 2, "Pass",style2)
    else:
        result_sheet.write(i, 0, in_rows[0],style2)
        result_sheet.write(i, 1, result)
        result_sheet.write(i, 2, "Fail",style)


workbook.save("exams_score_result.xls")
