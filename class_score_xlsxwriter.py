##################################################################################
#   Class Score computation and Bar chart plot in python using sample
#   data from Excel. XLRD IS USED TO IMPORT & XLSWRITER IS USED TO EXPORT TO EXCEL
#
#   Copyright 2018, Selassie Adiwokor, selassieadiwokor@gmail.com
###################################################################################
import xlrd
import xlsxwriter

print("========CLASS SCORE PROGRAM=====")
print("computation in progress....")

#path to excel file directory on desktop -- please change
loc = "/Users/macos/Desktop/simple Class Score/exams_score.xlsx"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
workbook = xlsxwriter.Workbook('xlwriter_result.xlsx')
res_array = []

for i in range(sheet.nrows):
    in_rows = sheet.row_values(i)
    if(i==0):
        new_arr = [in_rows[0],'Class 100% Score','Remarks']
        res_array.append(new_arr)
        continue
    result = round(in_rows[1]+in_rows[2],2)
    if(result>=80):
        new_arr = [in_rows[0],result,'Excellent']
        res_array.append(new_arr)
    elif(result>=70):
        new_arr = [in_rows[0],result,'Very Good']
        res_array.append(new_arr)
    elif(result>=60):
        new_arr = [in_rows[0],result,'Good']
        res_array.append(new_arr)
    elif(result>=50):
        new_arr = [in_rows[0],result,'Pass']
        res_array.append(new_arr)
    else:
        new_arr = [in_rows[0],result,'Fail']
        res_array.append(new_arr)

# print(es_array)

#add worksheet from a workbook object
worksheet = workbook.add_worksheet()
#create bold fonts
bold = workbook.add_format({'bold': 1})
#loop to insert data into excel
y = 0
arr_length = len(res_array) #array length
for x in range(arr_length):
    y +=1
    if y==1:
        worksheet.write_row('A'+str(y), res_array[x], bold)
    worksheet.write_row('A'+str(y), res_array[x])

#creating a chart
chart1 = workbook.add_chart({'type': 'column'}) #specify column | bar ----

#add a serie
chart1.add_series({
    'name':       ['Sheet1', 0, 1],
    'categories': ['Sheet1', 1, 0, arr_length-1, 0],
    'values':     ['Sheet1', 1, 1, arr_length-1, 1],
})

# Add a chart title
chart1.set_title ({'name': 'Class Score Graph'})

# Add x-axis label
chart1.set_x_axis({'name': 'Student Names'})

# Add y-axis label
chart1.set_y_axis({'name': 'Score in 100%'})

# Set an Excel chart style.
chart1.set_style(11)

# add chart to the worksheet
# the top-left corner of a chart
# is anchored to cell E2 .
worksheet.insert_chart('E2', chart1)
workbook.close()

print("==== computation completed ====")
