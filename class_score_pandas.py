##########################################################################
#   Class Score computation and Bar chart plot in python using sample
#   data from Excel. PANDAS DATAFRAME IS USED TO IMPORT & EXPORT TO EXCEL
#
#   Copyright 2018, Selassie Adiwokor, selassieadiwokor@gmail.com
###########################################################################
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

print("......Processing in progress....... ")
 
df = pd.read_excel('/Users/macos/Desktop/Simple Class Score/exams_score.xlsx', sheet_name='Sheet1')
 
# print(df.columns)

result_arr = []
remark_arr = []

# Add the class score and example score, store into an array & generate a remark array
# with appropriate conditions

for i in df.index:
    res = float(df['class score'][i])+ float(df['exam score'][i])
    result_arr.append(round(res,2))
    if(res>=80):
        remark_arr.append('Excellent')
    elif(res>=70):
        remark_arr.append('Very Good')
    elif(res>=60):
        remark_arr.append('Good')
    elif(res>=50):
        remark_arr.append('Pass')
    else:
        remark_arr.append('Fail')

# print(remark_arr)
# print(result_arr)

# Create a dictionary for data to be exported
exp_result = {'Name':df['name'],'Class Score':df['class score'],'Exam Score':df['exam score'],'Total Score (100%)':result_arr,'Remarks':remark_arr}
# print(exp_result)
df = pd.DataFrame(exp_result)
 
writer = ExcelWriter('Pandas_Result.xlsx')
df.to_excel(writer,'Sheet1',index=False)

# Access the XlsxWriter workbook and worksheet objects from the dataframe.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Create a chart object
chart = workbook.add_chart({'type': 'column'})

# Configure the series of the chart from the dataframe data.
chart.add_series({
    'name':       ['Sheet1', 0, 3],
    'categories': ['Sheet1', 1, 0,   len(result_arr)-1, 0],
    'values':     ['Sheet1', 1, 3, len(result_arr)-1, 3],
})

# Configure the chart axes
chart.set_x_axis({'name': 'Student Names'})
chart.set_y_axis({'name': 'Class Score', 'major_gridlines': {'visible': False}})

# Insert the chart into the worksheet.
worksheet.insert_chart('G2', chart)
# Save result and graph in excel
writer.save()

print("==== computation completed ====")