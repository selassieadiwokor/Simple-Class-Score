This program simply explore the use of excel to perform basic data computation and analysis in python3.

A simple excel file containing the the name, class score and exam score of some John Does were
used as a sample data. 

Python libraries such as (PANDAS, XLSXWRITER,XLRD,XLWT) were used to import and export data
from excel and back to excel. 

This was further updated with a Bar Chart of the exported result.
The programs present different alternative to plotting. 


class_score
==============
This program leverage XLRD,XLWT to read the data from excel perform the calculation (finds the 100% by
adding the class score to the exam score) and output the result in excl.

class_score_xlsxwriter
=======================
This program leverage XLRD & XLSWRITER as python libraries to read from excel and the write back to excel 
with a plotted Bar Chart of Scores against names.

class_score_pandas
===================
This program leverage PANDAS as python library to read and write to excel.

Note:
I will personally recommend the use of PANDAS since it's a great python library for big data analysis

Thank you.