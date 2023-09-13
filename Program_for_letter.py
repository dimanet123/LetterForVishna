import pandas as pnd
import numpy as np
import openpyxl
from datetime import date, time, timedelta
import glob
from datetime import datetime
from docx import Document
from docx.shared import Mm
from docx.enum.section import WD_ORIENT
from docx.enum.section import WD_SECTION
from openpyxl.styles import (
                        PatternFill, Border, Side, 
                        Alignment, Font, GradientFill
                        )
from openpyxl.utils.dataframe import dataframe_to_rows


aligment_day = Alignment(horizontal='right',
                    vertical='center',
                    text_rotation=90,
                    wrap_text=True,
                    shrink_to_fit=False,
                    indent=0)

table = openpyxl.Workbook()


MainTable = openpyxl.load_workbook("ListOfMarks.xlsx")
tableCheck = pnd.read_excel('ListOfStudents.xlsx', sheet_name="Worksheet")
tableCheck = tableCheck[tableCheck['Целевое обучение']=="Да - ЦДП"]

for sheetName in MainTable.sheetnames:
    timetableMain = pnd.read_excel('ListOfMarks.xlsx', sheet_name=sheetName)
    timetableMain = timetableMain[timetableMain['ФИО'].isin(tableCheck['Полное ФИО'])]
    sheet = table.create_sheet(sheetName)
    empty_column = pnd.Series([np.nan] * len(timetableMain), name='ФИО')
    timetableMain = timetableMain.drop('№', axis=1)
    timetableMain.insert(0, 'Группа',empty_column)
    timetableMain['Группа'] = sheetName
    for row in dataframe_to_rows(timetableMain, index=False, header=True):
        sheet.append(row)
    for i in range (15):
        sheet.cell(row = 1, column = 4+i).alignment = aligment_day
table.save('Оценки.xlsx')
    

