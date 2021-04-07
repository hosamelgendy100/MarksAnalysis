from openpyxl import Workbook, load_workbook
import arabic_reshaper
from bidi.algorithm import get_display
import re
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment


def ArabicPrint(valueToPrint):
    text = arabic_reshaper.reshape(valueToPrint)
    print(get_display(text))


def RemoveStrings(sequance):
    seq_type = type(sequance)
    return seq_type().join(filter(seq_type.isdigit, sequance))


def RemoveNumbers(word):
    regex = re.compile('[^a-zA-Z]')
    return regex.sub('', word)


# Import Workbook
workbook = load_workbook('8th.xlsx')
sheet = workbook.active

nameCN = "B5"
classCN = "C5"
schoolName = "مدرسة معيذر الابتدائية للبنين"

classCellValue = sheet[classCN].value

# get the first cell number
cellNumber = int(RemoveStrings(nameCN))-1
subjectsCN = ["E5", "G5", "I5", "K5", "M5", "O5", "Q5"]

# creats new Excel File
finishedWorkbook = Workbook()
# Removes the default created workbook
finishedWorkbook.remove(finishedWorkbook.active)


def CalculateTotal(subjects):
    total = 0
    for i in range(len(subjects)):
        subjects[i] = RemoveNumbers(subjects[i])
        subjects[i] = subjects[i] + str(cellNumber)
        total += sheet[subjects[i]].value
    return total


def AddNewSheet(name):
    global finishedWorkbook
    sheetName = str(name).replace('/', '-')
    newSheet = finishedWorkbook.create_sheet(sheetName)
    newSheet.sheet_view.rightToLeft = True

    newSheet.merge_cells('A1:B2')
    cell = newSheet.cell(1, 1, schoolName)
    cell.alignment = Alignment(horizontal='center', vertical='center')

    newSheet["A3"] = "م"
    newSheet["B3"] = "أسماء الطلاب"
    newSheet["C3"] = "التقييم الرئيسي"

    return newSheet


def AdjustSheetWidthAndAlign(newSheet):
    for col in newSheet.columns:
        for cell in col:
            try:
                cell.alignment = Alignment(horizontal='right')
            except:
                pass

    newSheet.column_dimensions['A'].width = 5
    newSheet.column_dimensions['B'].width = 50
    newSheet.column_dimensions['C'].width = 10


def AnalyzeData():
    global nameCN
    global classCN
    global cellNumber
    global classCellValue
    newSheetCellNumber = 3

    # Save Workbook
    newSheet = AddNewSheet(sheet[classCN].value)
    newSheet.sheet_view.rightToLeft = True

    while(sheet[nameCN].value != None):
        nameCN = RemoveNumbers(nameCN)
        classCN = RemoveNumbers(classCN)

        # Incerment Excel
        cellNumber = int(cellNumber) + 1
        nameCN = nameCN + str(cellNumber)
        classCN = classCN + str(cellNumber)

        if(sheet[nameCN].value == None):
            AdjustSheetWidthAndAlign(newSheet)
            return

        # Add New Tab
        if(classCellValue != sheet[classCN].value):
            classCellValue = sheet[classCN].value
            AdjustSheetWidthAndAlign(newSheet)
            newSheet = AddNewSheet(sheet[classCN].value)
            newSheetCellNumber = 3

        # Save data to excel workbook
        newSheetCellNumber += 1
        newSheet["A"+str(newSheetCellNumber)] = str(newSheetCellNumber-1)
        newSheet["B"+str(newSheetCellNumber)] = str(sheet[nameCN].value)
        newSheet["C"+str(newSheetCellNumber)] = str(CalculateTotal(subjectsCN))


AnalyzeData()
finishedWorkbook.save(filename="Analysis.xlsx")
#os.startfile('Analysis.xlsx')
