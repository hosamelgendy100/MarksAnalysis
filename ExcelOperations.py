import os
import pathlib
import re
import webbrowser
from openpyxl.styles import Alignment, alignment
from openpyxl.styles.borders import Side, Border
from AppConfiguration import CELL_FILL_SETTINGS, CELL_FONT_SETTINGS
import arabic_reshaper
from bidi.algorithm import get_display

def AddAverage(newSheet, newSheetCellNumber, titleRow):
    newSheet["B"+str(newSheetCellNumber+2)] = "متوسط التحصيل الأكاديمي"
    newSheet["C"+str(newSheetCellNumber+2)] = "=Average(C{firstVal}:C{lastVal})".format(
        firstVal=int(titleRow), lastVal=str(newSheetCellNumber))

    newSheet["C"+str(newSheetCellNumber+2)].number_format = '0.00'

    newSheet["B"+str(newSheetCellNumber+2)].fill = CELL_FILL_SETTINGS
    newSheet["C"+str(newSheetCellNumber+2)].fill = CELL_FILL_SETTINGS

    newSheet["B"+str(newSheetCellNumber+2)].font = CELL_FONT_SETTINGS
    newSheet["C"+str(newSheetCellNumber+2)].font = CELL_FONT_SETTINGS


def RemoveStrings(sequance):
    seq_type = type(sequance)
    return seq_type().join(filter(seq_type.isdigit, sequance))


def RemoveNumbers(word):
    regex = re.compile('[^a-zA-Z]')
    return regex.sub('', word)


def getSubjectsTotal(subjectsFullMark, subjectsCN):
    return subjectsFullMark * len(subjectsCN)



def AddRankeFunction(newSheet, lastValue):
    for i in range(5, lastValue+1):
        newSheet['E'+str(i)] = "=RANK(C{i},$C$5:$C${lastValue},0)".format(
            i=i, lastValue=lastValue)
        newSheet['E'+str(i)].font = CELL_FONT_SETTINGS


def ArabicPrint(valueToPrint):
    text = arabic_reshaper.reshape(valueToPrint)
    print(get_display(text))



def SaveExcelFile(finishedWorkbook, sheetTitle):
    # Get Desktop Path
    desktopPath = str(pathlib.Path.home()) + '/Desktop/نتائج التحليل/'
    if not os.path.exists(desktopPath):
        os.mkdir(desktopPath)

    i = 1
    excelFileName = str(sheetTitle) + '.xlsx'

    # Check if file exists
    while(os.access(desktopPath + excelFileName, os.R_OK | os.X_OK)):
        excelFileName = str(sheetTitle)
        excelFileName += '( ' + str(i) + ' )' + '.xlsx'
        i += 1

    finishedWorkbook.save(desktopPath+excelFileName)
    webbrowser.open(os.path.realpath(desktopPath))
    os.startfile(desktopPath+excelFileName)


def SetFontRange(workSheet, cell_range):
    for row in workSheet[cell_range]:
        for cell in row:
            cell.font = CELL_FONT_SETTINGS


def SetFillRange(workSheet, cell_range):
    for row in workSheet[cell_range]:
        for cell in row:
            cell.fill = CELL_FILL_SETTINGS


def SetAlignRange(workSheet, cell_range,alignPos):
    for row in workSheet[cell_range]:
        for cell in row:
            cell.alignment = alignPos



def AdjustSheetWidth(newSheet):
    newSheet.column_dimensions['A'].width = 5
    newSheet.column_dimensions['B'].width = 40
    newSheet.column_dimensions['E'].width = 7


def SetBorders(workSheet, cell_range):
    thin = Side(border_style="thin", color="000000")
    for row in workSheet[cell_range]:
        for cell in row:
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
