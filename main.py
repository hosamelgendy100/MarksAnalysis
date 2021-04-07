from openpyxl import Workbook, load_workbook
import arabic_reshaper
from bidi.algorithm import get_display
import re


def ArabicPrint(valueToPrint):
    text = arabic_reshaper.reshape(valueToPrint)
    print(get_display(text))

def RemoveStrings(sequance):
    seq_type= type(sequance)
    return seq_type().join(filter(seq_type.isdigit, sequance))

def RemoveNumbers(word):
    regex = re.compile('[^a-zA-Z]')
    return regex.sub('', word)



#Import Workbook
workbook = load_workbook('8th.xlsx')
sheet = workbook.active

nameCN = "B5"
classCN = "C5"

classCellValue = sheet[classCN].value

#get the first cell number
cellNumber = int(RemoveStrings(nameCN))-1
subjectsCN=["E5","G5","I5","K5","M5","O5","Q5"]

finshedWorkbook = Workbook()

def CalculateTotal(subjects):
    total=0
    for i in range(len(subjects)):
        subjects[i] = RemoveNumbers(subjects[i])
        subjects[i] = subjects[i] + str(cellNumber)
        total +=sheet[subjects[i]].value
    return total


def AddNewSheet(name):
    global finshedWorkbook
    sheetName = str(name).replace('/','-')
    print('sheetName: '+sheetName)
    newSheet = finshedWorkbook.create_sheet(sheetName)

    newSheet["A1"] = ''
    newSheet["B1"] = "الرحمن الرحيم!"

    return newSheet


def AnalyzeData():
  global nameCN
  global classCN
  global cellNumber
  global classCellValue
  
  #Save Workbook
  newSheet = AddNewSheet(sheet[classCN].value)

  print('---'+str(sheet[classCN].value)+'----')
  while(sheet[nameCN].value != None):
      nameCN = RemoveNumbers(nameCN)
      classCN = RemoveNumbers(classCN)

      #Incerment Excel
      cellNumber = int(cellNumber) + 1
      nameCN = nameCN + str(cellNumber)
      classCN = classCN + str(cellNumber)
      
      if(sheet[nameCN].value == None):return  

      #Add New Tab
      if(classCellValue != sheet[classCN].value): 
          classCellValue = sheet[classCN].value
          print('---'+str(sheet[classCN].value)+'----')

      #Save data to excel workbook
      newSheet["A1"] = str(sheet[nameCN].value)
      ArabicPrint("Name: " + str(sheet[nameCN].value))
      print("Total:"+ str(CalculateTotal(subjectsCN)))


  finshedWorkbook.save(filename="Analysis.xlsx")


       
AnalyzeData()
