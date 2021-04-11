from InformationScreen import titlesTxtBoxs
from ImportData import GetDataToAnalyse
from tkinter import *
from tkinter import messagebox
import DataBase
from AppConfiguration import ClearTab,GetOriginalWorkbook, tab2_AddInformation, tab3_Subjects, colors, notebook, colors, GET_SUBJECTS_FROM_DB


subjectsTxtBoxs = []
subjectsCN = []

fontSettings = ("Times New Roman", 16, 'bold')
txtWidth = 10
rowNumber = 1
subjectsCells = ["E5", "G5", "I5", "K5", "M5", "O5", "Q5","R5","H5","C5","F5","T5","U5"]


def DisplaySubjectsScreen():
    global subjectsTxtBoxs
    ClearTab(tab3_Subjects)
    subjectsTxtBoxs.clear()
    subjectsCN.clear()
    global rowNumber
    label1 = Label(tab3_Subjects, text="-- المواد الدراسية --", background=colors.bg, foreground=colors.bg2,
                    padx=80, pady=15, relief=RIDGE,  font=("Times New Roman", 30, 'bold'))
    label1.grid(row=0, column=0, pady=15, columnspan=10, padx=150)

    # label Title
    label1 = Label(tab3_Subjects, text=" هذه قيم لمواد افتراضية يمكنك حذف أو إضافة أي مواد تريدها",
                   background=colors.bg, pady=15, font=fontSettings, wraplength=400)
    label1.grid(row=1, column=0, pady=15, columnspan=10, padx=150)

    rowNumber += 2
    btnAddNewSubject = Button(
        tab3_Subjects, text="+ إضافة مادة ", font=fontSettings, bg=colors.bg2,
        foreground="white", border=2, cursor="hand2",
        command=lambda: ShowNewSubjectWindow())
    btnAddNewSubject.grid(row=rowNumber, column=1, pady=10, padx=6)

    label2 = Label(tab3_Subjects, text="رقم الخلية", background=colors.bg, foreground=colors.txtColor, pady=20,
                   font=fontSettings)
    label2.grid(row=rowNumber, column=2, sticky='E')

    label3 = Label(tab3_Subjects, text="إسم المادة", background=colors.bg, foreground=colors.txtColor, pady=20,
                   font=fontSettings)
    label3.grid(row=rowNumber, column=3, sticky='E')

    DisplaySubjectsTxtboxs()



def DisplaySubjectsTxtboxs():
    global subjectRow
    subjectRowNumber = rowNumber+1

    #to be deleted
    
    i = 0
    for subject in GET_SUBJECTS_FROM_DB():
        label = Label(tab3_Subjects, text=subject[1], background=colors.bg, foreground=colors.txtColor,
                      font=fontSettings)
        label.grid(row=subjectRowNumber, column=3, sticky='E')

        subjectNameTxtBox = Entry(
            tab3_Subjects, font=fontSettings, width=txtWidth)
        subjectNameTxtBox.grid(row=subjectRowNumber, column=2)

        subjectNameTxtBox.insert(0, subjectsCells[i])
        i += 1
        subjectsTxtBoxs.append(subjectNameTxtBox)

        btnRemoveFromSubjects = Button(tab3_Subjects, text="X", bg=colors.bg2,
                                       foreground="white", border=0, cursor="hand2", font=("Times New Roman", 15),
                                       command=lambda subject=subject: RemoveFromSubjects(subject))

        btnRemoveFromSubjects.grid(row=subjectRowNumber, column=4, pady=5, padx=10)
        subjectRowNumber += 1

    btnExport = Button(tab3_Subjects, text="إنشاء التحليل", font=fontSettings, bg=colors.bg2,
                       foreground="white", border=2, cursor="hand2",
                       command=lambda: BtnAnalyseData())
    btnExport.grid(row=subjectRowNumber, column=1, pady=6, padx=6)

    label = Label(tab3_Subjects, text='---------------------------------', background=colors.bg, foreground=colors.txtColor,
                  font=fontSettings)
    label.grid(row=subjectRowNumber+10, column=3, rowspan=10)


def ShowNewSubjectWindow():
    global newSubjectWindow
    newSubjectWindow = Tk()
    newSubjectWindow.title("Add New Subject")
    newSubjectWindow.geometry("400x300")
    newSubjectWindow.attributes('-topmost', 'true')

    label = Label(newSubjectWindow, text="إسم المادة ", foreground=colors.txtColor,
                  font=("Times New Roman", 20, 'bold'))
    label.grid(row=0, column=1, sticky='E', padx=10, pady=10)

    subjectNameTxtBox = Entry(newSubjectWindow, font=(
        "Times New Roman", 20, 'bold'), width=25)
    subjectNameTxtBox.grid(row=1, column=1, padx=10, pady=20)
    subjectNameTxtBox.focus()

    btnAddNewSubject = Button(
        newSubjectWindow, text="إضافة", font=fontSettings, bg=colors.bg2,
        foreground="white", border=1, cursor="hand2",
        command=lambda: AddNewSubject(subjectNameTxtBox.get()))
    btnAddNewSubject.grid(row=2, column=1, padx=10, pady=10)


def AddNewSubject(newSubjectName):
    if(newSubjectName == ''):
        messagebox.showerror('Error', 'لابد من إدخال إسم المادة')
        return

    if DataBase.InsertData(newSubjectName):
        newSubjectWindow.destroy()
        messagebox.showinfo("Saved", "تم إضافة المادة بنجاح")
        DisplaySubjectsScreen()
    else:
        messagebox.showerror("Error", "لم تتم إضافة المادة")


def RemoveFromSubjects(subject):
    msgBox = messagebox.askquestion(
        'Delete Subject', 'هل أنت متأكد من حذف هذه المادة؟')
    if msgBox == 'yes':
        if DataBase.DeleteData(subject):
            messagebox.showinfo("Deleted", "تم حذف المادة بنجاح")
            DisplaySubjectsScreen()
            
        else:
            messagebox.showerror("Error", "لم يتم حذف المادة")


def BtnAnalyseData():
    global subjectsCN
    originalWorkbook = GetOriginalWorkbook()
    if(CheckData(titlesTxtBoxs) == False):
        notebook.select(tab2_AddInformation)
        return
    if(CheckData(subjectsTxtBoxs) == False):
        notebook.select(tab3_Subjects)
        return

    ministryName = titlesTxtBoxs[0].get()
    schoolName = titlesTxtBoxs[1].get()
    sheetTitle = titlesTxtBoxs[2].get()
    studentNameCN = titlesTxtBoxs[3].get()
    studentClassCN = titlesTxtBoxs[4].get()
    bestStudentsMax = int(titlesTxtBoxs[5].get())
    subjectsFullMark = int(titlesTxtBoxs[6].get())

    
    for txtBox in subjectsTxtBoxs:
        subjectsCN.append(txtBox.get())

    try:
        GetDataToAnalyse(originalWorkbook, ministryName, schoolName, sheetTitle,
                         studentNameCN, studentClassCN, bestStudentsMax, subjectsCN, subjectsFullMark)
    except AssertionError as error:
        messagebox.showerror('Error', error)




def CheckData(dataToCheck):
    for txtBox in dataToCheck:
        txtBox.config({"background": 'white'})
        if(txtBox.get() == None or txtBox.get() == ''):
            messagebox.showerror('Error', 'برجاء إدخال جميع البيانات')
            txtBox.focus()
            txtBox.config({"background": 'red'})
            return False



