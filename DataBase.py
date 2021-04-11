import sqlite3
from tkinter import messagebox

connection = sqlite3.connect('StudentManagement.db')
cursor = connection.cursor()

table_name = "Subjects"
subject_name = "SubjectName"

def GetDataFromTable(tableName):
    cursor.execute("SELECT * FROM "+tableName)
    return cursor.fetchall()


def InsertData(subjectName):
    subjects = GetDataFromTable('Subjects')
    for subject in subjects:
        if subject[1] == subjectName:
            messagebox.showerror('Error','اسم المادة تم إدخاله من قبل')
            return
            

    connection.execute("INSERT INTO " + table_name + " ( " + subject_name + ") VALUES ('"+ subjectName +"')")
    connection.commit()
    return True


def DeleteData(subject):
    connection.execute("DELETE FROM Subjects Where SubjectName= '" + subject[1] + "';")
    connection.commit()
    return True
