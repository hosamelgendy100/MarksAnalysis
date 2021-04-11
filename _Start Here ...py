from SearchScreen import DisplaySerachScreen
from InformationScreen import DisplayInformationScreen
from tkinter import *
from AppConfiguration import ROOT, tab2_AddInformation, colors, welcomeTab, notebook, DEFAULT_VALUES
from SubjectsScreen import DisplaySubjectsScreen


fontSettings = ("Times New Roman", 18, 'bold')


def WelcomeScreen():
    # label Title
    label1 = Label(welcomeTab, text="-- برنامج تحليل درجات الطلاب --", background=colors.bg, foreground=colors.bg2,
                   padx=180, pady=15, relief=RIDGE,  font=("Times New Roman", 30, 'bold'))
    label1.grid(row=0, column=0, pady=15, columnspan=10, padx=150)

    # Admin Button
    loginBtn = Button(welcomeTab, text='ابـــــدأ', font=fontSettings, cursor="hand2",
                      width=15, foreground='white', background=colors.bg2,
                      command=lambda: LoginBtn())
    loginBtn.grid(row=3, column=5, pady=15)

    lblWelcome = Label(welcomeTab, text=DEFAULT_VALUES[7][2],
                       background=colors.bg, pady=15, font=fontSettings, wraplength=700)
    lblWelcome.grid(row=4, column=5, pady=10, padx=25, sticky='E')

def LoginBtn():
    notebook.select(tab2_AddInformation)


notebook.tab(2, state="disabled")
WelcomeScreen()
DisplayInformationScreen()
DisplaySubjectsScreen()
DisplaySerachScreen()
ROOT.mainloop()
