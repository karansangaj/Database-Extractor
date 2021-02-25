import pandas
import pyodbc
import pathlib
import ctypes
import platform
import PySimpleGUI as sg
from tabulate import tabulate

def make_dpi_aware():
    if int(platform.release()) >= 8:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)
make_dpi_aware()

sg.LOOK_AND_FEEL_TABLE['MyNewTheme'] = {'BACKGROUND': '#3c3f41',
                                        'TEXT': '#afb1b3',
                                        'INPUT': '#2b2b2b',
                                        'TEXT_INPUT': '#afb1b3',
                                        'SCROLL': '#3c3f41',
                                        'BUTTON': ('#afb1b3', '#3c3f41'),
                                        'PROGRESS': ('#3c3f41', '#3c3f41'),
                                        'BORDER': 0, 'SLIDER_DEPTH': 1, 'PROGRESS_DEPTH': 0,
                                        }
sg.theme('MyNewTheme')

layout = [
    [sg.Text('Extract DB  |',size=(13, 1), font='Courier 12 bold', justification='left'),
     sg.FileBrowse("File",target=(2,1),size=(4, 1)),
     sg.Text('',size=(61, 1), font='Courier 11 bold', justification='left'),
     sg.Submit("X", size=(2, 1))
     ],
    [sg.Text('---------------------------------------------------------------------------------------------------------------------------------', size=(70, 1), font='Broadway 11 bold')],
    [sg.Text('DataBase', size=(8, 1), font='SegoeUI 10 bold', justification='left'),sg.InputText(size=(55,1)),sg.Submit("Extract Contents", font='SegoeUI 10 bold')],
    [sg.Text('Table No', size=(8, 1), font='SegoeUI 10 bold', justification='left'), sg.InputText(size=(46,1)), sg.Submit("Extract Table", font='SegoeUI 10 bold'),sg.Text('|', size=(1, 1), font='SegoeUI 11 bold'), sg.Submit("Save File", font='SegoeUI 10 bold')],
    [sg.Text('')],
    [sg.Output(size=(80,18), key = '_output_', font='Courier  11 bold')],
    [sg.Text(" ")],
    [sg.Text('Developed by KARAN SANGAJ  ', font='SegoeUI 10 bold',size=(91, 2), justification='Right')],
]

window = sg.Window('Extract DB', layout, no_titlebar=True, grab_anywhere=True, keep_on_top = True)

while True:

    event, values = window.read()

    if event == 'X':
        window.close()

    if event == 'Extract Contents':
        
       window.FindElement('_output_').Update('')
       file = pathlib.Path("C:\\ProgramData\\regid.1991-06.com.microsoft\\regid.1991-06.com.microsoft Microsoft Access database engine 2016 (English).swidtag")

       if not file.exists():
           print(" ")
           print(" Microsoft Access Database Engine not Installed")
           print(" ")
           print(" Download Package from")
           print(" https://www.microsoft.com/en-us/download/details.aspx?id=54920")

       else:

           X = str(values[0])
           X1 = X.split('/')
           X2 = str(X1[-1])

           if not values[0]:
               print(" ")
               print(" Please Add File")

           elif (".accdb" in X2) == False:
               print(" ")
               print("File not supported")
               print(" ")
               print("Supported File Type : .aacdb")

           else:
               X=str(values[0])
               X1=X.split('/')
               X2=str(X1[-1])
               conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s'%X)
               cursor=conn.cursor()
               tableNames = [x[2] for x in cursor.tables().fetchall() if x[3] == 'TABLE']
               Data = pandas.DataFrame(tableNames,columns=['Extracted Tables'])
               tabulate.PRESERVE_WHITESPACE = True
               print(" ")
               print(tabulate(Data, tablefmt="pretty", headers="keys", showindex= True))

    if event == 'Extract Table':

        window.FindElement('_output_').Update('')
        file = pathlib.Path(
            "C:\\ProgramData\\regid.1991-06.com.microsoft\\regid.1991-06.com.microsoft Microsoft Access database engine 2016 (English).swidtag")

        if not file.exists():
            print(" ")
            print(" Microsoft Access Database Engine not Installed")
            print(" ")
            print(" Download Package from")
            print(" https://www.microsoft.com/en-us/download/details.aspx?id=54920")

        else:
            if not values[0]:
                print(" ")
                print(" Please Add File")

            elif not values[1]:
                print(" ")
                print(" Please Enter Table no")

            elif int(values[1]) >= len(Data.index):
                s1 = len(Data.index) - 1
                print(" ")
                print(" Only %s Tables Found in DataBase" % s1)

            else:
                try:
                    X = str(values[0])
                    X1 = X.split('/')
                    X2 = str(X1[-1])
                    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s' % X)
                    cursor = conn.cursor()
                    tableNames = [x[2] for x in cursor.tables().fetchall() if x[3] == 'TABLE']
                    Data = pandas.DataFrame(tableNames, columns=['Extracted Tables'])
                    X3 = int(values[1])
                    X4 = str(Data.iat[X3, 0])
                    SQLCommand = ("SELECT * FROM [%s].[%s]" % (X, X4))
                    data = pandas.read_sql(SQLCommand, conn)
                    tabulate.PRESERVE_WHITESPACE = True
                    print(" ")
                    print(" " + X4 + "  Details")
                    print(" ")
                    print(tabulate(data, tablefmt="pretty", headers="keys", showindex=True, colalign=("right",)))

                except:
                    print(" ")
                    print(" File Empty")

    if event == 'Save File':

        window.FindElement('_output_').Update('')

        if not values[0]:
            print(" ")
            print(" Please Add File")

        else:
            if not values[1]:
                print(" ")
                print(" Please Enter Table no")

            else:
                if int(values[1]) >= len(Data.index):
                    s1 = len(Data.index) - 1
                    print(" ")
                    print(" Only %s Tables Found in DataBase" % s1)

                else:
                    X = str(values[0])
                    X1 = X.split('/')
                    X2 = str(X1[-1])
                    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s' % X)
                    cursor = conn.cursor()
                    tableNames = [x[2] for x in cursor.tables().fetchall() if x[3] == 'TABLE']
                    Data = pandas.DataFrame(tableNames, columns=['Extracted Tables'])
                    X3 = int(values[1])
                    X4 = str(Data.iat[X3, 0])
                    SQLCommand = ("SELECT * FROM [%s].[%s]" % (X, X4))
                    data = pandas.read_sql(SQLCommand, conn)
                    tabulate.PRESERVE_WHITESPACE = True
                    outputFile = "%s.xlsx"%X3
                    data.to_excel(outputFile)
                    print(" ")
                    print(" File Saved as %s.xlsx to Application Folder"%X3)


window.close()