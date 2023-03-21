import json
import requests
from bs4 import BeautifulSoup
from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtWidgets import QFileDialog ,QLabel
from PyQt6.QtWidgets import QMessageBox
from importlib.resources import path
from sys import path_hooks
import sys
import pandas as pd
import time
from importlib.resources import path
from sys import path_hooks

class Ui_MainWindow(object):
    def setupUi(self, MainWindow,btn_val,gb_val,title,all_text_color,Start_button_text_color,wait_time,main_icon):
        self.all_text_color_val=all_text_color
        self.all_text_color="font: 75 13pt \"MS Shell Dlg 2\";color:rgb%s;"%(self.all_text_color_val)
        self.all_text_color_browse="color:rgb%s;"%(self.all_text_color_val)
        self.Start_button_text_color = Start_button_text_color
        self.wait_time = wait_time
        self.background_color_value = gb_val
        self.background_color = "background-color: rgb%s;"%(self.background_color_value)
        self.button_color_value = btn_val
        self.button_color = "background-color: rgb%s;color: rgb%s"%(self.button_color_value,self.Start_button_text_color)
        
        # MainWindow.resize(562, 600)
        MainWindow.setGeometry(500,100,500,500)
        MainWindow.setWindowTitle(title)
        MainWindow.setWindowIcon(QtGui.QIcon(main_icon))
        self.centralwidget = QtWidgets.QWidget(MainWindow)

        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        
        font = QtGui.QFont()
        font.setPointSize(14)
        self.frame.setFont(font)
        self.frame.setStyleSheet(self.background_color)
        self.frame.setFrameShape(QtWidgets.QFrame.Shape.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Shadow.Raised)
        self.frame.setObjectName("frame")

        #Send Button
        self.pushButton = QtWidgets.QPushButton(self.frame)
        self.pushButton.setGeometry(QtCore.QRect(185, 350, 120, 50))
        self.pushButton.setStyleSheet(self.button_color)
        self.pushButton.setObjectName("pushButton")

        #Excel Button
        self.toolButton = QtWidgets.QToolButton(self.frame)
        self.toolButton.setGeometry(QtCore.QRect(330, 200, 61, 31))
        self.toolButton.setObjectName("toolButton")
        self.toolButton.clicked.connect(self.exl_path)
        self.toolButton.setStyleSheet(self.all_text_color_browse)

        #Excel Lable
        self.label = QtWidgets.QLabel(self.frame)
        self.label.setGeometry(QtCore.QRect(80, 200, 141, 31))
        font = QtGui.QFont()
        font.setFamily("MS Shell Dlg 2")
        font.setPointSize(13)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(9)
        self.label.setFont(font)
        self.label.setStyleSheet(self.all_text_color)
        self.label.setObjectName("label")
        
        #Title
        self.label_4 = QtWidgets.QLabel(self.frame)
        self.label_4.setStyleSheet(self.all_text_color)
        self.label_4.setText(title)
        self.label_4.setGeometry(QtCore.QRect(150, 20, 241, 61))
        font = QtGui.QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")

        #Settings
        self.pushButton_2 = QtWidgets.QPushButton(self.frame)
        self.pushButton_2.setGeometry(QtCore.QRect(5, 440, 30, 30))
        self.pushButton_2.setText("")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("settings.png"), QtGui.QIcon.Mode.Normal, QtGui.QIcon.State.Off)
        self.pushButton_2.setIcon(icon)
        self.pushButton_2.setIconSize(QtCore.QSize(30, 30))
        self.pushButton_2.setCheckable(False)
        self.pushButton_2.setChecked(False)
        self.pushButton_2.setAutoDefault(False)
        self.pushButton_2.setDefault(False)
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.ChangeSettings)
        
        #Logo
        self.label_8 = QtWidgets.QLabel(self.frame)
        self.label_8.setGeometry(QtCore.QRect(50, 20, 91, 61))
        self.label_8.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.SizeBDiagCursor))
        self.label_8.setText("")
        self.label_8.setPixmap(QtGui.QPixmap(main_icon))
        self.label_8.setScaledContents(True)
        self.label_8.setObjectName("label_8")

        self.verticalLayout.addWidget(self.frame)
        MainWindow.setCentralWidget(self.centralwidget)
        

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        self.exl_file_path = None
       

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        
        
        self.label.setText(_translate("MainWindow", "Select Excel Sheet"))
        self.toolButton.setText(_translate("MainWindow", "Browse"))

        self.pushButton.setText(_translate("MainWindow", "Start"))
        self.pushButton.clicked.connect(self.Start_button)
        

    def ChangeSettings(self):
        print("Changing Setting")
        # osCommandString = "notepad.exe Config_file.txt"
        # os.system(osCommandString)

        # os.startfile('Config_file.txt')
        import subprocess
        import platform as pf
        if sys.platform == "win32":
            subprocess.call(['notepad.exe', 'Config_file.txt'])
        elif sys.platform == "darwin":
            subprocess.call(['open', '-a', 'TextEdit', 'Config_file.txt'])
        else:
            self.Pop_up_message("System Not Suported")



    def exl_path(self):
        print("Exl path Button")
        files, _ = QFileDialog.getOpenFileName(None, "Open File", "", "Excel File (*.xlsx)")
        self.exl_file_path = str(files)
        print(self.exl_file_path)
        
        if self.exl_file_path == "":
            print("No excel file selected.")
            return
        
        # data = pd.read_excel(self.exl_file_path)
        
        # all_columns = data.columns

        # self.lineEdit.setText(QFileDialog.getOpenFileName(None, "Open File", "Desktop", "Excel Workshee (*.xlsx)"))

    def Pop_up_message(self,msg_text,icon="Warning"):
        msg = QMessageBox()
        msg.setWindowTitle("Alert!")
        if(icon=="Success"):
            msg.setWindowIcon(QtGui.QIcon('check.png'))
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setIconPixmap(QtGui.QPixmap('check.png').scaled(20, 20))
        else:
            msg.setWindowIcon(QtGui.QIcon('warning.png'))
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setIconPixmap(QtGui.QPixmap('warning.png').scaled(20, 20))
        msg.setText(msg_text)
        x = msg.exec()  # this will show our messagebox
        
    def getDetails(self):
        comapny_details = []
        def get_company_details(cin):
            res = requests.get(f"https://www.falconebiz.com/company/{cin}")
            soup = BeautifulSoup(res.text, "lxml")
            tables = soup.find_all("table")

            # TODO Comapany Name in Disctionary
            # c_name = soup.get('li',attrs={'class':'breadcrumb-item active'})
            c_name = soup.find(class_='breadcrumb-item active')
            c_name=c_name.get_text()
            # print(c_name)

            #info, contact, directors, capital = tables
            info = tables[0]
            contact = tables[-3]
            directors = tables[-2]
            capital = tables[-1]

            # result = {}
            result = {'Comapany Name' : c_name}
            # print(result)
            for table in (info, contact):
                for row in table.findAll('tr'):
                    aux = row.findAll('td')
                    # print(str(aux[1].getText()))
                    result[aux[0].string] = str(aux[1].getText())
                    # result[aux[0].string] = str(aux[1].string).strip()

            result["directors"] = [
                {
                    # cell["data-label"]: cell.string
                    cell["data-label"]: cell.getText()
                    for cell in record.findAll('td')
                }
                for record in directors.find("tbody").findAll("tr")
            ]

            for row in capital.findAll('tr'):
                aux = list(row.find('td').stripped_strings)
                result[aux[0]] = ' '.join(aux[1:])

            return result
        #end of Function

        exl_file_path=self.exl_file_path


        def exl_read_write():
            wait_time = self.wait_time
            data = pd.read_excel(exl_file_path)
            if exl_file_path !=None: 
                data = pd.read_excel(exl_file_path)
                col_value=[]
                for name in data:
                    col_value.append(name)
                    # print(name)
                cin=""
                for i in col_value:
                    if i=='Cin' or i=='cin' or i=='CIN':
                        cin=i
                        # print(cin+"##"")
            else:
                print("Please Select Excel file first.")
                return
            # getting the names and the cins
            l1 = []
            for index, row in data.iterrows():
                l1.append(row.to_list())

            cins = data[cin]
            #here
            for i in range(len(cins)):
                try:
                    # for every record get the name and the cin addresses
                    l2=[]
                    for j in l1[i]:
                        j=str(j)
                        if 'nan' in j or 'NaN' in j:
                            j = "unknown"
                        l2.append(j)    

                    e = cins[i]
                    cin=str(e)
                    if 'nan' in cin or 'NaN' in cin:
                        cin = "unknown"
                    print(cin)

                    cd = get_company_details(cin)
                    for i, director in enumerate(cd["directors"]):
                        for key, value in director.items():
                            cd[key + "_" + str(i+1)] = value
                    del cd["directors"]
                    # print(
                    #     json.dumps(
                    #         cd,
                    #         indent=2,
                    #     )
                    # )
                    # print(col_value)
                    # print(df)
                    comapny_details.append(cd)  
                    print("waiting "+ str(wait_time) +" seconds")
                    time.sleep(int(wait_time))
                    print("************************ cycle completed ************************")

                except Exception as e:
                    print(e)    
            
        #End of Exl Path
        exl_read_write()
        # print(
        #     json.dumps(
        #         comapny_details,
        #         indent=2,
        #     ))
        df = pd.DataFrame(comapny_details)
        newFileName=exl_file_path.rsplit('.')[0]+" Result.xlsx"
        # print(newFileName)
        df.to_excel(newFileName)
        # cd = get_company_details("U55101DL2023PTC410401")
        # print(cd["Registration Number"])
        # print(
        #     json.dumps(
        #         get_company_details("U55101DL2023PTC410401"),
        #         indent=2,
        #     )
        # )

    def Start_button(self):

        if self.exl_file_path==None:
            self.Pop_up_message("Please Select Excel File")
            return
        
        self.getDetails()

        #all variables reset
        self.exl_file_path = None


        print("Process completed.")
        self.Pop_up_message("Excel with all Details Generated Successfully!","Success")

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    with open('Config_file.txt') as f:
        lines = f.readlines()
    
    title=lines[0].strip()
    title = title.split("=")[1]

    bg_color_set=lines[1].strip()
    bg_color_set = bg_color_set.split("=")[1]

    btn_color_set =  lines[2].strip()
    btn_color_set = btn_color_set.split("=")[1]

    all_text_color=lines[3].strip()
    all_text_color=all_text_color.split("=")[1]

    Start_button_text_color=lines[4].strip()
    Start_button_text_color=Start_button_text_color.split("=")[1]

    wait_time_in_sec=lines[5].strip()
    wait_time_in_sec=wait_time_in_sec.split("=")[1]

    main_icon=lines[6].strip()
    main_icon=main_icon.split("=")[1]

    myLabel= QLabel()
    myLabel.setAutoFillBackground(True) # This is important!!
    

    ui.setupUi(MainWindow,btn_color_set,bg_color_set,title,all_text_color,Start_button_text_color,wait_time_in_sec,main_icon)
    MainWindow.show()
    sys.exit(app.exec())
