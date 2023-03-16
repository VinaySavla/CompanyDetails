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



def get_company_details(cin):
    res = requests.get(f"https://www.falconebiz.com/company/{cin}")
    soup = BeautifulSoup(res.text, "lxml")
    tables = soup.find_all("table")

    #info, contact, directors, capital = tables
    info = tables[0]
    contact = tables[-3]
    directors = tables[-2]
    capital = tables[-1]

    result = {}
    for table in (info, contact):
        for row in table.findAll('tr'):
            aux = row.findAll('td')
            result[aux[0].string] = str(aux[1].string).strip()

    result["directors"] = [
        {
            cell["data-label"]: cell.string
            for cell in record.findAll('td')
        }
        for record in directors.find("tbody").findAll("tr")
    ]

    for row in capital.findAll('tr'):
        aux = list(row.find('td').stripped_strings)
        result[aux[0]] = ' '.join(aux[1:])

    return result
#end of Function

exl_file_path=sys.argv[1]
def exl_path():
    # exl_file_path=exl_file_path
    # print("Exl path Button")
    # files, _ = QFileDialog.getOpenFileName(None, "Open File", "", "PDF File (*.xlsx)")
    # self.exl_file_path = str(files)
    # value = self.comboBox.currentText()
    # print(value)

    print(exl_file_path)
    
    # if self.exl_file_path == "":
    #     print("No excel file selected.")
    #     return

    # self.comboBox.clear()
    data = pd.read_excel(exl_file_path)
    
    # all_columns = data.columns
    # self.comboBox.setCurrentText("")
    # self.lineEdit.setCurrentText
    # for val in all_columns:
        # self.comboBox.addItem(str(val))
    
    # self.lineEdit.setText(QFileDialog.getOpenFileName(None, "Open File", "Desktop", "Excel Workshee (*.xlsx)"))
    if exl_file_path !=None: 
        print(exl_file_path)
        data = pd.read_excel(exl_file_path)
        col_value=[]
        for name in data:
            col_value.append(name)
            print(name)
        cin=""
        for i in col_value:
            if i=='Cin' or i=='cin' or i=='CIN':
                cin=i
                # print(cin+"##"")
    else:
        print("Please Select Excel file first.")
        return
#End of Exl Path






cd = get_company_details("U55101DL2023PTC410401")
exl_path()
print(cd["Registration Number"])
# print(
#     json.dumps(
#         get_company_details("U55101DL2023PTC410401"),
#         indent=2,
#     )
# )

