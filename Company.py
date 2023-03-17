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
    result = {'comapany Name' : c_name}
    # print(result)
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


def exl_read_write():
    wait_time = 1
    # exl_file_path=exl_file_path
    # print("Exl path Button")
    # files, _ = QFileDialog.getOpenFileName(None, "Open File", "", "PDF File (*.xlsx)")
    # self.exl_file_path = str(files)
    # value = self.comboBox.currentText()
    # print(value)

    # print(exl_file_path)
    

    data = pd.read_excel(exl_file_path)
    
    
    # self.lineEdit.setText(QFileDialog.getOpenFileName(None, "Open File", "Desktop", "Excel Workshee (*.xlsx)"))
    if exl_file_path !=None: 
        # print(exl_file_path)
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
    
    # user_name = ""
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
df.to_excel(exl_file_path+" Result.xlsx")
print("Process Completed")
# cd = get_company_details("U55101DL2023PTC410401")
# print(cd["Registration Number"])
# print(
#     json.dumps(
#         get_company_details("U55101DL2023PTC410401"),
#         indent=2,
#     )
# )

