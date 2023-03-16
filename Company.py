import json
import requests
from bs4 import BeautifulSoup

from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtWidgets import QFileDialog ,QLabel
from PyQt6.QtWidgets import QMessageBox

from importlib.resources import path
from sys import path_hooks



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


cd = get_company_details("U55101DL2023PTC410401")
print(cd["Registration Number"])
# print(
#     json.dumps(
#         get_company_details("U55101DL2023PTC410401"),
#         indent=2,
#     )
# )

