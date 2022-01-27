from openpyxl import load_workbook, Workbook
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.writer.excel import save_workbook
import os
import subprocess


POSITION = {
    "AM" : [
        {
            "RANGE" : list(range(76, 136)) + list(range(137, 178)),
            "CD_TEXT" : "A",
            "NAME" : "B",
            "AMOUNT" : "C",
            "DETAIL" : "D"
        },
        {
            "RANGE" : list(range(76, 136)) + list(range(137, 178)),
            "CD_TEXT" : "E",
            "NAME" : "F",
            "AMOUNT" : "G",
            "DETAIL" : "H"
        }
    ],
    "PM" : [
         {
            "RANGE" : list(range(181, 232)),
            "CD_TEXT" : "A",
            "NAME" : "B",
            "AMOUNT" : "C",
            "DETAIL" : "D"
        },
        {
            "RANGE" : list(range(181, 232)),
            "CD_TEXT" : "E",
            "NAME" : "F",
            "AMOUNT" : "G",
            "DETAIL" : "H"
        }
    ]
}


def write(template, output_folder, accounting_date, data):
    output_file_path = os.path.join(output_folder, accounting_date) + ".xlsx"
    if os.path.isfile(output_file_path):
        os.remove(output_file_path)

    wb = load_workbook(template)
    wb.security =  WorkbookProtection(workbookPassword='6933')

    sheet = wb["수입지출결의서_양식"]
    sheet.protection.set_password('6933')
    position = POSITION["AM"]
    for i in range(0, len(data)):
        line = position[i%2]["RANGE"][int(i/2)]
        cd_text = position[i%2]["CD_TEXT"] + str(line)
        name = position[i%2]["NAME"] + str(line)
        amount = position[i%2]["AMOUNT"] + str(line)
        detail = position[i%2]["DETAIL"] + str(line)
        
        sheet[cd_text] = data[i][0]
        sheet[name] = data[i][1]
        sheet[amount] = data[i][2]
        sheet[detail] = data[i][3]
    
    save_workbook(wb, output_file_path)
    
    wb.close()
