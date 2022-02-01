from openpyxl import load_workbook
from openpyxl.writer.excel import save_workbook
import os

POSITION = {
    "AM" : [
        {
            "RANGE" : list(range(89, 146)) + list(range(148, 207)),
            "CD_TEXT" : "A",
            "NAME" : "B",
            "AMOUNT" : "C",
            "DETAIL" : "D"
        },
        {
            "RANGE" : list(range(89, 146)) + list(range(148, 207)),
            "CD_TEXT" : "E",
            "NAME" : "F",
            "AMOUNT" : "G",
            "DETAIL" : "H"
        }
    ],
    "PM" : [
         {
            "RANGE" : list(range(211, 268)),
            "CD_TEXT" : "A",
            "NAME" : "B",
            "AMOUNT" : "C",
            "DETAIL" : "D"
        },
        {
            "RANGE" : list(range(211, 268)),
            "CD_TEXT" : "E",
            "NAME" : "F",
            "AMOUNT" : "G",
            "DETAIL" : "H"
        }
    ],
    "EXPENSE_CASH": {
        "RANGE" : list(range(47, 59)),
        "FO_OUT_DATE" : "A",
        "CD_TEXT" : "B",
        "FO_AMOUNT" : "D",
        "FO_COMMENT" : "E"
    },
    "EXPENSE_BANK": {
        "RANGE" : list(range(65, 83)),
        "FO_OUT_DATE" : "A",
        "CD_TEXT" : "B",
        "FO_AMOUNT" : "D",
        "FO_COMMENT" : "E"
    },
}


def write(template, output_folder, accounting_date, data_import, data_expense_cash, data_expense_bank):
    output_file_path = os.path.join(output_folder, accounting_date) + ".xlsx"
    if os.path.isfile(output_file_path):
        os.remove(output_file_path)
    wb = load_workbook(template)
    sheet = wb["수입지출결의서_양식"]
    
    write_import(sheet, data_import)
    write_expense(sheet, data_expense_cash, "EXPENSE_CASH")
    write_expense(sheet, data_expense_bank, "EXPENSE_BANK")
    save_workbook(wb, output_file_path)
    
    wb.close()


def write_import(sheet, data):
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


def write_expense(sheet, data, expense_type):
    assert expense_type in ["EXPENSE_CASH", "EXPENSE_BANK"]
    position = POSITION[expense_type]
    for i in range(0, len(data)):
        line = position["RANGE"][i]
        fo_out_date = position["FO_OUT_DATE"] + str(line)
        cd_text = position["CD_TEXT"] + str(line)
        fo_amount = position["FO_AMOUNT"] + str(line)
        fo_comment = position["FO_COMMENT"] + str(line)

        sheet[fo_out_date] = data[i][0]
        sheet[cd_text] = data[i][1]
        sheet[fo_amount] = data[i][2]
        sheet[fo_comment] = data[i][3]

