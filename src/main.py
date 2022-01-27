
import db_connection
import write_xlsx

DATABASE=r'C:\Users\blues\workspace\accounting\samples\sample_DB\FBUCMSDATA.FDB'
OUTPUT_FOLDER = r"C:\Users\blues\workspace\accounting\outputs"
TEMPLATE = r"C:\Users\blues\workspace\accounting\template\template.xlsx"
ACCOUNTING_DATE = "2022-01-02"

data = db_connection.get_raw_data(DATABASE, ACCOUNTING_DATE)

write_xlsx.write(TEMPLATE, OUTPUT_FOLDER, ACCOUNTING_DATE, data)

