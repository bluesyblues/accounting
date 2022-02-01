
import src.db_connection as db_connection
import src.write_xlsx as write_xlsx
import configparser
import datetime

def main():
    config = configparser.ConfigParser()
    config.read("config.ini")
    config = config["DEFAULT"]
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    accounting_date = input(f"날짜를 입력하세요. 예) {today}\n")
    try:
        datetime.datetime.strptime(accounting_date, "%Y-%m-%d")
    except:
        print("날짜를 올바른 형식으로 입력해주세요")
        return
    import_data = db_connection.get_import_data(config["DATABASE"], accounting_date)
    expense_cash_data = db_connection.get_expense_cash_data(config["DATABASE"], accounting_date)
    expense_bank_data = db_connection.get_expense_bank_data(config["DATABASE"], accounting_date)
    write_xlsx.write(config["TEMPLATE"], config["OUTPUT_FOLDER"], accounting_date, import_data, expense_cash_data, expense_bank_data)

main()