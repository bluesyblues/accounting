
import src.db_connection as db_connection
import src.write_xlsx as write_xlsx
import configparser

def main():
    config = configparser.ConfigParser()
    config.read("config.ini")
    config = config["DEFAULT"]
    accounting_date = input()
    data = db_connection.get_raw_data(config["DATABASE"], accounting_date)
    write_xlsx.write(config["TEMPLATE"], config["OUTPUT_FOLDER"], accounting_date, data)

main()