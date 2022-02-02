import fdb

HOST = "127.0.0.1"
USER = "sysdba"
PASSWORD = "masterkey"

MAPPING_TABLE = {
    "영유아부헌금" : "영유아부",
    "유초등부헌금" : "유초등부",
    "중고등부헌금" : "중고등부",
    "청년부헌금" : "청년부",
    "주일헌금" : "주일",
    "감사헌금" : "감사",
    "십일조" : "십일조",
    "헌신예배헌금": "헌신예배",
    "건축헌금":"건축",
    "구제헌금":"구제",
    "선교헌금":"선교",
    "장학헌금":"장학",
    "목장헌금":"목장",
    "지정헌금":"지정",
    "특새헌금":"특새",
    "새벽기도회":"새벽기도회",
    "겟세마네":"겟세마네",
    "수요예배헌금":"수요예배",
    "심방헌금":"심방",
    "부활절":"부활절",
    "추수감사절":"추수감사절",
    "신년예배(송구영신)":"송구영신예배",
    "신년감사":"신년감사",
    "재직세미나":"재직세미나",
    "성미헌금":"성미",
    "구제헌금(지정)":"구제헌금(지정)",
    "선교헌금(지정)":"선교헌금(지정)",
    "금요예배헌금":"금요예배헌금",
    "성탄절":"성탄절"
}



def send_query(database, query_str):
    con = fdb.connect(
        host=HOST, 
        database=database, 
        user=USER, 
        password=PASSWORD,
        charset='UTF8'
    )
    cur = con.cursor()
    cur.execute(query_str)
    result = cur.fetchall()
    con.close()

    return result


def get_import_data(database, accounting_date):
    query_str = f"""
        SELECT
            CD.CD_TEXT,
            CM.CM_NAME,
            FI.FI_AMOUNT,
            FI.FI_COMM_GIFTER
        FROM
            (SELECT 
                FI_AMOUNT, 
                CASE WHEN FI_COMM_GIFTER = '통장' THEN '통장'
                ELSE FI_COMM_GIFTER END AS FI_COMM_GIFTER, 
                FI_IN_DATE, 
                FI_IMPORT_TYPE, 
                FI_DEBTOR_CM_SEQ,
                FI_ACCOUNT
            FROM FINANCE_IMPORT 
            WHERE FI_IN_DATE = '{accounting_date}') AS FI
        LEFT JOIN CHRIST_MEMBER AS CM
        ON FI.FI_DEBTOR_CM_SEQ = CM.CM_SEQ
        JOIN CHRIST_CODE AS CD
        ON FI.FI_ACCOUNT = CD.CD_SEQ
        ORDER BY CD.CD_TEXT
    """
    query_result = send_query(database, query_str)
    for i in range(0, len(query_result)):
        query_result[i] = list(query_result[i])
        simple_type = MAPPING_TABLE.get(query_result[i][0])
        if simple_type :
            query_result[i][0] = simple_type
        else:
            query_result[i][0] = "NULL"

    return query_result


def get_expense_cash_data(database, accounting_date):
    query_str = f"""
    SELECT
        FO.FO_OUT_DATE,
        CD.CD_TEXT,
        FO.FO_AMOUNT,
        FO.FO_COMMENT
    FROM (
        SELECT 
            FO_OUT_DATE,
            FO_ACCOUNT,
            FO_AMOUNT,
            FO_COMMENT
        FROM FINANCE_EXPENSE as fe 
        WHERE fo_out_date = '{accounting_date}' AND FO_COMMENT LIKE  '현금%'
    ) AS FO
    LEFT JOIN CHRIST_CODE AS CD
    ON FO.FO_ACCOUNT = CD.CD_SEQ
    """
    query_result = send_query(database, query_str)
    return query_result


def get_expense_bank_data(database, accounting_date):
    query_str = f"""
    SELECT
        FO.FO_OUT_DATE,
        CD.CD_TEXT,
        FO.FO_AMOUNT,
        FO.FO_COMMENT
    FROM (
        SELECT 
            FO_OUT_DATE,
            FO_ACCOUNT,
            FO_AMOUNT,
            FO_COMMENT
        FROM FINANCE_EXPENSE as fe 
        WHERE fo_out_date = '{accounting_date}' AND (NOT FO_COMMENT LIKE '현금%' OR FO_COMMENT IS NULL)
    ) AS FO
    LEFT JOIN CHRIST_CODE AS CD
    ON FO.FO_ACCOUNT = CD.CD_SEQ
    """
    query_result = send_query(database, query_str)
    return query_result

