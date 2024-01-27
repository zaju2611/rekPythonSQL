import psycopg2
import pyexcel as pe

def connect_to_database():
    dbname = input("Podaj nazwę bazy danych: ")
    user = input("Podaj nazwę użytkownika: ")
    password = input("Podaj hasło: ")
    port = input("Podaj numer portu: ")

    try:
        connection = psycopg2.connect(dbname=dbname, user=user, password=password, port=port)
        print("Połączono z bazą danych.")
        return connection
    except psycopg2.Error:
        print("Błąd połączenia z bazą. Spróbuj jeszcze raz...")
        return None

def execute_query(connection, query):
    cur = connection.cursor()
    cur.execute(query)
    records = cur.fetchall()
    cur.close()
    return records

def print_records(records):
    for record in records:
        print(record)

def save_records_to_xls(headers, records, dest_file_name):
    data = [headers] + records
    pe.save_as(array=data, dest_file_name=dest_file_name)
    print('Wyniki zapisane do pliku "', dest_file_name, '"')

def close_connection(connection):
    connection.commit()
    connection.close()

query = """
SELECT
    c.nazwa_klienta AS "NAME",
    DATE_PART('month', p.data_wplaty) AS "MONTH",
    SUM(p.kwota_wplaty) AS "SUM"
FROM
    clients c
JOIN
    debts d ON c.client_id = d.client_id1
JOIN
    payments p ON d.debt_id = p.debt_id1
GROUP BY
    c.nazwa_klienta, DATE_PART('month', p.data_wplaty)
ORDER BY
    c.nazwa_klienta;
"""

headers = ['Nazwa_klienta', 'Miesiac', 'Suma']

connection = connect_to_database()
records = execute_query(connection, query)
print_records(records)
save_records_to_xls(headers, records, 'results.xls')
close_connection(connection)