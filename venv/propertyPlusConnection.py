import pyodbc
import xlsxwriter
conn = pyodbc.connect("DSN=Seasonal;Description=Seasonal;UID=PropertyPlus;PWD=propertyplus;APP=Microsoft Office 2003;WSID=DON;Database=S1372;")

cursor = conn.cursor()
cursor.execute("select * from unit")
# Selection.QueryTable.Connection = "ODBC;DSN=Seasonal;Description=Seasonal;UID=PropertyPlus;PWD=propertyplus;APP=Microsoft Office 2003;WSID=DON;Database=S1372;"
row = cursor.fetchone()
cols = cursor.columns()
tables = cursor.tables()

for row in tables:
    print (row.table_name)

if row:
    print(str(row)+"\n")
