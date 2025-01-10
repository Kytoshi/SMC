import pandas as pd
import pyodbc

# Read data from Excel file
excel_file = 'C:\\Path\\To\\YourFile.xlsx'  # Update with your Excel file path
df = pd.read_excel(excel_file)

# Connect to SQL Server
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                      'SERVER=.\SQLEXPRESS;'  # Update with your server name
                      'DATABASE=YourDatabase;'  # Update with your database name
                      'Trusted_Connection=yes;')

# Insert data into SQL table
for index, row in df.iterrows():
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO YourTable (Column1, Column2, Column3)
        VALUES (?, ?, ?)
    """, row['Column1'], row['Column2'], row['Column3'])  # Update column names
    conn.commit()

conn.close()
print("Data imported successfully!")
