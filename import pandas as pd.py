import pandas as pd
import mysql.connector
import os

# --- SETTINGS ---
excel_file = "your_file.xlsm"        # Excel file path
sheet_name = "Sheet1"                # Change if needed
database = "carbon_emissions"        # Existing database

# Use filename (without extension) as table name
table_name = os.path.splitext(os.path.basename(excel_file))[0]

# --- STEP 1: Read the Excel sheet ---
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# --- STEP 2: Connect to MySQL ---
conn = mysql.connector.connect(
    host="localhost",         # or your server IP
    user="root",         # your MySQL username
    password="Nor@eb@ng99", # your MySQL password
    database=database         # connect directly to carbon_emissions
)
cursor = conn.cursor()

# --- STEP 3: Create table dynamically ---
# Add an auto-increment primary key + one column per Excel header
columns = ", ".join([f"`{col}` TEXT" for col in df.columns])
create_table_query = f"""
CREATE TABLE IF NOT EXISTS `{table_name}` (
    id INT AUTO_INCREMENT PRIMARY KEY,
    {columns}
)
"""
cursor.execute(create_table_query)

# --- STEP 4: Insert the Excel data ---
placeholders = ", ".join(["%s"] * len(df.columns))
insert_query = f"INSERT INTO `{table_name}` ({', '.join(df.columns)}) VALUES ({placeholders})"

for _, row in df.iterrows():
    cursor.execute(insert_query, tuple(row))

conn.commit()
conn.close()

print(f"✅ Data from {excel_file} → table `{table_name}` in DB `{database}`")
