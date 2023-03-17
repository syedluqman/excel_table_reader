import os
import pandas as pd
import sqlite3

# Set the directory path to search for Excel files
root_dir = "C:/Users/Username/Documents/"

# Connect to the SQLite database
conn = sqlite3.connect('example.db')

# Create a function to export a dataframe to a SQLite table
def export_to_sqlite(dataframe, table_name):
    dataframe.to_sql(table_name, conn, if_exists='replace')

# Loop through all files and directories within the root directory
for root, dirs, files in os.walk(root_dir):
    for file in files:
        # Check if the file is an Excel file
        if file.endswith(".xlsx") or file.endswith(".xls"):
            # If it is an Excel file, read all sheets in the file
            file_path = os.path.join(root, file)
            excel_data = pd.read_excel(file_path, sheet_name=None)
            # Loop through all the sheet names and export the dataframes to SQLite tables
            for sheet_name, sheet_data in excel_data.items():
                # Create a table name using the file name and sheet name
                table_name = f"{file[:-5]}_{sheet_name}"
                # Export the dataframe to the SQLite table
                export_to_sqlite(sheet_data, table_name)

# Commit changes and close the connection
conn.commit()
conn.close()
