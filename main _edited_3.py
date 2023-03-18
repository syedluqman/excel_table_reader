import os
import pandas as pd
import sqlite3
import matplotlib.pyplot as plt

# Set the directory path to search for Excel files
root_dir = r"D:\Excel Code"

# Create a connection to the SQLite database
conn = sqlite3.connect('excel collated.db')

# Create a function to export a dataframe to a SQLite table
def export_to_sqlite(dataframe, table_name, conn):
    # Drop rows with all empty cells
    dataframe.dropna(how='all', inplace=True)
    # Drop columns with all empty cells
    dataframe.dropna(axis=1, how='all', inplace=True)
    # Export the dataframe to the SQLite table
    dataframe.to_sql(table_name, conn, if_exists='replace', index=False)

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
                # Check if sheet_data contains any data
                if not sheet_data.empty:
                    # Add day, month, and year columns for date columns
                    for col in sheet_data.columns:
                        if sheet_data[col].dtype == "datetime64[ns]":
                            sheet_data[col + "_day"] = sheet_data[col].dt.day
                            sheet_data[col + "_month"] = sheet_data[col].dt.month
                            sheet_data[col + "_year"] = sheet_data[col].dt.year
                    # Create a table name using the file name and sheet name
                    table_name = f"{file[:-5]}_{sheet_name}"
                    # Export the dataframe to the SQLite table
                    export_to_sqlite(sheet_data, table_name, conn)

# Ask the user for variables to search for
search_vars = []
while True:
    var = input("Enter a variable to search for (or type 'done' when finished): ")
    if var.lower() == 'done':
        break
    search_vars.append(var)

# Connect to the SQLite database
conn = sqlite3.connect('excel collated.db')

# Search for tables that contain the variables
tables_with_vars = []
cursor = conn.cursor()
for table_name in cursor.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall():
    table_name = table_name[0]
    table = pd.read_sql_query(f"SELECT * FROM [{table_name}]", conn)
    if any(var in table.columns for var in search_vars):
        tables_with_vars.append(table_name)

# Ask the user to select a table
if len(tables_with_vars) == 0:
    print("No tables found that contain the specified variables.")
else:
    print("Tables found that contain the specified variables:")
    for i, table_name in enumerate(tables_with_vars):
        print(f"{i+1}. {table_name}")
    table_index = int(input("Enter the number of the table you want to use: ")) - 1
    selected_table_name = tables_with_vars[table_index]

    # Get the selected table and variables
    selected_table = pd.read_sql_query(f"SELECT * FROM [{selected_table_name}]", conn)
    selected_vars = [var for var in search_vars if var in selected_table.columns]

    # Perform the cross-tabulation and plot the results
    if len(selected_vars) == 0:
        print("None of the specified variables were found in the selected table.")
    else:
        if any(selected_table[var].dtype == 'datetime64[ns]' for var in selected_vars):
            date_var = [var for var in selected_vars if selected_table[var].dtype == 'datetime64[ns]'][0]
            pivot_table = pd.pivot_table(selected_table, values=selected_vars, index=date_var, aggfunc=np.sum)
            pivot_table.plot(kind='bar')
            plt.show()
            print(pivot_table)
        else:
            print("One of the selected variables must be a date variable to plot the results.")