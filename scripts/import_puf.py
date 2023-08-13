import pandas as pd
import pyodbc
from openpyxl import load_workbook
from configparser import ConfigParser

class CaseSensitiveConfigParser(ConfigParser):
    def optionxform(self, optionstr):
        return optionstr

# Read configuration from config.ini
config = CaseSensitiveConfigParser()
config.read('config.ini')

# Set up SQL Server connection
conn = pyodbc.connect(
    f'DRIVER=SQL Server;'
    f'SERVER={config["database"]["server"]};'
    f'DATABASE={config["database"]["database"]};'
    f'Trusted_Connection=yes;'
)

# Mapping of sheet names to database table names
table_mapping = dict(config['SheetToTableMapping'])
# Mapping of Excel columns to database columns
column_mapping = {}

for key in config['ColumnMapping']:
    column_mapping[key] = config.get('ColumnMapping', key)

# Mapping of Excel files to their respective sheets
excel_files = {
    'puf_2022': config['ExcelFiles']['puf_2022'],
    'puf_2021': config['ExcelFiles']['puf_2021'],
    'puf_2019': config['ExcelFiles']['puf_2019']
}

# Function to replace Yes/No values with bit
def map_yes_no_nan(value):
    if value in ['YES', 'Yes', 'yes', 'y']:
        return True
    else:
        return False

# Iterate over each Excel file and its sheets
for excel_key, excel_file_path in excel_files.items():
    excel_file = pd.ExcelFile(excel_file_path)
    for sheet_name in excel_file.sheet_names:
        # Get the Excel sheet as a DataFrame
        df = pd.read_excel(excel_file, sheet_name)
        
        # Get the corresponding database table name
        table_name = table_mapping.get(sheet_name, None)
        if table_name is None:
            print(f"Warning: No table mapping found for sheet '{sheet_name}'. Skipping...")
            continue

        # Extract the correct prefix from the column_mapping keys
        column_prefix = f'{sheet_name}_'
        
        # Filter column_mapping keys based on the prefix
        relevant_columns = {col_name: db_col for col_name, db_col in column_mapping.items() if col_name.startswith(column_prefix)}
        #print(relevant_columns)

        # Get data types from config for relevant columns
        relevant_db_column_names = [db_col for col_name, db_col in relevant_columns.items()]
        column_data_types = {db_col: config.get('ColumnDataTypes', db_col) for db_col in relevant_db_column_names}
        
        # Construct the CREATE TABLE query with specified data types
        create_table_columns = [f'{db_col} {data_type}' for db_col, data_type in column_data_types.items()]
        create_table_query = f"CREATE TABLE {table_name} ({', '.join(create_table_columns)})"
        
        with conn.cursor() as cursor:
            table_exists_query = f"IF OBJECT_ID('{table_name}', 'U') IS NULL SELECT 0 ELSE SELECT 1"
            cursor.execute(table_exists_query)
            table_exists = cursor.fetchone()[0]
            
            if table_exists:
                print(f"Table '{table_name}' already exists. Skipping CREATE TABLE statement.")
            else:
                cursor.execute(create_table_query)
                conn.commit()
            
        # Add the "measurement_yr" column
        measurement_yr = excel_key[-4:]  # Extract the last 4 numbers from the file name
        df['measurement_yr'] = measurement_yr

        # Replace columns values y/n with 1/0
        col_list = ['General-0014'
                    ,'General-0016'
                    ,'General-0085'
                    ,'General-0087'
                    ,'SA-0070']

        for col in col_list:
            if col in df.columns:
                df[col] = df[col].apply(map_yes_no_nan)
        
        # For specific sheets, add the "measure_abbr" column
        if sheet_name in ['hedishos_frm', 'hedishos_mui', 'hedishos_pao']:
            df['measure_abbr'] = sheet_name.split('_')[1]

        # Map Excel columns to database columns
        mapped_columns = [column_mapping.get(f'{sheet_name}_["{excel_col}"]', None) for excel_col in df.columns]
        
        # Rename DataFrame columns based on the column_mapping
        field_to_db_mapping = {excel_col: column_mapping.get(f'{sheet_name}_["{excel_col}"]', excel_col) for excel_col in df.columns}
        df.rename(columns=field_to_db_mapping, inplace=True)
        
        # Remove any columns that don't have a mapping
        mapped_columns = [col for col in mapped_columns if col is not None]
        print(" ")
        print(excel_file_path)
        print(table_name)
        print(df.dtypes)
        print(" ")

        # Insert data into the SQL table
        insert_query = f'INSERT INTO {table_name} ({", ".join(mapped_columns)}) VALUES ({", ".join(["?" for _ in range(len(mapped_columns))])})'

        # Execute data insertion for each row
        with conn.cursor() as cursor:
            for _, row in df.iterrows():
                values = [row[mapped_col] for mapped_col in mapped_columns]
                #print("Insert Query:", insert_query)
                #print("Mapped Columns:", mapped_columns)
                #print("Values:", values)
                cursor.execute(insert_query, values)
            conn.commit()

# Close the database connection
conn.close()
