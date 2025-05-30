# C:\Users\willy\Projects\bancarizacion\bancarizacion\core_logic.py
"""
Core logic for the Bancarizacion project.
This module will handle database interactions, data processing,
and Excel file generation.
"""
import configparser
import os
import mysql.connector
import openpyxl # For reading/writing .xlsx files

def get_db_config(config_file_path="db_config.ini"):
    """Reads database configuration from an INI file."""
    config = configparser.ConfigParser()
    # Ensure the path is relative to this file or an absolute path is provided
    # For now, let's assume config_file_path is relative to the project root
    # or an absolute path.
    # A more robust way would be to determine the project root dynamically.
    if not os.path.exists(config_file_path):
        # Try to find it relative to the project root if main.py is in root
        # This assumes core_logic.py is in a subdirectory 'bancarizacion'
        # and db_config.ini is in the parent directory (project root)
        proj_root_config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), config_file_path)
        if os.path.exists(proj_root_config_path):
            config_file_path = proj_root_config_path
        else:
            raise FileNotFoundError(f"Configuration file {config_file_path} not found.")
            
    config.read(config_file_path)
    if 'DATABASE' not in config:
        raise ValueError("DATABASE section not found in the configuration file.")
    return config['DATABASE']

def connect_to_db(db_config):
    """Connects to the MySQL database."""
    try:
        cnx = mysql.connector.connect(**db_config)
        print(f"Successfully connected to database: {db_config.get('database')}")
        return cnx
    except mysql.connector.Error as err:
        print(f"Error connecting to database: {err}")
        return None

def fetch_data_from_db(cnx, query):
    """Fetches data from the database using a SELECT query."""
    cursor = None
    try:
        cursor = cnx.cursor(dictionary=True) # dictionary=True to get results as dicts
        cursor.execute(query)
        results = cursor.fetchall()
        print(f"Fetched {len(results)} rows.")
        return results
    except mysql.connector.Error as err:
        print(f"Error fetching data: {err}")
        return None
    finally:
        if cursor:
            cursor.close()

def process_data(data):
    """Processes the fetched data. (Placeholder)"""
    print("Processing data...")
    # Example: Convert all string values to uppercase
    processed = []
    if data:
        for row in data:
            processed_row = {k: (v.upper() if isinstance(v, str) else v) for k, v in row.items()}
            processed.append(processed_row)
    print("Data processing complete.")
    return processed

def write_to_excel(data_rows, output_file_path):
    """Writes the processed data to an Excel .xlsx file."""
    if not data_rows:
        print("No data to write to Excel.")
        return

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    # Write headers (column names from the first row of data)
    headers = list(data_rows[0].keys())
    sheet.append(headers)
    
    # Write data rows
    for row in data_rows:
        sheet.append([row.get(header) for header in headers])
    
    try:
        # Ensure output directory exists
        output_dir = os.path.dirname(output_file_path)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created directory: {output_dir}")

        workbook.save(output_file_path)
        print(f"Data successfully written to {output_file_path}")
    except Exception as e:
        print(f"Error writing to Excel file: {e}")


def run_bancarizacion_process(config_path, output_excel_path):
    """
    Main function to run the bancarizacion process.
    """
    print("--- Bancarizacion Process Started ---")
    try:
        db_config = get_db_config(config_path)
        cnx = connect_to_db(db_config)
        
        if cnx and cnx.is_connected():
            # Replace with your actual query
            query = "SELECT * FROM your_table_name LIMIT 10;" 
            print(f"Executing query: {query}")
            raw_data = fetch_data_from_db(cnx, query)
            
            if raw_data:
                processed_data = process_data(raw_data)
                write_to_excel(processed_data, output_excel_path)
            else:
                print("No data fetched from the database.")
            
            cnx.close()
            print("Database connection closed.")
        else:
            print("Failed to connect to the database. Process aborted.")
            
    except FileNotFoundError as e:
        print(f"Configuration Error: {e}")
    except ValueError as e:
        print(f"Configuration Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    
    print("--- Bancarizacion Process Finished ---")
