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
from datetime import datetime # Added for potential use, though timestamp generation is in main.py for this feature

def get_db_config_core(config_file_path="db_config.ini"):
    """Reads database configuration from an INI file for core logic."""
    # Correctly locate db_config.ini relative to the project root
    # Assumes core_logic.py is in 'bancarizacion' subdirectory
    if not os.path.isabs(config_file_path):
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        config_file_path = os.path.join(project_root, config_file_path)

    if not os.path.exists(config_file_path):
        raise FileNotFoundError(f"Configuration file {config_file_path} not found.")

    config = configparser.ConfigParser()
    config.read(config_file_path)

    if 'mysql' not in config:
        raise ValueError(f"[mysql] section not found in {config_file_path}.")
    
    db_params = {}
    try:
        db_params['host'] = config.get('mysql', 'host')
        db_params['user'] = config.get('mysql', 'user')
        db_params['password'] = config.get('mysql', 'password')
        db_params['database'] = config.get('mysql', 'database')
        db_params['port'] = config.getint('mysql', 'port')
    except configparser.NoOptionError as e:
        raise ValueError(f"Missing option in [mysql] section of {config_file_path}: {e}")        
    
    return db_params

def connect_to_db(db_config_params):
    """Connects to the MySQL database using a dictionary of parameters."""
    try:
        # Ensure port is treated as integer if it's a string coming from certain configs
        if 'port' in db_config_params and isinstance(db_config_params['port'], str):
            db_config_params['port'] = int(db_config_params['port'])
            
        cnx = mysql.connector.connect(**db_config_params)
        print(f"Successfully connected to database: {db_config_params.get('database')} on host {db_config_params.get('host')}:{db_config_params.get('port')}")
        return cnx
    except mysql.connector.Error as err:
        print(f"Error connecting to database: {err}")
        # Consider re-raising the error or handling it more gracefully depending on application flow
        raise # Re-raise the exception to be caught by the caller
    except ValueError as ve:
        print(f"Configuration error for database connection: {ve}")
        raise # Re-raise for caller to handle

def fetch_data_from_db(cnx, query, params=None):
    """Fetches data from the database using a SELECT query and parameters."""
    cursor = None
    try:
        cursor = cnx.cursor(dictionary=True) # dictionary=True to get results as dicts
        cursor.execute(query, params)
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

def populate_excel_from_template(data_rows, template_file_path, output_file_path, column_order):
    """
    Populates an Excel template with data_rows and saves it to output_file_path.
    Assumes headers are already in the template and data should be appended after them.
    """
    if not data_rows:
        print("No data provided to populate Excel template.")
        return False

    try:
        if not os.path.exists(template_file_path):
            print(f"Error: Template file not found at {template_file_path}")
            return False

        workbook = openpyxl.load_workbook(template_file_path)
        sheet = workbook.active  # Assumes data goes into the active sheet

        # Append data rows based on the specified column_order
        for record in data_rows:
            row_values = [record.get(col_name, "") for col_name in column_order] # Use empty string for missing keys
            sheet.append(row_values)
        
        # Ensure output directory exists
        output_dir = os.path.dirname(output_file_path)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"Created directory: {output_dir}")

        workbook.save(output_file_path)
        print(f"Data successfully written to {output_file_path} using template {os.path.basename(template_file_path)}")
        return True
    except Exception as e:
        print(f"Error populating Excel template: {e}")
        return False

def run_bancarizacion_process(config_path, output_excel_path):
    """
    Main function to orchestrate the bancarizacion process.
    1. Reads DB config.
    2. Connects to DB.
    3. Fetches data (using a placeholder query for now).
    4. Processes data.
    5. Writes data to Excel.
    """
    try:
        db_params = get_db_config_core(config_file_path=config_path)
        cnx = connect_to_db(db_config_params=db_params)
        
        if cnx and cnx.is_connected():
            # Placeholder query - replace with actual query logic
            # For example, you might call get_sales_invoice_data here if it's the main data source
            query = "SELECT * FROM your_table_name LIMIT 10;" # Replace this
            print(f"Executing placeholder query: {query}")
            # data = fetch_data_from_db(cnx, query) # If using a generic fetch
            
            # Example of using the new specific function (if this is the primary goal)
            # current_year = 2025
            # current_month = 3
            # print(f"Fetching sales invoice data for {current_year}-{current_month:02d}")
            # data = get_sales_invoice_data(current_year, current_month, config_path) # This function is defined below

            # For now, let's assume the original flow with a generic query for run_bancarizacion_process
            # and main.py will call the new function directly for testing.
            # If get_sales_invoice_data is the *only* data needed, integrate its call here.
            # For demonstration, we'll keep it separate and main.py will call it.
            
            # This part would use a generic query or be replaced
            # For now, let's simulate no data from the placeholder to avoid errors
            data = [] 
            
            if data:
                processed_data = process_data(data)
                write_to_excel(processed_data, output_excel_path)
            else:
                print("No data fetched or placeholder query used. Skipping processing and Excel writing.")

    except FileNotFoundError as e:
        print(f"Configuration Error: {e}")
    except ValueError as e:
        print(f"Configuration Value Error: {e}")
    except mysql.connector.Error as e:
        print(f"Database Process Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred in run_bancarizacion_process: {e}")
    finally:
        if 'cnx' in locals() and cnx and cnx.is_connected():
            cnx.close()
            print("Database connection closed in run_bancarizacion_process.")

def get_sales_invoice_data(year, month, config_file_path="db_config.ini"):
    """
    Fetches sales invoice data for a given year and month.
    """
    db_params = None
    cnx = None
    try:
        db_params = get_db_config_core(config_file_path=config_file_path)
        cnx = connect_to_db(db_config_params=db_params)

        if not (cnx and cnx.is_connected()):
            print("Cannot fetch sales invoice data, database connection failed.")
            return None

        query = """
        SELECT
            2 AS contractType,
            1 AS transactionType,
            'VENTA DE MERCADERIA' AS contractObject,
            '' AS providerNit,
            '' AS contractNumber,
            f.fechaFac AS contractDate,
            f.total AS totalAmount,
            0 AS exchangeValue,
            1 AS numberOfInstallments,
            0 AS advanceAmount,
            '' AS exchangeObject,
            0 AS accumulatedAmount,
            f.idFactura AS invoiceId,
            f.pagada AS paid
        FROM
            factura f
        WHERE
            YEAR(f.fechaFac) = %s
            AND MONTH(f.fechaFac) = %s
            AND f.anulada = 0
            AND f.total >= 50000
        """
        params = (year, month)
        print(f"Executing query for sales invoices with year={year}, month={month}")
        data = fetch_data_from_db(cnx, query, params)
        
        if data is not None:
            print(f"Successfully fetched {len(data)} sales invoices.")
        else:
            print("No sales invoice data returned from query or an error occurred.")
        return data

    except FileNotFoundError as e:
        print(f"Configuration Error for sales invoice data: {e}")
        return None
    except ValueError as e:
        print(f"Configuration Value Error for sales invoice data: {e}")
        return None
    except mysql.connector.Error as e:
        print(f"Database Error fetching sales invoice data: {e}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred in get_sales_invoice_data: {e}")
        return None
    finally:
        if cnx and cnx.is_connected():
            cnx.close()
            print("Database connection closed in get_sales_invoice_data.")

# Example of how to use the Excel writing part if needed separately
# def generate_excel_report(data_to_write, output_path):
# write_to_excel(data_to_write, output_path)
