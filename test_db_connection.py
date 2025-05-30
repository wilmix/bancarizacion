import configparser
import os
import mysql.connector
from mysql.connector import Error

def get_db_config(config_path='db_config.ini'):
    """Reads database configuration from the .ini file."""
    if not os.path.exists(config_path):
        print(f"Error: Configuration file '{config_path}' not found.")
        return None
    
    config = configparser.ConfigParser()
    config.read(config_path)
    
    if 'mysql' not in config:
        print(f"Error: [mysql] section not found in '{config_path}'.")
        return None
        
    db_config = {}
    try:
        db_config['host'] = config.get('mysql', 'host')
        db_config['user'] = config.get('mysql', 'user')
        db_config['password'] = config.get('mysql', 'password')
        db_config['database'] = config.get('mysql', 'database')
        db_config['port'] = config.getint('mysql', 'port') # Read port as an integer
    except configparser.NoOptionError as e:
        print(f"Error: Missing option in [mysql] section of '{config_path}': {e}")
        return None
        
    return db_config

def test_mysql_connection(config_path='db_config.ini'):
    """Tests the MySQL database connection using credentials from config_path."""
    db_params = get_db_config(config_path)
    
    if not db_params:
        return False

    try:
        print(f"Attempting to connect to database '{db_params['database']}' on host '{db_params['host']}' with user '{db_params['user']}'...")
        connection = mysql.connector.connect(
            host=db_params['host'],
            user=db_params['user'],
            password=db_params['password'],
            database=db_params['database'],
            port=db_params['port']  # Add port to the connection parameters
        )
        if connection.is_connected():
            db_Info = connection.get_server_info()
            print(f"Successfully connected to MySQL Server version {db_Info}")
            cursor = connection.cursor()
            cursor.execute("SELECT DATABASE();")
            record = cursor.fetchone()
            print(f"You're connected to database: {record[0]}")
            
            # Test a simple query
            cursor.execute("SHOW TABLES;")
            tables = cursor.fetchall()
            print("\nTables in the database:")
            if tables:
                for table in tables:
                    print(f"- {table[0]}")
            else:
                print("(No tables found)")
            
            return True
    except Error as e:
        print(f"Error while connecting to MySQL: {e}")
        return False
    finally:
        if 'connection' in locals() and connection.is_connected():
            cursor.close()
            connection.close()
            print("MySQL connection is closed.")

if __name__ == "__main__":
    project_root = os.path.dirname(os.path.abspath(__file__))
    config_file_path = os.path.join(project_root, "db_config.ini")
    
    print("--- Testing Database Connection ---")
    if not os.path.exists(config_file_path):
        print(f"CRITICAL: Database configuration file '{config_file_path}' does not exist.")
        print("Please ensure 'db_config.ini' is present in the project root and correctly configured.")
    else:
        test_mysql_connection(config_path=config_file_path)
    print("--- Database Connection Test Finished ---")
