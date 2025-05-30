# C:\Users\willy\Projects\bancarizacion\main.py
"""
Main executable script for the Bancarizacion project.
This script orchestrates the overall process.
"""
import os
from bancarizacion.core_logic import run_bancarizacion_process

if __name__ == "__main__":
    print("--- Starting Bancarizacion Application ---")
    
    # Define paths relative to this main.py script
    project_root = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(project_root, "db_config.ini")
    output_file = os.path.join(project_root, "data", "output", "bancarizacion_report.xlsx")
    
    print(f"Using config file: {config_file}")
    print(f"Output will be saved to: {output_file}")
    
    run_bancarizacion_process(config_path=config_file, output_excel_path=output_file)
    
    print("--- Bancarizacion Application Finished ---")
