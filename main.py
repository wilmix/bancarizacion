# C:\Users\willy\Projects\bancarizacion\main.py
"""
Main executable script for the Bancarizacion project.
This script orchestrates the overall process.
"""
import os
import sys # Import sys to access command-line arguments
from bancarizacion.core_logic import (
    get_sales_invoice_data, populate_excel_from_template, get_auxiliary_sales_data, write_to_excel,
    process_zipped_contracts_excel  # Added for processing zipped contracts Excel
)
from datetime import datetime

def process_contratos(project_root, config_file):
    """Processes Sales Invoice Data (Contratos)."""
    print("\\n--- Processing Sales Invoice Data (Contratos) ---")
    target_year_contratos = 2025
    target_month_contratos = 3
    print(f"Requesting Contratos data for Year: {target_year_contratos}, Month: {target_month_contratos}")
    
    sales_data_contratos = get_sales_invoice_data(year=target_year_contratos, month=target_month_contratos, config_file_path=config_file)
    
    if sales_data_contratos is not None:
        if sales_data_contratos:
            print(f"Successfully retrieved {len(sales_data_contratos)} records for Contratos.")
            
            siat_column_names_contratos = [
                "N°", "TIPO DE CONTRATO O ACUERDO", "TIPO TRANSACCIÓN",
                "OBJETO DEL CONTRATO O ACUERDO", "NIT/CI PROVEEDOR", "NÚMERO CONTRATO O ACUERDO",
                "FECHA DE CONTRATO O ACUERDO", "IMPORTE TOTAL DEL CONTRATO O ACUERDO (BS)",
                "VALOR DE LA PERMUTA (BS)", "CANTIDAD DE CUOTAS",
                "IMPORTE ADELANTADO, IMPORTE RETENIDO, DESCUENTOS U OTROS (BS)",
                "OBJETO PERMUTA", "MONTO ACUMULADO (BS)"
            ]

            data_for_siat_template_contratos = []
            for index, record in enumerate(sales_data_contratos):
                numero_secuencial = index + 1
                tipo_contrato_val = record.get('contractType')
                tipo_transaccion_val = record.get('transactionType')
                objeto_contrato_val = str(record.get('contractObject', ''))[:100]
                nit_proveedor_val = '' 
                numero_contrato_val = str(record.get('contractNumber', ''))
                if tipo_contrato_val == 2:
                    numero_contrato_val = ''
                fecha_contrato_dt = record.get('contractDate')
                fecha_contrato_str = fecha_contrato_dt.strftime("%d/%m/%Y") if fecha_contrato_dt else ''
                importe_total_val = record.get('totalAmount')
                valor_permuta_raw = record.get('exchangeValue', 0)
                valor_permuta_val = valor_permuta_raw if valor_permuta_raw != 0 else ''
                cantidad_cuotas_val = record.get('numberOfInstallments', 1)
                importe_adelantado_raw = record.get('advanceAmount', 0)
                importe_adelantado_val = importe_adelantado_raw if importe_adelantado_raw else 0
                objeto_permuta_val = str(record.get('exchangeObject', ''))[:100]
                monto_acumulado_val = record.get('accumulatedAmount', 0)
                if monto_acumulado_val == '' or monto_acumulado_val is None:
                    monto_acumulado_val = 0

                row_dict_for_siat = {
                    "N°": numero_secuencial,
                    "TIPO DE CONTRATO O ACUERDO": tipo_contrato_val,
                    "TIPO TRANSACCIÓN": tipo_transaccion_val,
                    "OBJETO DEL CONTRATO O ACUERDO": objeto_contrato_val,
                    "NIT/CI PROVEEDOR": nit_proveedor_val,
                    "NÚMERO CONTRATO O ACUERDO": numero_contrato_val,
                    "FECHA DE CONTRATO O ACUERDO": fecha_contrato_str,
                    "IMPORTE TOTAL DEL CONTRATO O ACUERDO (BS)": importe_total_val,
                    "VALOR DE LA PERMUTA (BS)": valor_permuta_val,
                    "CANTIDAD DE CUOTAS": cantidad_cuotas_val,
                    "IMPORTE ADELANTADO, IMPORTE RETENIDO, DESCUENTOS U OTROS (BS)": importe_adelantado_val,
                    "OBJETO PERMUTA": objeto_permuta_val,
                    "MONTO ACUMULADO (BS)": monto_acumulado_val
                }
                data_for_siat_template_contratos.append(row_dict_for_siat)

            template_name_contratos = "PlantillaContratos.xlsx"
            template_path_contratos = os.path.join(project_root, "data", template_name_contratos) 
            
            timestamp_contratos = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_excel_name_contratos = f"{os.path.splitext(template_name_contratos)[0]}_SIAT_{timestamp_contratos}.xlsx"
            output_excel_full_path_contratos = os.path.join(project_root, "data", "output", output_excel_name_contratos)

            print(f"\\nAttempting to populate template '{template_name_contratos}' with {len(data_for_siat_template_contratos)} records for SIAT format.")
            success_contratos = populate_excel_from_template(data_for_siat_template_contratos, template_path_contratos, output_excel_full_path_contratos, siat_column_names_contratos)
            if success_contratos:
                print(f"Contratos Excel template populated and saved to: {output_excel_full_path_contratos}")
            else:
                print(f"Failed to populate Contratos Excel template. Check logs.")
        else:
            print("No sales invoice records found for Contratos for the specified period.")
    else:
        print("Failed to retrieve sales invoice data for Contratos. Check logs for errors.")

def process_auxiliary_sales(project_root, config_file):
    """Processes Auxiliary Sales Data (Registro Auxiliar de Ventas) and writes to a new Excel file."""
    print("\\\\n\\\\n--- Processing Auxiliary Sales Data (Registro Auxiliar de Ventas) ---")
    target_year_aux_ventas = 2025
    target_month_aux_ventas = 3
    print(f"Requesting Auxiliary Sales data for Year: {target_year_aux_ventas}, Month: {target_month_aux_ventas}")

    aux_sales_data = get_auxiliary_sales_data(year=target_year_aux_ventas, month=target_month_aux_ventas, config_file_path=config_file)

    if aux_sales_data is not None:
        if aux_sales_data:
            print(f"Successfully retrieved {len(aux_sales_data)} records for Auxiliary Sales.")
            
            # No specific column pre-definition needed if write_to_excel derives headers from data.
            # if aux_sales_data:
            #     aux_sales_column_names = list(aux_sales_data[0].keys())
            # else:
            #     aux_sales_column_names = []

            output_aux_excel_name = f"AuxiliarySalesData_Raw_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            output_aux_excel_full_path = os.path.join(project_root, "data", "output", output_aux_excel_name)
            
            print(f"\\\\nAttempting to save raw auxiliary sales data to a new Excel file: {output_aux_excel_full_path}")
            
            # Call write_to_excel to create a new file with the data
            success_aux_ventas = write_to_excel(aux_sales_data, output_aux_excel_full_path)
            
            if success_aux_ventas:
                print(f"Auxiliary Sales data successfully written to: {output_aux_excel_full_path}")
            else:
                print(f"Failed to write Auxiliary Sales data to Excel. Check logs.")
        else:
            print("No auxiliary sales records found for the specified period.")
    else:
        print("Failed to retrieve auxiliary sales data. Check logs for errors.")

def process_zipped_contracts(project_root):
    """Processes zipped contract data from ContratosXlsx.zip."""
    print("\n--- Processing Zipped Contracts Data ---")
    
    # Path to the zip file
    zip_file_path = os.path.join(project_root, "data", "ContratosXlsx.zip")
    
    # Process the zip file and get filtered contract data
    contract_data = process_zipped_contracts_excel(zip_file_path)
    
    if contract_data is not None:
        if contract_data:
            print(f"Successfully processed {len(contract_data)} pending contracts from the zip file.")
            
            # Define output file path
            output_excel_name = f"FilteredContracts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            output_excel_path = os.path.join(project_root, "data", "output", output_excel_name)
            
            # Write the filtered data to an Excel file
            print(f"\nWriting filtered contract data to Excel: {output_excel_path}")
            write_to_excel(contract_data, output_excel_path)
            print(f"Filtered contract data saved to: {output_excel_path}")
        else:
            print("No pending contracts found in the zip file after filtering.")
    else:
        print("Failed to process the zipped contracts file. Check logs for errors.")

if __name__ == "__main__":
    print("--- Starting Bancarizacion Application ---")
    
    project_root = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(project_root, "db_config.ini")
    # output_file is not directly used here anymore as each function defines its own output.
    
    print(f"Using config file: {config_file}")    # Check command-line arguments
    args = sys.argv[1:] # Get arguments, excluding the script name
    
    if not args:
        # No arguments provided, run all processes
        print("No specific process requested, running 'contratos', 'auxventas', and 'zipcontratos'.")
        process_contratos(project_root, config_file)
        process_auxiliary_sales(project_root, config_file)
        process_zipped_contracts(project_root)
    elif "contratos" in args:
        print("Processing 'contratos' requested.")
        process_contratos(project_root, config_file)
    elif "auxventas" in args:
        print("Processing 'auxventas' requested.")
        process_auxiliary_sales(project_root, config_file)
    elif "zipcontratos" in args:
        print("Processing 'zipcontratos' requested.")
        process_zipped_contracts(project_root)
    else:
        print("Invalid argument. Available options: 'contratos', 'auxventas', 'zipcontratos', or no argument to run all.")

    print("\\n--- Bancarizacion Application Finished ---")
