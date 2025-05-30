# C:\Users\willy\Projects\bancarizacion\main.py
"""
Main executable script for the Bancarizacion project.
This script orchestrates the overall process.
"""
import os
from bancarizacion.core_logic import run_bancarizacion_process, get_sales_invoice_data, populate_excel_from_template
from datetime import datetime

if __name__ == "__main__":
    print("--- Starting Bancarizacion Application ---")
    
    # Define paths relative to this main.py script
    project_root = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(project_root, "db_config.ini")
    output_file = os.path.join(project_root, "data", "output", "bancarizacion_report.xlsx")
    
    print(f"Using config file: {config_file}")
    print(f"Output will be saved to: {output_file}")
    
    # --- Original process call (can be commented out if only testing the new query) ---
    # print("\n--- Running Full Bancarizacion Process (Placeholder) ---")
    # run_bancarizacion_process(config_path=config_file, output_excel_path=output_file)
    
    # --- New: Fetch and display sales invoice data for a specific year and month ---
    print("\n--- Fetching Sales Invoice Data --- ")
    target_year = 2025
    target_month = 3
    print(f"Requesting data for Year: {target_year}, Month: {target_month}")
    
    sales_data = get_sales_invoice_data(year=target_year, month=target_month, config_file_path=config_file)
    
    if sales_data is not None:
        if sales_data: # If list is not empty
            print(f"Successfully retrieved {len(sales_data)} records.")
            
            # --- Define SIAT column names (these should match headers in PlantillaContratos.xlsx for clarity) ---
            siat_column_names = [
                "N°", "TIPO DE CONTRATO O ACUERDO", "TIPO TRANSACCIÓN",
                "OBJETO DEL CONTRATO O ACUERDO", "NIT/CI PROVEEDOR", "NÚMERO CONTRATO O ACUERDO",
                "FECHA DE CONTRATO O ACUERDO", "IMPORTE TOTAL DEL CONTRATO O ACUERDO (BS)",
                "VALOR DE LA PERMUTA (BS)", "CANTIDAD DE CUOTAS",
                "IMPORTE ADELANTADO, IMPORTE RETENIDO, DESCUENTOS U OTROS (BS)",
                "OBJETO PERMUTA", "MONTO ACUMULADO (BS)"
            ]

            # --- Transform sales_data to SIAT format ---
            data_for_siat_template = []
            for index, record in enumerate(sales_data):
                # Original data keys from get_sales_invoice_data:
                # contractType, transactionType, contractObject, providerNit, contractNumber,
                # contractDate, totalAmount, exchangeValue, numberOfInstallments,
                # advanceAmount, exchangeObject, accumulatedAmount, invoiceId, paid

                numero_secuencial = index + 1
                
                # contractType from query is 2 (Verbal)
                # transactionType from query is 1 (Con Factura)
                tipo_contrato_val = record.get('contractType')
                tipo_transaccion_val = record.get('transactionType')
                
                objeto_contrato_val = str(record.get('contractObject', ''))[:100]
                
                # For sales, NIT/CI PROVEEDOR is blank
                nit_proveedor_val = '' 
                
                # For verbal contracts (type 2), NÚMERO CONTRATO is blank
                numero_contrato_val = str(record.get('contractNumber', ''))
                if tipo_contrato_val == 2: # 2 is 'Verbal'
                    numero_contrato_val = ''
                
                fecha_contrato_dt = record.get('contractDate')
                fecha_contrato_str = fecha_contrato_dt.strftime("%d/%m/%Y") if fecha_contrato_dt else ''
                
                importe_total_val = record.get('totalAmount') # Should be numeric
                
                valor_permuta_raw = record.get('exchangeValue', 0)
                valor_permuta_val = valor_permuta_raw if valor_permuta_raw != 0 else ''
                
                cantidad_cuotas_val = record.get('numberOfInstallments', 1)
                
                importe_adelantado_raw = record.get('advanceAmount', 0)
                # If advanceAmount is 0 or not present, set to 0, otherwise use the value
                importe_adelantado_val = importe_adelantado_raw if importe_adelantado_raw else 0
                
                objeto_permuta_val = str(record.get('exchangeObject', ''))[:100]
                
                # MONTO ACUMULADO should be 0 if no other information
                monto_acumulado_val = record.get('accumulatedAmount', 0) # Assuming 'accumulatedAmount' is the key
                if monto_acumulado_val == '' or monto_acumulado_val is None: # Or if it's not in the record
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
                data_for_siat_template.append(row_dict_for_siat)

            # --- Populate Excel Template with SIAT formatted data ---
            template_name = "PlantillaContratos.xlsx"
            template_path = os.path.join(project_root, "data", template_name) 
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_excel_name = f"{os.path.splitext(template_name)[0]}_SIAT_{timestamp}.xlsx" # Added _SIAT_
            output_excel_full_path = os.path.join(project_root, "data", "output", output_excel_name)

            print(f"\nAttempting to populate template '{template_name}' with {len(data_for_siat_template)} records for SIAT format.")
            # The 'populate_excel_from_template' function expects a list of dicts (data_for_siat_template)
            # and a list of keys (siat_column_names) to determine the order and which data to pick.
            success = populate_excel_from_template(data_for_siat_template, template_path, output_excel_full_path, siat_column_names)
            if success:
                print(f"Excel template populated and saved to: {output_excel_full_path}")
            else:
                print(f"Failed to populate Excel template. Check logs.")

        else:
            print("No sales invoice records found for the specified period.")
    else:
        print("Failed to retrieve sales invoice data. Check logs for errors.")

    print("\n--- Bancarizacion Application Finished --- ")
