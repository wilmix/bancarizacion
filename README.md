# Bancarizacion Project

This project handles the bancarizacion process, which involves:
1.  Connecting to a MySQL database using credentials from `db_config.ini`.
2.  Running SELECT queries to fetch data.
3.  Processing the retrieved data.
4.  Generating Excel (.xlsx) spreadsheets with the processed data into the `data/output/` directory.

## Setup

1.  **Clone the repository (if applicable) or ensure you are in the project root directory `C:\Users\willy\Projects\bancarizacion\`.**
2.  **Create a Python virtual environment:**
    ```bash
    python -m venv venv
    ```
3.  **Activate the virtual environment:**
    *   Windows (Command Prompt or PowerShell):
        ```bash
        .\venv\Scripts\activate
        ```
    *   Git Bash or similar (Linux/macOS style):
        ```bash
        source venv/Scripts/activate 
        ```
4.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
5.  **Configure database access:**
    *   Copy `db_config.ini.example` to `db_config.ini`.
    *   Edit `db_config.ini` with your actual MySQL database credentials.

## Usage

Run the main script from the project root directory (`C:\\Users\\willy\\Projects\\bancarizacion\\`):
```bash
python main.py
```
This will execute the process defined in `bancarizacion/core_logic.py`:
1. Connect to the MySQL database using credentials from `db_config.ini`.
2. Fetch sales invoice data for a predefined year and month (e.g., 2025, month 3) where the total is >= 50,000.
3. Populate the `data/PlantillaContratos.xlsx` template with the fetched data.
4. Save the populated Excel file to the `data/output/` directory with a timestamp in the filename (e.g., `PlantillaContratos_YYYYMMDD_HHMMSS.xlsx`).

## Database Query for Sales Invoices

The script `main.py` calls `get_sales_invoice_data` in `bancarizacion/core_logic.py`, which executes a SQL query similar to the following to retrieve sales data meeting bancarizacion criteria (total >= 50,000):

```sql
SELECT
    2 AS contractType,          -- Tipo de Contrato (e.g., Venta de Bienes)
    1 AS transactionType,       -- Tipo de Transacción (e.g., Ingreso)
    'VENTA DE MERCADERIA' AS contractObject, -- Objeto del Contrato
    '' AS providerNit,           -- NIT del Proveedor (si aplica, para compras)
    '' AS contractNumber,        -- Número de Contrato (si existe)
    f.fechaFac AS contractDate,  -- Fecha del Contrato/Factura
    f.total AS totalAmount,      -- Importe Total del Contrato/Factura
    0 AS exchangeValue,         -- Valor de Permuta (si aplica)
    1 AS numberOfInstallments,  -- Cantidad de Cuotas (si aplica)
    0 AS advanceAmount,         -- Importe Adelantado (si aplica)
    '' AS exchangeObject,        -- Objeto de la Permuta (si aplica)
    0 AS accumulatedAmount,     -- Monto Acumulado (para control interno)
    f.idFactura AS invoiceId     -- ID de la Factura (para referencia)
    -- f.pagada AS paid          -- Estado de pago (disponible pero no usado en la plantilla)
FROM
    factura f
WHERE
    YEAR(f.fechaFac) = %s -- Parameter for Year
    AND MONTH(f.fechaFac) = %s -- Parameter for Month
    AND f.anulada = 0
    AND f.total >= 50000;
```

## Excel Template Population

The `populate_excel_from_template` function in `core_logic.py` is responsible for:
- Loading the `data/PlantillaContratos.xlsx`.
- Appending rows from the fetched sales data into the active sheet.
- Mapping data to columns in a predefined order, excluding the `paid` status.
- Saving the new file to `data/output/PlantillaContratos_<timestamp>.xlsx`.

Make sure the column headers in your `data/PlantillaContratos.xlsx` template correspond to the data being selected by the query and the `column_mapping_order` defined in `main.py`.
