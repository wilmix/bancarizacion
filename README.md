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

Run the main script from the project root directory (`C:\Users\willy\Projects\bancarizacion\`):
```bash
python main.py
```
This will execute the process defined in `bancarizacion/core_logic.py`, connect to the database, fetch data (you'll need to update the placeholder query in `core_logic.py`), process it, and save an Excel report to `data/output/bancarizacion_report.xlsx`.
