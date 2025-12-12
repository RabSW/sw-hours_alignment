# SW Hours Alignment

SQL and SharePoint integration script for hours alignment.

## Setup

1. **Install dependencies:**
   ```bash
   uv sync
   ```

2. **Configure environment variables:**
   - Copy `.env.example` to `.env`
   - Fill in your SQL Server and SharePoint credentials

3. **Run the script:**
   ```bash
   uv run python src/main.py
   ```

## Configuration

### SQL Server
- `SQL_SERVER`: Your SQL Server address
- `SQL_DATABASE`: Database name
- `SQL_USERNAME`: SQL authentication username
- `SQL_PASSWORD`: SQL authentication password
- `SQL_DRIVER`: ODBC driver (default: "ODBC Driver 17 for SQL Server")

### SharePoint
- `SHAREPOINT_SITE_URL`: Full URL to your SharePoint site
- `SHAREPOINT_USERNAME`: Your SharePoint/Microsoft 365 email
- `SHAREPOINT_PASSWORD`: Your SharePoint/Microsoft 365 password
- `SHAREPOINT_FILE_PATH`: Path to the file on SharePoint (e.g., "/Shared Documents/file.xlsx")

## Usage

The script provides two main classes:

### SQLConnection
```python
sql_conn = SQLConnection(server, database, username, password, driver)
sql_conn.connect()
df = sql_conn.execute_query("SELECT * FROM table")
sql_conn.close()
```

### SharePointConnection
```python
sp_conn = SharePointConnection(site_url, username, password)
sp_conn.connect()
sp_conn.download_file("/path/to/file.xlsx", "./local_file.xlsx")
sp_conn.upload_file("./local_file.xlsx", "/Shared Documents")
```

## Quick Start

1. Install dependencies:
   ```bash
   uv sync
   ```

2. Create `.env` file with your credentials (see `config.example.txt`)

3. Edit `src/main.py` to uncomment the example queries you want to use

4. Run:
   ```bash
   uv run python src/main.py
   ```

## Docker

To run in Docker:
```bash
docker build -t sw-hours-alignment .
docker run --env-file .env sw-hours-alignment
```

## Requirements

- Python 3.10+
- SQL Server with ODBC Driver 17 installed
- SharePoint/Microsoft 365 account with appropriate permissions

## Notes

- The script uses environment variables for security (never commit `.env` to git)
- SQL queries return Pandas DataFrames for easy data manipulation
- SharePoint operations support both download and upload
