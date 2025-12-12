"""
Main orchestration script for SQL and SharePoint integration.
"""

import os
from dotenv import load_dotenv
from data import SQLConnection, SharePointConnection


def setup_connections():
    """Initialize SQL and SharePoint connections from environment variables."""
    load_dotenv()
    
    sql_conn = SQLConnection(
        server=os.getenv("SQL_SERVER"),
        database=os.getenv("SQL_DATABASE"),
        username=os.getenv("SQL_USERNAME"),
        password=os.getenv("SQL_PASSWORD"),
        driver=os.getenv("SQL_DRIVER")
    )
    
    sp_conn = SharePointConnection(
        site_url=os.getenv("SHAREPOINT_SITE_URL"),
        username=os.getenv("SHAREPOINT_USERNAME"),
        password=os.getenv("SHAREPOINT_PASSWORD")
    )
    
    return sql_conn, sp_conn


def process_data(sql_conn, sp_conn):
    """Main business logic - customize this for your needs."""
    
    # Example: Query data from SQL
    # df = sql_conn.execute_query("SELECT TOP 10 * FROM your_table")
    
    # Example: Download file from SharePoint
    # sp_conn.download_file(
    #     file_path=os.getenv("SHAREPOINT_FILE_PATH"),
    #     local_path="./downloaded_file.xlsx"
    # )
    
    # Example: Process data and upload results
    # df.to_excel("./results.xlsx", index=False)
    # sp_conn.upload_file("./results.xlsx", "/Shared Documents")
    
    pass


def main():
    """Main entry point."""
    
    sql_conn, sp_conn = setup_connections()
    
    try:
        # Connect to services
        sql_conn.connect()
        sp_conn.connect()
        
        # Run business logic
        process_data(sql_conn, sp_conn)
        
    except Exception as e:
        print(f"Error: {e}")
        raise
    
    finally:
        # Clean up
        sql_conn.close()


if __name__ == "__main__":
    main()

