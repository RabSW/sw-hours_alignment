"""
Data Sources Module
Handles all data connections: SQL Server and SharePoint
"""

import pandas as pd
import pyodbc
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File
import io
import config


def connect_to_sql():
    """Create SQL Server connection"""
    conn_str = (
        f"DRIVER={{{config.SQL_DRIVER}}};"
        f"SERVER={config.SQL_SERVER};"
        f"DATABASE={config.SQL_DATABASE};"
        f"UID={config.SQL_USERNAME};"
        f"PWD={config.SQL_PASSWORD}"
    )
    return pyodbc.connect(conn_str)


def get_billable_data(month_filter=None):
    """
    Extract billable data from SQL Server
    
    Args:
        month_filter: Optional month filter in YYYY-MM format
    
    Returns:
        DataFrame with columns: Hours, BillableRate, BillableAmount, Date, CustomerName, EmployeeName
    """
    conn = connect_to_sql()
    
    query = """
    SELECT 
        f.Hours,
        f.BillableRate,
        f.BillableAmount,
        f.Date,
        c.CustomerName,
        e.EmployeeName
    FROM PowerBIData.FactTable_HARVEST_Actual f
    JOIN PowerBIData.DimCustomer_Tabular_Flat c ON f.CustomerKey = c.CustomerKey
    JOIN PowerBIData.DimEmployee_Tabular_Flat e ON f.EmployeeKey = e.EmployeeKey
    WHERE f.IsBillableKey = 1
    """
    
    if month_filter:
        query += f" AND FORMAT(f.Date, 'yyyy-MM') = '{month_filter}'"
    
    df = pd.read_sql(query, conn)
    conn.close()
    
    return df


def parse_sharepoint_file(file_bytes):
    """
    Parse SharePoint Excel file into structured data
    
    Args:
        file_bytes: BytesIO object containing Excel file
    
    Returns:
        Tuple of (employees_df, regular_pricing_df, fcc_pricing_df)
    """
    df_raw = pd.read_excel(file_bytes, sheet_name=0, header=None)
    
    # Extract employee list (rows 118+, columns E=4 and F=5)
    employees = []
    for i in range(118, len(df_raw)):
        name = df_raw.iloc[i, 4]
        title = df_raw.iloc[i, 5]
        if pd.notna(name) and pd.notna(title) and name not in ['Consulent', 'NUMBERS']:
            employees.append({
                'name': str(name).strip(),
                'title': str(title).strip()
            })
    
    employees_df = pd.DataFrame(employees)
    
    # Extract regular customer pricing (rows 6-109)
    regular_pricing = []
    for i in range(6, 110):
        customer = df_raw.iloc[i, 0]
        if pd.notna(customer):
            regular_pricing.append({
                'customer': str(customer).strip(),
                'Junior Consultant': df_raw.iloc[i, 2],
                'Consultant': df_raw.iloc[i, 3],
                'Senior Consultant': df_raw.iloc[i, 4],
                'Principal Consultant': df_raw.iloc[i, 5],
                'Data Scientist': df_raw.iloc[i, 6],
                'Support': df_raw.iloc[i, 7]
            })
    
    regular_df = pd.DataFrame(regular_pricing)
    
    # Extract FCC customer pricing (rows 115-138)
    fcc_pricing = []
    for i in range(115, 139):
        customer = df_raw.iloc[i, 0]
        price = df_raw.iloc[i, 1]
        if pd.notna(customer) and customer != 'Customer':
            fcc_pricing.append({
                'customer': str(customer).strip(),
                'Consultant': price
            })
    
    fcc_df = pd.DataFrame(fcc_pricing)
    
    return employees_df, regular_df, fcc_df


def get_sharepoint_data():
    """
    Read and parse local Excel file (previously from SharePoint)
    
    Returns:
        Tuple of (employees_df, regular_pricing_df, fcc_pricing_df)
    """
    import os
    
    # Path to local Excel file (relative to this script)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, "Hourly rate 2026.xlsx")
    
    if not os.path.exists(file_path):
        raise FileNotFoundError(
            f"Excel file not found at: {file_path}\n"
            f"Please ensure 'Hourly rate 2026.xlsx' is in the src/Data/ folder"
        )
    
    # Read file into BytesIO for consistent parsing
    with open(file_path, 'rb') as f:
        bytes_file = io.BytesIO(f.read())
    
    return parse_sharepoint_file(bytes_file)