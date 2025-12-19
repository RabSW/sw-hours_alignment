"""
SQL Database Explorer
Run this first to verify your SQL connection and understand the data
"""

import pandas as pd
from Data.data_sources import connect_to_sql


def test_connection():
    """Test SQL Server connection"""
    try:
        conn = connect_to_sql()
        print("✓ SQL Server connection successful!")
        conn.close()
        return True
    except Exception as e:
        print(f"✗ Connection failed: {e}")
        return False


def explore_tables():
    """Show sample data from the tables"""
    
    if not test_connection():
        return
    
    conn = connect_to_sql()
    
    print("\n" + "=" * 80)
    print("FACT TABLE - Sample Billable Entries")
    print("=" * 80)
    
    query = """
    SELECT TOP 5
        Hours,
        BillableRate,
        BillableAmount,
        Date,
        CustomerKey,
        EmployeeKey
    FROM PowerBIData.FactTable_HARVEST_Actual
    WHERE IsBillableKey = 1
    ORDER BY Date DESC
    """
    
    df = pd.read_sql(query, conn)
    print(df)
    
    print("\n" + "=" * 80)
    print("EMPLOYEES - Sample Employee Names")
    print("=" * 80)
    
    query = """
    SELECT TOP 10
        EmployeeKey,
        EmployeeName
    FROM PowerBIData.DimEmployee_Tabular_Flat
    WHERE EmployeeName IS NOT NULL
    """
    
    df = pd.read_sql(query, conn)
    print(df)
    
    print("\n" + "=" * 80)
    print("CUSTOMERS - Sample Customer Names")
    print("=" * 80)
    
    query = """
    SELECT TOP 10
        CustomerKey,
        CustomerName
    FROM PowerBIData.DimCustomer_Tabular_Flat
    WHERE CustomerName IS NOT NULL
    """
    
    df = pd.read_sql(query, conn)
    print(df)
    
    print("\n" + "=" * 80)
    print("JOINED DATA - Sample Complete Records")
    print("=" * 80)
    
    query = """
    SELECT TOP 5
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
    ORDER BY f.Date DESC
    """
    
    df = pd.read_sql(query, conn)
    print(df)
    
    conn.close()
    
    print("\n" + "=" * 80)
    print("Connection and queries are working correctly!")
    print("You can now run: python main.py")
    print("=" * 80)


if __name__ == "__main__":
    explore_tables()