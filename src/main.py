"""
Main Script
Orchestrates the entire hours reconciliation process
"""

from Data.data_sources import get_billable_data, get_sharepoint_data
from Workflow.reconciliation import reconcile_data
from Workflow.report import create_report


def main():
    """Main execution"""
    print("=" * 60)
    print("BILLING RECONCILIATION")
    print("=" * 60)
    
    # Get month filter
    month_filter = input("Enter month (YYYY-MM) or press Enter for all: ").strip()
    if not month_filter:
        month_filter = None
    
    # Step 1: Extract SQL data
    print("\n1. Extracting billable data from SQL...")
    sql_df = get_billable_data(month_filter)
    print(f"   Found {len(sql_df)} billable entries")
    
    # Step 2: Get pricing data from local file
    print("\n2. Reading pricing data from local file...")
    employees_df, regular_df, fcc_df = get_sharepoint_data()
    print(f"   Loaded {len(employees_df)} employees")
    print(f"   Loaded {len(regular_df)} regular customers")
    print(f"   Loaded {len(fcc_df)} FCC customers")
    
    # Step 3: Reconcile
    print("\n3. Reconciling data...")
    results_df, unmatched_employees, unmatched_customers = reconcile_data(
        sql_df, employees_df, regular_df, fcc_df
    )
    
    # Calculate statistics
    discrepancies = results_df[abs(results_df['discrepancy_pct']) > 1]
    print(f"   Found {len(discrepancies)} entries with >1% discrepancy")
    print(f"   Total discrepancy: {results_df['discrepancy'].sum():,.2f}")
    
    if unmatched_employees:
        print(f"   WARNING: {len(set(unmatched_employees))} unmatched employees")
    if unmatched_customers:
        print(f"   WARNING: {len(set(unmatched_customers))} unmatched customers")
    
    # Step 4: Generate report
    print("\n4. Generating report...")
    output_file = create_report(results_df, unmatched_employees, unmatched_customers)
    print(f"   Report saved: {output_file}")
    
    print("\n" + "=" * 60)
    print("RECONCILIATION COMPLETE")
    print("=" * 60)


if __name__ == "__main__":
    main()