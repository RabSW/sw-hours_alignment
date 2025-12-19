"""
Report Module
Generates Excel reports with reconciliation results
"""

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from datetime import datetime


def create_report(results_df, unmatched_employees, unmatched_customers):
    """
    Generate Excel report with reconciliation results
    
    Args:
        results_df: DataFrame with reconciliation results
        unmatched_employees: List of tuples (employee_name, match_score)
        unmatched_customers: List of tuples (customer_name, match_score)
    
    Returns:
        Filename of generated report
    """
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = f'billing_reconciliation_{timestamp}.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        
        # Summary sheet
        _create_summary_sheet(writer, results_df, unmatched_employees, unmatched_customers)
        
        # Discrepancies sheet
        _create_discrepancies_sheet(writer, results_df)
        
        # All records sheet
        _create_all_records_sheet(writer, results_df)
        
        # Unmatched employees sheet
        if unmatched_employees:
            _create_unmatched_employees_sheet(writer, unmatched_employees)
        
        # Unmatched customers sheet
        if unmatched_customers:
            _create_unmatched_customers_sheet(writer, unmatched_customers)
    
    return output_file


def _create_summary_sheet(writer, results_df, unmatched_employees, unmatched_customers):
    """Create summary sheet with overall statistics"""
    total_billed = results_df['amount_billed'].sum()
    total_expected = results_df['expected_amount'].sum()
    total_discrepancy = results_df['discrepancy'].sum()
    
    summary = pd.DataFrame({
        'Metric': [
            'Total Records',
            'Total Amount Billed',
            'Total Expected Amount',
            'Total Discrepancy',
            'Records with Discrepancies (>1%)',
            'Unmatched Employees',
            'Unmatched Customers'
        ],
        'Value': [
            len(results_df),
            f"{total_billed:,.2f}",
            f"{total_expected:,.2f}",
            f"{total_discrepancy:,.2f}",
            len(results_df[abs(results_df['discrepancy_pct']) > 1]),
            len(set(unmatched_employees)),
            len(set(unmatched_customers))
        ]
    })
    summary.to_excel(writer, sheet_name='Summary', index=False)


def _create_discrepancies_sheet(writer, results_df):
    """Create sheet with only discrepant records"""
    discrepancies = results_df[abs(results_df['discrepancy_pct']) > 1].copy()
    
    if len(discrepancies) == 0:
        # Create empty sheet with message
        pd.DataFrame({
            'Message': ['No discrepancies found (all within 1%)']
        }).to_excel(writer, sheet_name='Discrepancies', index=False)
        return
    
    disc_output = discrepancies[[
        'sql_customer', 'sql_employee', 'employee_title', 'normalized_rank',
        'hours', 'rate_charged', 'expected_rate', 'amount_billed',
        'expected_amount', 'discrepancy', 'discrepancy_pct'
    ]].sort_values('discrepancy', key=abs, ascending=False)
    
    disc_output.to_excel(writer, sheet_name='Discrepancies', index=False)
    
    # Highlight discrepancies in red
    worksheet = writer.sheets['Discrepancies']
    red_fill = PatternFill(start_color='FFB6C1', end_color='FFB6C1', fill_type='solid')
    
    for row_idx in range(2, len(disc_output) + 2):
        worksheet[f'J{row_idx}'].fill = red_fill  # discrepancy column
        worksheet[f'K{row_idx}'].fill = red_fill  # discrepancy_pct column


def _create_all_records_sheet(writer, results_df):
    """Create sheet with all records"""
    all_output = results_df[[
        'date', 'sql_customer', 'sql_employee', 'employee_title',
        'normalized_rank', 'price_rank_used', 'hours', 'rate_charged',
        'expected_rate', 'amount_billed', 'expected_amount',
        'discrepancy', 'discrepancy_pct'
    ]].sort_values('date')
    
    all_output.to_excel(writer, sheet_name='All Records', index=False)


def _create_unmatched_employees_sheet(writer, unmatched_employees):
    """Create sheet with unmatched employees"""
    unmatched_emp_df = pd.DataFrame(
        list(set(unmatched_employees)),
        columns=['Employee', 'Match Score']
    ).sort_values('Match Score')
    
    unmatched_emp_df.to_excel(writer, sheet_name='Unmatched Employees', index=False)


def _create_unmatched_customers_sheet(writer, unmatched_customers):
    """Create sheet with unmatched customers"""
    unmatched_cust_df = pd.DataFrame(
        list(set(unmatched_customers)),
        columns=['Customer', 'Match Score']
    ).sort_values('Match Score')
    
    unmatched_cust_df.to_excel(writer, sheet_name='Unmatched Customers', index=False)