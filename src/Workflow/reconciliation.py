"""
Reconciliation Module
Performs the actual reconciliation between SQL data and SharePoint pricing
"""

import pandas as pd
from Workflow.matching import match_employee, match_customer, normalize_title, get_price_with_fallback


def reconcile_data(sql_df, employees_df, regular_df, fcc_df):
    """
    Reconcile SQL billable data against SharePoint pricing
    
    Args:
        sql_df: DataFrame with SQL billable data
        employees_df: DataFrame with SharePoint employees
        regular_df: DataFrame with regular customer pricing
        fcc_df: DataFrame with FCC customer pricing
    
    Returns:
        Tuple of (results_df, unmatched_employees, unmatched_customers)
    """
    results = []
    unmatched_employees = []
    unmatched_customers = []
    
    for idx, row in sql_df.iterrows():
        result = {
            'sql_customer': row['CustomerName'],
            'sql_employee': row['EmployeeName'],
            'date': row['Date'],
            'hours': row['Hours'],
            'rate_charged': row['BillableRate'],
            'amount_billed': row['BillableAmount']
        }
        
        # Match employee
        sp_employee, sp_title, emp_score = match_employee(
            row['EmployeeName'],
            employees_df
        )
        
        if sp_employee:
            result['matched_employee'] = sp_employee
            result['employee_title'] = sp_title
            result['employee_match_score'] = emp_score
            
            # Normalize title to rank
            normalized_rank = normalize_title(sp_title)
            result['normalized_rank'] = normalized_rank
        else:
            unmatched_employees.append((row['EmployeeName'], emp_score))
            result['matched_employee'] = None
            result['employee_title'] = None
            result['normalized_rank'] = 'Consultant'  # Default fallback
        
        # Match customer
        customer_type, sp_customer, cust_score = match_customer(
            row['CustomerName'],
            regular_df,
            fcc_df
        )
        
        if customer_type:
            result['customer_type'] = customer_type
            result['matched_customer'] = sp_customer
            result['customer_match_score'] = cust_score
            
            # Get pricing
            if customer_type == 'FCC':
                # FCC uses Consultant price for all ranks
                pricing_row = fcc_df[
                    fcc_df['customer'].str.strip().str.upper() == sp_customer
                ].iloc[0]
                expected_rate = pricing_row['Consultant']
                result['expected_rate'] = expected_rate
                result['price_rank_used'] = 'Consultant (FCC)'
            else:
                # Regular customer - get price for rank with fallback
                pricing_row = regular_df[
                    regular_df['customer'].str.strip().str.upper() == sp_customer
                ].iloc[0]
                expected_rate, rank_used = get_price_with_fallback(
                    pricing_row,
                    result['normalized_rank']
                )
                result['expected_rate'] = expected_rate
                result['price_rank_used'] = rank_used or result['normalized_rank']
        else:
            unmatched_customers.append((row['CustomerName'], cust_score))
            result['matched_customer'] = None
            result['expected_rate'] = None
        
        # Calculate discrepancy
        if result['expected_rate']:
            result['expected_amount'] = result['hours'] * result['expected_rate']
            result['discrepancy'] = result['amount_billed'] - result['expected_amount']
            
            if result['expected_amount'] != 0:
                result['discrepancy_pct'] = (
                    result['discrepancy'] / result['expected_amount'] * 100
                )
            else:
                result['discrepancy_pct'] = 0
        else:
            result['expected_amount'] = None
            result['discrepancy'] = None
            result['discrepancy_pct'] = None
        
        results.append(result)
    
    return pd.DataFrame(results), unmatched_employees, unmatched_customers