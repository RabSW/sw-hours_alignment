"""
Matching Module
Handles fuzzy matching for employees, customers, and title normalization
"""

from fuzzywuzzy import fuzz, process
import pandas as pd
import config


def normalize_title(title):
    """
    Normalize employee title to pricing rank
    
    Args:
        title: Employee title from SharePoint
    
    Returns:
        Normalized rank (e.g., "Senior Consultant")
    """
    title_lower = title.lower().strip()
    
    # Direct match first
    if title_lower in config.TITLE_TO_RANK:
        return config.TITLE_TO_RANK[title_lower]
    
    # Partial matching for complex titles
    for key, rank in config.TITLE_TO_RANK.items():
        if key in title_lower:
            return rank
    
    # Default to Consultant if no match
    return "Consultant"


def match_employee(sql_name, employees_df):
    """
    Match SQL employee name to SharePoint employee using fuzzy matching
    
    Handles formats like:
    - SQL: "SKA - Sam K. Andersen"
    - SharePoint: "Sam K- Andersen"
    
    Args:
        sql_name: Employee name from SQL (with initials)
        employees_df: DataFrame with SharePoint employees
    
    Returns:
        Tuple of (matched_name, title, match_score)
    """
    # Extract just the name part (after the dash and initials)
    if ' - ' in sql_name:
        name_part = sql_name.split(' - ', 1)[1].strip()
    else:
        name_part = sql_name.strip()
    
    # Remove dots from initials for better matching
    name_normalized = name_part.replace('.', '')
    
    # Find best match
    sharepoint_names = employees_df['name'].tolist()
    
    if not sharepoint_names:
        return None, None, 0
    
    match, score = process.extractOne(
        name_normalized,
        sharepoint_names,
        scorer=fuzz.token_sort_ratio
    )
    
    if score >= config.EMPLOYEE_MATCH_THRESHOLD:
        matched_row = employees_df[employees_df['name'] == match].iloc[0]
        return matched_row['name'], matched_row['title'], score
    
    return None, None, score


def match_customer(sql_customer, regular_df, fcc_df):
    """
    Match SQL customer name to SharePoint customer using fuzzy matching
    
    Args:
        sql_customer: Customer name from SQL
        regular_df: DataFrame with regular customer pricing
        fcc_df: DataFrame with FCC customer pricing
    
    Returns:
        Tuple of (customer_type, matched_name, match_score)
        customer_type is either 'FCC', 'Regular', or None
    """
    sql_customer_clean = sql_customer.strip().upper()
    
    # Try FCC customers first
    fcc_customers = fcc_df['customer'].str.strip().str.upper().tolist()
    if fcc_customers:
        match, score = process.extractOne(
            sql_customer_clean,
            fcc_customers,
            scorer=fuzz.ratio
        )
        if score >= config.CUSTOMER_MATCH_THRESHOLD:
            return 'FCC', match, score
    
    # Try regular customers
    regular_customers = regular_df['customer'].str.strip().str.upper().tolist()
    
    if not regular_customers:
        return None, None, 0
    
    match, score = process.extractOne(
        sql_customer_clean,
        regular_customers,
        scorer=fuzz.ratio
    )
    
    if score >= config.CUSTOMER_MATCH_THRESHOLD:
        return 'Regular', match, score
    
    return None, None, score


def get_price_with_fallback(pricing_row, rank):
    """
    Get price for rank, falling back to next available rank if 0 or missing
    
    Fallback order: Principal → Senior → Consultant → Junior
    
    Args:
        pricing_row: Row from pricing DataFrame
        rank: Requested rank (e.g., "Senior Consultant")
    
    Returns:
        Tuple of (price, rank_used)
    """
    rank_order = [
        'Principal Consultant',
        'Senior Consultant',
        'Consultant',
        'Junior Consultant'
    ]
    
    # Start from requested rank
    if rank in rank_order:
        start_idx = rank_order.index(rank)
    else:
        start_idx = 2  # Default to Consultant
    
    for check_rank in rank_order[start_idx:]:
        if check_rank in pricing_row:
            price = pricing_row[check_rank]
            if pd.notna(price) and price > 0:
                return price, check_rank
    
    return None, None