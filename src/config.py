"""
Configuration Module
Loads sensitive credentials from .env file and defines matching rules
"""

import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Database Configuration (from .env)
SQL_SERVER = os.getenv('SQL_SERVER')
SQL_DATABASE = os.getenv('SQL_DATABASE')
SQL_USERNAME = os.getenv('SQL_USERNAME')
SQL_PASSWORD = os.getenv('SQL_PASSWORD')
SQL_DRIVER = os.getenv('SQL_DRIVER', 'ODBC Driver 17 for SQL Server')

# SharePoint Configuration (from .env)
SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_SITE_URL')
SHAREPOINT_USERNAME = os.getenv('SHAREPOINT_USERNAME')
SHAREPOINT_PASSWORD = os.getenv('SHAREPOINT_PASSWORD')
SHAREPOINT_FILE_PATH = os.getenv('SHAREPOINT_FILE_PATH')

# Fuzzy matching thresholds (0-100, higher = stricter)
EMPLOYEE_MATCH_THRESHOLD = 80
CUSTOMER_MATCH_THRESHOLD = 85

# Title normalization mapping
TITLE_TO_RANK = {
    # Junior roles -> Junior Consultant pricing
    "junior": "Junior Consultant",
    "junior dev": "Junior Consultant",
    "junior ba": "Junior Consultant",
    
    # Consultant roles -> Consultant pricing
    "consultant": "Consultant",
    "dev": "Consultant",
    "support": "Consultant",  # Support uses Consultant pricing
    
    # Senior roles -> Senior Consultant pricing
    "senior": "Senior Consultant",
    "senior dev": "Senior Consultant",
    "senior ba": "Senior Consultant",
    "tl": "Senior Consultant",
    "senior ba/tl": "Senior Consultant",
    
    # Principal roles -> Principal Consultant pricing
    "principal": "Principal Consultant",
    "principal dev": "Principal Consultant",
    "principal ba": "Principal Consultant",
    
    # Data Scientist -> Data Scientist pricing
    "data scientist": "Data Scientist"
}