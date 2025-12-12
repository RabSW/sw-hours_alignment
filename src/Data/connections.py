"""
Database and SharePoint connection handlers.
"""

import os
from sqlalchemy import create_engine, text
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import pandas as pd


class SQLConnection:
    """Handles SQL Server database connections."""
    
    def __init__(self, server: str, database: str, username: str, password: str, driver: str):
        self.server = server
        self.database = database
        self.username = username
        self.password = password
        self.driver = driver
        self.engine = None
    
    def connect(self):
        """Establish connection to SQL Server."""
        connection_string = (
            f"mssql+pyodbc://{self.username}:{self.password}"
            f"@{self.server}/{self.database}?driver={self.driver}"
        )
        self.engine = create_engine(connection_string)
        print(f"Connected to SQL Server: {self.server}/{self.database}")
        return self.engine
    
    def execute_query(self, query: str) -> pd.DataFrame:
        """Execute a SQL query and return results as DataFrame."""
        if not self.engine:
            raise ConnectionError("Database not connected. Call connect() first.")
        
        with self.engine.connect() as connection:
            result = pd.read_sql(text(query), connection)
        return result
    
    def close(self):
        """Close the database connection."""
        if self.engine:
            self.engine.dispose()
            print("SQL connection closed")


class SharePointConnection:
    """Handles SharePoint file operations."""
    
    def __init__(self, site_url: str, username: str, password: str):
        self.site_url = site_url
        self.username = username
        self.password = password
        self.ctx = None
    
    def connect(self):
        """Establish connection to SharePoint."""
        credentials = UserCredential(self.username, self.password)
        self.ctx = ClientContext(self.site_url).with_credentials(credentials)
        print(f"Connected to SharePoint: {self.site_url}")
        return self.ctx
    
    def download_file(self, file_path: str, local_path: str):
        """Download a file from SharePoint."""
        if not self.ctx:
            raise ConnectionError("SharePoint not connected. Call connect() first.")
        
        file = self.ctx.web.get_file_by_server_relative_url(file_path)
        
        with open(local_path, "wb") as local_file:
            file.download(local_file).execute_query()
        
        print(f"Downloaded: {file_path} -> {local_path}")
    
    def upload_file(self, local_path: str, sharepoint_folder: str):
        """Upload a file to SharePoint."""
        if not self.ctx:
            raise ConnectionError("SharePoint not connected. Call connect() first.")
        
        with open(local_path, "rb") as content_file:
            file_content = content_file.read()
        
        target_folder = self.ctx.web.get_folder_by_server_relative_url(sharepoint_folder)
        file_name = os.path.basename(local_path)
        target_folder.upload_file(file_name, file_content).execute_query()
        
        print(f"Uploaded: {local_path} -> {sharepoint_folder}/{file_name}")

