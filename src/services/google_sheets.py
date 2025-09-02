"""
Google Sheets service for Nusuk Automation Tool
Handles reading from and writing to Google Sheets
"""

import os
from typing import List, Dict, Optional, Any
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from utils.config import Config
from utils.logger import logger

class GoogleSheetsService:
    """Service for interacting with Google Sheets"""
    
    def __init__(self):
        """Initialize the Google Sheets service"""
        try:
            # Load configuration
            self.sheet_id = Config.GOOGLE_SHEETS_ID
            self.credentials_file = Config.GOOGLE_CREDENTIALS_FILE
            
            if not self.sheet_id or not self.credentials_file:
                raise ValueError("Missing Google Sheets configuration")
            
            # Initialize the service
            self.service = self._initialize_service()
            logger.log_success("Google Sheets service initialized")
            
        except Exception as e:
            logger.log_error_with_context(e, "Initializing Google Sheets service")
            raise
    
    def _initialize_service(self):
        """Initialize Google Sheets API service"""
        try:
            # Load credentials
            credentials = Credentials.from_service_account_file(
                self.credentials_file,
                scopes=['https://www.googleapis.com/auth/spreadsheets']
            )
            
            # Build the service
            service = build('sheets', 'v4', credentials=credentials)
            return service
            
        except Exception as e:
            logger.log_error_with_context(e, "Building Google Sheets service")
            raise
    
    def read_sheet_data(self, range_name: str = "A:N") -> List[List[Any]]:
        """Read data from Google Sheet"""
        try:
            logger.log_step("Reading Google Sheet data", "STARTED", f"Range: {range_name}")
            
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.sheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            logger.log_success("Google Sheet data read", f"Retrieved {len(values)} rows")
            return values
            
        except HttpError as e:
            logger.log_error_with_context(e, "Google Sheets read operation")
            raise
    
    def get_pending_records(self) -> List[Dict[str, Any]]:
        """Get all pending records from the Sample tab"""
        try:
            logger.log_step("Fetching pending records", "STARTED")
            
            # Specify the Sample tab explicitly - use proper Google Sheets syntax
            range_name = "'Sample '!A:N"  # Note the single quotes around tab name WITH trailing space
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.sheet_id, range=range_name
            ).execute()
            
            values = result.get('values', [])
            if not values:
                logger.warning("No data found in Sample tab")
                return []
            
            # First row contains headers
            headers = values[0]
            logger.info(f"Found {len(headers)} columns in Sample tab")
            
            # Find STATUS column (case-insensitive)
            status_column_index = None
            for i, header in enumerate(headers):
                if header and 'status' in header.lower():
                    status_column_index = i
                    logger.log_success("STATUS column found", f"Column {i+1}: {header}")
                    break
            
            if status_column_index is None:
                logger.error("STATUS column not found in Sample tab")
                logger.info("Available columns: " + ", ".join([f"{i+1}:{header}" for i, header in enumerate(headers)]))
                return []
            
            # Find pending records
            pending_records = []
            for row_index, row in enumerate(values[1:], start=2):  # Start from row 2 (skip header)
                if len(row) > status_column_index:
                    status = row[status_column_index]
                    if status and status.lower() in ['pending', 'pending']:
                        # Create record with all available data
                        record_data = {}
                        for i, value in enumerate(row):
                            if i < len(headers):
                                record_data[headers[i]] = value
                            else:
                                record_data[f'Column_{i+1}'] = value
                        
                        record = {
                            'row_number': row_index,
                            'data': record_data
                        }
                        pending_records.append(record)
                        logger.info(f"Found pending record at row {row_index}")
            
            logger.log_success("Pending records fetched", f"Found {len(pending_records)} records")
            return pending_records
            
        except Exception as e:
            logger.log_error_with_context(e, "Fetching pending records")
            return []
    
    def update_record_status(self, row_number: int, status: str):
        """Update status for a specific record in the Sample tab"""
        try:
            # Status is in column H (8th column, 0-indexed)
            column_letter = 'H'
            range_name = f"'Sample '!{column_letter}{row_number}"
            
            body = {
                'values': [[status]]
            }
            
            result = self.service.spreadsheets().values().update(
                spreadsheetId=self.sheet_id,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
            
            logger.log_success("Status updated in Sample tab", f"Row {row_number}: {status}")
            return True
            
        except Exception as e:
            logger.log_error_with_context(e, f"Updating status for row {row_number}")
            return False
    
    def add_qr_code_url(self, row_number: int, qr_code_url: str):
        """Add QR code URL to the specified row"""
        try:
            logger.log_step("Adding QR code URL", "STARTED", f"Row {row_number}")
            
            # Find QR_CODE_URL column
            headers = self.read_sheet_data("A1:N1")[0]
            qr_col_idx = None
            for i, header in enumerate(headers):
                if 'QR_CODE_URL' in header.upper() or 'QR' in header.upper():
                    qr_col_idx = i
                    break
            
            if qr_col_idx is None:
                # If QR column doesn't exist, add it to column M
                qr_col_idx = 12  # Column M
            
            qr_range = f"{chr(65 + qr_col_idx)}{row_number}"
            
            body = {
                'valueInputOption': 'RAW',
                'data': [{
                    'range': qr_range,
                    'values': [[qr_code_url]]
                }]
            }
            
            self.service.spreadsheets().values().batchUpdate(
                spreadsheetId=self.sheet_id,
                body=body
            ).execute()
            
            logger.log_success("QR code URL added", f"Row {row_number}")
            
        except Exception as e:
            logger.log_error_with_context(e, f"Adding QR code URL for row {row_number}")
            raise
    
    def log_error(self, row_number: int, error_message: str):
        """Log error for a specific record in the Sample tab"""
        try:
            # Error log is in column N (14th column, 0-indexed)
            column_letter = 'N'
            range_name = f"'Sample '!{column_letter}{row_number}"
            
            body = {
                'values': [[error_message]]
            }
            
            result = self.service.spreadsheets().values().update(
                spreadsheetId=self.sheet_id,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
            
            logger.log_success("Error logged in Sample tab", f"Row {row_number}")
            return True
            
        except Exception as e:
            logger.log_error_with_context(e, f"Logging error for row {row_number}")
            return False
    
    def test_connection(self) -> bool:
        """Test Google Sheets connection"""
        try:
            logger.log_step("Testing Google Sheets connection", "STARTED")
            
            # Try to read first row
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.sheet_id,
                range="A1:N1"
            ).execute()
            
            values = result.get('values', [])
            if values:
                logger.log_success("Google Sheets connection test", f"Headers: {values[0]}")
                return True
            else:
                logger.warning("Google Sheets connection test - No data found")
                return False
                
        except Exception as e:
            logger.log_error_with_context(e, "Google Sheets connection test")
            return False

    def update_email_used(self, row_number: int, email: str):
        """Update email used for a specific record in the Sample tab"""
        try:
            # Email used is in column M (13th column, 0-indexed)
            column_letter = 'M'
            range_name = f"'Sample '!{column_letter}{row_number}"
            
            body = {
                'values': [[email]]
            }
            
            result = self.service.spreadsheets().values().update(
                spreadsheetId=self.sheet_id,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()
            
            logger.log_success("Email updated in Sample tab", f"Row {row_number}: {email}")
            return True
            
        except Exception as e:
            logger.log_error_with_context(e, f"Updating email for row {row_number}")
            return False
