"""
Temporary Email Service for Nusuk Automation Tool
Uses GuerrillaMail API for temporary email generation and OTP retrieval
"""

import requests
import time
import re
from typing import Optional, Dict, Any
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from utils.config import Config
from utils.logger import logger

class TempEmailService:
    """Service for temporary email operations using GuerrillaMail"""
    
    def __init__(self):
        self.session = requests.Session()
        self.email_address = None
        self.sid_token = None
        self.base_url = "https://api.guerrillamail.com/ajax"
        
    def generate_email(self) -> str:
        """Generate a new temporary email address"""
        try:
            logger.log_step("Generating temporary email", "STARTED")
            
            # Try alternative endpoint first
            try:
                response = self.session.get("https://api.guerrillamail.com/ajax/get_email_address")
                response.raise_for_status()
            except:
                # Fallback to main site
                response = self.session.get("https://www.guerrillamail.com/ajax/get_email_address")
                response.raise_for_status()
            
            data = response.json()
            self.email_address = data.get('email_addr')
            self.sid_token = data.get('sid_token')
            
            if not self.email_address:
                raise ValueError("Failed to generate email address")
            
            logger.log_success("Temporary email generated", f"Email: {self.email_address}")
            return self.email_address
            
        except Exception as e:
            logger.log_error_with_context(e, "Generating temporary email")
            raise
        

    
    def check_for_emails(self, max_wait_time: int = 60) -> list:
        """Check for new emails with timeout"""
        try:
            logger.log_step("Checking for emails", "STARTED")
            
            start_time = time.time()
            while time.time() - start_time < max_wait_time:
                # Get email list
                params = {
                    'f': 'check_email',
                    'sid_token': self.sid_token
                }
                
                response = self.session.get(f"{self.base_url}/check_email", params=params)
                response.raise_for_status()
                
                data = response.json()
                emails = data.get('list', [])
                
                if emails:
                    logger.log_success("Emails found", f"Found {len(emails)} emails")
                    return emails
                
                # Wait before next check
                time.sleep(5)
            
            logger.warning("No emails received within timeout period")
            return []
            
        except Exception as e:
            logger.log_error_with_context(e, "Checking for emails")
            raise
    
    def get_email_content(self, email_id: str) -> Optional[Dict[str, Any]]:
        """Get content of a specific email"""
        try:
            logger.log_step("Getting email content", "STARTED", f"Email ID: {email_id}")
            
            params = {
                'f': 'fetch_email',
                'email_id': email_id,
                'sid_token': self.sid_token
            }
            
            response = self.session.get(f"{self.base_url}/fetch_email", params=params)
            response.raise_for_status()
            
            data = response.json()
            
            if data.get('error'):
                logger.error(f"Error fetching email: {data.get('error')}")
                return None
            
            email_content = {
                'subject': data.get('mail_subject', ''),
                'body': data.get('mail_body', ''),
                'from': data.get('mail_from', ''),
                'timestamp': data.get('mail_timestamp', ''),
                'id': email_id
            }
            
            logger.log_success("Email content retrieved", f"Subject: {email_content['subject']}")
            return email_content
            
        except Exception as e:
            logger.log_error_with_context(e, f"Getting email content for ID {email_id}")
            return None
    
    def extract_otp_from_email(self, email_content: Dict[str, Any]) -> Optional[str]:
        """Extract OTP from email content"""
        try:
            logger.log_step("Extracting OTP from email", "STARTED")
            
            body = email_content.get('body', '')
            subject = email_content.get('subject', '')
            
            # Common OTP patterns
            otp_patterns = [
                r'\b\d{4,6}\b',  # 4-6 digit numbers
                r'OTP[:\s]*(\d{4,6})',  # OTP: 123456
                r'verification[:\s]*(\d{4,6})',  # verification: 123456
                r'code[:\s]*(\d{4,6})',  # code: 123456
                r'(\d{4,6})',  # Any 4-6 digit number
            ]
            
            # Search in body first
            for pattern in otp_patterns:
                matches = re.findall(pattern, body, re.IGNORECASE)
                if matches:
                    otp = matches[0] if isinstance(matches[0], str) else str(matches[0])
                    logger.log_success("OTP extracted from email body", f"OTP: {otp}")
                    return otp
            
            # Search in subject
            for pattern in otp_patterns:
                matches = re.findall(pattern, subject, re.IGNORECASE)
                if matches:
                    otp = matches[0] if isinstance(matches[0], str) else str(matches[0])
                    logger.log_success("OTP extracted from email subject", f"OTP: {otp}")
                    return otp
            
            logger.warning("No OTP found in email content")
            return None
            
        except Exception as e:
            logger.log_error_with_context(e, "Extracting OTP from email")
            return None
    
    def wait_for_otp_email(self, max_wait_time: int = 120) -> Optional[str]:
        """Wait for OTP email and extract the OTP"""
        try:
            logger.log_step("Waiting for OTP email", "STARTED", f"Max wait: {max_wait_time}s")
            
            start_time = time.time()
            while time.time() - start_time < max_wait_time:
                # Check for new emails
                emails = self.check_for_emails(max_wait_time=10)
                
                for email in emails:
                    email_id = email.get('mail_id')
                    if email_id:
                        # Get email content
                        email_content = self.get_email_content(email_id)
                        if email_content:
                            # Try to extract OTP
                            otp = self.extract_otp_from_email(email_content)
                            if otp:
                                logger.log_success("OTP email received and processed", f"OTP: {otp}")
                                return otp
                
                # Wait before next check
                time.sleep(10)
            
            logger.warning("OTP email not received within timeout period")
            return None
            
        except Exception as e:
            logger.log_error_with_context(e, "Waiting for OTP email")
            return None
    
    def get_email_info(self) -> Dict[str, Any]:
        """Get current email information"""
        return {
            'email_address': self.email_address,
            'sid_token': self.sid_token
        }
    
    def test_service(self) -> bool:
        """Test the temporary email service"""
        try:
            logger.log_step("Testing temporary email service", "STARTED")
            
            # Generate email
            email = self.generate_email()
            if not email:
                return False
            
            # Check if email is valid format
            if '@' not in email:
                logger.error("Generated email is not in valid format")
                return False
            
            logger.log_success("Temporary email service test", f"Generated email: {email}")
            return True
            
        except Exception as e:
            logger.log_error_with_context(e, "Temporary email service test")
            return False
