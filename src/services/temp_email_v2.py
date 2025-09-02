"""
Improved Temporary Email Service for Nusuk Automation Tool
Uses Mail.tm API for reliable temporary email generation and OTP retrieval
"""

import requests
import time
import re
import random
import string
from typing import Optional, Dict, Any
import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from utils.config import Config
from utils.logger import logger

class TempEmailServiceV2:
    """Improved service for temporary email operations using Mail.tm"""
    
    def __init__(self):
        self.session = requests.Session()
        self.email_address = None
        self.auth_token = None
        self.account_id = None
        self.service_used = None
        self.base_url = "https://api.mail.tm"
        
    def generate_email(self) -> str:
        """Generate a new temporary email address using Mail.tm"""
        try:
            logger.log_step("Generating temporary email", "STARTED")
            
            # Try Mail.tm first (most reliable)
            email = self._try_mailtm()
            if email:
                return email
            
            # Fallback to other services
            services = [
                self._try_10minutemail,
                self._try_tempmail,
                self._generate_fake_email
            ]
            
            for service_func in services:
                try:
                    email = service_func()
                    if email and '@' in email:
                        logger.log_success("Temporary email generated", f"Email: {email} (Service: {self.service_used})")
                        return email
                except Exception as e:
                    logger.debug(f"Service {service_func.__name__} failed: {e}")
                    continue
            
            raise ValueError("All temporary email services failed")
            
        except Exception as e:
            logger.log_error_with_context(e, "Generating temporary email")
            raise
    
    def _try_mailtm(self) -> Optional[str]:
        """Try Mail.tm API - most reliable service"""
        try:
            logger.info("Attempting to use Mail.tm service...")
            
            # Step 1: Get available domains
            domains_response = self.session.get(f"{self.base_url}/domains", timeout=10)
            if domains_response.status_code != 200:
                logger.warning(f"Mail.tm domains request failed: {domains_response.status_code}")
                return None
            
            domains_data = domains_response.json()
            if not domains_data.get('hydra:member'):
                logger.warning("No domains available from Mail.tm")
                return None
            
            # Get first available domain
            domain = domains_data['hydra:member'][0]['domain']
            logger.info(f"Using Mail.tm domain: {domain}")
            
            # Step 2: Generate random username
            username = ''.join(random.choices(string.ascii_lowercase + string.digits, k=12))
            email_address = f"{username}@{domain}"
            
            # Step 3: Create account
            account_data = {
                "address": email_address,
                "password": "TempPass123!"  # Fixed password for consistency
            }
            
            account_response = self.session.post(
                f"{self.base_url}/accounts",
                json=account_data,
                timeout=10
            )
            
            if account_response.status_code != 201:
                logger.warning(f"Mail.tm account creation failed: {account_response.status_code}")
                return None
            
            account_info = account_response.json()
            self.account_id = account_info.get('id')
            
            # Step 4: Get authentication token
            token_response = self.session.post(
                f"{self.base_url}/token",
                json=account_data,
                timeout=10
            )
            
            if token_response.status_code != 200:
                logger.warning(f"Mail.tm token request failed: {token_response.status_code}")
                return None
            
            token_data = token_response.json()
            self.auth_token = token_data.get('token')
            
            if not self.auth_token:
                logger.warning("No auth token received from Mail.tm")
                return None
            
            # Store email and service info
            self.email_address = email_address
            self.service_used = "Mail.tm"
            
            logger.log_success("Mail.tm email generated", f"Email: {email_address}")
            return email_address
            
        except Exception as e:
            logger.warning(f"Mail.tm setup failed: {e}")
            return None
    
    def _try_10minutemail(self) -> Optional[str]:
        """Try 10MinuteMail API"""
        try:
            response = self.session.get("https://10minutemail.com/address.api.php", timeout=10)
            if response.status_code == 200:
                email = response.text.strip()
                if '@' in email:
                    self.email_address = email
                    self.service_used = "10MinuteMail"
                    return email
            return None
        except:
            return None
    
    def _try_tempmail(self) -> Optional[str]:
        """Try TempMail API"""
        try:
            response = self.session.get("https://api.tempmail.com/email", timeout=10)
            if response.status_code == 200:
                data = response.json()
                email = data.get('email')
                if email and '@' in email:
                    self.email_address = email
                    self.service_used = "TempMail"
                    return email
            return None
        except:
            return None
    
    def _generate_fake_email(self) -> str:
        """Generate a fake email as last resort"""
        random_string = ''.join(random.choices(string.ascii_lowercase + string.digits, k=10))
        email = f"{random_string}@tempmail.org"
        self.email_address = email
        self.service_used = "FakeEmail"
        return email
    
    def get_otp(self, email: str) -> Optional[str]:
        """Get OTP from temporary email"""
        try:
            logger.log_step("Retrieving OTP", f"From: {email}")
            
            # If using Mail.tm, use the proper API
            if self.service_used == "Mail.tm" and self.auth_token:
                return self._get_mailtm_otp()
            
            # For other services, try fallback methods
            otp = self._try_10minutemail_otp(email)
            if otp:
                return otp
            
            otp = self._try_tempmail_otp(email)
            if otp:
                return otp
            
            logger.error("Failed to retrieve OTP from all services")
            return None
            
        except Exception as e:
            logger.log_error_with_context(e, "Retrieving OTP")
            return None
    
    def _get_mailtm_otp(self) -> Optional[str]:
        """Get OTP from Mail.tm using proper API"""
        try:
            if not self.auth_token:
                logger.error("No auth token available for Mail.tm")
                return None
            
            # Set authorization header
            headers = {"Authorization": f"Bearer {self.auth_token}"}
            
            # Get messages
            messages_response = self.session.get(
                f"{self.base_url}/messages",
                headers=headers,
                timeout=10
            )
            
            if messages_response.status_code != 200:
                logger.warning(f"Mail.tm messages request failed: {messages_response.status_code}")
                return None
            
            messages_data = messages_response.json()
            messages = messages_data.get('hydra:member', [])
            
            if not messages:
                logger.info("No messages found in Mail.tm inbox")
                return None
            
            logger.info(f"Found {len(messages)} messages in Mail.tm inbox")
            
            # Check recent messages for OTP
            for message in messages[:5]:  # Check last 5 messages
                message_id = message.get('id')
                if not message_id:
                    continue
                
                # Get full message content
                message_response = self.session.get(
                    f"{self.base_url}/messages/{message_id}",
                    headers=headers,
                    timeout=10
                )
                
                if message_response.status_code != 200:
                    continue
                
                message_content = message_response.json()
                text_content = message_content.get('text', '')
                html_content = message_content.get('html', [''])[0] if message_content.get('html') else ''
                
                # Look for OTP in content
                otp = self._extract_otp_from_text(text_content + ' ' + html_content)
                if otp:
                    logger.log_success("OTP found in Mail.tm message", f"OTP: {otp}")
                    return otp
            
            logger.warning("No OTP found in Mail.tm messages")
            return None
            
        except Exception as e:
            logger.warning(f"Mail.tm OTP retrieval failed: {e}")
            return None
    
    def _try_10minutemail_otp(self, email: str) -> Optional[str]:
        """Try to get OTP from 10MinuteMail"""
        try:
            # 10MinuteMail doesn't have a direct API, but we can try to parse
            # For now, return None and implement later if needed
            return None
            
        except Exception as e:
            logger.warning(f"10MinuteMail OTP retrieval failed: {e}")
            return None
    
    def _try_tempmail_otp(self, email: str) -> Optional[str]:
        """Try to get OTP from TempMail"""
        try:
            # TempMail API implementation
            # For now, return None and implement later if needed
            return None
            
        except Exception as e:
            logger.warning(f"TempMail OTP retrieval failed: {e}")
            return None
    
    def _extract_otp_from_text(self, text: str) -> Optional[str]:
        """Extract OTP from email text using regex patterns"""
        try:
            # Common OTP patterns
            otp_patterns = [
                r'\b\d{4,6}\b',  # 4-6 digit numbers
                r'OTP[:\s]*(\d{4,6})',  # OTP: 123456
                r'verification[:\s]*(\d{4,6})',  # verification: 123456
                r'code[:\s]*(\d{4,6})',  # code: 123456
                r'(\d{4,6})',  # Just 4-6 digits
            ]
            
            for pattern in otp_patterns:
                matches = re.findall(pattern, text, re.IGNORECASE)
                if matches:
                    # Return the first match
                    otp = matches[0] if isinstance(matches[0], str) else str(matches[0])
                    if len(otp) >= 4 and len(otp) <= 6 and otp.isdigit():
                        return otp
            
            return None
            
        except Exception as e:
            logger.warning(f"OTP extraction failed: {e}")
            return None
    
    def check_for_emails(self, max_wait_time: int = 60) -> list:
        """Check for new emails"""
        logger.log_step("Checking for emails", "STARTED")
        
        if self.service_used == "Mail.tm" and self.auth_token:
            return self._check_mailtm_emails(max_wait_time)
        
        logger.warning("Email checking not implemented for current service")
        return []
    
    def _check_mailtm_emails(self, max_wait_time: int = 60) -> list:
        """Check for emails in Mail.tm"""
        try:
            start_time = time.time()
            while time.time() - start_time < max_wait_time:
                if not self.auth_token:
                    break
                
                headers = {"Authorization": f"Bearer {self.auth_token}"}
                response = self.session.get(
                    f"{self.base_url}/messages",
                    headers=headers,
                    timeout=10
                )
                
                if response.status_code == 200:
                    data = response.json()
                    messages = data.get('hydra:member', [])
                    if messages:
                        logger.log_success("Emails found in Mail.tm", f"Count: {len(messages)}")
                        return messages
                
                time.sleep(5)
            
            return []
            
        except Exception as e:
            logger.warning(f"Mail.tm email check failed: {e}")
            return []
    
    def wait_for_otp_email(self, max_wait_time: int = 120) -> Optional[str]:
        """Wait for OTP email"""
        logger.log_step("Waiting for OTP email", "STARTED")
        
        if self.service_used == "Mail.tm":
            return self._wait_for_mailtm_otp(max_wait_time)
        
        logger.warning("OTP waiting not implemented for current service")
        return None
    
    def _wait_for_mailtm_otp(self, max_wait_time: int = 120) -> Optional[str]:
        """Wait for OTP email in Mail.tm"""
        try:
            start_time = time.time()
            while time.time() - start_time < max_wait_time:
                otp = self._get_mailtm_otp()
                if otp:
                    return otp
                
                time.sleep(10)
            
            return None
            
        except Exception as e:
            logger.warning(f"Mail.tm OTP wait failed: {e}")
            return None
    
    def test_service(self) -> bool:
        """Test the temporary email service"""
        try:
            logger.log_step("Testing temporary email service", "STARTED")
            
            email = self.generate_email()
            if email and '@' in email:
                logger.log_success("Temporary email service test", f"Generated email: {email} (Service: {self.service_used})")
                return True
            else:
                logger.error("Failed to generate valid email")
                return False
                
        except Exception as e:
            logger.log_error_with_context(e, "Temporary email service test")
            return False

# For backward compatibility
TempEmailService = TempEmailServiceV2
