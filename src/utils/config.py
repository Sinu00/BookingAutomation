"""
Configuration utility for Nusuk Automation Tool
Handles environment variables and configuration settings
"""

import os
from dotenv import load_dotenv
from typing import Optional

# Load environment variables from .env file
load_dotenv()

class Config:
    """Configuration class for Nusuk Automation"""
    
    # Google Sheets Configuration
    GOOGLE_SHEETS_ID = os.getenv('GOOGLE_SHEETS_ID')
    GOOGLE_CREDENTIALS_FILE = os.getenv('GOOGLE_CREDENTIALS_FILE', 'google_credentials.json')
    
    # Temporary Email Service
    TEMP_EMAIL_API_KEY = os.getenv('TEMP_EMAIL_API_KEY')
    TEMP_EMAIL_SERVICE = os.getenv('TEMP_EMAIL_SERVICE', 'mailtm')
    
    # WhatsApp Configuration
    WHATSAPP_API_TOKEN = os.getenv('WHATSAPP_API_TOKEN')
    WHATSAPP_PHONE_NUMBER = os.getenv('WHATSAPP_PHONE_NUMBER')
    WHATSAPP_GROUP_ID = os.getenv('WHATSAPP_GROUP_ID')
    
    # Nusuk Website Configuration
    NUSUK_BASE_URL = os.getenv('NUSUK_BASE_URL', 'https://services.nusuk.sa/nusuk-svc/auth/register')
    HEADLESS_MODE = os.getenv('HEADLESS_MODE', 'False').lower() == 'true'
    
    # Scheduling Configuration
    CHECK_INTERVAL_MINUTES = int(os.getenv('CHECK_INTERVAL_MINUTES', '30'))
    QR_RETRIEVAL_BUFFER_HOURS = int(os.getenv('QR_RETRIEVAL_BUFFER_HOURS', '2'))
    
    # Logging Configuration
    LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO')
    LOG_FILE = os.getenv('LOG_FILE', 'logs/nusuk_automation.log')
    
    # Automation Settings
    DELAY_BETWEEN_ACTIONS = int(os.getenv('DELAY_BETWEEN_ACTIONS', '2'))
    MAX_RETRIES = int(os.getenv('MAX_RETRIES', '3'))
    TIMEOUT_SECONDS = int(os.getenv('TIMEOUT_SECONDS', '30'))
    
    @classmethod
    def validate_config(cls) -> bool:
        """Validate that all required configuration is present"""
        required_fields = [
            'GOOGLE_SHEETS_ID',
            'GOOGLE_CREDENTIALS_FILE'
        ]
        
        missing_fields = []
        for field in required_fields:
            if not getattr(cls, field):
                missing_fields.append(field)
        
        if missing_fields:
            print(f"❌ Missing required configuration: {', '.join(missing_fields)}")
            print("Please check your .env file and ensure all required fields are set.")
            return False
        
        return True
    
    @classmethod
    def print_config(cls):
        """Print current configuration (without sensitive data)"""
        print("🔧 Current Configuration:")
        print(f"  Google Sheets ID: {cls.GOOGLE_SHEETS_ID[:20]}..." if cls.GOOGLE_SHEETS_ID else "Not set")
        print(f"  Credentials File: {cls.GOOGLE_CREDENTIALS_FILE}")
        print(f"  Temp Email Service: {cls.TEMP_EMAIL_SERVICE}")
        print(f"  Nusuk URL: {cls.NUSUK_BASE_URL}")
        print(f"  Headless Mode: {cls.HEADLESS_MODE}")
        print(f"  Check Interval: {cls.CHECK_INTERVAL_MINUTES} minutes")
        print(f"  QR Buffer Hours: {cls.QR_RETRIEVAL_BUFFER_HOURS}")
        print(f"  Log Level: {cls.LOG_LEVEL}")
        print(f"  Max Retries: {cls.MAX_RETRIES}")
        print(f"  Timeout: {cls.TIMEOUT_SECONDS} seconds")
