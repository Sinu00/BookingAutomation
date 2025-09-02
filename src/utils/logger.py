"""
Logging utility for Nusuk Automation Tool
Provides centralized logging functionality
"""

import logging
import os
from datetime import datetime
from typing import Optional
from .config import Config

class Logger:
    """Centralized logging for Nusuk Automation"""
    
    _instance = None
    _logger = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(Logger, cls).__new__(cls)
            cls._instance._setup_logger()
        return cls._instance
    
    def _setup_logger(self):
        """Setup the logger with proper configuration"""
        # Create logs directory if it doesn't exist
        log_dir = os.path.dirname(Config.LOG_FILE)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # Create logger
        self._logger = logging.getLogger('NusukAutomation')
        self._logger.setLevel(getattr(logging, Config.LOG_LEVEL.upper()))
        
        # Clear existing handlers
        self._logger.handlers.clear()
        
        # Create formatter
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # File handler
        file_handler = logging.FileHandler(Config.LOG_FILE, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        self._logger.addHandler(file_handler)
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)
        self._logger.addHandler(console_handler)
    
    def info(self, message: str):
        """Log info message"""
        self._logger.info(message)
    
    def error(self, message: str):
        """Log error message"""
        self._logger.error(message)
    
    def warning(self, message: str):
        """Log warning message"""
        self._logger.warning(message)
    
    def debug(self, message: str):
        """Log debug message"""
        self._logger.debug(message)
    
    def critical(self, message: str):
        """Log critical message"""
        self._logger.critical(message)
    
    def log_step(self, step: str, status: str = "STARTED", details: Optional[str] = None):
        """Log automation step with status"""
        message = f"STEP: {step} - {status}"
        if details:
            message += f" - {details}"
        self.info(message)
    
    def log_error_with_context(self, error: Exception, context: str):
        """Log error with context information"""
        self.error(f"ERROR in {context}: {str(error)}")
        self.debug(f"Error type: {type(error).__name__}")
    
    def log_success(self, operation: str, details: Optional[str] = None):
        """Log successful operation"""
        message = f"✅ SUCCESS: {operation}"
        if details:
            message += f" - {details}"
        self.info(message)
    
    def log_failure(self, operation: str, error: str):
        """Log failed operation"""
        self.error(f"❌ FAILED: {operation} - {error}")

# Global logger instance
logger = Logger()
