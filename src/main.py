#!/usr/bin/env python3
"""
Main script to run the Nusuk automation
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from automation.nusuk_bot import NusukBot
from utils.logger import logger

def main():
    """Main automation runner"""
    try:
        logger.log_step("Starting Nusuk Automation", "STARTED")
        
        # Create bot instance (non-headless for visibility)
        bot = NusukBot(headless=False)
        
        # Run the automation directly (no website access test)
        logger.info("Starting automation directly...")
        bot.run_automation()
        
    except Exception as e:
        logger.log_error_with_context(e, "Running main automation")
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()
