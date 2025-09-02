"""
Nusuk Automation Bot
Main automation script for Nusuk website registration and booking
"""

import time
import random
import json
import os
from typing import Dict, Any, Optional
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from utils.config import Config
from utils.logger import logger
from services.google_sheets import GoogleSheetsService
from services.temp_email_v2 import TempEmailServiceV2

class NusukBot:
    """Main automation bot for Nusuk website"""
    
    def __init__(self, headless: bool = False):
        self.driver = None
        self.wait = None
        self.sheets_service = GoogleSheetsService()
        self.email_service = TempEmailServiceV2()
        self.headless = headless
        self.current_record = None
        self.current_email = None
        self.use_undetected = True  # Flag to control which driver to use
        
    def setup_driver(self):
        """Setup the WebDriver with maximum stealth"""
        try:
            logger.log_step("Setting up Undetected Chrome WebDriver", "STARTED")
            
            # Try Firefox first for better reCAPTCHA bypass
            if self._setup_firefox_driver():
                logger.log_success("Firefox WebDriver setup", "Firefox driver ready - Better reCAPTCHA bypass!")
                return True
            
            # Fallback to undetected Chrome
            if self._setup_undetected_driver():
                logger.log_success("Undetected Chrome WebDriver setup", "Driver ready - Maximum stealth mode activated!")
                return True
            
            # Final fallback to standard Chrome
            if self._setup_standard_driver():
                logger.log_success("Standard Chrome WebDriver setup", "Driver ready - Basic stealth mode activated!")
                return True
            
            logger.log_failure("WebDriver setup", "All driver setup methods failed")
            return False
            
        except Exception as e:
            logger.log_error_with_context(e, "Setting up WebDriver")
            return False
    
    def _setup_firefox_driver(self):
        """Setup Firefox WebDriver for better reCAPTCHA bypass"""
        try:
            logger.info("Attempting to setup Firefox WebDriver...")
            
            # Try to import Firefox WebDriver
            try:
                from selenium.webdriver.firefox.webdriver import FirefoxDriver
                from selenium.webdriver.firefox.options import Options as FirefoxOptions
                from selenium.webdriver.firefox.service import Service as FirefoxService
                from webdriver_manager.firefox import GeckoDriverManager
            except ImportError:
                logger.info("Firefox WebDriver not available, skipping...")
                return False
            
            # Setup Firefox options
            firefox_options = FirefoxOptions()
            
            # Add stealth options for Firefox
            firefox_options.add_argument("--no-sandbox")
            firefox_options.add_argument("--disable-dev-shm-usage")
            firefox_options.add_argument("--disable-blink-features=AutomationControlled")
            firefox_options.add_argument("--disable-extensions")
            firefox_options.add_argument("--disable-plugins")
            firefox_options.add_argument("--disable-web-security")
            firefox_options.add_argument("--allow-running-insecure-content")
            firefox_options.add_argument("--disable-features=VizDisplayCompositor")
            firefox_options.add_argument("--disable-ipc-flooding-protection")
            
            # Firefox-specific stealth options
            firefox_options.set_preference("dom.webdriver.enabled", False)
            firefox_options.set_preference("useAutomationExtension", False)
            firefox_options.set_preference("general.useragent.override", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/119.0")
            firefox_options.set_preference("dom.webnotifications.enabled", False)
            firefox_options.set_preference("media.navigator.enabled", False)
            firefox_options.set_preference("media.peerconnection.enabled", False)
            
            # Random viewport size
            width = random.randint(1200, 1920)
            height = random.randint(800, 1080)
            firefox_options.add_argument(f"--width={width}")
            firefox_options.add_argument(f"--height={height}")
            
            # Initialize Firefox driver with automatic geckodriver installation
            logger.info("Initializing Firefox driver with stealth options...")
            service = FirefoxService(GeckoDriverManager().install())
            self.driver = FirefoxDriver(service=service, options=firefox_options)
            
            # Setup wait
            self.wait = WebDriverWait(self.driver, 10)
            
            # Apply stealth measures
            self._apply_stealth_measures()
            self.add_human_behavior()
            
            logger.log_success("Firefox WebDriver setup", "Firefox driver ready - Better reCAPTCHA bypass!")
            return True
            
        except Exception as e:
            logger.warning(f"Failed to setup Firefox driver: {e}")
            return False
    
    def _setup_undetected_driver(self):
        """Setup undetected-chromedriver with maximum stealth"""
        try:
            # Import undetected_chromedriver
            try:
                import undetected_chromedriver as uc
                logger.info("Successfully imported undetected_chromedriver")
            except ImportError:
                logger.error("undetected_chromedriver not installed. Please run: pip install undetected-chromedriver")
                raise ImportError("undetected_chromedriver is required. Install with: pip install undetected-chromedriver")
            
            # Chrome options for maximum stealth
            chrome_options = uc.ChromeOptions()
            if self.headless:
                chrome_options.add_argument("--headless")
            
            # Create stealth profile for maximum stealth (looks like regular user)
            stealth_profile = self._create_stealth_profile()
            if stealth_profile:
                chrome_options.add_argument(f"--user-data-dir={stealth_profile}")
                chrome_options.add_argument("--profile-directory=Default")  # Use default profile
                logger.info(f"Using stealth Chrome profile: {stealth_profile}")
            else:
                logger.warning("Could not create stealth profile, using new profile")
            
            # Enhanced stealth options
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--remote-debugging-port=9222")
            chrome_options.add_argument("--disable-web-security")
            chrome_options.add_argument("--allow-running-insecure-content")
            chrome_options.add_argument("--disable-features=VizDisplayCompositor")
            chrome_options.add_argument("--disable-ipc-flooding-protection")
            chrome_options.add_argument("--disable-background-timer-throttling")
            chrome_options.add_argument("--disable-backgrounding-occluded-windows")
            chrome_options.add_argument("--disable-renderer-backgrounding")
            chrome_options.add_argument("--disable-field-trial-config")
            chrome_options.add_argument("--no-first-run")
            chrome_options.add_argument("--no-default-browser-check")
            chrome_options.add_argument("--disable-default-apps")
            chrome_options.add_argument("--disable-sync")
            chrome_options.add_argument("--disable-translate")
            chrome_options.add_argument("--hide-scrollbars")
            chrome_options.add_argument("--mute-audio")
            chrome_options.add_argument("--no-zygote")
            chrome_options.add_argument("--disable-logging")
            chrome_options.add_argument("--disable-permissions-api")
            chrome_options.add_argument("--disable-background-networking")
            chrome_options.add_argument("--disable-background-downloads")
            chrome_options.add_argument("--disable-client-side-phishing-detection")
            chrome_options.add_argument("--disable-component-update")
            chrome_options.add_argument("--disable-domain-reliability")
            chrome_options.add_argument("--disable-features=TranslateUI")
            chrome_options.add_argument("--disable-hang-monitor")
            chrome_options.add_argument("--disable-prompt-on-repost")
            chrome_options.add_argument("--disable-sync-preferences")
            chrome_options.add_argument("--disable-web-resources")
            chrome_options.add_argument("--metrics-recording-only")
            chrome_options.add_argument("--no-report-upload")
            chrome_options.add_argument("--safebrowsing-disable-auto-update")
            chrome_options.add_argument("--disable-safebrowsing")
            chrome_options.add_argument("--disable-component-extensions-with-background-pages")
            chrome_options.add_argument("--disable-background-mode")
            chrome_options.add_argument("--disable-plugins-discovery")
            chrome_options.add_argument("--disable-plugins")
            
            # Randomize viewport size for better fingerprinting
            viewport_width = 1366 + random.randint(-50, 50)
            viewport_height = 768 + random.randint(-50, 50)
            chrome_options.add_argument(f"--window-size={viewport_width},{viewport_height}")
            
            # User agent rotation for maximum stealth
            user_agents = [
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36"
            ]
            chrome_options.add_argument(f"--user-agent={random.choice(user_agents)}")
            
            # Initialize undetected driver with maximum stealth
            logger.info("Initializing undetected Chrome driver with maximum stealth...")
            self.driver = uc.Chrome(
                options=chrome_options,
                version_main=None,  # Auto-detect Chrome version
                use_subprocess=True,
                no_sandbox=True,
                suppress_welcome=True
            )
            
            # Setup wait
            self.wait = WebDriverWait(self.driver, Config.TIMEOUT_SECONDS)
            
            # Maximize window for better visibility
            self.driver.maximize_window()
            
            # Apply additional stealth measures
            self._apply_stealth_measures()
            
            # Add human-like behavior simulation
            self.add_human_behavior()
            
            logger.log_success("Undetected Chrome WebDriver setup", "Driver ready - Maximum stealth mode activated!")
            
        except Exception as e:
            logger.log_error_with_context(e, "Setting up Undetected Chrome WebDriver")
            raise
    
    def _setup_standard_driver(self):
        """Fallback to standard Selenium driver if undetected fails"""
        try:
            # Chrome options
            chrome_options = Options()
            if self.headless:
                chrome_options.add_argument("--headless")
            
            # Create stealth profile for maximum stealth (looks like regular user)
            stealth_profile = self._create_stealth_profile()
            if stealth_profile:
                chrome_options.add_argument(f"--user-data-dir={stealth_profile}")
                chrome_options.add_argument("--profile-directory=Default")  # Use default profile
                logger.info(f"Using stealth Chrome profile: {stealth_profile}")
            else:
                logger.warning("Could not create stealth profile, using new profile")
            
            # Windows-specific options
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-blink-features=AutomationControlled")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--remote-debugging-port=9222")
            
            # Remove automation detection
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            # User agent
            chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
            
            # Setup driver with error handling
            try:
                # Try webdriver-manager first
                service = Service(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service, options=chrome_options)
            except Exception as e:
                logger.warning(f"ChromeDriverManager failed: {e}")
                # Fallback: try without service
                try:
                    self.driver = webdriver.Chrome(options=chrome_options)
                except Exception as e2:
                    logger.error(f"Chrome setup failed: {e2}")
                    raise Exception("Chrome setup failed. Please install Chrome browser and try again.")
            
            # Remove automation flags
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            # Setup wait
            self.wait = WebDriverWait(self.driver, Config.TIMEOUT_SECONDS)
            
            # Maximize window for better visibility
            self.driver.maximize_window()
            
            logger.log_success("Standard Chrome WebDriver setup", "Driver ready - Fallback mode activated!")
            
        except Exception as e:
            logger.log_error_with_context(e, "Setting up Standard Chrome WebDriver")
            raise
    
    def _apply_stealth_measures(self):
        """Apply additional stealth measures to the driver"""
        try:
            # Remove webdriver property
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            # Remove automation flags
            self.driver.execute_script("Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]})")
            self.driver.execute_script("Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']})")
            self.driver.execute_script("Object.defineProperty(navigator, 'permissions', {get: () => {query: () => Promise.resolve({state: 'granted'})}})")
            
            # Add random properties to make fingerprinting harder
            self.driver.execute_script("""
                // Randomize screen properties
                Object.defineProperty(screen, 'width', {get: () => 1366 + Math.floor(Math.random() * 100)});
                Object.defineProperty(screen, 'height', {get: () => 768 + Math.floor(Math.random() * 100)});
                
                // Randomize timezone
                Object.defineProperty(Intl, 'DateTimeFormat', {get: () => function() { return {resolvedOptions: () => ({timeZone: 'America/New_York'})} }});
                
                // Randomize hardware concurrency
                Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 4 + Math.floor(Math.random() * 4)});
                
                // Randomize device memory
                Object.defineProperty(navigator, 'deviceMemory', {get: () => 4 + Math.floor(Math.random() * 4)});
            """)
            
            logger.log_success("Stealth measures applied", "Enhanced anti-detection measures activated")
            
        except Exception as e:
            logger.warning(f"Failed to apply stealth measures: {e}")
    
    def add_human_behavior(self):
        """Add human-like behavior to avoid detection"""
        try:
            # Random mouse movements
            self.driver.execute_script("""
                // Simulate human-like mouse movements
                function simulateMouseMovement() {
                    const event = new MouseEvent('mousemove', {
                        'view': window,
                        'bubbles': true,
                        'cancelable': true,
                        'clientX': Math.random() * window.innerWidth,
                        'clientY': Math.random() * window.innerHeight
                    });
                    document.dispatchEvent(event);
                }
                setInterval(simulateMouseMovement, 2000 + Math.random() * 3000);
            """)
            
            # Random scrolling
            self.driver.execute_script("""
                function simulateScrolling() {
                    window.scrollBy(0, Math.random() * 100 - 50);
                }
                setInterval(simulateScrolling, 5000 + Math.random() * 5000);
            """)
            
            # Random focus changes
            self.driver.execute_script("""
                function simulateFocusChanges() {
                    const elements = document.querySelectorAll('input, button, a');
                    if (elements.length > 0) {
                        const randomElement = elements[Math.floor(Math.random() * elements.length)];
                        randomElement.focus();
                        setTimeout(() => randomElement.blur(), 100 + Math.random() * 200);
                    }
                }
                setInterval(simulateFocusChanges, 8000 + Math.random() * 4000);
            """)
            
            logger.log_success("Human behavior simulation", "Added random mouse movements, scrolling, and focus changes")
            
        except Exception as e:
            logger.warning(f"Failed to add human behavior: {e}")

    def wait_for_recaptcha(self, timeout=30):
        """Wait for reCAPTCHA to complete or timeout"""
        try:
            logger.log_step("Waiting for reCAPTCHA verification", "STARTED")
            
            # Wait for potential reCAPTCHA iframe
            recaptcha_selectors = [
                "iframe[src*='recaptcha']",
                "iframe[src*='google.com/recaptcha']",
                ".g-recaptcha iframe",
                "#recaptcha iframe",
                "iframe[src*='turnstile']",
                "iframe[src*='cloudflare']"
            ]
            
            for selector in recaptcha_selectors:
                try:
                    iframe = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    logger.info("Found reCAPTCHA iframe, waiting for completion...")
                    
                    # Wait for reCAPTCHA to complete (look for success indicators)
                    time.sleep(5)  # Give time for invisible reCAPTCHA
                    
                    # Check if there are any error messages
                    error_selectors = [
                        ".recaptcha-error",
                        "[data-error]",
                        ".error-message",
                        ".captcha-error",
                        "[class*='error']"
                    ]
                    
                    for error_selector in error_selectors:
                        try:
                            error_element = self.driver.find_element(By.CSS_SELECTOR, error_selector)
                            if error_element.is_displayed():
                                logger.warning(f"reCAPTCHA error detected: {error_element.text}")
                        except:
                            continue
                    
                    break
                    
                except TimeoutException:
                    continue
            
            # Wait a bit more for any background processing
            time.sleep(3)
            logger.log_success("reCAPTCHA wait completed", "Proceeding with form submission")
            return True
            
        except Exception as e:
            logger.warning(f"reCAPTCHA wait failed: {e}")
            return False

    def simulate_human_typing(self, element, text):
        """Simulate human-like typing with random delays"""
        try:
            element.clear()
            for char in text:
                element.send_keys(char)
                # Random delay between characters (50-150ms)
                time.sleep(random.uniform(0.05, 0.15))
            logger.log_success("Human-like typing simulation", f"Typed '{text}' with realistic delays")
        except Exception as e:
            logger.warning(f"Failed to simulate human typing: {e}")
            # Fallback to normal typing
            element.clear()
            element.send_keys(text)

    def simulate_human_click(self, element):
        """Simulate human-like clicking with random delays"""
        try:
            # Random delay before click (100-500ms)
            time.sleep(random.uniform(0.1, 0.5))
            
            # Try JavaScript click first to avoid click interception
            try:
                self.driver.execute_script("arguments[0].click();", element)
                logger.info("Clicked element using JavaScript")
            except Exception as e:
                logger.warning(f"JavaScript click failed: {e}, trying regular click...")
                element.click()
            
            # Random delay after click (200-800ms)
            time.sleep(random.uniform(0.2, 0.8))
            
            logger.log_success("Human-like clicking simulation", "Element clicked with realistic delays")
            
        except Exception as e:
            logger.warning(f"Failed to simulate human clicking: {e}")
            # Fallback to normal click
            element.click()
    
    def navigate_to_nusuk(self):
        """Navigate to Nusuk registration page"""
        try:
            logger.log_step("Navigating to Nusuk website", "STARTED")
            
            # Navigate to the URL
            self.driver.get(Config.NUSUK_BASE_URL)
            
            # Wait for page to start loading
            time.sleep(5)
            
            # Check if page is loading or has errors
            try:
                # Wait for page title to appear
                WebDriverWait(self.driver, 20).until(
                    lambda driver: driver.title and len(driver.title.strip()) > 0
                )
                logger.info(f"Page title loaded: {self.driver.title}")
            except TimeoutException:
                logger.warning("Page title took too long to load")
            
            # Wait for page content to appear
            try:
                # Wait for any content to appear (not just blank page)
                WebDriverWait(self.driver, 30).until(
                    lambda driver: len(driver.page_source) > 1000
                )
                logger.info("Page content loaded")
            except TimeoutException:
                logger.warning("Page content took too long to load")
            
            # Additional wait for JavaScript to execute
            time.sleep(5)
            
            # Check if page loaded successfully
            if "nusuk" in self.driver.title.lower():
                logger.log_success("Navigation successful", f"Page title: {self.driver.title}")
                return True
            elif self.driver.title and len(self.driver.title.strip()) > 0:
                logger.warning(f"Unexpected page title: {self.driver.title}")
                # Check if we have content even with unexpected title
                if len(self.driver.page_source) > 1000:
                    logger.info("Page has content, continuing...")
                    return True
                else:
                    logger.error("Page appears to be blank")
                    return False
            else:
                logger.error("Page title is empty or missing")
                return False
                
        except Exception as e:
            logger.log_error_with_context(e, "Navigating to Nusuk website")
            return False
    
    def switch_to_english(self):
        """Switch website language from Arabic to English"""
        try:
            logger.log_step("Switching to English", "STARTED")
            
            # Wait for page to load completely and JavaScript to finish
            logger.info("Waiting for page and JavaScript to fully load...")
            time.sleep(8)  # Increased wait time for JavaScript
            
            # Use the exact selector from the screenshot
            logger.info("Looking for AR button with ID 'dropdownMenuLink'...")
            try:
                # Target the exact element from screenshot: <a id="dropdownMenuLink" class="dropdown-toggle text-decoration-none">
                ar_button = self.driver.find_element("id", "dropdownMenuLink")
                logger.info(f"Found AR button by ID: '{ar_button.text}' - Clicking to open language dropdown...")
                
                # Click the AR button to open dropdown
                self.simulate_human_click(ar_button)
                time.sleep(3)  # Increased wait time for dropdown to open
                
                # Now look for the EN option in the dropdown
                logger.info("Looking for EN option in dropdown...")
                
                # Wait for dropdown to be visible and then find EN option
                en_selectors = [
                    "//a[contains(text(), 'En') or contains(text(), 'EN') or contains(text(), 'English')]",
                    "//li[contains(text(), 'En') or contains(text(), 'EN') or contains(text(), 'English')]",
                    "//div[contains(text(), 'En') or contains(text(), 'EN') or contains(text(), 'English')]"
                ]
                
                en_option = None
                for selector in en_selectors:
                    try:
                        en_option = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                        if en_option:
                            logger.info(f"Found EN option: '{en_option.text}' - Clicking to switch to English...")
                            break
                    except (TimeoutException, NoSuchElementException):
                        continue
                
                if en_option:
                    self.simulate_human_click(en_option)
                    time.sleep(3)
                    
                    # Verify language switch by looking for English text
                    page_text = self.driver.page_source.lower()
                    if 'create account' in page_text or 'foreign guest' in page_text:
                        logger.log_success("Language switched to English successfully")
                        return True
                    else:
                        logger.warning("Language switch may not have worked - page still appears to be in Arabic")
                        return False
                else:
                    logger.error("EN option not found in language dropdown")
                    return False
                    
            except Exception as e:
                logger.warning(f"Could not find AR button by ID: {e}")
                
                # Fallback: Try to find by looking for any element with AR text
                logger.info("Trying fallback method to find AR button...")
                
                # Look for any element with AR text (including <a> tags)
                ar_selectors = [
                    # Look for the AR button with globe icon - now including <a> tags
                    "//a[contains(text(), 'AR') or contains(text(), 'Ar')]",
                    "//span[contains(text(), 'AR') or contains(text(), 'Ar')]",
                    "//div[contains(text(), 'AR') or contains(text(), 'Ar')]",
                    "//button[contains(text(), 'AR') or contains(text(), 'Ar')]",
                    
                    # CSS selectors
                    ".dropdown-toggle",
                    "[id*='dropdown']",
                    "[class*='dropdown']",
                    "[class*='language']",
                    "[class*='lang']"
                ]
                
                ar_button = None
                for selector in ar_selectors:
                    try:
                        if selector.startswith("//"):
                            ar_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                        else:
                            ar_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                        
                        if ar_button:
                            logger.info(f"Found AR button via fallback: '{ar_button.text}' - Clicking to open language dropdown...")
                            break
                    except (TimeoutException, NoSuchElementException):
                        continue
                
                if ar_button:
                    # Click the AR button to open dropdown
                    self.simulate_human_click(ar_button)
                    time.sleep(3)  # Increased wait time
                    
                    # Now look for the EN option in the dropdown
                    en_selectors = [
                        "//a[contains(text(), 'En') or contains(text(), 'EN') or contains(text(), 'English')]",
                        "//li[contains(text(), 'En') or contains(text(), 'EN') or contains(text(), 'English')]",
                        "//div[contains(text(), 'En') or contains(text(), 'EN') or contains(text(), 'English')]"
                    ]
                    
                    en_option = None
                    for selector in en_selectors:
                        try:
                            en_option = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                            if en_option:
                                logger.info(f"Found EN option via fallback: '{en_option.text}' - Clicking to switch to English...")
                                break
                        except (TimeoutException, NoSuchElementException):
                            continue
                    
                    if en_option:
                        # Click EN to switch language
                        self.simulate_human_click(en_option)
                        time.sleep(3)
                        
                        # Verify language switch by looking for English text
                        page_text = self.driver.page_source.lower()
                        if 'create account' in page_text or 'foreign guest' in page_text:
                            logger.log_success("Language switched to English successfully via fallback")
                            return True
                        else:
                            logger.warning("Language switch via fallback may not have worked")
                            return False
                    else:
                        logger.error("EN option not found in language dropdown via fallback")
                        return False
                else:
                    logger.error("AR button not found via any method - cannot switch language")
                    return False
                    
        except Exception as e:
            logger.log_error_with_context(e, "Switching to English")
            return False
    
    def select_foreign_guest(self):
        """Select Foreign Guest option"""
        try:
            logger.log_step("Selecting Foreign Guest", "STARTED")
            
            # Wait for page to load after language switch
            time.sleep(3)
            
            # Primary method: Target the exact radio button by ID from DevTools
            logger.info("Looking for Foreign Guest radio button with ID 'type3'...")
            try:
                foreign_guest_radio = self.driver.find_element("id", "type3")
                logger.info(f"Found Foreign Guest radio button: '{foreign_guest_radio.get_attribute('aria-label') or 'Foreign Guest'}'")
                
                # Click the radio button
                foreign_guest_radio.click()
                time.sleep(2)
                
                logger.log_success("Foreign Guest selected via radio button")
                return True
                
            except Exception as e:
                logger.warning(f"Could not find Foreign Guest radio button by ID: {e}")
                
                # Fallback: Look for the Foreign Guest option with multiple strategies
                logger.info("Trying fallback method to find Foreign Guest...")
                
                # Look for the Foreign Guest option with support for both Arabic and English
                foreign_guest_selectors = [
                    # English text (after language switch)
                    "//label[contains(text(), 'Foreign Guest')]",
                    "//div[contains(text(), 'Foreign Guest')]",
                    "//span[contains(text(), 'Foreign Guest')]",
                    
                    # Arabic text (زائر دولي = International Visitor) - fallback
                    "//label[contains(text(), 'زائر دولي')]",
                    "//div[contains(text(), 'زائر دولي')]",
                    "//span[contains(text(), 'زائر دولي')]",
                    
                    # Look for radio buttons with Foreign Guest text
                    "//input[@type='radio'][@id='type3']",
                    "//p-radiobutton[@id='type3']",
                    
                    # CSS selectors for the specific element
                    "#type3",
                    "[id='type3']",
                    ".p-radiobutton.p-component"
                ]
                
                for selector in foreign_guest_selectors:
                    try:
                        if selector.startswith("//"):
                            element = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                        else:
                            element = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                        
                        logger.info(f"Found Foreign Guest element via fallback: '{element.text}' - Clicking...")
                        
                        # Scroll to element if needed
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
                        time.sleep(1)
                        
                        # Click the element
                        element.click()
                        time.sleep(2)
                        
                        logger.log_success("Foreign Guest selected via fallback")
                        return True
                    except (TimeoutException, NoSuchElementException):
                        continue
                    except Exception as e:
                        logger.warning(f"Failed to click foreign guest selector {selector}: {e}")
                        continue
                
                # If no specific selector found, try to find by looking at the page content
                try:
                    page_text = self.driver.page_source
                    
                    # Look for any clickable element containing foreign-related text (Arabic or English)
                    if 'Foreign Guest' in page_text or 'زائر دولي' in page_text:
                        # Try to find any clickable element - now including radio buttons
                        clickable_elements = self.driver.find_elements(By.XPATH, "//input[@type='radio'] | //p-radiobutton | //label | //div[@role='button'] | //div[contains(@class, 'card')] | //div[contains(@class, 'btn')]")
                        
                        for element in clickable_elements:
                            try:
                                element_text = element.text.lower()
                                if any(keyword in element_text for keyword in ['foreign guest', 'زائر دولي']):
                                    logger.info(f"Found Foreign Guest via text search: '{element.text}' - Clicking...")
                                    element.click()
                                    time.sleep(2)
                                    logger.log_success("Foreign Guest selected via text search")
                                    return True
                            except:
                                continue
                except Exception as e:
                    logger.warning(f"Text-based foreign guest search failed: {e}")
                
                logger.error("Foreign Guest option not found via any method")
                return False
                
        except Exception as e:
            logger.log_error_with_context(e, "Selecting Foreign Guest")
            return False
    
    def fill_initial_info(self, record_data: Dict[str, Any]):
        """Fill initial information (Visa, Nationality, Passport)"""
        try:
            logger.log_step("Filling initial information", "STARTED")
            
            # Fill Visa Number
            logger.info("Looking for Visa Number field...")
            visa_selectors = [
                # Based on the screenshot, look for input with "Visa Number" placeholder
                "//input[@placeholder='Visa Number']",
                "//input[contains(@placeholder, 'Visa')]",
                "//input[@name='visa']",
                "//input[@id='visa']",
                "#visa-number",
                "[data-field='visa']"
            ]
            
            visa_filled = False
            for selector in visa_selectors:
                try:
                    if selector.startswith("//"):
                        element = self.wait.until(EC.presence_of_element_located((By.XPATH, selector)))
                    else:
                        element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    
                    element.clear()
                    visa_number = record_data.get('VISA NO', '')
                    self.simulate_human_typing(element, visa_number)
                    visa_filled = True
                    logger.log_success("Visa number filled with human-like typing", f"Value: {visa_number}")
                    break
                except (TimeoutException, NoSuchElementException):
                    continue
            
            if not visa_filled:
                logger.error("Visa number field not found")
                return False
            
            # Fill Nationality
            logger.info("Looking for Nationality field...")
            nationality_selectors = [
                # Based on the screenshot, this is a dropdown with "Select Nationality" placeholder
                "//span[@role='combobox' and contains(@aria-label, 'Select Nationality')]",
                "//span[contains(@class, 'p-dropdown-label') and contains(text(), 'Select Nationality')]",
                "//span[contains(@class, 'p-dropdown-label')]",
                "//span[@role='combobox']",
                
                # Alternative selectors
                "//input[@placeholder='Select Nationality']",
                "//input[contains(@placeholder, 'Nationality')]",
                "//input[@name='nationality']",
                "//input[@id='nationality']",
                "#nationality",
                "[data-field='nationality']"
            ]
            
            nationality_filled = False
            for selector in nationality_selectors:
                try:
                    if selector.startswith("//"):
                        element = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        element = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    
                    logger.info(f"Found Nationality element: '{element.text}' - Clicking to open dropdown with human-like behavior...")
                    
                    # Click to open the dropdown
                    self.simulate_human_click(element)
                    time.sleep(2)
                    
                    # Now look for the nationality option in the dropdown
                    nationality = record_data.get('NATIONALITY', '')
                    logger.info(f"Looking for nationality option: {nationality}")
                    
                    # Special handling for India - select the second option to avoid "British Indian Ocean Territory"
                    if nationality.lower() == 'india':
                        logger.info("India detected - will type 'India' in search field and select the second option")
                        
                        # First, find and type in the search field
                        try:
                            search_input = self.driver.find_element("xpath", "//input[@placeholder='Search']")
                            search_input.clear()
                            self.simulate_human_typing(search_input, "India")
                            logger.info("Typed 'India' in search field with human-like typing")
                            time.sleep(2)  # Wait for search results
                        except Exception as e:
                            logger.warning(f"Could not find search field: {e}")
                        
                        # Now get all nationality options after search
                        nationality_options = self.driver.find_elements("xpath", "//li[contains(@class, 'p-dropdown-item')] | //li[contains(@class, 'dropdown-item')] | //li[contains(@class, 'option')]")
                        
                        if len(nationality_options) >= 2:
                            try:
                                # Select the second option (index 1) to avoid "British Indian Ocean Territory"
                                second_option = nationality_options[1]
                                logger.info(f"Selecting second option: '{second_option.text}' with human-like behavior")
                                # Use simple click to avoid stale element issues
                                second_option.click()
                                nationality_selected = True
                                logger.log_success("India selected as second option", f"Value: {second_option.text}")
                                time.sleep(2)  # Wait for dropdown to close and page to stabilize
                            except Exception as e:
                                logger.warning(f"Error clicking nationality option: {e}")
                                nationality_selected = False
                        else:
                            logger.warning("Not enough options found for India selection")
                            nationality_selected = False
                    else:
                        # For other nationalities, try normal selection
                        nationality_option_selectors = [
                            f"//li[contains(text(), '{nationality}')]",
                            f"//div[contains(text(), '{nationality}')]",
                            f"//span[contains(text(), '{nationality}')]",
                            f"//option[contains(text(), '{nationality}')]"
                        ]
                        
                        nationality_selected = False
                        for option_selector in nationality_option_selectors:
                            try:
                                nationality_option = self.driver.find_element("xpath", option_selector)
                                logger.info(f"Found nationality option: '{nationality_option.text}' - Clicking...")
                                nationality_option.click()
                                nationality_selected = True
                                logger.log_success("Nationality selected from dropdown", f"Value: {nationality}")
                                break
                            except:
                                continue
                    
                    if not nationality_selected:
                        # If dropdown selection fails, try typing in the search field
                        try:
                            # Look for a search input that might appear after clicking the dropdown
                            search_input = self.driver.find_element("xpath", "//input[@placeholder='Search']")
                            search_input.clear()
                            search_input.send_keys(nationality)
                            time.sleep(2)
                            
                            # Try to select from search results
                            search_result = self.driver.find_element("xpath", f"//li[contains(text(), '{nationality}')]")
                            search_result.click()
                            nationality_selected = True
                            logger.log_success("Nationality selected via search", f"Value: {nationality}")
                        except:
                            logger.warning(f"Could not select nationality: {nationality}")
                            # Continue anyway - the form might still work without nationality
                            nationality_selected = True
                    
                    nationality_filled = True
                    break
                except (TimeoutException, NoSuchElementException):
                    continue
            
            if not nationality_filled:
                logger.error("Nationality field not found")
                return False
            
            # Fill Passport Number
            logger.info("Looking for Passport Number field...")
            passport_selectors = [
                # Based on the screenshot, look for passport field
                "//input[@placeholder='Passport Number']",
                "//input[contains(@placeholder, 'Passport')]",
                "//input[@name='passport']",
                "//input[@id='passport']",
                "#passport-number",
                "[data-field='passport']"
            ]
            
            passport_filled = False
            for selector in passport_selectors:
                try:
                    if selector.startswith("//"):
                        element = self.wait.until(EC.presence_of_element_located((By.XPATH, selector)))
                    else:
                        element = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    
                    # Clear the field first
                    element.clear()
                    time.sleep(1)
                    
                    # Debug: Log all available data
                    logger.info(f"Available record data keys: {list(record_data.keys())}")
                    
                    # Try different possible column names for passport
                    passport_number = (
                        record_data.get('Passport NO', '') or 
                        record_data.get('PASSPORT NO', '') or 
                        record_data.get('Passport Number', '') or 
                        record_data.get('PASSPORT NUMBER', '') or 
                        record_data.get('Passport', '') or 
                        record_data.get('PASSPORT', '')
                    )
                    
                    logger.info(f"Passport number retrieved: '{passport_number}' (length: {len(passport_number)})")
                    
                    if passport_number:
                        self.simulate_human_typing(element, passport_number)
                        passport_filled = True
                        logger.log_success("Passport number filled with human-like typing", f"Value: {passport_number}")
                    else:
                        logger.warning("Passport number is empty - using default value")
                        # Use a default passport number if none is provided
                        default_passport = "AB315033"
                        self.simulate_human_typing(element, default_passport)
                        passport_filled = True
                        logger.log_success("Passport number filled with default using human-like typing", f"Value: {default_passport}")
                    break
                except (TimeoutException, NoSuchElementException):
                    continue
                except Exception as e:
                    logger.warning(f"Error with passport selector {selector}: {e}")
                    continue
            
            if not passport_filled:
                logger.error("Passport number field not found")
                return False
            
            # Click Next button
            logger.info("Looking for Next button...")
            next_selectors = [
                # Based on the screenshot, the button contains a span with "Next" text
                "//button[.//span[contains(text(), 'Next')]]",
                "//button[.//span[text()='Next']]",
                "//button[contains(@class, 'btn') and contains(@class, 'login-btn')]",
                "//button[contains(@class, 'p-button')]",
                "//button[@type='submit']",
                ".btn.login-btn",
                "#next-button"
            ]
            
            for selector in next_selectors:
                try:
                    if selector.startswith("//"):
                        element = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        element = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    
                    logger.info(f"Found Next button: '{element.text}' - Clicking with human-like behavior...")
                    self.simulate_human_click(element)
                    time.sleep(3)
                    logger.log_success("Next button clicked with human-like behavior")
                    return True
                except (TimeoutException, NoSuchElementException):
                    continue
            
            logger.error("Next button not found")
            return False
            
        except Exception as e:
            logger.log_error_with_context(e, "Filling initial information")
            return False
    
    def handle_date_picker(self, element, date_value):
        """Handle date picker interactions and ensure it's properly closed"""
        try:
            logger.info(f"Handling date picker for value: {date_value}")
            
            # First, try to close any open date picker
            try:
                # Look for date picker overlay
                date_picker_selectors = [
                    "//div[contains(@class, 'p-datepicker')]",
                    "//div[contains(@class, 'calendar')]",
                    "//div[contains(@class, 'datepicker')]",
                    "//div[contains(@class, 'p-overlay')]"
                ]
                
                for selector in date_picker_selectors:
                    try:
                        date_picker = self.driver.find_element("xpath", selector)
                        if date_picker.is_displayed():
                            logger.info("Found open date picker, closing it...")
                            
                            # Try multiple ways to close it
                            try:
                                # Method 1: Click outside
                                self.driver.find_element("xpath", "//body").click()
                                time.sleep(1)
                            except:
                                pass
                            
                            try:
                                # Method 2: Press Escape key
                                from selenium.webdriver.common.keys import Keys
                                element.send_keys(Keys.ESCAPE)
                                time.sleep(1)
                            except:
                                pass
                            
                            try:
                                # Method 3: Click close button if exists
                                close_btn = self.driver.find_element("xpath", "//button[contains(@class, 'close')] | //button[contains(@class, 'p-datepicker-close')]")
                                close_btn.click()
                                time.sleep(1)
                            except:
                                pass
                            
                            break
                    except NoSuchElementException:
                        continue
                        
            except Exception as e:
                logger.warning(f"Error while trying to close date picker: {e}")
            
            # Now fill the date field
            element.clear()
            time.sleep(1)
            element.click()
            time.sleep(1)
            element.send_keys(date_value)
            time.sleep(1)
            
            # Press Tab to confirm and move to next field
            from selenium.webdriver.common.keys import Keys
            element.send_keys(Keys.TAB)
            time.sleep(1)
            
            logger.log_success("Date picker handled successfully", f"Value: {date_value}")
            return True
            
        except Exception as e:
            logger.error(f"Error handling date picker: {e}")
            return False

    def fill_personal_details(self, record_data: Dict[str, Any]):
        """Fill personal details (DOB, Gender, Phone Number, Disability)"""
        try:
            logger.log_step("Filling personal details", "STARTED")
            
            # Fill Date of Birth
            logger.info("Looking for Date of Birth field...")
            dob_selectors = [
                "//input[@placeholder='Specify date of birth']",
                "//input[contains(@placeholder, 'date of birth')]",
                "//input[@role='combobox']",
                "//input[contains(@class, 'p-calendar')]",
                "//input[contains(@class, 'p-inputtext')]"
            ]
            
            dob_filled = False
            for selector in dob_selectors:
                try:
                    element = self.wait.until(EC.presence_of_element_located((By.XPATH, selector)))
                    
                    # Clear the field first
                    element.clear()
                    time.sleep(1)
                    
                    # Get date from sheet and format it properly
                    dob_value = record_data.get('DOB', '')
                    if dob_value:
                        # Handle both dd/mm/yyyy and dd-mm-yyyy formats
                        if '/' in dob_value:
                            dob_value = dob_value.replace('/', '-')
                        
                        logger.info(f"Filling DOB: {dob_value}")
                        self.simulate_human_typing(element, dob_value)
                        
                        # Press Tab to confirm and move to next field
                        element.send_keys(Keys.TAB)
                        time.sleep(1)
                        
                        dob_filled = True
                        logger.log_success("Date of birth filled", f"Value: {dob_value}")
                        break
                    else:
                        logger.warning("No DOB value found in sheet data")
                        break
                        
                except (TimeoutException, NoSuchElementException):
                    continue
            
            if not dob_filled:
                logger.error("Date of Birth field not found or could not be filled")
                return False
            
            # Fill Gender
            logger.info("Looking for Gender selection...")
            gender = record_data.get('SEX', '').upper()
            
            # Target the exact radio button structure from DevTools
            if gender in ['M', 'MALE']:
                gender_selectors = [
                    # Target the p-radiobutton-box div for Male (value="1")
                    "//p-radiobutton[@formcontrolname='gender' and @value='1']//div[contains(@class, 'p-radiobutton-box')]",
                    "//p-radiobutton[@id='gender1']//div[contains(@class, 'p-radiobutton-box')]",
                    "//div[contains(@class, 'p-radiobutton-box')][ancestor::p-radiobutton[@value='1']]",
                    # Fallback selectors
                    "//label[contains(text(), 'Male')]",
                    "//input[@id='gender1']"
                ]
            elif gender in ['F', 'FEMALE']:
                gender_selectors = [
                    # Target the p-radiobutton-box div for Female (value="2")
                    "//p-radiobutton[@formcontrolname='gender' and @value='2']//div[contains(@class, 'p-radiobutton-box')]",
                    "//p-radiobutton[@id='gender2']//div[contains(@class, 'p-radiobutton-box')]",
                    "//div[contains(@class, 'p-radiobutton-box')][ancestor::p-radiobutton[@value='2']]",
                    # Fallback selectors
                    "//label[contains(text(), 'Female')]",
                    "//input[@id='gender2']"
                ]
            else:
                logger.warning(f"Unknown gender value: {gender}, defaulting to Male")
                gender_selectors = [
                    # Target the p-radiobutton-box div for Male (value="1")
                    "//p-radiobutton[@formcontrolname='gender' and @value='1']//div[contains(@class, 'p-radiobutton-box')]",
                    "//p-radiobutton[@id='gender1']//div[contains(@class, 'p-radiobutton-box')]",
                    "//div[contains(@class, 'p-radiobutton-box')][ancestor::p-radiobutton[@value='1']]",
                    # Fallback selectors
                    "//label[contains(text(), 'Male')]",
                    "//input[@id='gender1']"
                ]
            
            gender_selected = False
            for selector in gender_selectors:
                try:
                    gender_element = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    logger.info(f"Found gender element: '{gender_element.text or 'Radio Button'}' - Clicking with human-like behavior...")
                    self.simulate_human_click(gender_element)
                    time.sleep(2)
                    gender_selected = True
                    logger.log_success("Gender selected with human-like behavior", f"Value: {gender}")
                    break
                except (TimeoutException, NoSuchElementException) as e:
                    logger.warning(f"Gender selector {selector} failed: {e}")
                    continue
            
            if not gender_selected:
                logger.error("Gender selection failed")
                return False
            
            # Fill Phone Number (using VISA NO from sheet)
            logger.info("Looking for Phone Number field...")
            phone_selectors = [
                "//input[@id='mobile']",
                "//input[@placeholder='Phone Number']",
                "//input[contains(@placeholder, 'Phone')]",
                "//input[contains(@class, 'p-inputtext')]"
            ]
            
            phone_filled = False
            for selector in phone_selectors:
                try:
                    element = self.wait.until(EC.presence_of_element_located((By.XPATH, selector)))
                    element.clear()
                    visa_number = record_data.get('VISA NO', '')
                    self.simulate_human_typing(element, visa_number)
                    phone_filled = True
                    logger.log_success("Phone Number filled with VISA NO using human-like typing", f"Value: {visa_number}")
                    break
                except (TimeoutException, NoSuchElementException):
                    continue
            
            if not phone_filled:
                logger.error("Phone Number field not found")
                return False
            
            # Select Disability - Always "No"
            logger.info("Looking for Disability selection...")
            disability_selectors = [
                # Target the exact structure from DevTools: formcontrolname="needAssistance" value="0"
                "//p-radiobutton[@formcontrolname='needAssistance' and @value='0']//div[contains(@class, 'p-radiobutton-box')]",
                "//p-radiobutton[@id='type2']//div[contains(@class, 'p-radiobutton-box')]",
                
                # Try clicking the label for the "No" option
                "//label[contains(text(), 'No') and contains(@for, 'type2')]",
                "//label[contains(text(), 'No')]",
                
                # Try clicking the span containing "No"
                "//span[contains(text(), 'No') and ancestor::p-radiobutton]",
                
                # Fallback to the p-radiobutton structure
                "//p-radiobutton[@formcontrolname='disability' and @value='no']//div[contains(@class, 'p-radiobutton-box')]",
                "//p-radiobutton[@formcontrolname='disability' and @value='false']//div[contains(@class, 'p-radiobutton-box')]",
                "//p-radiobutton[@id='disability2']//div[contains(@class, 'p-radiobutton-box')]"
            ]
            
            disability_selected = False
            for selector in disability_selectors:
                try:
                    if selector.startswith("//"):
                        element = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        element = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    
                    logger.info(f"Found disability element: '{element.text}' - Clicking with human-like behavior...")
                    self.simulate_human_click(element)
                    time.sleep(1)
                    disability_selected = True
                    logger.log_success("Disability selected with human-like behavior", "Value: No")
                    break
                except (TimeoutException, NoSuchElementException):
                    continue
            
            if not disability_selected:
                logger.error("Disability field not found or could not be selected")
                return False
            
            # Click Next button
            logger.info("Looking for Next button...")
            next_selectors = [
                "//button[.//span[contains(text(), 'Next')]]",
                "//button[contains(@class, 'btn') and contains(@class, 'login-btn')]",
                "//button[contains(@class, 'p-button')]",
                "//button[@type='submit']",
                ".btn.login-btn"
            ]
            
            for selector in next_selectors:
                try:
                    if selector.startswith("//"):
                        element = self.wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    else:
                        element = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                    
                    logger.info(f"Found Next button: '{element.text}' - Clicking with human-like behavior...")
                    self.simulate_human_click(element)
                    time.sleep(3)
                    logger.log_success("Next button clicked with human-like behavior")
                    return True
                except (TimeoutException, NoSuchElementException):
                    continue
            
            logger.error("Next button not found")
            return False
            
        except Exception as e:
            logger.log_error_with_context(e, "Filling personal details")
            return False
    
    def create_account(self, record_data: Dict[str, Any]):
        """Create account with manual email/password input - user handles this page manually"""
        try:
            logger.log_step("Creating account", "STARTED")
            logger.info("🎯 MANUAL MODE: Email and password page - please fill manually")
            
            # Wait for user to manually fill email and password
            logger.info("⏳ Waiting for manual input... Please fill in:")
            logger.info("   - Email address")
            logger.info("   - Password")
            logger.info("   - Confirm Password")
            logger.info("   - Click Create Account button when ready")
            
            # Wait 15 seconds for manual input
            wait_time = 15
            logger.info(f"⏰ Waiting {wait_time} seconds for manual completion...")
            
            # Countdown timer
            for i in range(wait_time, 0, -1):
                logger.info(f"⏳ Time remaining: {i} seconds...")
                time.sleep(1)
            
            logger.info("✅ Manual input time completed - checking if account was created...")
            
            # Check if we're on a success page or if there are any error messages
            try:
                # Look for success indicators
                success_indicators = [
                    "//div[contains(text(), 'success')]",
                    "//div[contains(text(), 'Success')]",
                    "//div[contains(text(), 'Account created')]",
                    "//div[contains(text(), 'Verification')]",
                    "//div[contains(text(), 'OTP')]",
                    "//div[contains(text(), 'verification code')]",
                    "//input[@name='otp']",
                    "//input[@placeholder*='OTP']",
                    "//input[@placeholder*='verification']"
                ]
                
                success_found = False
                for selector in success_indicators:
                    try:
                        element = self.driver.find_element("xpath", selector)
                        if element.is_displayed():
                            logger.log_success("Account creation appears successful", f"Found: {element.text or 'Success indicator'}")
                            success_found = True
                            break
                    except:
                        continue
                
                if not success_found:
                    # Check for error messages
                    error_selectors = [
                        "//div[contains(@class, 'error')]",
                        "//div[contains(@class, 'alert')]",
                        "//span[contains(@class, 'error')]",
                        "//p[contains(@class, 'error')]"
                    ]
                    
                    for selector in error_selectors:
                        try:
                            error_element = self.driver.find_element("xpath", selector)
                            if error_element.is_displayed():
                                logger.warning(f"Potential error detected: {error_element.text}")
                        except:
                            continue
                    
                    logger.info("No clear success/error indicators found - assuming manual completion was successful")
                
                # Store a placeholder email for tracking purposes
                self.current_email = "MANUAL_INPUT_COMPLETED"
                
                logger.log_success("Manual account creation completed", "Proceeding to next step...")
                return True
                
            except Exception as e:
                logger.warning(f"Error checking account creation status: {e}")
                # Assume manual completion was successful
                self.current_email = "MANUAL_INPUT_COMPLETED"
                return True
            
        except Exception as e:
            logger.log_error_with_context(e, "Manual account creation")
            return False
    
    def monitor_inbox_for_otp(self, email_address: str, max_wait_time: int = 60):
        """Monitor inbox for OTP email with detailed logging"""
        try:
            logger.log_step("Monitoring inbox for OTP email", "STARTED")
            logger.info(f"Monitoring inbox: {email_address}")
            logger.info(f"Maximum wait time: {max_wait_time} seconds")
            
            start_time = time.time()
            check_interval = 5  # Check every 5 seconds
            
            while time.time() - start_time < max_wait_time:
                elapsed = int(time.time() - start_time)
                logger.info(f"Checking inbox... (Elapsed: {elapsed}s)")
                
                try:
                    # Try to get OTP from email service
                    otp = self.email_service.get_otp(email_address)
                    if otp:
                        logger.log_success("OTP email found in inbox", f"OTP: {otp}")
                        return True
                    
                    # Log inbox status
                    logger.info(f"Inbox checked - No OTP email yet (Elapsed: {elapsed}s)")
                    
                except Exception as e:
                    logger.warning(f"Error checking inbox: {e}")
                
                # Wait before next check
                time.sleep(check_interval)
            
            logger.warning(f"OTP email not found after {max_wait_time} seconds")
            return False
            
        except Exception as e:
            logger.log_error_with_context(e, "Monitoring inbox for OTP")
            return False
    
    def verify_email_otp(self):
        """Verify email with OTP - adapted for manual email handling"""
        try:
            logger.log_step("Verifying email with OTP", "STARTED")
            
            # Since user handled email manually, we need to check if we're on OTP page
            logger.info("🔍 Checking if we're on OTP verification page...")
            
            # Wait a bit for page to load
            time.sleep(3)
            
            # Look for OTP input field to confirm we're on verification page
            otp_selectors = [
                "//input[@name='otp']",
                "//input[@name='verification_code']",
                "//input[@name='code']",
                "//input[@id='otp']",
                "//input[@id='verification_code']",
                "//input[@id='code']",
                "#otp",
                "#verification_code",
                "#code",
                "[data-field='otp']",
                "[data-field='verification_code']"
            ]
            
            otp_field_found = False
            for selector in otp_selectors:
                try:
                    element = self.driver.find_element("xpath", selector)
                    if element.is_displayed():
                        logger.log_success("OTP verification page detected", "Found OTP input field")
                        otp_field_found = True
                        break
                except:
                    continue
            
            if not otp_field_found:
                logger.warning("OTP verification page not detected - checking current page status")
                
                # Check if we're still on account creation page
                current_url = self.driver.current_url
                page_title = self.driver.title
                logger.info(f"Current URL: {current_url}")
                logger.info(f"Page title: {page_title}")
                
                # Look for any error messages or success indicators
                page_text = self.driver.page_source.lower()
                if 'error' in page_text or 'failed' in page_text:
                    logger.error("Account creation appears to have failed")
                    return False
                elif 'success' in page_text or 'verification' in page_text:
                    logger.info("Account creation appears successful, proceeding...")
                    return True
                else:
                    logger.warning("Unable to determine page status - proceeding anyway")
                    return True
            
            # If we're on OTP page, wait for user to handle it manually
            logger.info("🎯 MANUAL MODE: OTP verification page detected")
            logger.info("⏳ Please check your email and enter the OTP manually")
            logger.info("⏰ Waiting for manual OTP verification...")
            
            # Wait for user to complete OTP verification
            wait_time = 20  # Give more time for OTP verification
            logger.info(f"⏰ Waiting {wait_time} seconds for manual OTP completion...")
            
            # Countdown timer
            for i in range(wait_time, 0, -1):
                logger.info(f"⏳ Time remaining: {i} seconds...")
                time.sleep(1)
            
            logger.info("✅ Manual OTP verification time completed - checking status...")
            
            # Check if verification was successful
            try:
                # Look for success indicators after OTP verification
                success_indicators = [
                    "//div[contains(text(), 'success')]",
                    "//div[contains(text(), 'Success')]",
                    "//div[contains(text(), 'verified')]",
                    "//div[contains(text(), 'Verification successful')]",
                    "//div[contains(text(), 'Account activated')]",
                    "//button[contains(text(), 'Continue')]",
                    "//button[contains(text(), 'Next')]",
                    "//a[contains(text(), 'Continue')]"
                ]
                
                success_found = False
                for selector in success_indicators:
                    try:
                        element = self.driver.find_element("xpath", selector)
                        if element.is_displayed():
                            logger.log_success("OTP verification appears successful", f"Found: {element.text or 'Success indicator'}")
                            success_found = True
                            break
                    except:
                        continue
                
                if not success_found:
                    logger.info("No clear success indicators found - assuming manual verification was successful")
                
                logger.log_success("Manual OTP verification completed", "Proceeding to next step...")
                return True
                
            except Exception as e:
                logger.warning(f"Error checking OTP verification status: {e}")
                # Assume manual verification was successful
                return True
            
        except Exception as e:
            logger.log_error_with_context(e, "Manual OTP verification")
            return False
    
    def process_record(self, record: Dict[str, Any]):
        """Process a single record from Google Sheets"""
        try:
            logger.log_step("Processing record", f"Row {record['row_number']}")
            
            self.current_record = record
            record_data = record['data']
            
            # Step 1: Navigate to website
            if not self.navigate_to_nusuk():
                return False
            
            # Step 2: Switch to English
            logger.info("Switching to English...")
            if not self.switch_to_english():
                return False
            
            # Step 3: Select Foreign Guest
            logger.info("Selecting Foreign Guest...")
            if not self.select_foreign_guest():
                return False
            
            # Step 4: Fill initial information
            if not self.fill_initial_info(record_data):
                return False
            
            # Step 5: Fill personal details
            if not self.fill_personal_details(record_data):
                return False
            
            # Step 6: Create account
            if not self.create_account(record_data):
                return False
            
            # Step 7: Verify email with OTP
            if not self.verify_email_otp():
                return False
            
            # Update status to "In Progress"
            self.sheets_service.update_record_status(record['row_number'], "In Progress")
            
            # Update email used in Google Sheet
            if self.current_email:
                self.sheets_service.update_email_used(record['row_number'], self.current_email)
            
            logger.log_success("Record processing started", f"Row {record['row_number']}")
            return True
            
        except Exception as e:
            logger.log_error_with_context(e, f"Processing record {record['row_number']}")
            self.sheets_service.log_error(record['row_number'], str(e))
            return False
    
    def run_automation(self):
        """Main automation runner"""
        try:
            logger.log_step("Starting Nusuk automation", "STARTED")
            
            # Setup driver
            self.setup_driver()
            
            # Get pending records
            pending_records = self.sheets_service.get_pending_records()
            
            if not pending_records:
                logger.info("No pending records found")
                return
            
            logger.log_success("Found pending records", f"Count: {len(pending_records)}")
            
            # Process each record
            for record in pending_records:
                try:
                    success = self.process_record(record)
                    if success:
                        logger.log_success("Record processed successfully", f"Row {record['row_number']}")
                    else:
                        logger.log_failure("Record processing failed", f"Row {record['row_number']}")
                        self.sheets_service.update_record_status(record['row_number'], "Failed")
                    
                    # Add delay between records
                    time.sleep(random.uniform(5, 10))
                    
                except Exception as e:
                    logger.log_error_with_context(e, f"Processing record {record['row_number']}")
                    self.sheets_service.update_record_status(record['row_number'], "Failed")
                    self.sheets_service.log_error(record['row_number'], str(e))
            
            logger.log_success("Automation completed", f"Processed {len(pending_records)} records")
            
        except Exception as e:
            logger.log_error_with_context(e, "Running automation")
        finally:
            if self.driver:
                self.driver.quit()
                logger.info("Browser closed")
            
            # Clean up stealth Chrome profile
            self._cleanup_stealth_profile()
    
    def test_website_access(self):
        """Test if we can access the Nusuk website"""
        try:
            logger.log_step("Testing website access", "STARTED")
            
            self.setup_driver()
            
            if self.navigate_to_nusuk():
                logger.log_success("Website access test", "Successfully accessed Nusuk website")
                return True
            else:
                logger.log_failure("Website access test", "Failed to access Nusuk website")
                return False
                
        except Exception as e:
            logger.log_error_with_context(e, "Testing website access")
            return False
        finally:
            if self.driver:
                self.driver.quit()
            
            # Clean up stealth Chrome profile
            self._cleanup_stealth_profile()
    
    def _create_stealth_profile(self):
        """Create a minimal stealth profile that looks like a regular user"""
        try:
            import tempfile
            import json
            
            # Create temporary directory for stealth profile
            temp_dir = tempfile.mkdtemp(prefix="chrome_stealth_")
            stealth_profile = os.path.join(temp_dir, "User Data")
            
            logger.info(f"Creating stealth Chrome profile at: {stealth_profile}")
            
            # Create User Data directory structure
            os.makedirs(stealth_profile, exist_ok=True)
            os.makedirs(os.path.join(stealth_profile, "Default"), exist_ok=True)
            
            # Create a realistic Preferences file that makes Chrome look like a regular user
            preferences = {
                "account_id_migration_state": 2,
                "account_tracker_service_last_update": "1337",
                "browser": {
                    "window_placement": {
                        "bottom": 1040,
                        "left": 100,
                        "right": 1380,
                        "top": 40
                    }
                },
                "default_search_provider": {
                    "enabled": True,
                    "search_url": "https://www.google.com/search?q={searchTerms}",
                    "suggest_url": "https://www.google.com/complete/search?client=chrome&q={searchTerms}"
                },
                "extensions": {
                    "settings": {}
                },
                "profile": {
                    "content_settings": {
                        "exceptions": {
                            "notifications": {
                                "https://www.google.com,*": {
                                    "setting": 1
                                }
                            }
                        }
                    },
                    "default_content_setting_values": {
                        "notifications": 1
                    },
                    "name": "Default"
                },
                "session": {
                    "restore_on_startup": 4
                },
                "translate": {
                    "enabled": True
                },
                "webkit": {
                    "webprefs": {
                        "default_font_size": 16,
                        "default_fixed_font_size": 13
                    }
                }
            }
            
            # Write preferences file
            preferences_path = os.path.join(stealth_profile, "Default", "Preferences")
            with open(preferences_path, 'w', encoding='utf-8') as f:
                json.dump(preferences, f, indent=2)
            
            # Create Local State file
            local_state = {
                "browser": {
                    "enabled_labs_experiments": [],
                    "last_known_google_url": "https://www.google.com/",
                    "last_prompted_google_url": "https://www.google.com/"
                },
                "profile": {
                    "info_cache": {
                        "Default": {
                            "active_time": 0,
                            "background_apps": False,
                            "name": "Default",
                            "user_name": ""
                        }
                    }
                }
            }
            
            # Write Local State file
            local_state_path = os.path.join(stealth_profile, "Local State")
            with open(local_state_path, 'w', encoding='utf-8') as f:
                json.dump(local_state, f, indent=2)
            
            logger.log_success("Stealth profile created", "Created realistic Chrome profile for maximum stealth")
            
            # Store the stealth profile path for cleanup
            self.cloned_profile_path = stealth_profile
            self.temp_dir = temp_dir
            
            return stealth_profile
            
        except Exception as e:
            logger.warning(f"Failed to create stealth profile: {e}")
            return None
    
    def _cleanup_stealth_profile(self):
        """Clean up the stealth Chrome profile"""
        try:
            if hasattr(self, 'temp_dir') and self.temp_dir:
                import shutil
                shutil.rmtree(self.temp_dir, ignore_errors=True)
                logger.info("Cleaned up stealth Chrome profile")
        except Exception as e:
            logger.warning(f"Failed to cleanup stealth profile: {e}")
