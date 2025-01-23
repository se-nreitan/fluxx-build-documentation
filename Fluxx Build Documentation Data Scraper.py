# HOW TO USE:
# 1. Install Selenium using pip: pip install selenium
# 2. Download ChromeDriver matching your Chrome version from:
#    https://googlechromelabs.github.io/chrome-for-testing/
# 3. Place chromedriver.exe in the drivers/chromedriver_win32 folder
# 4. Install webdriver-manager: pip install webdriver-manager

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from getpass import getpass
from urllib.parse import urlparse
import sys
import platform
import os
import re
import subprocess
import json
import zipfile
import requests
import winreg
import shutil
import time
from selenium.common.exceptions import TimeoutException
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime

def get_chrome_path():
    """Get installed Chrome path from registry"""
    try:
        # Check common Chrome locations
        possible_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ]
        
        # Try registry first
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe")
            chrome_path = winreg.QueryValue(key, None)
            if os.path.exists(chrome_path):
                return chrome_path
        except:
            pass
        
        # Try common paths
        for path in possible_paths:
            if os.path.exists(path):
                return path
                
        return None
    except:
        return None

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_url():
    """Get and format the Fluxx URL from user input"""
    while True:
        try:
            url = input("\nEnter Fluxx URL (e.g., example.fluxx.io): ").strip()
            if not url:
                print("URL cannot be empty. Please try again.")
                continue
                
            if not url.startswith(('http://', 'https://')):
                url = 'https://' + url
                
            print(f"Using URL: {url}")
            return url
            
        except Exception as e:
            print(f"Error with URL input: {str(e)}")
            print("Please try again.")

def get_chrome_version():
    """Get Chrome version from registry."""
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Google\Chrome\BLBeacon')
        version = winreg.QueryValueEx(key, 'version')[0]
        return version
    except:
        return "unknown"

def get_driver_version(driver_path):
    """Get ChromeDriver version by running the executable with --version flag."""
    try:
        result = subprocess.run([driver_path, '--version'], 
                              capture_output=True, 
                              text=True)
        version = result.stdout.split()[1]  # Format: "ChromeDriver XX.X.XXXX.XX"
        return version.split('.')[0]  # Return major version
    except:
        return None

def setup_webdriver():
    try:
        print("\nSetting up Chrome WebDriver...")
        
        # Use bundled ChromeDriver
        driver_path = get_resource_path("chromedriver.exe")
        chrome_binary_path = get_resource_path(os.path.join("chrome-portable", "chrome.exe"))
        
        print(f"Looking for Chrome at: {chrome_binary_path}")
        
        # Basic Chrome setup
        options = webdriver.ChromeOptions()
        options.binary_location = chrome_binary_path
        options.add_argument('--no-sandbox')  # Required for navigation
        options.add_argument('--disable-dev-shm-usage')  # Required for navigation
        
        print("Initializing Chrome...")
        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=options)
        
        # Test navigation to make sure it works
        print("Testing navigation...")
        driver.get("https://www.google.com")
        if driver.current_url == "data:,":
            raise Exception("Chrome navigation not working properly")
            
        print("Chrome initialized successfully!")
        return driver
        
    except Exception as e:
        print("\nError setting up Chrome WebDriver:")
        print(str(e))
        print("\nChrome path:", chrome_binary_path)
        print("Driver path:", driver_path)
        input("\nPress Enter to exit...")
        raise

def get_credentials():
    """Get username and password from user securely."""
    print("\nPlease enter your Fluxx credentials:")
    username = input("Username: ").strip()
    password = getpass("Password: ")  # getpass hides the password while typing
    return username, password

def handle_login(driver):
    """Handle the login process with browser interaction."""
    try:
        print("\nWaiting for login page to load...")
        print("\nPlease follow these steps:")
        print("1. In the browser window:")
        print("   - Enter your credentials for", driver.current_url)
        print("   - Click 'Sign In' or 'Log In'")
        print("\n2. If you have multiple profiles:")
        print("   - Select your Admin profile from the profile selection screen")
        print("   - Wait for the profile to load")
        print("\n3. Once you see the dashboard screen:")
        print("   - Ensure all elements have loaded")
        print("   - Return to this window")
        print("   - Press Enter to continue")
        
        # Wait for initial page load
        wait = WebDriverWait(driver, 20)
        
        # Let user handle login manually
        input("\nPress Enter when you're on the dashboard screen...")
        
        # Verify we're logged in (not on login page)
        try:
            current_url = driver.current_url
            if 'login' in current_url.lower():
                print("\nError: Still on login page. Please make sure you're logged in.")
                input("Press Enter to exit...")
                return False
            else:
                print("Login verified!")
                return True
        except Exception as e:
            print("\nError verifying login status:")
            print(str(e))
            print("Current URL:", driver.current_url)
            input("Press Enter to exit...")
            return False
        
    except Exception as e:
        print("\nError during login process:")
        print(str(e))
        print("\nCurrent URL:", driver.current_url)
        print("\nDebug info:")
        print(f"Page title: {driver.title}")
        input("Press Enter to exit...")
        return False

def scrape_fluxx_data(driver, url):
    try:
        # Clean up the URL
        url = url.strip()
        if url.startswith('@'):
            url = url[1:]
        if not url.startswith(('http://', 'https://')):
            url = 'https://' + url
            
        print(f"\nNavigating to: {url}")
        
        # Try multiple navigation methods
        try:
            print("Attempting direct navigation...")
            driver.get(url)
            if driver.current_url == "data:,":
                raise Exception("Direct navigation failed")
        except:
            print("Trying JavaScript navigation...")
            driver.execute_script(f"window.location.href = '{url}';")
            time.sleep(2)
            
        current_url = driver.current_url
        print(f"Current URL: {current_url}")
        
        if current_url == "data:,":
            raise Exception("Unable to navigate to the URL")
            
        # Handle login
        if not handle_login(driver):
            print("Login failed. Exiting...")
            return None
            
        # TODO: Add actual scraping logic here
        pass
        
    except Exception as e:
        print(f"\nError during navigation/scraping:")
        print(str(e))
        print(f"Attempted URL: {url}")
        print(f"Current URL: {driver.current_url if driver else 'Not available'}")
        input("\nPress Enter to continue...")
        return None

def wait_for_dashboard(driver, timeout=60):
    """Wait for dashboard to load and verify we're logged in"""
    try:
        print("\nWaiting for dashboard to load...")
        print("Please:")
        print("1. Log in with your credentials")
        print("2. Select your Admin profile (if applicable)")
        print("3. System will automatically proceed when dashboard loads")
        
        wait = WebDriverWait(driver, timeout)
        
        # Wait for the specific Admin Panel link
        admin_link = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a.to-admin-panel[href="/?db=config"]'))
        )
        print("Dashboard loaded successfully!")
        return True
        
    except TimeoutException:
        print("Timed out waiting for dashboard to load")
        return False
    except Exception as e:
        print(f"Error waiting for dashboard: {str(e)}")
        return False

def navigate_to_admin(driver):
    """Navigate to Admin Panel"""
    try:
        wait = WebDriverWait(driver, 10)
        
        # Wait for and click the specific Admin Panel link
        admin_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.to-admin-panel[href="/?db=config"]'))
        )
        print("Found Admin Panel button")
        admin_button.click()
        
        # Wait for admin page to load
        wait.until(EC.url_contains('db=config'))
        print("Successfully navigated to Admin Panel")
        return True
        
    except Exception as e:
        print(f"Error navigating to Admin Panel: {str(e)}")
        return False

def wait_for_forms_and_parse(driver, max_retries=3):
    """Wait for Forms section to load and parse content with retry logic"""
    try:
        print("\nParsing Forms section...")
        
        wait = WebDriverWait(driver, 10)
        
        # First verify Forms is selected
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@id='dashboard-picker']//li[contains(@class,'selected')]//a[text()='Forms']")
        ))
        
        while True:  # Loop until user confirms correct model count
            # Get model count by counting list-label elements that contain model names
            model_list = driver.find_elements(By.CSS_SELECTOR, 
                "#admin-navigation #iconList > ul > li.list-label div.link")
            model_count = len(model_list)
            
            # Ask user to verify count
            print(f"\nFound {model_count} models in the Forms section.")
            verify = input("Does this count appear correct? (y/n): ").strip().lower()
            
            if verify == 'y':
                print("\nProceeding with parsing...")
                # Now get the full UL elements for processing
                model_uls = driver.find_elements(By.CSS_SELECTOR, "#iconList > ul.toggle-class")
                break
            else:
                print("\nWaiting for page to fully load...")
                retry = input("Press Enter when ready to rescan, or 'q' to quit: ").strip().lower()
                if retry == 'q':
                    return None
                continue
        
        models = {}
        
        # Process each model
        for model_ul in model_uls:
            try:
                # Get model name
                model_name = model_ul.find_element(By.CSS_SELECTOR, 
                    "li.list-label div.link").get_attribute("textContent").strip()
                
                print(f"\nProcessing model: {model_name}")
                models[model_name] = {'themes': {}}
                
                # Get themes (excluding utility items)
                theme_items = model_ul.find_elements(By.CSS_SELECTOR, 
                    "li.icon:not(.new-theme):not(.retired-themes):not(.export-view):not(.filter-view):not(.viz-view)")
                
                for theme in theme_items:
                    try:
                        # Get theme name
                        theme_name = theme.find_element(By.CSS_SELECTOR, 
                            "a.link span.label").get_attribute("textContent").strip()
                        
                        if theme_name in ['New Theme', 'Retired Themes', 'Export', 'Filter', 'Visualizations']:
                            continue
                            
                        print(f"  Theme: {theme_name}")
                        models[model_name]['themes'][theme_name] = {'views': []}
                        
                        # Get views if they exist
                        try:
                            views = theme.find_elements(By.CSS_SELECTOR,
                                "div.listing ul.list li.entry:not(.non-entry) div.label")
                            
                            for view in views:
                                view_name = view.get_attribute("textContent").strip()
                                if view_name and view_name != 'New View':
                                    print(f"    View: {view_name}")
                                    models[model_name]['themes'][theme_name]['views'].append(view_name)
                        except:
                            continue
                            
                    except Exception as e:
                        print(f"Error processing theme: {str(e)}")
                        continue
                        
            except Exception as e:
                print(f"Error processing model: {str(e)}")
                continue
        
        return models
        
    except Exception as e:
        print(f"Error parsing Forms section: {str(e)}")
        return None

def generate_word_document(models_data):
    """Generate a Word document with the extracted model data"""
    try:
        # Create document
        doc = Document()
        
        # Add title
        title = doc.add_heading('Fluxx Build Documentation', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add timestamp
        timestamp = doc.add_paragraph()
        timestamp.add_run(f'Generated: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        doc.add_paragraph()  # Add spacing
        
        # Add content for each model
        for model_name, model_data in models_data.items():
            # Add model heading
            model_heading = doc.add_heading(f'Model: {model_name}', level=1)
            
            # Add themes and views
            for theme_name, theme_data in model_data['themes'].items():
                # Add theme
                theme_para = doc.add_paragraph()
                theme_para.add_run('Theme: ').bold = True
                theme_para.add_run(theme_name)
                
                # Add views
                for view_name in theme_data['views']:
                    view_para = doc.add_paragraph()
                    view_para.style = 'List Bullet'
                    view_para.add_run('View: ').bold = True
                    view_para.add_run(view_name)
            
            # Add spacing between models
            doc.add_paragraph()
        
        # Save the document
        timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f'fluxx_documentation_{timestamp_str}.docx'
        doc.save(filename)
        print(f"\nWord document saved as: {filename}")
        return filename
        
    except Exception as e:
        print(f"\nError generating Word document: {str(e)}")
        return None

def main():
    try:
        while True:  # Main program loop
            # Clear screen
            os.system('cls' if os.name == 'nt' else 'clear')
            print("=== Fluxx Build Documentation Data Scraper ===\n")
            
            # Get URL
            url = input("Enter Fluxx URL (e.g., example.fluxx.io): ").strip()
            if not url.startswith(('http://', 'https://')):
                url = 'https://' + url
            print(f"Using URL: {url}")
            
            # Find Chrome installation
            chrome_path = get_chrome_path()
            if not chrome_path:
                print("\nError: Chrome not found. Please install Google Chrome.")
                input("Press Enter to exit...")
                return
            
            print(f"\nFound Chrome at: {chrome_path}")
            
            # Setup ChromeDriver
            driver_path = get_resource_path("chromedriver.exe")
            
            # Chrome options
            options = webdriver.ChromeOptions()
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--log-level=3')  # Suppress most Chrome logs
            options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
            options.add_experimental_option('useAutomationExtension', False)
            
            # Create temp profile
            temp_dir = os.path.join(os.getcwd(), 'chrome_temp')
            if not os.path.exists(temp_dir):
                os.makedirs(temp_dir)
            options.add_argument(f'--user-data-dir={temp_dir}')
            
            print("\nStarting Chrome...")
            service = Service(driver_path, log_path='NUL')  # Suppress ChromeDriver logs
            driver = webdriver.Chrome(service=service, options=options)
            
            print(f"Navigating to {url}")
            driver.get(url)
            
            current_url = driver.current_url
            print(f"Current URL: {current_url}")
            
            if current_url == "data:,":
                print("\nNavigation failed. Please check your internet connection.")
                return
            
            # Wait for dashboard and navigate to Admin Panel
            if not wait_for_dashboard(driver):
                print("\nError: Could not detect dashboard load.")
                input("Press Enter to exit...")
                return
            
            print("\nDashboard detected! Navigating to Admin Panel...")
            
            # Navigate to Admin Panel
            if not navigate_to_admin(driver):
                print("\nError: Could not navigate to Admin Panel.")
                input("Press Enter to exit...")
                return
            
            # Parse the Forms section
            while True:  # Options loop
                # Get the data
                models_data = wait_for_forms_and_parse(driver)
                if not models_data:
                    print("\nError: Could not parse Forms section.")
                    retry_choice = input("Would you like to return to the main menu? (y/n): ").strip().lower()
                    if retry_choice == 'y':
                        break  # Break to main menu
                    else:
                        input("Press Enter to exit...")
                        return

                print("\nWhat would you like to do?")
                print("1. Generate Word document")
                print("2. Re-run scan")
                print("3. Exit")
                
                choice = input("\nEnter your choice (1-3): ").strip()
                
                if choice == '1':
                    doc_filename = generate_word_document(models_data)
                    if doc_filename:
                        print(f"\nDocumentation has been saved to: {doc_filename}")
                    
                    print("\nWould you like to:")
                    print("1. Return to options")
                    print("2. Exit")
                    sub_choice = input("\nEnter your choice (1-2): ").strip()
                    if sub_choice == '2':
                        return
                        
                elif choice == '2':
                    print("\nRe-running scan...")
                    continue  # Continue the loop to re-run wait_for_forms_and_parse
                    
                elif choice == '3':
                    return
                    
                else:
                    print("\nInvalid choice. Please try again.")
            
        input("\nPress Enter when finished to close Chrome...")
        
    except Exception as e:
        print(f"\nError: {str(e)}")
        input("Press Enter to exit...")
    finally:
        if 'driver' in locals():
            driver.quit()
            # Clean up temp directory
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except:
                pass

if __name__ == "__main__":
    main()

