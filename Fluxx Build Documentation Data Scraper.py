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
import threading

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

def print_divider():
    """Print a visual divider line"""
    print("\n" + "=" * 50 + "\n")

def wait_with_spinner(message, action_func, *args, **kwargs):
    """Execute a function while showing a loading spinner"""
    stop_spinner = threading.Event()
    spinner_thread = threading.Thread(
        target=show_spinner, 
        args=(stop_spinner, message)
    )
    spinner_thread.start()
    
    try:
        result = action_func(*args, **kwargs)
        stop_spinner.set()
        spinner_thread.join()
        return result
    except Exception as e:
        stop_spinner.set()
        spinner_thread.join()
        raise e

def wait_for_dashboard(driver, timeout=60):
    """Wait for dashboard to load and verify we're logged in"""
    try:
        print("Waiting for dashboard to load...")
        print("Please:")
        print("1. Log in with your credentials")
        print("2. Select your Admin profile (if applicable)")
        print("3. System will automatically proceed when dashboard loads\n")
        
        def wait_for_admin():
            wait = WebDriverWait(driver, timeout)
            return wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'a.to-admin-panel[href="/?db=config"]'))
            )
            
        wait_with_spinner("Waiting for successful login...", wait_for_admin)
        print("\nDashboard loaded successfully!")
        return True
        
    except TimeoutException:
        print("Timed out waiting for login")
        return False
    except Exception as e:
        print(f"Error waiting for login: {str(e)}")
        return False

def navigate_to_admin(driver):
    """Navigate to Admin Panel"""
    try:
        def nav_action():
            wait = WebDriverWait(driver, 10)
            admin_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.to-admin-panel[href="/?db=config"]'))
            )
            print("Found Admin Panel button")
            admin_button.click()
            wait.until(EC.url_contains('db=config'))
            
        wait_with_spinner("Navigating to Admin Panel...", nav_action)
        print("Successfully navigated to Admin Panel")
        return True
        
    except Exception as e:
        print(f"Error navigating to Admin Panel: {str(e)}")
        return False

# HTML Structure Reference for Forms Section:
#
# Models:
# - Located in: #iconList > ul
# - Model name from: ul[id] attribute (replace _ with space, title case)
# Example: <ul id="custom_ui_enhancements" class="toggle-class open" data-click-when-opened=".scroll-to-card">
#
# Themes:
# - Located in: ul > li.icon[data-card-uid]
# - Theme name from: li.icon > a.link.scroll-to-card > span.label[role='menuitem']
# Example:
# <li class="icon" data-card-uid="19980">
#   <a class="link scroll-to-card" href="#fluxx-card-26">
#     <span class="label" role="menuitem">Animation Examples</span>
#
# Views:
# - Container: li.icon > div.listing[data-type='listing'][data-src='/stencils']
# - View items: div.listing > ul.list > li.entry:not(.non-entry)
# - View name from: li.entry > a.to-detail > div.label
# Example:
# <div class="listing" data-type="listing" data-src="/stencils" ...>
#   <ul class="list">
#     <li class="active entry selected" data-model-id="35727">
#       <a class="to-detail" href="/stencils/35727">
#         <div class="label">Gallery</div>

def wait_for_forms_and_parse(driver, max_retries=3):
    """Wait for Forms section to load and parse content with retry logic"""
    try:
        print("\nParsing Forms section...")
        
        wait = WebDriverWait(driver, 10)
        
        # First verify Forms section exists
        wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "#iconList")
        ))
        
        while True:  # Loop until user confirms correct model count
            # Get model count by finding all model ULs in the iconList
            model_list = driver.find_elements(By.CSS_SELECTOR, "#iconList > ul")
            # Filter out any empty or invalid ULs
            model_list = [model for model in model_list if model.get_attribute("id")]
            model_count = len(model_list)
            
            # Ask user to verify count
            print(f"\nFound {model_count} models in the Forms section.")
            verify = input("Does this count appear correct? (y/n): ").strip().lower()
            
            if verify == 'y':
                break
            else:
                print("\nPlease wait for all models to load and try again...")
                time.sleep(2)
                continue
                
        # Initialize dictionary to store model data
        models = {}
        
        # Process each model
        for model_ul in model_list:
            try:
                # Get model name from the UL id attribute
                model_id = model_ul.get_attribute("id")
                if not model_id:
                    continue
                    
                model_name = model_id.replace('_', ' ').title()
                
                # Find the "New Theme" link to extract model type
                model_type = None
                try:
                    # Try multiple selectors to find model type
                    selectors = [
                        "a.link.to-modal[href*='model_theme[model_type]']",
                        "a.link[href*='model_theme[model_type]']",
                        "a[href*='model_theme[model_type]']"
                    ]
                    
                    for selector in selectors:
                        try:
                            new_theme_link = model_ul.find_element(By.CSS_SELECTOR, selector)
                            if new_theme_link:
                                href = new_theme_link.get_attribute("href")
                                # Extract model type from URL parameter
                                match = re.search(r'model_theme\[model_type\]=(\w+)', href)
                                if match:
                                    model_type = match.group(1)
                                    break
                        except:
                            continue
                except:
                    # If we can't find the model type, continue without it
                    pass
                
                # Check if model is dynamic by checking if type starts with MacModelTypeDyn
                is_dynamic = model_type and model_type.startswith('MacModelTypeDyn')
                
                # Print model name with type and dynamic indicator if available
                output = f"\nModel: {model_name}"
                if model_type:
                    output += f" ({model_type})"
                if is_dynamic:
                    output += " - Dynamic Model"
                print(output)
                
                models[model_name] = {
                    'type': model_type,
                    'is_dynamic': is_dynamic,
                    'themes': {}
                }
                
                # Find themes within this model's UL - looking for li.icon elements
                theme_items = model_ul.find_elements(By.CSS_SELECTOR, 
                    "li.icon[data-card-uid]")
                
                for theme in theme_items:
                    try:
                        # Get theme name from the exact path: li.icon > a.link > span.label
                        theme_label = theme.find_element(By.CSS_SELECTOR, 
                            "a.link.scroll-to-card > span.label[role='menuitem']")
                        theme_name = theme_label.get_attribute("textContent").strip()
                        
                        if theme_name and theme_name not in ['New Theme', 'Retired Themes', 'Export', 'Filter', 'Visualizations']:
                            print(f"  Theme: {theme_name}")
                            models[model_name]['themes'][theme_name] = {'views': []}
                            
                            # Find the listing div that contains views - using exact path
                            listing_div = theme.find_element(By.CSS_SELECTOR, 
                                "div.listing[data-type='listing'][data-src='/stencils']")
                            
                            # Get all view entries from the list - using exact path
                            views = listing_div.find_elements(By.CSS_SELECTOR,
                                "ul.list > li.entry:not(.non-entry)")
                            
                            for view in views:
                                try:
                                    # Get view name using exact path: li.entry > a.to-detail > div.label
                                    label_div = view.find_element(By.CSS_SELECTOR, 
                                        "a.to-detail > div.label")
                                    view_name = label_div.get_attribute("textContent").strip()
                                    
                                    if view_name and view_name != 'New View':
                                        print(f"    View: {view_name}")
                                        models[model_name]['themes'][theme_name]['views'].append(view_name)
                                except Exception as e:
                                    print(f"    Error getting view name: {str(e)}")
                                    continue
                                    
                    except Exception as e:
                        print(f"Error processing theme: {str(e)}")
                        continue
                        
            except Exception as e:
                print(f"Error processing model {model_name if 'model_name' in locals() else 'Unknown'}: {str(e)}")
                continue
        
        return models
        
    except Exception as e:
        print(f"\nError parsing Forms section: {str(e)}")
        return None

def generate_word_document(models_data, site_url=None):
    """Generate a Word document using the Social Edge template format"""
    try:
        # Create document
        doc = Document()
        
        # Title
        doc.add_heading('Fluxx Build Documentation', 0)
        doc.add_paragraph('via Social Edge Consulting').italic = True
        
        # Add TOC placeholder
        doc.add_paragraph()
        
        # Client Information section
        doc.add_heading('Client Information', 1)
        table = doc.add_table(rows=7, cols=2)
        table.style = 'Table Grid'
        
        # Add client info headers and populate URL
        headers = ['URLs:', 'Product Enhancements:', 'Integrations:', 
                  'Social Edge Implementation Team:', 'Project Team:', 
                  'Admins:', 'Notes:']
        for i, header in enumerate(headers):
            cells = table.rows[i].cells
            cells[0].text = header
            cells[0].paragraphs[0].runs[0].bold = True
            
            # Add URL to the first row if available
            if i == 0 and site_url:
                cells[1].text = site_url
        
        # Models Section
        doc.add_page_break()
        doc.add_heading('Models', 1)
        
        # Add each model
        for model_name, model_data in models_data.items():
            try:
                # Model table
                table = doc.add_table(rows=9, cols=2)
                table.style = 'Table Grid'
                
                # Model name header with type and dynamic indicator
                header_row = table.rows[0]
                if len(header_row.cells) >= 2:  # Verify we have enough cells
                    header_row.cells[0].merge(header_row.cells[1])
                    model_type = model_data.get('type', '')
                    is_dynamic = model_data.get('is_dynamic', False)
                    header_text = model_name
                    if model_type:
                        header_text += f" ({model_type})"
                    if is_dynamic:
                        header_text += " - Dynamic Model"
                    header_row.cells[0].text = header_text
                    header_row.cells[0].paragraphs[0].style = doc.styles['Heading 3']
                
                # Themes section
                if len(table.rows) > 1:  # Verify we have a second row
                    row = table.rows[1].cells
                    if len(row) >= 2:  # Verify we have both cells
                        row[0].text = 'Themes:'
                        row[0].paragraphs[0].runs[0].bold = True
                        
                        # Add themes and views
                        themes_text = []
                        for theme_name, theme_data in model_data['themes'].items():
                            theme_section = [f"\n{theme_name}"]
                            for view in theme_data['views']:
                                theme_section.append(f"• {view}")
                            themes_text.append('\n'.join(theme_section))
                        
                        row[1].text = f"{len(model_data['themes'])} Themes Built\n"
                        row[1].text += '\n'.join(themes_text)
                
                # Add other rows as placeholders
                placeholders = [
                    'Workflow:', 'Add Card Menu:', 'Method:', 
                    'Before New / After Create:', 
                    'Before Validation / After Enter / Guard Instructions:',
                    'Documents:', 'Embedded Cards / Dynamic Relationships:',
                    'Notes:'
                ]
                
                for i, placeholder in enumerate(placeholders, start=2):
                    if i < len(table.rows):  # Verify row exists
                        cells = table.rows[i].cells
                        if len(cells) >= 2:  # Verify cells exist
                            cells[0].text = placeholder
                            cells[0].paragraphs[0].runs[0].bold = True
                
                # Add page break between models
                if model_name != list(models_data.keys())[-1]:
                    doc.add_page_break()
                    
            except Exception as e:
                print(f"\nError processing model {model_name}: {str(e)}")
                continue
        
        # Add remaining sections as placeholders
        doc.add_page_break()
        doc.add_heading('Portals', 1)
        # ... portal tables would go here
        
        doc.add_page_break()
        doc.add_heading('Other Build Considerations', 1)
        # ... other considerations table would go here
        
        # Save the document
        timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f'fluxx_documentation_{timestamp_str}.docx'
        doc.save(filename)
        print(f"\nWord document saved as: {filename}")
        return filename
        
    except Exception as e:
        print(f"\nError generating Word document: {str(e)}")
        return None

def show_spinner(stop_event, message=""):
    """Show a simple spinner animation with a message"""
    spinner = ['|', '/', '-', '\\']  # Simple ASCII spinner
    i = 0
    while not stop_event.is_set():
        print(f"\r{spinner[i]} {message}", end='', flush=True)
        i = (i + 1) % len(spinner)
        time.sleep(0.1)
    print("\r", end='', flush=True)  # Clear the spinner line

def print_header():
    """Print the application header with clean formatting"""
    os.system('cls' if os.name == 'nt' else 'clear')
    
    # Social Edge logo ASCII art
    print("""
                                                                                           
                             
                    ;X.            
                  ;XX+             
                .XX+  ..:.         
              .XX+  xXXxxXXX       
            .XXx  xXX.    .XX      
            XX  ;XX. .XX+  XX      
            Xx  .  .XXX  +XX.      
            ;XXx.;XXx  ;XX;        
              .+++:  :XX;          
                   :XX+            
                  +X+            
                                                           
       Social Edge Consulting - 2025
   nick.reitan@socialedgeconsulting.com                    
                                                                   
    """)
    
    print("  Fluxx Build Documentation Automated Tool    ")
    print_divider()

def check_chrome_and_driver():
    """Diagnose Chrome and ChromeDriver versions and compatibility"""
    try:
        print("Please wait while your Google Chrome version is verified...")
        print("(This will take a few moments)")
        print_divider()
        
        # Create a threading event to control the spinner
        stop_spinner = threading.Event()
        spinner_thread = threading.Thread(
            target=show_spinner, 
            args=(stop_spinner, "Checking Chrome setup...")
        )
        spinner_thread.start()
        
        # Get Chrome version
        chrome_path = get_chrome_path()
        if not chrome_path:
            stop_spinner.set()
            spinner_thread.join()
            print("Error: Chrome not found. Please install Google Chrome.")
            return False
            
        chrome_version = get_chrome_version()
        
        # Update spinner message
        stop_spinner.set()
        spinner_thread.join()
        print(f"Chrome version detected: {chrome_version}")
        print_divider()
        
        # Check if chromedriver exists and is compatible
        driver_path = os.path.join(os.getcwd(), "chromedriver.exe")
        if os.path.exists(driver_path):
            stop_spinner = threading.Event()
            spinner_thread = threading.Thread(
                target=show_spinner, 
                args=(stop_spinner, "Verifying ChromeDriver compatibility...")
            )
            spinner_thread.start()
            
            try:
                service = Service(driver_path, log_path='NUL')
                options = webdriver.ChromeOptions()
                options.add_argument('--headless')
                options.add_argument('--log-level=3')  # Suppress console messages
                options.add_experimental_option('excludeSwitches', ['enable-logging'])
                driver = webdriver.Chrome(service=service, options=options)
                driver.quit()
                stop_spinner.set()
                spinner_thread.join()
                print("\nSetup complete!")
                print_divider()
                time.sleep(0.5)
                return True
            except Exception:
                stop_spinner.set()
                spinner_thread.join()
                print("\nUpdating ChromeDriver to match Chrome version...")
        else:
            print("\nSetting up ChromeDriver...")
            
        # Download matching ChromeDriver if needed
        stop_spinner = threading.Event()
        spinner_thread = threading.Thread(
            target=show_spinner, 
            args=(stop_spinner, "Downloading and installing ChromeDriver...")
        )
        spinner_thread.start()
        
        chrome_major = chrome_version.split('.')[0]
        driver_url = f"https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/{chrome_version}/win64/chromedriver-win64.zip"
        
        try:
            response = requests.get(driver_url)
            response.raise_for_status()
            
            with open("chromedriver.zip", "wb") as f:
                f.write(response.content)
                
            with zipfile.ZipFile("chromedriver.zip", "r") as zip_ref:
                zip_ref.extractall()
                
            if os.path.exists("chromedriver.exe"):
                os.remove("chromedriver.exe")
            shutil.move("chromedriver-win64/chromedriver.exe", "chromedriver.exe")
            
            # Clean up
            os.remove("chromedriver.zip")
            shutil.rmtree("chromedriver-win64")
            
            stop_spinner.set()
            spinner_thread.join()
            print("Setup complete!")
            return True
            
        except Exception as e:
            stop_spinner.set()
            spinner_thread.join()
            print("\nError: Unable to setup ChromeDriver automatically")
            print("Please download ChromeDriver manually:")
            print(f"1. Visit: https://googlechromelabs.github.io/chrome-for-testing/")
            print(f"2. Download version matching Chrome {chrome_major}")
            print("3. Extract chromedriver.exe to the same folder as this program")
            return False
            
    except Exception as e:
        if 'stop_spinner' in locals() and not stop_spinner.is_set():
            stop_spinner.set()
            spinner_thread.join()
        print("\nError: Unable to verify Chrome setup")
        return False

def validate_fluxx_url(url):
    """Validate and format Fluxx URL"""
    url = url.strip().lower()
    
    # Remove any protocol prefixes
    if url.startswith(('http://', 'https://')):
        url = url.replace('http://', '').replace('https://', '')
    
    # Remove any trailing slashes
    url = url.rstrip('/')
    
    # Check if URL ends with .fluxx.io
    if not url.endswith('.fluxx.io'):
        if '.fluxx.io' in url:
            # Extract the part before .fluxx.io if it exists
            url = url.split('.fluxx.io')[0] + '.fluxx.io'
        else:
            # Append .fluxx.io if missing
            url = url + '.fluxx.io'
    
    return f'https://{url}'

def main():
    try:
        print_header()
        
        # Check Chrome setup
        if not check_chrome_and_driver():
            input("\nPress Enter to exit...")
            return
        
        # Clear screen and show fresh header before URL input
        print_header()
        
        # Get and validate URL
        while True:
            url = input("Enter Fluxx Instance Name or URL (e.g., 'example' or example.fluxx.io): ").strip()
            if not url:
                print("URL cannot be empty. Please try again.")
                continue
                
            url = validate_fluxx_url(url)
            print(f"\nUsing URL: {url}")
            verify = input("Is this correct? (y/n): ").strip().lower()
            if verify == 'y':
                break
            print()  # Add blank line before retry
            
        # Setup Chrome
        print("Starting Chrome...")
        driver_path = get_resource_path("chromedriver.exe")
        options = webdriver.ChromeOptions()
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--log-level=3')
        options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
        options.add_experimental_option('useAutomationExtension', False)
        
        # Create temp profile
        temp_dir = os.path.join(os.getcwd(), 'chrome_temp')
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        options.add_argument(f'--user-data-dir={temp_dir}')
        
        service = Service(driver_path, log_path='NUL')  # Suppress ChromeDriver logs
        driver = webdriver.Chrome(service=service, options=options)
        
        print(f"Navigating to {url}")
        driver.get(url)
        
        current_url = driver.current_url
        print(f"Current URL: {current_url}\n")  # Add newline after URL
        
        if current_url == "data:,":
            print("Navigation failed. Please check your internet connection.")
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
                print_divider()
                retry_choice = input("Would you like to return to the main menu? (y/n): ").strip().lower()
                if retry_choice == 'y':
                    break  # Break to main menu
                else:
                    input("Press Enter to exit...")
                    return

            print_divider()
            print("Available Actions:")
            print("1. Generate Word document")
            print("2. Re-run scan")
            print("3. Exit")
            
            choice = input("\nEnter your choice (1-3): ").strip()
            
            if choice == '1':
                doc_filename = wait_with_spinner(
                    "Generating Word document...", 
                    generate_word_document, 
                    models_data,
                    site_url=url  # Pass the URL to the document generator
                )
                if doc_filename:
                    print(f"\nDocumentation has been saved to: {doc_filename}")
                
                print_divider()
                print("Would you like to:")
                print("1. Return to options")
                print("2. Exit")
                sub_choice = input("\nEnter your choice (1-2): ").strip()
                if sub_choice == '2':
                    return
                    
            elif choice == '2':
                print("\nRe-running scan...")
                continue
                
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

