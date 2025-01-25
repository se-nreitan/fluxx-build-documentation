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
        print("\n" + "=" * 80)
        print("\nLogin Instructions:")
        print("\n1. Enter your credentials in the browser window")
        print("2. Click 'Sign In' or 'Log In'")
        print("3. If prompted, select your Admin profile")
        print("\nIMPORTANT:")
        print("- Wait for the dashboard to fully load")
        print("- Once logged in, avoid clicking any elements in the browser")
        print("- The script will automatically navigate to required sections")
        print("\n" + "=" * 80)
        
        def wait_for_admin():
            wait = WebDriverWait(driver, timeout)
            return wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'a.to-admin-panel[href="/?db=config"]'))
            )
            
        wait_with_spinner("Waiting for login...", wait_for_admin)
        print("\nLogin successful!")
        return True
        
    except TimeoutException:
        print("\nLogin timed out")
        return False
    except Exception as e:
        print(f"\nLogin error: {str(e)}")
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
        # Clear screen and show header
        os.system('cls' if os.name == 'nt' else 'clear')
        print("\n" + "=" * 80)
        print("\n                     Scanning Models and Themes")
        print("\n" + "=" * 80)
        
        wait = WebDriverWait(driver, 10)
        
        # First verify Forms section exists
        wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "#iconList")
        ))
        
        # Wait for at least one model to be loaded
        try:
            wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "#iconList > ul[id]")) > 0)
        except TimeoutException:
            wait = WebDriverWait(driver, 30)
            wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "#iconList > ul[id]")) > 0)
        
        time.sleep(2)
        
        while True:
            model_list = driver.find_elements(By.CSS_SELECTOR, "#iconList > ul[id]")
            model_list = [model for model in model_list if model.get_attribute("id")]
            model_count = len(model_list)
            
            # Get the height of the iconList container
            icon_list = driver.find_element(By.CSS_SELECTOR, "#iconList")
            list_height = driver.execute_script("return arguments[0].scrollHeight", icon_list)
            
            # Scroll through the list to ensure all models are loaded
            current_scroll = 0
            scroll_step = 500
            
            while current_scroll < list_height:
                driver.execute_script(f"arguments[0].scrollTop = {current_scroll}", icon_list)
                time.sleep(0.5)
                current_scroll += scroll_step
            
            driver.execute_script("arguments[0].scrollTop = 0", icon_list)
            
            model_list = driver.find_elements(By.CSS_SELECTOR, "#iconList > ul[id]")
            model_list = [model for model in model_list if model.get_attribute("id")]
            new_count = len(model_list)
            
            if new_count > model_count:
                model_count = new_count
                continue
            
            print(f"\nFound {model_count} models.")
            verify = input("Does this count appear correct? (y/n): ").strip().lower()
            
            if verify == 'y':
                break
            print("\nWaiting for more models to load...")
            time.sleep(3)
                
        # Initialize dictionary to store model data
        models = {}
        current_model = 0
        total_models = len(model_list)
        
        # Print initial progress bar
        sys.stdout.write("\rScanning Models: [--------------------------------------------------] 0.0% (0/{})".format(total_models))
        sys.stdout.flush()
        
        # Process each model
        for model_ul in model_list:
            try:
                current_model += 1
                progress = (current_model / total_models) * 100
                
                # Create progress bar
                bar_length = 50
                filled_length = int(bar_length * current_model // total_models)
                bar = '=' * filled_length + '-' * (bar_length - filled_length)
                
                # Update progress bar
                sys.stdout.write('\r' + ' ' * 100)
                sys.stdout.write('\rScanning Models: [{0}] {1:.1f}% ({2}/{3})'.format(
                    bar, progress, current_model, total_models
                ))
                sys.stdout.flush()
                
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
                    pass
                
                # Check if model is dynamic
                is_dynamic = model_type and model_type.startswith('MacModelTypeDyn')
                
                models[model_name] = {
                    'type': model_type,
                    'is_dynamic': is_dynamic,
                    'themes': {}
                }
                
                # Find themes within this model's UL
                theme_items = model_ul.find_elements(By.CSS_SELECTOR, 
                    "li.icon[data-card-uid]")
                
                for theme in theme_items:
                    try:
                        theme_label = theme.find_element(By.CSS_SELECTOR, 
                            "a.link.scroll-to-card > span.label[role='menuitem']")
                        theme_name = theme_label.get_attribute("textContent").strip()
                        
                        if theme_name and theme_name not in ['New Theme', 'Retired Themes', 'Export', 'Filter', 'Visualizations']:
                            models[model_name]['themes'][theme_name] = {'views': []}
                            
                            # Find and process views
                            listing_div = theme.find_element(By.CSS_SELECTOR, 
                                "div.listing[data-type='listing'][data-src='/stencils']")
                            
                            views = listing_div.find_elements(By.CSS_SELECTOR,
                                "ul.list > li.entry:not(.non-entry)")
                            
                            for view in views:
                                try:
                                    label_div = view.find_element(By.CSS_SELECTOR, 
                                        "a.to-detail > div.label")
                                    view_name = label_div.get_attribute("textContent").strip()
                                    
                                    if view_name and view_name != 'New View':
                                        models[model_name]['themes'][theme_name]['views'].append(view_name)
                                except:
                                    continue
                    except:
                        continue
                        
            except:
                continue
        
        # Show completion
        sys.stdout.write('\r' + ' ' * 100)
        sys.stdout.write('\rModel scanning complete!')
        sys.stdout.flush()
        print("\n" + "=" * 80)
        
        return models
        
    except Exception as e:
        print(f"\nError parsing Forms section: {str(e)}")
        return None

def generate_word_document(models_data, site_url=None):
    """Generate a Word document using the Social Edge template format"""
    try:
        # Create document
        doc = Document()
        
        # Set standard margins (1 inch = 1440 twips)
        sections = doc.sections
        for section in sections:
            # Set all margins to 1 inch
            section.left_margin = 1440000 // 1000  # 1 inch
            section.right_margin = 1440000 // 1000  # 1 inch
            section.top_margin = 1440000 // 1000  # 1 inch
            section.bottom_margin = 1440000 // 1000  # 1 inch
            # Set page dimensions
            section.page_width = Pt(8.5 * 72)  # 8.5 inches
            section.page_height = Pt(11 * 72)  # 11 inches
        
        # Set default font and styles
        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)
        style.paragraph_format.space_after = Pt(12)  # Add spacing after paragraphs
        
        # Update heading styles
        for heading_level in range(1, 4):
            style = doc.styles[f'Heading {heading_level}']
            style.font.name = 'Calibri'
            style.font.bold = True
            style.paragraph_format.space_before = Pt(18)  # Add spacing before headings
            style.paragraph_format.space_after = Pt(12)  # Add spacing after headings
            if heading_level == 1:
                style.font.size = Pt(16)
            elif heading_level == 2:
                style.font.size = Pt(14)
            else:
                style.font.size = Pt(12)
        
        # Title with spacing
        doc.add_paragraph()  # Add space before title
        title = doc.add_heading('Fluxx Build Documentation', 0)
        title.style.font.name = 'Calibri'
        title.style.font.size = Pt(20)
        title.style.font.bold = True
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        subtitle = doc.add_paragraph('via Social Edge Consulting')
        subtitle.style = doc.styles['Normal']
        subtitle.runs[0].italic = True
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add space after subtitle
        doc.add_paragraph()
        doc.add_paragraph()
        
        # Client Information section
        info_heading = doc.add_heading('Client Information', 1)
        info_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        # Calculate table dimensions
        available_width = section.page_width - (section.left_margin + section.right_margin)
        table_width = int(available_width * 0.9)  # 90% of available width
        
        # Client Info table
        table = doc.add_table(rows=7, cols=2)
        table.style = 'Table Grid'
        table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        table.width = table_width
        
        # Set column widths
        col_widths = [0.3, 0.7]  # 30% and 70% of table width
        for i, width in enumerate(col_widths):
            for cell in table.columns[i].cells:
                cell.width = int(table_width * width)
        
        # Add client info headers and populate URL
        headers = ['URLs:', 'Product Enhancements:', 'Integrations:', 
                  'Social Edge Implementation Team:', 'Project Team:', 
                  'Admins:', 'Notes:']
        for i, header in enumerate(headers):
            cells = table.rows[i].cells
            cells[0].text = header
            cells[0].paragraphs[0].runs[0].bold = True
            cells[0].paragraphs[0].runs[0].font.name = 'Calibri'
            
            # Add URL to the first row if available
            if i == 0 and site_url:
                cells[1].text = site_url
                cells[1].paragraphs[0].runs[0].font.name = 'Calibri'
            
            # Add padding to cells
            for cell in cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.space_before = Pt(6)
                    paragraph.paragraph_format.space_after = Pt(6)
        
        # Models Section
        doc.add_page_break()
        models_heading = doc.add_heading('Models', 1)
        models_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        # Add each model
        for model_name, model_data in models_data.items():
            try:
                # Add space before model table
                doc.add_paragraph()
                
                # Model table
                table = doc.add_table(rows=9, cols=2)
                table.style = 'Table Grid'
                table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                table.width = table_width
                
                # Set column widths
                for i, width in enumerate(col_widths):
                    for cell in table.columns[i].cells:
                        cell.width = int(table_width * width)
                
                # Model name header with type and dynamic indicator
                header_row = table.rows[0]
                if len(header_row.cells) >= 2:
                    header_row.cells[0].merge(header_row.cells[1])
                    model_type = model_data.get('type', '')
                    is_dynamic = model_data.get('is_dynamic', False)
                    header_text = model_name
                    if model_type:
                        header_text += f" ({model_type})"
                    if is_dynamic:
                        header_text += " - Dynamic Model"
                    header_cell = header_row.cells[0]
                    header_cell.text = header_text
                    header_para = header_cell.paragraphs[0]
                    header_para.style = doc.styles['Heading 3']
                    header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in header_para.runs:
                        run.font.name = 'Calibri'
                
                # Themes section
                if len(table.rows) > 1:
                    row = table.rows[1].cells
                    if len(row) >= 2:
                        row[0].text = 'Themes:'
                        row[0].paragraphs[0].runs[0].bold = True
                        row[0].paragraphs[0].runs[0].font.name = 'Calibri'
                        
                        # Add themes and views
                        themes_text = []
                        for theme_name, theme_data in model_data['themes'].items():
                            theme_section = [f"\n{theme_name}"]
                            
                            # Add views
                            for view in theme_data['views']:
                                theme_section.append(f"• {view}")
                            
                            themes_text.append('\n'.join(theme_section))
                        
                        row[1].text = f"{len(model_data['themes'])} Themes Built\n"
                        row[1].paragraphs[0].runs[0].font.name = 'Calibri'
                        row[1].paragraphs[0].runs[0].bold = True
                        
                        # Add themes list
                        themes_para = row[1].add_paragraph('\n'.join(themes_text))
                        for run in themes_para.runs:
                            run.font.name = 'Calibri'
                
                # Add other rows with placeholders
                placeholders = [
                    'Workflow:', 'Add Card Menu:', 'Method:', 
                    'Before New / After Create:', 
                    'Before Validation / After Enter / Guard Instructions:',
                    'Documents:', 'Embedded Cards / Dynamic Relationships:',
                    'Notes:'
                ]
                
                for i, placeholder in enumerate(placeholders, start=2):
                    if i < len(table.rows):
                        cells = table.rows[i].cells
                        if len(cells) >= 2:
                            cells[0].text = placeholder
                            cells[0].paragraphs[0].runs[0].bold = True
                            cells[0].paragraphs[0].runs[0].font.name = 'Calibri'
                            
                            # Add theme code to Before New / After Create cell
                            if placeholder == 'Before New / After Create:':
                                code_text = []
                                for theme_name, theme_data in model_data['themes'].items():
                                    if 'code' in theme_data:
                                        code = theme_data['code']
                                        theme_code = []
                                        
                                        # Add theme name as header
                                        theme_code.append(f"\n{theme_name}:")
                                        
                                        # Add Current Before New code
                                        current_before = code.get('current_before_new', 'N/A')
                                        if current_before != 'N/A':
                                            theme_code.append("\nCurrent Before New Block:")
                                            theme_code.append(current_before)
                                            
                                        # Add Draft Before New code
                                        draft_before = code.get('draft_before_new', 'N/A')
                                        if draft_before != 'N/A':
                                            theme_code.append("\nDraft Before New Block:")
                                            theme_code.append(draft_before)
                                            
                                        # Add Current After Create code
                                        current_after = code.get('current_after_create', 'N/A')
                                        if current_after != 'N/A':
                                            theme_code.append("\nCurrent After Create Block:")
                                            theme_code.append(current_after)
                                            
                                        # Add Draft After Create code
                                        draft_after = code.get('draft_after_create', 'N/A')
                                        if draft_after != 'N/A':
                                            theme_code.append("\nDraft After Create Block:")
                                            theme_code.append(draft_after)
                                            
                                        if any(code != 'N/A' for code in [current_before, draft_before, current_after, draft_after]):
                                            code_text.append('\n'.join(theme_code))
                                
                                if code_text:
                                    code_para = cells[1].add_paragraph('\n'.join(code_text))
                                    for run in code_para.runs:
                                        run.font.name = 'Consolas'  # Use monospace font for code
                                else:
                                    cells[1].text = 'No code blocks configured'
                                    cells[1].paragraphs[0].runs[0].font.name = 'Calibri'
                
                # Add space after model table
                if model_name != list(models_data.keys())[-1]:
                    doc.add_page_break()
                    
            except Exception as e:
                print(f"\nError processing model {model_name}: {str(e)}")
                continue
        
        # Add remaining sections
        doc.add_page_break()
        portals_heading = doc.add_heading('Portals', 1)
        portals_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        doc.add_paragraph()  # Add space after heading
        
        doc.add_page_break()
        other_heading = doc.add_heading('Other Build Considerations', 1)
        other_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        doc.add_paragraph()  # Add space after heading
        
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

def wait_for_modal_load(driver, timeout=10):
    """Wait for modal to fully load"""
    try:
        wait = WebDriverWait(driver, timeout)
        modal = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div.modal.new-modal.area[style*='opacity: 1']")
        ))
        # Additional wait for content
        wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "div.modal textarea.code-to-submit")
        ))
        return modal
    except TimeoutException:
        return None

def safely_close_modal(driver, timeout=10):
    """Safely close the modal window"""
    try:
        wait = WebDriverWait(driver, timeout)
        close_button = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "a.close-modal")
        ))
        close_button.click()
        # Wait for modal to disappear
        wait.until(EC.invisibility_of_element_located(
            (By.CSS_SELECTOR, "div.modal.new-modal.area")
        ))
        return True
    except:
        return False

def close_all_models(driver):
    """Close all open models"""
    try:
        # Find all open models
        open_models = driver.find_elements(By.CSS_SELECTOR, "ul.toggle-class.open")
        for model in open_models:
            # Remove the open class
            driver.execute_script("arguments[0].classList.remove('open');", model)
        time.sleep(0.5)  # Brief pause to let animations complete
    except Exception as e:
        print(f"Warning: Could not close all models: {str(e)}")

def ensure_model_open(driver, model_ul, model_name):
    """Ensure a model is open and ready for processing"""
    try:
        # Check if model is already open
        if 'open' not in model_ul.get_attribute('class').split():
            # Find and click the model header
            model_header = model_ul.find_element(By.CSS_SELECTOR, "li.list-label div.link.is-admin")
            driver.execute_script("arguments[0].click();", model_header)
            
            # Wait for the open class to appear
            wait = WebDriverWait(driver, 10)
            wait.until(lambda d: 'open' in model_ul.get_attribute('class').split())
            time.sleep(1)  # Additional pause to let content load
            
            # Verify the model opened
            if 'open' not in model_ul.get_attribute('class').split():
                return False
        return True
    except Exception as e:
        return False

def get_theme_code(driver, theme_element, model_ul, model_name):
    """Get Before/After code for a theme"""
    try:
        # Ensure model is open before processing themes
        if not ensure_model_open(driver, model_ul, model_name):
            return None
        
        # Now ensure the theme is visible by clicking the theme name
        try:
            theme_link = theme_element.find_element(By.CSS_SELECTOR, 
                "a.link.scroll-to-card")
            theme_link.click()
            time.sleep(1)
        except Exception as e:
            return None
            
        # Wait for theme to be visible and expanded
        wait = WebDriverWait(driver, 10)
        wait.until(EC.visibility_of(theme_element))
        
        # Find and click the gear icon
        try:
            gear_icon = theme_element.find_element(By.CSS_SELECTOR, 
                "a.to-modal.open-config[data-on-success='matchListItem,close']")
            driver.execute_script("arguments[0].scrollIntoView(true);", gear_icon)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", gear_icon)
        except Exception as e:
            return None
        
        # Wait for modal to load
        modal = wait_for_modal_load(driver)
        if not modal:
            return None
            
        # Find all code textareas with correct IDs
        code_blocks = {
            'current_before_new': {
                'id': 'model_theme_unsafe_before_new_block',
                'label': 'Current Before New Block'
            },
            'draft_before_new': {
                'id': 'model_theme_draft_before_new_block',
                'label': 'Draft Before New Block'
            },
            'current_after_create': {
                'id': 'model_theme_unsafe_after_create_block',
                'label': 'Current After Create Block'
            },
            'draft_after_create': {
                'id': 'model_theme_draft_after_create_block',
                'label': 'Draft After Create Block'
            }
        }
        
        code_data = {}
        for key, block in code_blocks.items():
            try:
                textarea = modal.find_element(By.CSS_SELECTOR, f"textarea#{block['id']}")
                code = textarea.get_attribute("value").strip()
                code_data[key] = code if code else "N/A"
            except:
                code_data[key] = "N/A"
        
        # Safely close the modal
        safely_close_modal(driver)
        return code_data
        
    except Exception as e:
        try:
            safely_close_modal(driver)
        except:
            pass
        return None

def gather_theme_code(driver, models):
    """Gather Before/After code for all themes"""
    print("\n" + "=" * 80)
    print("\n                     Theme Code Gathering Process")
    print("\n" + "=" * 80 + "\n")
    print("This process will gather Before/After code from all themes automatically.")
    print("This may take several minutes to complete depending on the number of models.")
    print("\nIMPORTANT:")
    print("- You can stop the process at any time by pressing Ctrl+C")
    print("- Please do not interact with the browser while the process is running")
    print("- The browser will automatically handle all interactions")
    print("\n" + "-" * 80)
    verify = input("\nWould you like to proceed with gathering code? (y/n): ").strip().lower()
    
    if verify != 'y':
        print("\nSkipping code gathering process.")
        return models
    
    total_models = len(models)
    current_model = 0
    
    # Clear screen and show initial progress
    os.system('cls' if os.name == 'nt' else 'clear')
    print("\n" + "=" * 80)
    print("\n                     Theme Code Gathering Process")
    print("\nPress Ctrl+C to stop the process at any time")
    print("Please wait while code is gathered from all themes...")
    print("\n" + "=" * 80)
    
    # Print initial progress bar line
    sys.stdout.write("\rProcessing Models: [--------------------------------------------------] 0.0% (0/{})".format(total_models))
    sys.stdout.flush()
    
    try:
        for model_name, model_data in models.items():
            current_model += 1
            progress = (current_model / total_models) * 100
            
            # Create progress bar
            bar_length = 50
            filled_length = int(bar_length * current_model // total_models)
            bar = '=' * filled_length + '-' * (bar_length - filled_length)
            
            # Update progress bar (stay on same line, no extra newlines)
            sys.stdout.write('\r' + ' ' * 100)  # Clear current line
            sys.stdout.write('\rProcessing Models: [{0}] {1:.1f}% ({2}/{3})'.format(
                bar, progress, current_model, total_models
            ))
            sys.stdout.flush()
            
            try:
                model_ul = driver.find_element(By.CSS_SELECTOR, f"ul#{model_name.lower().replace(' ', '_')}")
                
                # Ensure model is open
                if not ensure_model_open(driver, model_ul, model_name):
                    continue
                
                for theme_name, theme_data in model_data['themes'].items():
                    theme_elements = model_ul.find_elements(By.CSS_SELECTOR, "li.icon[data-card-uid]")
                    
                    for theme_element in theme_elements:
                        try:
                            label = theme_element.find_element(By.CSS_SELECTOR, 
                                "a.link.scroll-to-card span.label").text
                            
                            if label == theme_name:
                                code_data = get_theme_code(driver, theme_element, model_ul, model_name)
                                if code_data:
                                    models[model_name]['themes'][theme_name]['code'] = code_data
                                break
                        except:
                            continue
                
                # Close model after processing
                if 'open' in model_ul.get_attribute('class').split():
                    model_header = model_ul.find_element(By.CSS_SELECTOR, "li.list-label div.link.is-admin")
                    driver.execute_script("arguments[0].click();", model_header)
                    time.sleep(1)
                        
            except:
                continue
        
        # Clear line and show completion message
        sys.stdout.write('\r' + ' ' * 100)  # Clear current line
        sys.stdout.write('\rCode gathering process complete!')
        sys.stdout.flush()
        print("\n" + "=" * 80)
        return models
        
    except KeyboardInterrupt:
        # Handle Ctrl+C gracefully
        sys.stdout.write('\r' + ' ' * 100)  # Clear current line
        sys.stdout.write('\rProcess stopped by user.')
        sys.stdout.flush()
        print("\n" + "=" * 80)
        return models

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
                    
            # Gather theme code if requested
            models_data = gather_theme_code(driver, models_data)

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

