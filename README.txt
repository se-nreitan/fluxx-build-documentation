=== Fluxx Build Documentation Data Scraper ===

This tool automates the documentation of Fluxx build configurations by extracting model, theme, and view information from the Admin Panel.

Features:
- Automated navigation to Admin Panel
- Interactive model count verification
- Detailed extraction of:
  - All Models from the Forms section
  - Themes associated with each Model
  - Views configured for each Theme
- Real-time console output showing extraction progress
- Generation of formatted Word document with timestamp
- Option to re-scan without restarting the program

Output:
- Console display showing hierarchical structure:
  - Model names
  - Theme names under each Model
  - View names under each Theme
- Word document (fluxx_documentation_YYYYMMDD_HHMMSS.docx) containing formatted documentation

Requirements:
- Google Chrome browser
- Internet connection
- Fluxx admin credentials
- Appropriate permissions to access Admin Panel

Usage:
1. Run the executable
2. Enter your Fluxx URL (e.g., example.fluxx.io)
3. Log in when prompted
4. Select Admin profile if applicable
5. Verify the model count when prompted
6. Choose to:
   - Generate Word document
   - Re-run scan
   - Exit

Instructions:
1. Double-click Fluxx_Scraper.exe
2. Enter your Fluxx site URL
3. Log in with your credentials when the browser opens
4. Wait for the dashboard to load
5. Confirm the number of models found
6. Review the extracted data
7. Generate documentation or re-scan as needed

The Word document will be named 'fluxx_documentation_YYYYMMDD_HHMMSS.docx' with the current timestamp.

Note: This application uses Selenium WebDriver to automate Chrome browser interactions. It includes error handling and verification steps to ensure accurate data extraction.

If you encounter any issues:
- Verify you have admin access to Fluxx
- Ensure all models are loaded before confirming the count
- Try re-scanning if any data appears missing 