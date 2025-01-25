=== Fluxx Build Documentation Tool ===

This tool automates the documentation of Fluxx build configurations by scanning the Admin Panel and generating a standardized Word document.

Key Features:
- Automated scanning of Forms section in Admin Panel
- Extraction of Models, Themes, and Views
- Collection of Before/After code blocks from each Theme
- Generation of formatted Word documentation

Requirements:
- Google Chrome browser
- Internet connection
- Fluxx admin credentials

Quick Start:
1. Run the tool
2. Enter your Fluxx URL (e.g., example.fluxx.io)
3. Log in when prompted
4. Wait for model scanning to complete
5. Confirm to gather code from themes
6. Generate Word documentation

Output:
The tool generates a Word document (fluxx_documentation_YYYYMMDD_HHMMSS.docx) containing:
- Client Information section
- Detailed Model documentation including:
  - Model name and type
  - Associated Themes and Views
  - Before New and After Create code blocks
- Portals section
- Other Build Considerations section

Note: Please do not interact with the browser while the tool is running. The process is automated and will handle all necessary interactions. 