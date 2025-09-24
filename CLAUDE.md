# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a BNI (Business Network International) Lifetime data analysis project focused on scraping and analyzing TYFCB (Thank You For Closed Business) data from BNI Connect Global. The project contains multiple Python scripts for web scraping, data processing, and report generation.

## Key Scripts

### BNI-Lifetime-Selenuim-V5.py (Latest)
- **Main functionality**: Web scraping using Selenium to extract TYFCB Given Report data from BNI Connect Global
- **Dependencies**: selenium, webdriver-manager, pandas, openpyxl, csv, datetime
- **Key features**:
  - Automated Chrome WebDriver setup
  - Login automation with password masking
  - Iframe handling for nested report data
  - CSV and Excel export functionality
  - Screenshot capture for debugging

### BNI-Lifetime-V2.py (HTTP Requests approach)
- **Alternative approach**: Uses requests and BeautifulSoup for web scraping
- **Dependencies**: requests, beautifulsoup4, getpass, pandas, openpyxl
- **Purpose**: HTTP-based scraping as fallback to Selenium approach

### BNI-Lifetime.py (Data Analysis)
- **Purpose**: Excel data processing and growth analysis
- **Functionality**: Analyzes BNI member business growth over 12-month periods
- **Output**: Generates growth reports with month-over-month changes

## Architecture

### Web Scraping Architecture
- **Primary method**: Selenium WebDriver (Chrome) with automated driver management
- **Fallback method**: HTTP requests with session handling
- **Target site**: https://www.bniconnectglobal.com
- **Authentication**: Username/password login with CSRF token handling
- **Data extraction**: Multi-level iframe navigation to access embedded reports

### Data Processing Flow
1. Login to BNI Connect Global
2. Navigate to Dashboard → Lifetime section
3. Click Review button for TYFCB Given reports
4. Extract table data from nested iframes
5. Export to CSV and Excel formats
6. Generate summary reports with screenshots

### Key Functions
- `setup_driver()`: Chrome WebDriver initialization with anti-detection features
- `get_tyfcb_given_report_data()`: Complex iframe navigation and data extraction
- `export_tyfcb_given_report()`: Automated Excel export functionality
- `login_and_get_tyfcb()`: Main orchestration function

## Data Structure

### Input Data (Excel files in 00-ข้อมูลดิบจากggsheet/)
- BNI member lifetime business data
- Columns: 'กรุณาระบุชื่อ' (Name), 'ยอดธุรกิจ Lifetime' (Lifetime Business Amount), 'Timestamp'

### Output Files
- `tyfcb_given_report.csv`: Raw scraped data in CSV format
- `BNI_Growth_Report_.xlsx`: Processed growth analysis
- Screenshot files: `tyfcb_given_report.png`, `final_screen.png`, `table_screenshot.png`

## Dependencies Installation

```bash
pip install selenium webdriver-manager pandas openpyxl beautifulsoup4 requests
```

## Running the Scripts

### For Web Scraping (Latest version):
```bash
python BNI-Lifetime-Selenuim-V5.py
```

### For Data Analysis:
```bash
python BNI-Lifetime.py
```

## Technical Notes

### Selenium Configuration
- Uses Chrome with maximized window and anti-automation detection features
- Implements retry logic for WebDriver installation failures
- Handles Windows-specific password input with asterisk masking

### Iframe Handling
- Supports nested iframe navigation (main iframe → nested iframe)
- Implements fallback strategies for different iframe structures
- Uses explicit waits for element loading

### Error Handling
- Comprehensive exception handling with screenshot capture
- Detailed logging for debugging iframe and element location issues
- Fallback mechanisms for different website structures

### Security Considerations
- Password input masking for console applications
- No hardcoded credentials in source code
- Session management for authenticated requests