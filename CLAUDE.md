# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Selenium-based automated form filling system for a Chinese medical reporting website (y.chinadtc.org.cn). It reads data from Excel files and automatically fills/submits web forms with login support and diagnosis code lookup via API.

## Commands

### Running the Application
```bash
python main.py
```
The application now uses a GUI for configuration input. Users will see:
1. Configuration GUI for login credentials, month, function type, and data file selection
2. Confirmation dialog before starting automatic form filling
3. Confirmation dialog after completion for manual report submission

### Testing GUI Only
```bash
python test_gui.py
```

### Installing Dependencies
```bash
pip install -r requirements.txt
```

### Creating Sample Data
```bash
python create_sample_data.py
```

### Running Demo Script
```bash
python selenium_demo.py
```

## Architecture

### Module Structure

The codebase follows a modular architecture with clear separation of concerns:

- **main.py**: Entry point that orchestrates the entire workflow:
  1. Display GUI to collect configuration (username, password, month, function type, data file)
  2. Load config.yaml and override with GUI inputs
  3. Login flow
  4. Post-login confirmation handling
  5. Month selection
  6. Entry button navigation
  7. Function button selection (outpatient/emergency)
  8. Show confirmation dialog for manual pre-filling tasks
  9. Form filling and submission
  10. Show confirmation dialog for manual report submission
  11. Result export

- **src/gui.py**: Tkinter-based GUI module with:
  - ConfigGUI: Main configuration input window with validation
  - ConfirmationDialog: Modal dialogs for user confirmation during execution
  - Input fields: username, password, month, function type, data file path
  - File browser integration
  - Input validation (required fields, date format, file existence)

- **src/driver_manager.py**: Manages Chrome WebDriver lifecycle, handles headless mode configuration

- **src/login_handler.py**: Handles login flow with configurable element locators and success verification

- **src/form_filler.py**: Core form filling logic with:
  - Multi-type element support (input, select, radio, textarea, button)
  - Special handling for diagnosis fields (queries API to get diagnosis codes)
  - JavaScript-based element interaction to avoid occlusion issues
  - Alert dialog handling after submission

- **src/data_reader.py**: Reads Excel/CSV files using pandas, converts to dictionary format

- **src/result_exporter.py**: Exports processing results to timestamped Excel files

### Configuration-Driven Design

The system is heavily configuration-driven via `config/config.yaml`:

- **login**: Login credentials and element locators
- **month_selection**: Report month configuration
- **entry_button**: Navigation button configuration
- **function_button**: Function type selection (outpatient vs emergency)
- **functions**: Per-function configuration with separate:
  - Data file paths
  - Form element mappings
  - Element locators

### Key Implementation Details

**Element Locator Pattern**: All element interactions use a standardized pattern:
```python
LOCATOR_MAP = {
    "id": By.ID,
    "name": By.NAME,
    "xpath": By.XPATH,
    "css_selector": By.CSS_SELECTOR,
    # ...
}
```

**Diagnosis Lookup Flow** (form_filler.py:346-417):
- Uses HTTP requests to query `http://y.chinadtc.org.cn/entering/dict/search_dict`
- Copies browser cookies (PHPSESSID) to maintain session
- Sends multipart/form-data request with diagnosis keyword
- Parses JSON response and fills diagnosis name + code fields
- JavaScript is used to set diagnosis values to avoid triggering onfocus events

**Special Field Handling** (form_filler.py:74-113):
- `科室` field: Strips spaces and "门诊" text
- `年龄` field: Separates numeric age from unit suffix, fills both fields
- `诊断` field: Splits comma-separated diagnoses (max 5), queries each via API
- `注射剂` field: Auto-fills quantity field when "有" is selected

**JavaScript Interaction Strategy**:
- Confirmation buttons use `driver.execute_script("arguments[0].click()", element)` to avoid element occlusion (main.py:123-125, 244, 330)
- Readonly inputs (month selection) use JavaScript to set value and trigger change events (main.py:217-222)
- Diagnosis fields remove onfocus/onblur handlers before setting values (form_filler.py:161-170)

**Alert Handling Pattern**:
After clicking submit buttons, the code waits 1 second then handles alert dialogs:
```python
time.sleep(1)  # Wait for alert
alert = driver.switch_to.alert
alert.accept()
```

### Data Flow

1. GUI displays → User inputs credentials, month, function type, file path
2. GUI validates inputs → returns config dictionary
3. Config.yaml loaded → GUI values override specific fields (login credentials, month, function type, file path)
4. Excel file read → pandas DataFrame → list of dictionaries
5. Each row dictionary passed to FormFiller
6. FormFiller maps dictionary keys to form element locators from config
7. Special processing for diagnosis field (API lookup)
8. Form submission with alert handling
9. Results collected with status/message
10. Export all results to timestamped Excel file

### User Interaction Points

The application has three GUI-based user interaction points:
1. **Startup (gui.py:ConfigGUI)**: Collects configuration before execution starts
   - Can cancel to abort the entire process
2. **Pre-filling confirmation (main.py:593)**: Modal dialog waits for user to manually input report date and total volume
   - Cannot be closed without confirming
3. **Post-filling confirmation (main.py:668)**: Modal dialog waits for user to manually submit report
   - Cannot be closed without confirming

## Configuration Notes

When modifying form element mappings in config.yaml:
- Excel column names must exactly match keys in `form_elements` (spaces and special characters matter)
- For diagnosis fields, use pattern: `诊断N_ID`, `诊断N_名称`, `诊断N_编码` (N=1-5)
- Radio button options require value mapping in config (e.g., `男: "0"`, `女: "1"`)
- The `type` field in function_button determines which function config is used

## ChromeDriver

The project expects ChromeDriver at: `chromedriver-win64/chromedriver.exe`
The version must match the installed Chrome browser version.
