# Autogui-Excel
Autogui Excel automtion

# Excel Automation Script

This script automates data entry into an application using Python and various libraries.

## Requirements

- [openpyxl](https://openpyxl.readthedocs.io/en/stable/): Used for working with Excel files.
- [pyautogui](https://pyautogui.readthedocs.io/en/latest/): Used for controlling the mouse and keyboard to automate interactions.
- [keyboard](https://github.com/boppreh/keyboard): Used for sending keyboard input.
- [mouseinfo](https://github.com/boppreh/mouseinfo): Used for obtaining mouse information.

## Usage

1. **Install Dependencies:**
   ```bash
   pip install openpyxl pyautogui keyboard mouseinfo
Run the Script:
  python script.py
Script Execution:

  The script loads an Excel workbook named 'fonte.xlsx'.
  It iterates through each row in the 'Planilha1' sheet.
  For each row, it clicks on specific coordinates on the screen, entering data from the corresponding columns.
Notes
Ensure that the screen coordinates are accurate for your specific application.
Customize the script based on the structure of your Excel file and application.



Feel free to customize the documentation based on additional details, usage instructions, or any specific information related to your project.
