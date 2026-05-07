
# CCI Automation V1.00

**Authors**: Hoang Nguyen and Ming Li  
**Current Version**: V1.00  
**Python Version**: 3.11.0

## Project Overview

CCI Automation V1.00 is a tool designed to streamline and automate processes related to CCI reporting. This application simplifies data handling, report generation, and ensures consistent formatting across reports. It’s particularly suited for departments requiring accurate and efficient data processing workflows.

## Key Features

- Automates CCI reporting processes with minimal user interaction.
- Integration with Excel for input/output operations.
- Configured for single-executable deployment for ease of use.

## Installation and Setup

### Requirements

The application requires the following Python packages:

 - matplotlib==3.10.8
 - numpy==2.4.3
 - openpyxl==3.1.5
 - pandas==3.0.1
 - PyQt5==5.15.11
 - PyQt5-Qt5==5.15.2
 - PyQt5-sip==12.18.0
 - PySide6==6.10.2
 - PySide6-Addons==6.10.2
 - PySide6-Essentials==6.10.2
 - pytest==9.0.2
 - shiboken6==6.10.2
 - XlsxWriter==3.2.9

Install the dependencies using the following command:

```bash
pip install -r requirements.txt
```

### Building the Executable

To build the application into a standalone executable, use the following `pyinstaller` command. This will create an optimized `.exe` file named "CCI Automation V1.00" and exclude certain dependencies for a smaller file size.

```bash
python -m pyinstaller --onefile --strip --name "CCI Automation V1.00" --windowed "CCI_Automation.py" --exclude PyQt5 --exclude matplotlib
```

**Note**: Ensure that PyInstaller is installed as listed in the `requirements.txt`.

## Usage

After building the executable, run it by double-clicking or executing in the command line. Follow on-screen instructions to load data and begin the automation process.

## Version History

- **V1.00** - Initial release with foundational CCI reporting functionalities.

## License

This project was developed by Hoang Nguyen and Ming Li for internal use. Redistribution or modification without permission is prohibited.

