# Safety Automation V1.01

**Author**: Hoang Nguyen  
**Current Version**: V1.01

## Project Overview

Safety Automation V1.01 is a utility designed for the Safety Department to automate formatting and report generation for Adverse Event (AE) reports. This application processes and organizes data into a standardized format, making AE report management more efficient and reducing manual intervention.

## Key Features

- Automated formatting for AE reports to ensure consistency.
- Supports integration with Excel files for data processing.
- Lightweight, with all necessary functions packaged in a single executable.

## Installation and Setup

### Requirements

This application requires the following Python packages:

- pandas==2.2.1
- numpy==1.25.2
- openpyxl==3.1.2
- PySide6==6.5.2
- PySide6-Addons==6.5.2
- PySide6-Essentials==6.5.2
- ruff==0.4.1
- pyinstaller==6.11.0
- XlsxWriter>=3.1.2

You can install the required packages using the following command:

```bash
pip install -r requirements.txt
```

### Building the Executable

To compile the project into a single executable, use the `pyinstaller` command provided below. This command will generate a single `.exe` file named "Safety Automation V1.01" and exclude unnecessary dependencies to optimize the executable size.

```bash
pyinstaller --onefile --strip --name "Safety Automation V1.01" --windowed "Safety Automation.py" --exclude PyQt5 --exclude requests --exclude matplotlib
```

**Note**: Ensure you have PyInstaller installed, as specified in the `requirements.txt`.

## Files in the Repository

- **Safety Automation.py**: The main script that drives the automation process.
- **util.py**: Contains utility functions to support AE report formatting.
- **AECoreListing.py**: Manages AE core listing configurations.
- **requirements.txt**: Lists all required Python packages for this project.

## Usage

Once the executable is built, run it by double-clicking or executing it in the command line. Follow the on-screen instructions to load AE reports and initiate the formatting process.

## Version History

- **V1.01** - Initial release with basic AE report formatting functionalities.

## License

This project is developed by Hoang Nguyen for internal use in the Safety Department. Redistribution or modification without permission is prohibited.

## Contact

For further inquiries or support, please contact Hoang Nguyen at hoang.nguyen@pennmedicine.upenn.edu
