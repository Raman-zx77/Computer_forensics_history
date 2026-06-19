# Computer_forensics_history_report
# Computer Forensics Report Generator

A Python-based Windows forensic analysis tool that collects system artifacts and generates a PDF report for investigation and educational purposes.

## Features

- Collects saved Wi-Fi profiles
- Extracts Wi-Fi passwords (if accessible)
- Detects connected and previously connected USB devices
- Lists system user accounts
- Retrieves recently used applications
- Shows installed software
- Shows potentially uninstalled software entries
- Collects recently opened documents
- Generates a professional PDF forensic report

## Technologies Used

- Python 3.x
- WMI
- ReportLab
- Windows Registry (winreg)
- PowerShell
- Windows COM Interfaces

## Requirements

- Windows 10/11
- Python 3.8+
- Administrator privileges (recommended)

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/computer-forensics-report-generator.git
cd computer-forensics-report-generator
```

### 2. Install Dependencies

```bash
pip install wmi
pip install reportlab
pip install pywin32
```

Or

```bash
pip install -r requirements.txt
```

## Usage

Run the script:

```bash
python Computer-forensic-2.py
```

The program will:

1. Collect forensic artifacts.
2. Analyze system information.
3. Generate a PDF report.
4. Save the report inside:

```text
report/forensic_report.pdf
```

## Output

The generated PDF contains:

- Wi-Fi Information
- USB Device History
- System Usernames
- Recently Used Applications
- Recently Installed Software
- Recently Uninstalled Software
- Recently Opened Documents

## Project Structure

```text
.
├── Computer-forensic-2.py
├── report/
│   └── forensic_report.pdf
├── requirements.txt
└── README.md
```

## Ethical Use Notice

This project is intended for:

- Digital Forensics Education
- Cybersecurity Learning
- Incident Response Practice
- Authorized System Analysis

Do not use this tool on systems without proper authorization.

## Future Improvements

- Browser history analysis
- Event log collection
- Timeline generation
- Hash calculation for files
- Export to CSV/JSON
- GUI version

## Author

Raman Sharma
Cybersecurity Student
