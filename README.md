# Salary Slip Generator

A professional desktop application developed using Python to automate salary slip generation from Excel data and export employee-wise PDF payslips with professional formatting.

---

## Features

- Excel-based payroll processing
- Automatic PDF salary slip generation
- Professional payslip formatting
- Month-wise folder organization
- GUI desktop application using Tkinter
- Unicode ₹ symbol support
- Automated salary calculations
- Robust error handling & validation
- EXE application support using PyInstaller

---

## Tech Stack

- Python
- Tkinter
- Pandas
- ReportLab
- PyInstaller

---

## Project Structure

SalaryTool/
│
├── salary_tool.py
├── DejaVuSans.ttf
├── requirements.txt
├── README.md
│
├── dist/
│   └── salary_tool.exe
│
└── salary_slips/
    ├── May_2026/
    └── June_2026/

---

## Installation

### Clone Repository

git clone https://github.com/iakashmall/salary-slip-generator.git

### Move into Project Directory

cd salary-slip-generator

### Install Dependencies

pip install -r requirements.txt

---

## Run Application

python salary_tool.py

---

## Generate EXE File

pyinstaller --onefile --windowed --add-data "DejaVuSans.ttf;." salary_tool.py

---

## Excel Columns Required

Employee Name, Employee ID, UAN, PF No, Designation, ESI No, Department, Bank Name, DOJ, Bank A/C No, Gross Wage, Total Working Days, Paid Days, LOP Days, Leaves Taken, Basic Wage, HRA, Transport Allowances, Medical Allowances, Mobile & Internet, Books & Periodicals, Club, Other Allowance, EPF, ESIC, Professional Tax, Loan Recovery, TDS, Pay Month

---

## Highlights

- Automated HR payroll workflow
- Dynamic PDF generation
- Real-world business use case
- Desktop software packaging
- Clean GUI-based workflow
- Structured error handling

---

## Future Enhancements

- Company logo support
- Email integration
- Password-protected salary slips
- Database integration
- Employee portal

---

## Author

Akash Mall

GitHub: https://github.com/iakashmall
