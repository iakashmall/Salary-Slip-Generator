import os
import sys
import pandas as pd

from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer
)

from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# =========================================================
# GET FONT PATH FOR EXE
# =========================================================

if getattr(sys, 'frozen', False):

    base_path = sys._MEIPASS

else:

    base_path = os.path.abspath(".")

font_path = os.path.join(
    base_path,
    "DejaVuSans.ttf"
)

# =========================================================
# REGISTER FONT
# =========================================================

pdfmetrics.registerFont(
    TTFont(
        'DejaVu',
        font_path
    )
)

# =========================================================
# COMPANY DETAILS
# =========================================================

COMPANY_NAME = "SHARIKA ENTERPRISES LIMITED"

COMPANY_ADDRESS = (
    "C-504, ATS BOUQUET, SECTOR 132, NOIDA 201305"
)

COMPANY_WEBSITE = (
    "www.sharikaindia.com"
)

# =========================================================
# MAIN FUNCTION
# =========================================================

def generate_salary_slips():

    # =====================================================
    # SELECT EXCEL FILE
    # =====================================================

    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not file_path:
        return

    # =====================================================
    # READ EXCEL FILE
    # =====================================================

    try:

        df = pd.read_excel(file_path)

        # CLEAN COLUMN NAMES

        df.columns = (
            df.columns
            .str.strip()
            .str.replace("\n", " ")
        )

    except Exception as e:

        messagebox.showerror(
            "Excel Error",
            f"Unable to read Excel file:\n\n{e}"
        )

        return

    # =====================================================
    # REQUIRED COLUMNS
    # =====================================================

    required_columns = [

        "Employee Name",
        "Employee ID",
        "UAN",
        "PF No",
        "Designation",
        "ESI No",
        "Department",
        "Bank Name",
        "DOJ",
        "Bank A/C No",

        "Gross Wage",
        "Total Working Days",
        "Paid Days",
        "LOP Days",
        "Leaves Taken",

        "Basic Wage",
        "HRA",
        "Transport Allowances",
        "Medical Allowances",
        "Mobile & Internet",
        "Books & Periodicals",
        "Club",
        "Other Allowance",

        "EPF",
        "ESIC",
        "Professional Tax",
        "Loan Recovery",
        "TDS",

        "Pay Month"
    ]

    # =====================================================
    # CHECK MISSING COLUMNS
    # =====================================================

    missing_columns = [

        col for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:

        messagebox.showerror(
            "Missing Columns",
            "The following columns are missing:\n\n" +
            "\n".join(missing_columns)
        )

        return

    # =====================================================
    # MAIN OUTPUT FOLDER
    # =====================================================

    main_output_folder = "salary_slips"

    if not os.path.exists(main_output_folder):
        os.makedirs(main_output_folder)

    # =====================================================
    # STYLES
    # =====================================================

    styles = getSampleStyleSheet()

    generated_count = 0

    # =====================================================
    # PROCESS EACH EMPLOYEE
    # =====================================================

    for index, row in df.iterrows():

        try:

            # =================================================
            # EMPLOYEE DETAILS
            # =================================================

            employee_name = str(row["Employee Name"])
            employee_id = str(row["Employee ID"])

            uan = str(row["UAN"])
            pf_no = str(row["PF No"])

            designation = str(row["Designation"])
            esi_no = str(row["ESI No"])

            department = str(row["Department"])
            bank_name = str(row["Bank Name"])

            doj = pd.to_datetime(
                row["DOJ"]
            ).strftime("%d-%m-%Y")

            bank_ac = str(row["Bank A/C No"])

            # =================================================
            # PAY MONTH
            # =================================================

            pay_month = pd.to_datetime(
                row["Pay Month"]
            ).strftime("%B_%Y")

            display_month = pay_month.replace(
                "_",
                " "
            )

            # =================================================
            # WORK DETAILS
            # =================================================

            gross_wage = float(row["Gross Wage"])

            total_working_days = row["Total Working Days"]

            paid_days = row["Paid Days"]

            lop_days = row["LOP Days"]

            leaves_taken = row["Leaves Taken"]

            # =================================================
            # EARNINGS
            # =================================================

            basic_wage = float(row["Basic Wage"])

            hra = float(row["HRA"])

            transport = float(
                row["Transport Allowances"]
            )

            medical = float(
                row["Medical Allowances"]
            )

            mobile = float(
                row["Mobile & Internet"]
            )

            books = float(
                row["Books & Periodicals"]
            )

            club = float(row["Club"])

            other_allowance = float(
                row["Other Allowance"]
            )

            # =================================================
            # DEDUCTIONS
            # =================================================

            epf = float(row["EPF"])

            esic = float(row["ESIC"])

            professional_tax = float(
                row["Professional Tax"]
            )

            loan_recovery = float(
                row["Loan Recovery"]
            )

            tds = float(row["TDS"])

            # =================================================
            # CALCULATIONS
            # =================================================

            total_earnings = (

                basic_wage +
                hra +
                transport +
                medical +
                mobile +
                books +
                club +
                other_allowance
            )

            total_deductions = (

                epf +
                esic +
                professional_tax +
                loan_recovery +
                tds
            )

            net_salary = (

                total_earnings -
                total_deductions
            )

            # =================================================
            # GENERATED DATE
            # =================================================

            current_date = datetime.now().strftime(
                "%d-%m-%Y"
            )

            # =================================================
            # SAFE FILE NAME
            # =================================================

            safe_name = employee_name.replace(
                " ",
                "_"
            )

            # =================================================
            # CREATE MONTH FOLDER
            # =================================================

            month_output_folder = os.path.join(
                main_output_folder,
                pay_month
            )

            if not os.path.exists(month_output_folder):
                os.makedirs(month_output_folder)

            # =================================================
            # PDF FILE PATH
            # =================================================

            pdf_file = os.path.join(

                month_output_folder,

                f"{employee_id}_{safe_name}.pdf"
            )

            # =================================================
            # PDF DOCUMENT
            # =================================================

            doc = SimpleDocTemplate(

                pdf_file,

                pagesize=A4,

                rightMargin=20,
                leftMargin=20,

                topMargin=20,
                bottomMargin=20
            )

            elements = []

            # =================================================
            # HEADER
            # =================================================

            company = Paragraph(
                f'''
                <para align=center>
                <font color="red" size="20">
                <b>{COMPANY_NAME}</b>
                </font>
                </para>
                ''',
                styles["Normal"]
            )

            address = Paragraph(
                f'''
                <para align=center>
                <font size="10">
                {COMPANY_ADDRESS}
                <br/>
                <font color="blue">
                {COMPANY_WEBSITE}
                </font>
                </font>
                </para>
                ''',
                styles["Normal"]
            )

            title = Paragraph(
                f'''
                <para align=center>
                <font size="13">
                <b>PAY SLIP FOR {display_month}</b>
                </font>
                </para>
                ''',
                styles["Normal"]
            )

            generated = Paragraph(
                f'''
                <para align=right>
                Generated On : {current_date}
                </para>
                ''',
                styles["Normal"]
            )

            elements.append(company)

            elements.append(Spacer(1, 10))

            elements.append(address)

            elements.append(Spacer(1, 12))

            elements.append(title)

            elements.append(Spacer(1, 10))

            elements.append(generated)

            elements.append(Spacer(1, 12))

            # =================================================
            # EMPLOYEE TABLE
            # =================================================

            employee_data = [

                ["Name of Employee",
                 employee_name,

                 "UAN",
                 uan],

                ["Employee ID",
                 employee_id,

                 "PF No",
                 pf_no],

                ["Designation",
                 designation,

                 "ESI No",
                 esi_no],

                ["Department",
                 department,

                 "Bank Name",
                 bank_name],

                ["DOJ",
                 doj,

                 "Bank A/C No",
                 bank_ac],

                ["Gross Wage",
                 f"₹ {gross_wage:,.2f}",

                 "", ""],

                ["Total Working Days",
                 total_working_days,

                 "Paid Days",
                 paid_days],

                ["LOP Days",
                 lop_days,

                 "Leaves Taken",
                 leaves_taken]
            ]

            emp_table = Table(

                employee_data,

                colWidths=[130, 140, 130, 140]
            )

            emp_table.setStyle(TableStyle([

                ('GRID',
                 (0, 0),
                 (-1, -1),
                 1,
                 colors.black),

                ('BACKGROUND',
                 (0, 0),
                 (0, -1),
                 colors.HexColor("#d9d9d9")),

                ('BACKGROUND',
                 (2, 0),
                 (2, -1),
                 colors.HexColor("#d9d9d9")),

                ('FONTNAME',
                 (0, 0),
                 (-1, -1),
                 'DejaVu'),

                ('BOTTOMPADDING',
                 (0, 0),
                 (-1, -1),
                 6),

                ('TOPPADDING',
                 (0, 0),
                 (-1, -1),
                 6),

                ('FONTSIZE',
                 (0, 0),
                 (-1, -1),
                 9),
            ]))

            elements.append(emp_table)

            elements.append(Spacer(1, 18))

            # =================================================
            # SALARY TABLE
            # =================================================

            salary_data = [

                ["Earnings",
                 "Amount",

                 "Deductions",
                 "Amount"],

                ["Basic Wage",
                 f"₹ {basic_wage:,.2f}",

                 "EPF",
                 f"₹ {epf:,.2f}"],

                ["HRA",
                 f"₹ {hra:,.2f}",

                 "ESIC / Health In.",
                 f"₹ {esic:,.2f}"],

                ["Transport Allowances",
                 f"₹ {transport:,.2f}",

                 "Professional Tax",
                 f"₹ {professional_tax:,.2f}"],

                ["Medical Allowances",
                 f"₹ {medical:,.2f}",

                 "Loan Recovery",
                 f"₹ {loan_recovery:,.2f}"],

                ["Mobile & Internet",
                 f"₹ {mobile:,.2f}",

                 "TDS",
                 f"₹ {tds:,.2f}"],

                ["Books & Periodicals",
                 f"₹ {books:,.2f}",

                 "", ""],

                ["Club",
                 f"₹ {club:,.2f}",

                 "", ""],

                ["Other Allowance",
                 f"₹ {other_allowance:,.2f}",

                 "", ""],

                ["TOTAL EARNINGS",
                 f"₹ {total_earnings:,.2f}",

                 "TOTAL DEDUCTIONS",
                 f"₹ {total_deductions:,.2f}"]
            ]

            salary_table = Table(

                salary_data,

                colWidths=[180, 90, 180, 90]
            )

            salary_table.setStyle(TableStyle([

                ('GRID',
                 (0, 0),
                 (-1, -1),
                 1,
                 colors.black),

                ('BACKGROUND',
                 (0, 0),
                 (-1, 0),
                 colors.HexColor("#d9d9d9")),

                ('FONTNAME',
                 (0, 0),
                 (-1, -1),
                 'DejaVu'),

                ('FONTSIZE',
                 (0, 0),
                 (-1, -1),
                 9),

                ('BOTTOMPADDING',
                 (0, 0),
                 (-1, -1),
                 6),

                ('TOPPADDING',
                 (0, 0),
                 (-1, -1),
                 6),

                ('ALIGN',
                 (1, 1),
                 (-1, -1),
                 'CENTER'),
            ]))

            elements.append(salary_table)

            elements.append(Spacer(1, 18))

            # =================================================
            # NET SALARY BOX
            # =================================================

            net_salary_table = Table(

                [[
                    f"NET SALARY : ₹ {net_salary:,.2f}"
                ]],

                colWidths=[540]
            )

            net_salary_table.setStyle(TableStyle([

                ('GRID',
                 (0, 0),
                 (-1, -1),
                 1,
                 colors.black),

                ('FONTNAME',
                 (0, 0),
                 (-1, -1),
                 'DejaVu'),

                ('FONTSIZE',
                 (0, 0),
                 (-1, -1),
                 12),

                ('ALIGN',
                 (0, 0),
                 (-1, -1),
                 'CENTER'),

                ('BOTTOMPADDING',
                 (0, 0),
                 (-1, -1),
                 8),

                ('TOPPADDING',
                 (0, 0),
                 (-1, -1),
                 8),
            ]))

            elements.append(net_salary_table)

            elements.append(Spacer(1, 20))

            # =================================================
            # NOTE
            # =================================================

            note = Paragraph(
                '''
                <font face="DejaVu">
                <b>Note:</b>
                This is a company-issued system generated salary slip.
                </font>
                ''',
                styles["Normal"]
            )

            elements.append(note)

            elements.append(Spacer(1, 40))

            # =================================================
            # SIGNATURE
            # =================================================

            signature = Paragraph(
                '''
                <para align=right>
                <font face="DejaVu">
                For M/s Sharika Enterprises Limited
                <br/><br/><br/>
                <b>Ms. Sushmita Roy</b>
                <br/>
                Sr. Executive – HR
                </font>
                </para>
                ''',
                styles["Normal"]
            )

            elements.append(signature)

            # =================================================
            # BUILD PDF
            # =================================================

            doc.build(elements)

            generated_count += 1

        except Exception as e:

            print(
                f"Error in row {index + 1}: {e}"
            )

    # =====================================================
    # SUCCESS MESSAGE
    # =====================================================

    messagebox.showinfo(
        "Success",
        f"{generated_count} Salary Slips Generated Successfully!"
    )

# =========================================================
# GUI WINDOW
# =========================================================

root = tk.Tk()

root.title("Salary Slip Generator")

root.geometry("500x300")

root.configure(bg="#f2f2f2")

root.resizable(False, False)

# =========================================================
# TITLE
# =========================================================

heading = tk.Label(

    root,

    text="Salary Slip Generator",

    font=("Arial", 20, "bold"),

    bg="#f2f2f2",

    fg="#1f4e79"
)

heading.pack(pady=30)

# =========================================================
# DESCRIPTION
# =========================================================

info = tk.Label(

    root,

    text="Select Excel File and Generate Salary Slips",

    font=("Arial", 11),

    bg="#f2f2f2"
)

info.pack(pady=10)

# =========================================================
# BUTTON
# =========================================================

generate_button = tk.Button(

    root,

    text="Generate Salary Slips",

    font=("Arial", 12, "bold"),

    bg="#1f77b4",

    fg="white",

    padx=20,

    pady=10,

    command=generate_salary_slips
)

generate_button.pack(pady=30)

# =========================================================
# FOOTER
# =========================================================

footer = tk.Label(

    root,

    text="Sharika Enterprises Limited",

    font=("Arial", 9),

    bg="#f2f2f2",

    fg="gray"
)

footer.pack(side="bottom", pady=10)

# =========================================================
# RUN APPLICATION
# =========================================================

root.mainloop()