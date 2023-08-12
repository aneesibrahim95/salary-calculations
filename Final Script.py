import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Inches
from dependencies import *


dfx = pd.read_excel("biometric.xls", sheet_name=None)
df = pd.concat(dfx.values(), ignore_index=True)

es = pd.read_excel("Employee Structure.xlsx")



# Create a new Word document for writing the salary statement
doc = Document()

# Create an empty DataFrame to store all the detailed work reports
all_work_reports = pd.DataFrame()

for i in range(int(es.loc[0, 'Total number of employees'])):
    namee = es.loc[i, "Name"]
    basic_pay = es.loc[i, "Basic Pay"]
    work_hours = es.loc[i, "Working Hours Per day"]
    weekly_offs = es.loc[i, "Weekly offs per month"]
    # leaves = es.loc[i, "Total leaves Taken"]
    advance = es.loc[i, "Advance Taken"]
    pay_per_day = basic_pay / 30
    pay_per_hour = pay_per_day / work_hours
    number_of_days_in_a_month = es.loc[0,"Number of days in the month"]

    employee_name, total_work_duration, total_working_days, work_punch_detailing = workdur(pd, df, namee)
    leaves = number_of_days_in_a_month-total_working_days

    employee_name = employee_name.split(" : ")
    employee_name = employee_name[1]

    overtime = (total_work_duration - total_working_days * work_hours)
    overtime_pay = overtime * pay_per_hour

    leaves_pending = (weekly_offs - leaves) if leaves <= weekly_offs else 0
    leaves_pending_pay = leaves_pending * pay_per_day

    payless_leaves = leaves - weekly_offs if leaves >= weekly_offs else 0
    payless_leaves_cut = payless_leaves * pay_per_day

    salary = basic_pay + overtime_pay + leaves_pending_pay - payless_leaves_cut - advance

    output_structure = {
        employee_name: ["Parameter", "Pay"],
        "Basic Pay" : [f"Rs {round(basic_pay, 2)}", f"Rs {round(basic_pay, 2)}"],
        "Overtime": [f"{round(overtime, 2)}hrs", f"Rs {round(overtime_pay, 2)}"],
        "Monthly Leave": [leaves if leaves <= weekly_offs else weekly_offs, f"Rs {0}"],
        "Extra Leave": [payless_leaves, f"Rs {round(payless_leaves_cut, 2)}"],
        "Leave Pending": [leaves_pending, f"Rs {round(leaves_pending_pay, 2)}"],
        "Advance": [f"Rs {round(advance, 2)}", f"Rs {round(advance, 2)}"],
        "Total Pay": ["", f"Rs {round(salary, 2)}"]
    }
    print(output_structure)
    # Call the function to add a table for this employee
    add_table_to_doc(doc, employee_name, output_structure)


    # Convert work_punch_detailing list into a DataFrame (for writing the detailed work report)
    # Anywhere in the 2D list, if no data is available, change it to L for leave
    for i in work_punch_detailing:
        for j in range(len(i)):
            if pd.isna(i[j]):
                i[j] = "L"
    work_punch_df = pd.DataFrame(work_punch_detailing, columns=["Date", "In Time", "Out Time", "Duration (hours)"])
    # Add a column for the employee name in the DataFrame
    work_punch_df["Employee Name"] = employee_name

    # Concatenate the work punch DataFrame with all_work_reports (for writing the detailed work report)
    all_work_reports = pd.concat([all_work_reports, work_punch_df], ignore_index=True)



# Save the Word document
doc.save("Salary Summary.docx")


# Create an ExcelWriter to save the reports in the same file (for writing the detailed work report)
output_file = "Detailed_Work_Reports.xlsx"
with pd.ExcelWriter(output_file, engine="xlsxwriter") as excel_writer:
    # Write all the detailed work reports to a new sheet in the Excel file
    sheet_name = "All Detailed Work Reports"
    all_work_reports.to_excel(excel_writer, sheet_name=sheet_name, index=False)

