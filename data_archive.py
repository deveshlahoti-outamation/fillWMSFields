import pandas as pd
import os
import sys

# Get the directory of the .exe file
exe_dir = os.path.dirname(sys.executable)

# Build the path to the excel file
excel_file = os.path.join(exe_dir, 'WMS Information.xlsx')

df = pd.read_excel(excel_file, dtype={"Invoice Number": str})


def check_case_completion(invoice_number):
    print("checking if new case...")
    global df
    invoice_numbers = set(df['Invoice Number'].str.strip())
    if invoice_number in invoice_numbers:
        return False
    else:
        return True


def add_data_to_excel(invoice_number, patient_number, patient_name, state, data):
    if len(data) < 5:
        print(f"Error: expected data to have at least 5 elements, but got {len(data)} elements.")
        return

    new_row = pd.DataFrame({
               "Invoice Number": [invoice_number],
               "Patient Number": [patient_number],
               "Patient Name": [patient_name],
               "State": [state],
               "DOB": [data[0]],
               "Physician Name": [data[1]],
               "Physician Phone": [data[2]],
               "Physician Fax": [data[3]],
               "Request": [data[4]]
    })

    global df
    df = pd.concat([df, new_row], ignore_index=True)
    df["DOB"] = df["DOB"].str.replace("Output: ", "", regex=False)

    df.to_excel(excel_file, index=False)

