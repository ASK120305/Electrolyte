import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from datetime import datetime

def process_file():
    cols_needed = [
        "Case Number", "Customer Name", "Street", "Zip/Postal Code", "Customer Complaint",
        "Product Description", "LineItem Status", "Technician Name", "Created Date"
    ]
    root = tk.Tk()
    root.withdraw()
    print("\nPlease select the CSV file to process (a dialog will appear)...")
    input_path = filedialog.askopenfilename(
        title="Select the CSV file to process",
        filetypes=[("CSV Files", "*.csv")]
    )
    if not input_path:
        print("No file selected. Returning to main menu.")
        return
    try:
        try:
            df = pd.read_csv(input_path, encoding="utf-8")
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(input_path, encoding="latin1")
            except Exception:
                df = pd.read_csv(input_path, encoding="cp1252")
    except Exception as e:
        print(f"Failed to read the CSV file: {e}")
        return
    missing_cols = [col for col in cols_needed if col not in df.columns]
    if missing_cols:
        print(f"Missing columns in input file: {', '.join(missing_cols)}")
        return
    df_filtered = df[df["LineItem Status"] == "New"].copy()
    output_df = df_filtered[cols_needed].copy()
    output_df["remarks"] = ""
    output_df["Created Date"] = pd.to_datetime(output_df["Created Date"], dayfirst=True, errors="coerce")
    output_df["SLA"] = (datetime.today().date() - output_df["Created Date"].dt.date).apply(lambda x: x.days if pd.notnull(x) else None)
    output_df = output_df.sort_values("SLA", ascending=False)
    out_cols_with_sla = [
        "Case Number", "SLA", "Customer Name", "Street", "Zip/Postal Code",
        "Customer Complaint", "Product Description", "LineItem Status", "Technician Name", "remarks"
    ]
    print("Please select where to save the Excel file (a dialog will appear)...")
    output_path = filedialog.asksaveasfilename(
        title="Save the Excel file as",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not output_path:
        print("No output file selected. Returning to main menu.")
        return
    try:
        temp_cols = out_cols_with_sla.copy()
        temp_cols.insert(temp_cols.index("SLA")+1, "Created Date")
        output_df[temp_cols].to_excel(output_path, index=False, sheet_name="Filtered Data")
        # Sheet 2: Pivot Table by Technician Name (rows) x SLA (columns)
        pivot = pd.pivot_table(
            output_df,
            values="Case Number",
            index="Technician Name",
            columns="SLA",
            aggfunc="count",
            fill_value=0,
            margins=True,
            margins_name="Grand Total"
        )
        if "Grand Total" in pivot.columns:
            cols = [c for c in pivot.columns if c != "Grand Total"] + ["Grand Total"]
            pivot = pivot[cols]
        if "Grand Total" in pivot.index:
            pivot_no_total = pivot.drop("Grand Total", axis=0)
            pivot_total = pivot.loc[["Grand Total"]]
            pivot_no_total = pivot_no_total.sort_values("Grand Total", ascending=False)
            pivot = pd.concat([pivot_no_total, pivot_total])
        else:
            pivot = pivot.sort_values("Grand Total", ascending=False)
        with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            pivot.to_excel(writer, sheet_name="Pivot Summary")
        wb = load_workbook(output_path)
        ws1 = wb["Filtered Data"]
        ws2 = wb["Pivot Summary"]
        headers = [cell.value for cell in ws1[1]]
        created_col_idx = headers.index("Created Date") + 1
        ws1.delete_cols(created_col_idx)
        fixed_height = 60
        for ws in [ws1, ws2]:
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            for i in range(1, ws.max_row + 1):
                ws.row_dimensions[i].height = fixed_height
        header_fill = PatternFill(start_color="FFF200", end_color="FFF200", fill_type="solid")
        for cell in ws1[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
        for cell in ws2[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
        for ws in [ws1, ws2]:
            for col in ws.columns:
                max_length = 0
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[col[0].column_letter].width = min(max_length + 2, 40)
        for row in ws2.iter_rows():
            if row[0].value == "Grand Total":
                for cell in row:
                    cell.font = Font(bold=True)
            if ws2.cell(row=1, column=row[0].column).value == "Grand Total":
                for cell in row:
                    cell.font = Font(bold=True)
        wb.save(output_path)
        print(f"Success! Output saved to {output_path}")
    except Exception as e:
        print(f"Failed to save or format the Excel file: {e}")

def main():
    print("Welcome to the CSV to Excel Converter with SLA Calculated in Python, Sorted by SLA, and Pivot Table in Sheet 2.")
    while True:
        proceed = input("\nDo you want to process a file? (y/n): ").strip().lower()
        if proceed == 'y':
            process_file()
        elif proceed == 'n':
            print("Goodbye!")
            break
        else:
            print("Please enter 'y' or 'n'.")

if __name__ == "__main__":
    main()