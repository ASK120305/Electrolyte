# CSV to Excel SLA Converter

This tool processes a CSV file, filters and sorts data by SLA (Service Level Agreement in days), and generates an Excel file with two sheets:

- **Sheet 1: Filtered Data** — All rows with `LineItem Status` as "New", sorted by SLA (highest to lowest), and with the following columns:  
  Case Number, SLA, Customer Name, Street, Zip/Postal Code, Customer Complaint, Product Description, LineItem Status, Technician Name, remarks.

- **Sheet 2: Pivot Summary** — A pivot table with Technician Name as rows, SLA as columns, and a count of cases, including grand totals.

---

## Features

- Reads a CSV with the required columns.
- Filters for rows where `LineItem Status` is `"New"`.
- Calculates `SLA` as the number of days from "Created Date" to today (in Python, not Excel).
- Sorts the output by SLA (highest to lowest).
- Removes the "Created Date" column from the final output.
- Applies formatting for better readability.
- Generates a pivot table by Technician Name and SLA in Sheet 2.

---

## How to Use

1. **Run the Program**  
   You will be prompted to select a CSV file and then choose where to save the resulting Excel file.

2. **CSV Requirements**  
   The CSV must include at least these columns (case-sensitive):  
   - Case Number
   - Customer Name
   - Street
   - Zip/Postal Code
   - Customer Complaint
   - Product Description
   - LineItem Status
   - Technician Name
   - Created Date (format: DD/MM/YYYY)

3. **Output**  
   The Excel file will have:
   - **Sheet 1:** Filtered and sorted data, with SLA as an integer (days).
   - **Sheet 2:** Pivot table summary (Technician Name x SLA).

---

## Troubleshooting

- If you see an error about date conversion, check that your "Created Date" column in the CSV is in `DD/MM/YYYY` format.
- All SLA calculations are performed in Python, so no Excel formulas are used (no #VALUE! errors).
- If you get an error about missing columns, check that your CSV has all required columns listed above.



