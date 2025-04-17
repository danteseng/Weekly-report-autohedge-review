import os
import xlwings as xw

def open_master_file(search_path=r"C:\Weekly Report Bot"):
    """
    Searches the given directory for an Excel file whose name contains
    both 'Autohedge' and 'master'. If found, opens it as the master file
    and returns the xlwings Book object.
    """
    for filename in os.listdir(search_path):
        # Check if the filename contains both keywords and is an Excel file
        if ("Autohedge" in filename and "master" in filename and
            filename.endswith((".xlsx", ".xlsm", ".xls"))):
            full_path = os.path.join(search_path, filename)
            master_wb = xw.Book(full_path)
            print(f"Opened '{filename}' as the master workbook.")
            return master_wb
    
    print("No Excel file found with both 'Autohedge' and 'master' in its name.")
    return None


def open_weekly_meeting_files_and_copy_sums(search_path=r"C:\Weekly Report Bot"):
    """
    1. Searches 'search_path' for Excel/CSV files whose names contain '(Weekly meeting 7 day)'.
    2. For each matching file:
       - Opens it via xlwings.
       - Inserts/updates a =SUM(M2:M{last_row}) formula below the last used row in Column M.
       - Reads the total sum from that cell.
    3. Collects (filename, total_sum) pairs in a list.
    4. Creates a new workbook, pastes the data into Sheet1, and saves it.
    5. Returns the new summary workbook object.
    """

    sums_data = []  # will hold tuples of (filename, total_sum)

    for filename in os.listdir(search_path):
        # Only consider Excel or CSV files
        if filename.endswith((".xlsx", ".xlsm", ".xls", ".csv")) and "(Weekly meeting 7 day)" in filename:
            full_path = os.path.join(search_path, filename)
            print(f"Opening: {filename}")
            
            # Open the file with xlwings
            wb = xw.Book(full_path)
            sht = wb.sheets[0]  # assume data is on the first sheet

            # Find the last used row in Column M
            last_row = sht.range("M" + str(sht.cells.last_cell.row)).end("up").row

            # Insert the SUM formula in the next row
            sum_cell = sht.range(f"M{last_row + 1}")
            sum_cell.formula = f"=SUM(M2:M{last_row})"

            # Label the total in Column L
            sht.range(f"L{last_row + 1}").value = "Total"

            # Save after adding the formula so it's recalculated
            wb.save()

            # Read the sum value (xlwings should recalc automatically after save)
            total_sum = sum_cell.value

            # Store the (filename, total_sum) in the list
            sums_data.append((filename, total_sum))

            # Close the file (or leave it open if desired)
            wb.close()

    if not sums_data:
        print("No files found containing '(Weekly meeting 7 day)' in their name.")
        return None
    else:
        print(f"Successfully processed {len(sums_data)} workbook(s).")

    # 4) Create a new workbook for the summary
    summary_wb = xw.Book()  # creates a new blank workbook
    summary_sht = summary_wb.sheets[0]
    summary_sht.name = "Summary"

    # Write header
    summary_sht.range("A1").value = "File Name"
    summary_sht.range("B1").value = "Total Sum"

    # Paste the (filename, sum) pairs
    row = 2
    for fname, total in sums_data:
        summary_sht.range(row, 1).value = fname
        summary_sht.range(row, 2).value = total
        row += 1

    # Save the new summary workbook
    summary_file_path = os.path.join(search_path, "Weekly_Meeting_Sums.xlsx")
    summary_wb.save(summary_file_path)
    print(f"Summary workbook created and saved at: {summary_file_path}")

    return summary_wb


# Example usage:
if __name__ == "__main__":
    # 1) Open the master file (if needed)
    wb_master = open_master_file()
    if wb_master:
        # Optionally do something with the master file here
        master_sheet = wb_master.sheets[0]
        # ...

    # 2) Open Weekly Meeting files, sum Column M, and copy sums to a new workbook
    summary_wb = open_weekly_meeting_files_and_copy_sums()
    if summary_wb:
        print("New summary workbook is open. Check 'Weekly_Meeting_Sums.xlsx'.")
    
    # Optional: close the master workbook if you don't need it
    # if wb_master:
    #     wb_master.close()
    #     wb_master.app.quit()
