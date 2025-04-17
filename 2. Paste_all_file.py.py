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

def open_weekly_meeting_files(search_path=r"C:\Weekly Report Bot"):
    """
    1. Searches the given directory for Excel/CSV files whose names contain
       '(Weekly meeting 7 day)'.
    2. Opens them via xlwings and returns the opened workbook objects in a list.
    3. For each opened workbook, adds a =SUM(M2:M{last_row}) formula in Column M
       just below the last used row.
    """
    opened_books = []
    
    for filename in os.listdir(search_path):
        # Only consider Excel and CSV files
        if filename.endswith((".xlsx", ".xlsm", ".xls", ".csv")):
            if "(Weekly meeting 7 day)" in filename:
                full_path = os.path.join(search_path, filename)
                print(f"Opening: {filename}")
                
                # Open the file with xlwings
                wb = xw.Book(full_path)
                opened_books.append(wb)

                # Optional: If CSV, it opens as an Excel workbook with one sheet
                # We'll assume the data is on the first sheet:
                sht = wb.sheets[0]

                # Find the last used row in Column M
                last_row = sht.range("M" + str(sht.cells.last_cell.row)).end("up").row

                # Insert the SUM formula below the last used row
                sum_cell = sht.range(f"M{last_row + 1}")
                sum_cell.formula = f"=SUM(M2:M{last_row})"
                
                # Optionally label the total in Column L
                sht.range(f"L{last_row + 1}").value = "Total"

                # (Optional) save after adding the formula
                wb.save()

    if not opened_books:
        print("No files found containing '(Weekly meeting 7 day)' in their name.")
    else:
        print(f"Successfully opened and updated {len(opened_books)} workbook(s).")
    
    return opened_books

# Example usage:
if __name__ == "__main__":
    # 1) Open the master file
    wb_master = open_master_file()
    if wb_master:
        # You can work with wb_master if needed, e.g.:
        master_sheet = wb_master.sheets[0]
        # ...

    # 2) Open and process all Weekly Meeting Martingale files
    wb_list = open_weekly_meeting_files()
    for wb in wb_list:
        print(f"Processed workbook: {wb.name}")

    # Optionally close everything or leave them open for inspection:
    # for wb in wb_list:
    #     wb.close()
    # if wb_master:
    #     wb_master.close()
    # if wb_list or wb_master:
    #     wb_list[0].app.quit()  # closes the entire Excel app if desired
