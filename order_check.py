import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from io import BytesIO
from datetime import datetime
from collections import OrderedDict

CENTER = Alignment(horizontal="center", vertical="center")

def main():
    st.title("Excel Generator")
    st.write("""
             Upload the requriment files to the corresponding box and 
             provide the information to generate the summary excel sheet.

             Please ensure all the uploaded file is in excel format.
             
             Try download the excel file again if first download is fail. 
             Do not need to refresh the page.
            """)

    excel_upload_section()
    express_file = st.session_state.get("excel_file_1")
    stock_file = st.session_state.get("excel_file_2")

    if express_file is not None:
        express_data, bill_numbers, total = get_express_data(express_file)
        express_data = summarize_by_barcode_and_code(express_data)
        st.write(a)

    if stock_file is not None:
        stock_data = get_stock_data(stock_file)
        st.write(stock_data)

    excel_file = generate_excel(express_data)
    excel_file = update_user_input_title(excel_file)
    excel_file = get_datetime(excel_file)
    excel_file = get_branch_number_and_version(excel_file)
    excel_file = update_bill_numbers_and_total_profit(excel_file, bill_numbers, total)

    st.download_button(
        label="⬇️ Download Excel File",
        data=excel_file,
        file_name="example_layout.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# --- Excel generation helper functions ---

def generate_excel(data_rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Row 1
    ws.merge_cells("A1:G1")
    ws["A1"] = "My Big Title"
    ws["A1"].alignment = CENTER

    # Row 2
    ws["A2"] = "บิล:"
    ws.merge_cells("B2:D2")
    ws["B2"] = "Row 2: A-D merged"

    ws["E2"] = "Row 2: E"
    ws["E2"].alignment = CENTER

    ws.merge_cells("F2:G2")
    ws["F2"] = "Row 2: F-G merged"

    # Row 3
    ws.merge_cells("A3:E3")
    ws["A3"] = "Row 3: A-E merged"

    ws.merge_cells("F3:G3")
    ws["F3"] = "Row 3: F-G merged"
    ws["F3"].alignment = CENTER

    # Row 4
    ws.merge_cells("A4:D4")
    ws["A4"] = "Row 4: A-D merged"

    ws.merge_cells("E4:G4")
    ws["E4"] = "Row 4: E-G merged"\

    ws["A5"] = "NO."
    ws["B5"] = "บาร์โค้ด"
    ws["C5"] = "ชื่อสินค้า"
    ws["D5"] = "จำนวน"
    ws["E5"] = "STOCK"
    ws["F5"] = "แพ็ค"
    ws["G5"] = "จัดสินค้า"
    for row in ws["A5:F5"]:
        for cell in row:
            cell.alignment = CENTER

    # Main body (row 5+)
    # start_row = 6
    # for i, row_values in enumerate(data_rows):
    #     excel_row = start_row + i
    #     for col_index in range(7):
    #         col_letter = get_column_letter(col_index + 1)
    #         cell = ws[f"{col_letter}{excel_row}"]
    #         cell.value = row_values[col_index] if col_index < len(row_values) else ""
    #         cell.alignment = CENTER

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def update_user_input_title(excel_file):
    """
    Shows a text box. Whatever the user types is stored automatically
    (no save button) and returned exactly as entered.
    """

    st.subheader("Title")

    excel_file.seek(0)
    wb = load_workbook(excel_file)
    ws = wb.active

    st.text_input(
        "Enter title:",
        key = "user_title",
        placeholder="Enter the title here",
    )

    # Always store the latest raw text
    ws["A1"] = st.session_state.get("user_title", "")


    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def get_datetime(excel_file):
    """
    Show date & time inputs for the Excel file.
    - Defaults to current date & time on first run
    - User can edit any part they want
    - Values are stored live in st.session_state["date"] and ["time"]
    - No need to return anything; you can read from session_state later.
    """
    st.subheader("Date & Time")

    excel_file.seek(0)
    wb = load_workbook(excel_file)
    ws = wb.active

    now = datetime.now()

    # Set defaults only once (so user edits are not overwritten on rerun)
    if "date" not in st.session_state:
        st.session_state["date"] = now.strftime("%Y-%m-%d")   # e.g. 2026-01-07
    if "time" not in st.session_state:
        st.session_state["time"] = now.strftime("%H:%M:%S")   # e.g. 14:32:05

    # These inputs are live and bound to session_state
    st.text_input(
        "Date (YYYY-MM-DD)",
        key="date",
    )

    st.text_input(
        "Time (HH:MM:SS)",
        key="time",
    )

    ws["F2"] = "วันที่  " + st.session_state["date"]
    ws["E2"] = st.session_state["time"]

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def get_branch_number_and_version(excel_file):
    """
    Get the branch number and the version of the file 
    from user input and put in the excel
    """
    st.subheader("Branch Number & Version")

    excel_file.seek(0)
    wb = load_workbook(excel_file)
    ws = wb.active

    st.text_input(
        "Enter the branch number:",
        key = "branch_number",
        placeholder="Enter the branch number here",
    )

    st.text_input(
        "Enter the version:",
        key = "version",
        placeholder="Enter the version number here",
    )

    ws["A3"] = "เขต  " + st.session_state["branch_number"]
    ws["F3"] = st.session_state["version"]

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def update_bill_numbers_and_total_profit(excel_file, bill_numbers, total):
    excel_file.seek(0)
    wb = load_workbook(excel_file)
    ws = wb.active

    ws["B2"] = bill_numbers[0] + " – " + bill_numbers[-1]
    ws["E4"] = "จำนวนบิล          " + str(len(bill_numbers)) + "      บิล"
    ws["A4"] = "รวม                         " + total + "  บาท"

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# --- Data obtain & analysis helper functions

def excel_upload_section():
    """
    Show an interface that lets the user upload two Excel files.
    The uploaded files are stored in:
        st.session_state["excel_file_1"]
        st.session_state["excel_file_2"]
    so you can use them later in your code.
    """

    st.subheader("Upload Excel Files")

    file1 = st.file_uploader(
        "Upload the Excel file from Express Accounting.",
        type=["xlsx", "xlsm", "xls"],
        key="excel_upload_1",
    )

    file2 = st.file_uploader(
        "Upload the product stock file",
        type=["xlsx", "xlsm", "xls"],
        key="excel_upload_2",
    )

    # Store them in session_state so other blocks can use them
    if file1 is not None:
        st.session_state["excel_file_1"] = file1

    if file2 is not None:
        st.session_state["excel_file_2"] = file2

    # # (Optional) small status display
    # if "excel_file_1" in st.session_state:
    #     st.write("✅ First file uploaded:", st.session_state["excel_file_1"].name)

    # if "excel_file_2" in st.session_state:
    #     st.write("✅ Second file uploaded:", st.session_state["excel_file_2"].name)

def get_express_data(uploaded_file):
    """
    Given an uploaded Excel file, find the SECOND 'horizontal line' row
    (a row where any cell contains only '-' characters like '--------')
    and return all rows of data below that line.

    Returns:
        list[list]: list of rows; each row is a list of cell values.
    """
    # Make sure we're at the start of the file
    try:
        uploaded_file.seek(0)
    except Exception:
        # Some file-like objects may not have seek; ignore if so
        pass

    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb.active  # or wb["SheetName"] if you want a specific sheet

    max_row = ws.max_row
    max_col = ws.max_column

    separator_rows = []

    for r in range(1, max_row + 1):
        row_has_dash_line = False

        for c in range(1, max_col + 1):
            cell_value = ws.cell(row=r, column=c).value

            if isinstance(cell_value, str):
                stripped = cell_value.strip()
                # check "purely horizontal line", e.g. "-----" or " -------- "
                if stripped and set(stripped) == {"-"}:
                    row_has_dash_line = True
                    break  # we already know this row is a separator

        if row_has_dash_line:
            separator_rows.append(r)

    # Need at least 2 separator rows
    if len(separator_rows) < 2:
        return []  # nothing to return, can't find second horizontal line

    second_sep_row = separator_rows[1]
    start_row = second_sep_row + 1

    data = []
    for r in range(start_row, max_row + 1):
        row_values = [
            ws.cell(row=r, column=c).value
            for c in range(1, max_col + 1)
        ]

        # Optionally skip completely empty rows
        if all(v in (None, "") for v in row_values):
            continue

        split_row = row_values[0].split()
        data.append(split_row)

    return treat_express_data(data)

def treat_express_data(data):
    bill_number = ""
    bill_number_collection = []
    index = 0

    while (index < len(data) and data[index][0] != "รวมทั้งสิ้น"):
        if (data[index][0] != bill_number):
            bill_number = data[index][0]
            bill_number_collection.append(bill_number)
            del data[index]
            index -= 1
    
        index += 1

    if (data[index][0] == "รวมทั้งสิ้น"):
            total = (data[index][-1])
            data = data[0 : index]
            del bill_number_collection[-1]
            return data, bill_number_collection, total  

    print(bill_number_collection)
    raise ValueError("Cannot find รวมทั้งสิ้น, check the input file.")

def extract_pack_qty_from_row(row):
    """
    Find all cells in the row that contain '.แพ็ค' and
    sum the numbers before '.แพ็ค'.

    E.g. '55.แพ็ค' -> 55, '8.แพ็ค' -> 8.
    If nothing found or parse fails, returns 0.
    """
    total = 0.0
    for cell in row:
        if isinstance(cell, str) and ".แพ็ค" in cell:
            before = cell.split(".แพ็ค")[0].strip()
            before = before.replace(",", "")
            if not before:
                continue
            try:
                total += float(before)
            except ValueError:
                continue
    return total

def summarize_by_barcode_and_code(data_rows):
    """
    Group rows by (barcode, item_code) = (row[2], row[3])
    and sum all 'X.แพ็ค' quantities for each group.

    Returns:
        list of dicts:
          {
            'barcode': ...,
            'item_code': ...,
            'sum_qty': ...,
          }
    """
    summaries = OrderedDict()  # to keep order of first appearance

    for row in data_rows:
        if len(row) < 4:
            continue  # need at least [bill, line, barcode, item_code]

        barcode = row[2]
        item_code = row[3]

        if barcode is None or item_code is None:
            continue

        key = (barcode, item_code)

        if key not in summaries:
            summaries[key] = {
                "barcode": barcode,
                "item_code": item_code,
                "sum_qty": 0.0,
            }

        qty = extract_pack_qty_from_row(row)
        summaries[key]["sum_qty"] += qty

    return list(summaries.values())

def get_stock_data(uploaded_file):
    """
    Given the uploaded stock Excel file, search the barcode that listed before through the file.
    (which is appear at the second column in the stock file)
    and return all rows of barcode with the stock number (stored in the sixth column)

    Returns:
        List: a list of the stock in the order of barcode searching order.
    """
    # Make sure we're at the start of the file
    try:
        uploaded_file.seek(0)
    except Exception:
        # Some file-like objects may not have seek; ignore if so
        pass

    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb.active  # or wb["SheetName"] if you want a specific sheet
    max_row = ws.max_row
    print(max_row)
    max_col = ws.max_column

    data = []
    for r in range (1, max_row+1):
        row_value = [
            ws.cell(row=r, column=c).value
            for c in [2, 3, 6]
        ]

        # Optionally skip completely empty rows
        if all(v in (None, "") for v in row_value):
            continue

        if not (
        isinstance(row_value[0], (int, float)) or
        (isinstance(row_value[0], str) and row_value[0].strip().isdigit())
        ):
            continue

        data.append(row_value)

    return data

main()