import re
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles import Border, Side
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo
from collections import OrderedDict

CENTER = Alignment(horizontal="center", vertical="center")

BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

def main():
    st.title("Excel Generator")
    st.write("""
             Upload the requriment files to the corresponding box and 
             provide the information to generate the summary excel sheet.

             Please ensure all the uploaded file is in excel format.
             
             Try download the excel file again if first download is fail. 
             Do not need to refresh the page.
            """)

    def banana():
        st.success("🍌 Banana function called")
        st.write("This is banana logic")

    def peach():
        st.success("🍑 Peach function called")
        st.write("This is peach logic")

    st.subheader("กรุณาเลือกลูกค้า")
    choice = st.radio(
        "กรุณาเลือกลูกค้า",
        ["ร้านย่อย", "GBH", "DH", "HP"],
        index=None,  # nothing selected initially
        label_visibility="collapsed"
    )

    if choice == "ร้านย่อย":
        ThaiName()
    elif choice == "GBH":
        banana()
    elif choice == "DH":
        peach()
    elif choice == "HP":
        banana()

# --- Entrance function for different companies with specific programme logic ---

def ThaiName():
    ExcelUploadSection()
    express_file = st.session_state.get("excel_file_1")
    stock_file = st.session_state.get("excel_file_2")

    if express_file is not None and stock_file is not None:
        express_data, bill_numbers, total = GetExpressData(express_file)
        express_data = SummariseByBarcode(express_data)
        stock_data = GetStockData(stock_file)

        excel_file = GenerateExcel()
        excel_file = UpdateUserInputTitle(excel_file)
        excel_file = GetDateTime(excel_file)
        excel_file = GetBranchNumberAndVersion(excel_file)
        excel_file = UpdateBillNumberAndTotalProfit(excel_file, bill_numbers, total)
        excel_file = WriteMainData(excel_file, express_data, stock_data)
        excel_file = AdjustExcelColWidthAndAddBorder(excel_file)
        
        st.subheader("Download the Excel file")
        agree = st.toggle("I confirm I am not a robot")
        st.download_button(
            label="⬇️ Download Excel File",
            data=excel_file,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled = not agree
        )

def GBH():
    st.write("GBH")

def DH():
    st.write("DH")

def HP():
    st.write("HP") 

# --- Excel generation helper functions ---

def GenerateExcel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Row 1
    ws.merge_cells("A1:G1")
    ws["A1"] = "My Big Title"
    ws["A1"].alignment = CENTER

    # Row 2
    ws["A2"] = "บิล:"
    ws["A2"].font = Font(size=13, bold=True, color="9933FF")
    ws["A2"].fill = PatternFill(
        fill_type="solid",
        start_color="E2EFDA",
        end_color="E2EFDA",
    )

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
    for row in ws["A5:G5"]:
        for cell in row:
            cell.alignment = CENTER
            cell.font = Font(size=12, bold=True)

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

def UpdateUserInputTitle(excel_file):
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
    ws["A1"].font = Font(size=32, color="6600CC")
    ws["A1"].fill = PatternFill(
        fill_type="solid",
        start_color="FFAAFF",
        end_color="FFAAFF",
    )

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def GetDateTime(excel_file):
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

    now = datetime.now(ZoneInfo("Asia/Bangkok"))

    # Set defaults only once (so user edits are not overwritten on rerun)
    if "date" not in st.session_state:
        st.session_state["date"] = now.strftime("%Y-%m-%d")   # e.g. 2026-01-07
    if "time" not in st.session_state:
        st.session_state["time"] = now.strftime("%H:%M")   # e.g. 14:32:05

    date_val = st.date_input("Date", value=now.date(), key="date_input")
    time_val = st.time_input("Time", value=now.time(), key="time_input", step=300)

    st.session_state["date"] = date_val.strftime("%d/%m/") + str(date_val.year + 543)
    st.session_state["time"] = time_val.strftime("%H:%M")

    ws["E2"] = st.session_state["time"]
    ws["E2"].font = Font(size=14, color="000000")
    ws["E2"].fill = PatternFill(
        fill_type="solid",
        start_color="FFC000",
        end_color="FFC000",
    )

    ws["F2"] = "วันที่   " + st.session_state["date"]
    ws["F2"].font = Font(size=16, color="FF0000")
    ws["F2"].fill = PatternFill(
        fill_type="solid",
        start_color="FCE4D6",
        end_color="FCE4D6",
    )

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def GetBranchNumberAndVersion(excel_file):
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

    ws["A3"] = "เขต:  " + st.session_state["branch_number"]
    ws["A3"].font = Font(size=18, color="CC00FF")
    ws["A3"].fill = PatternFill(
        fill_type="solid",
        start_color="FFCCFF",
        end_color="FFCCFF",
    )

    ws["F3"] = st.session_state["version"]
    ws["F3"].font = Font(size=21, bold=True, color="0000FF")
    ws["F3"].fill = PatternFill(
        fill_type="solid",
        start_color="97DCFF",
        end_color="97DCFF",
    )
    #97DCFF

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def UpdateBillNumberAndTotalProfit(excel_file, bill_numbers, total):
    excel_file.seek(0)
    wb = load_workbook(excel_file)
    ws = wb.active

    ws["B2"] = bill_numbers[0] + " – " + bill_numbers[-1]
    ws["B2"].font = Font(size=21, color="0000FF")
    ws["B2"].fill = PatternFill(
        fill_type="solid",
        start_color="E2EFDA",
        end_color="E2EFDA",
    )

    ws["A4"] = "รวม                                         " + total + "   บาท"
    ws["A4"].font = Font(size=16, color="0066FF")
    ws["A4"].fill = PatternFill(
        fill_type="solid",
        start_color="CCCCFF",
        end_color="CCCCFF",
    )

    ws["E4"] = "จำนวนบิล         " + str(len(bill_numbers)) + "    บิล"
    ws["E4"].font = Font(size=14, color="FF0066")
    ws["E4"].fill = PatternFill(
        fill_type="solid",
        start_color="E2EFDA",
        end_color="E2EFDA",
    )

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def WriteMainData(excel_file, express_data, stock_data):
    excel_file.seek(0)
    wb = load_workbook(excel_file)
    ws = wb.active

    number = 0
    for row, item in enumerate(express_data):
        end = len(stock_data)
        excel_row = row + 6
        number += 1

        cell = ws[f"A{excel_row}"]
        cell.value = number
        cell.alignment = CENTER
        cell.font = Font(size=12)

        ws[f"D{excel_row}"].value = item["sum_qty"]
        ws[f"D{excel_row}"].alignment = CENTER
        ws[f"D{excel_row}"].font = Font(size=12)

        if "_" in item["barcode"]:
            before, _, after = item["barcode"].partition("_")
            ws[f"B{excel_row}"].value = before
            ws[f"B{excel_row}"].alignment = CENTER
            ws[f"B{excel_row}"].font = Font(size=12)
            ws[f"B{excel_row}"].fill = PatternFill(
                fill_type="solid",
                start_color="FF4A0B",
                end_color="FF4A0B",
            )

            ws[f"C{excel_row}"] = after
            ws[f"C{excel_row}"].alignment = CENTER
            ws[f"C{excel_row}"].font = Font(size=12)
            continue

        ws[f"B{excel_row}"].value = item["barcode"]
        ws[f"B{excel_row}"].alignment = CENTER
        ws[f"B{excel_row}"].font = Font(size=12)

        found = False
        for search in range (0, end):
            barcode_bill = SafeInt(item.get("barcode"))
            barcode_stock = SafeInt(stock_data[search][0])

            if barcode_bill is not None and barcode_stock is not None:
                if (barcode_bill == barcode_stock):
                    ws[f"C{excel_row}"].value = stock_data[search][1]
                    ws[f"C{excel_row}"].alignment = CENTER
                    ws[f"C{excel_row}"].font = Font(size=12)
                    
                    ws[f"E{excel_row}"].value = stock_data[search][2]
                    ws[f"E{excel_row}"].alignment = CENTER
                    ws[f"E{excel_row}"].font = Font(size=12)
                    del stock_data[search]
                    found = True
                    break

        if not found:
            ws[f"C{excel_row}"].value = "Cannot find the barcode.\nUpdate the main sheet of the stock file."
            ws[f"C{excel_row}"].alignment = CENTER
            ws[f"C{excel_row}"].font = Font(size=12)
            ws[f"C{excel_row}"].fill = PatternFill(
                fill_type="solid",
                start_color="FF4A0B",
                end_color="FF4A0B",
            )

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def AdjustExcelColWidthAndAddBorder(excel_file):
    excel_file.seek(0)
    wb = load_workbook(excel_file)
    ws = wb.active

    # Choose the row and column want to autosize
    min_col = 1
    max_col = ws.max_column
    min_row = 5
    max_row = ws.max_row

    for col in range(min_col, max_col + 1):
        col_letter = get_column_letter(col)
        max_len = 0

        for row in range(min_row, max_row + 1):
            value = ws.cell(row=row, column=col).value
            ws.cell(row=row, column=col).border = BORDER
            if value is None:
                continue

            # Convert to string for measuring
            text = str(value)

            # Optional: ignore very long multi-line values
            text = text.split("\n")[0]

            if len(text) > max_len:
                max_len = len(text)

        # Add padding, and cap width so it doesn't get crazy wide
        padding = 6 if col != 1 else 3
        ws.column_dimensions[col_letter].width = min(max_len + padding, 60)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# --- Data obtain & analysis helper functions

def ExcelUploadSection():
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

def GetExpressData(uploaded_file):
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

    return TreatExpressData(data)

def TreatExpressData(data):
    bill_number = ""
    bill_number_collection = []
    index = 0

    for row in data:
        if len(row) < 3:
            continue

        third = str(row[2]).strip()

        # already good
        if third.isdigit():
            continue

        # try to repair "1234.SOMETHING" => ["1234", "SOMETHING"]
        if "." in third:
            left, _, right = third.partition(".")

            # only split if left is digits AND right looks like a real next-field (not just decimals)
            if left.isdigit() and right and not right.isdigit():
                row[2] = left
                row.insert(3, right)  # shift the rest to the right
                continue

        if not third.isdigit():
            row[2] = "0000000000000_"+ third
            row.insert(3, third)

    while (index < len(data) and data[index][0] != "รวมทั้งสิ้น"):
        if len(data[index]) < 5:
            del data[index]
            continue
        
        checking_bill = data[index][0]
        checking_bill = re.sub(r'[^A-Za-z0-9]', '', checking_bill)

        if (checking_bill != bill_number):
            bill_number = checking_bill
            bill_number_collection.append(bill_number)
            del data[index]
            index -= 1
    
        index += 1

    if (data[index][0] == "รวมทั้งสิ้น"):
            total = (data[index][-1])
            data = data[0 : index]
            return data, bill_number_collection, total  

    raise ValueError("Cannot find รวมทั้งสิ้น, check the input file.")

def ExtractPackQtyFromRow(row):
    """
    Find all cells in the row that contain '.แพ็ค' and
    sum the numbers before '.แพ็ค'.

    E.g. '55.แพ็ค' -> 55, '8.แพ็ค' -> 8.
    If nothing found or parse fails, returns 0.
    """
    qty = 0.0
    for cell in row:
        if isinstance(cell, str) and ".แพ็ค" in cell:
            before = cell.split(".แพ็ค")[0].strip()
            before = before.replace(",", "")
            if not before:
                continue
            try:
                qty += float(before)
            except ValueError:
                continue
    return qty

def SummariseByBarcode(data_rows):
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
        if barcode is None:
            continue

        if barcode not in summaries:
            summaries[barcode] = {
                "barcode": barcode,
                "sum_qty": 0.0,
            }

        qty = ExtractPackQtyFromRow(row)
        summaries[barcode]["sum_qty"] += qty

    return list(summaries.values())

def GetStockData(uploaded_file):
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
    ws = wb.worksheets[0]  # or wb["SheetName"] if you want a specific sheet
    max_row = ws.max_row

    data = []
    for r in range (2, max_row+1):
        row_value = [
            ws.cell(row=r, column=c).value
            for c in [2, 3, 6]
        ]

        # Optionally skip completely empty rows
        if all(v in (None, "") for v in row_value):
            continue
        
        data.append(row_value)

    return data

# --- General helper function ---

def SafeInt(x):
    try:
        return int(x)
    except (ValueError, TypeError):
        return None

main()