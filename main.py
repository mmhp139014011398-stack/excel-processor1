from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from openpyxl import load_workbook
from copy import copy
from pathlib import Path
import json
import shutil
import uuid


app = FastAPI(title="Excel Template Processor")


BASE_DIR = Path(__file__).resolve().parent

# نام فایل Template ثابت
TEMPLATE_FILE = BASE_DIR / "الماس تجارت - Copy.xlsx"

# پوشه موقت
TEMP_DIR = BASE_DIR / "temp"
TEMP_DIR.mkdir(exist_ok=True)


# ---------------------------------------------------------
# Helper functions
# ---------------------------------------------------------

def copy_cell_style(source, target):
    """
    Copy Excel cell formatting without changing the value.
    """
    if source.has_style:
        target._style = copy(source._style)

    target.number_format = source.number_format
    target.alignment = copy(source.alignment)
    target.protection = copy(source.protection)


def copy_row(ws, source_row, target_row):
    """
    Copy an entire row's formatting and dimensions.
    """
    for col in range(1, ws.max_column + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)

        copy_cell_style(source, target)

    if source_row in ws.row_dimensions:
        ws.row_dimensions[target_row].height = (
            ws.row_dimensions[source_row].height
        )


def safe_value(value):
    """
    Convert None to empty string.
    """
    if value is None:
        return ""

    return value


# ---------------------------------------------------------
# Excel generation
# ---------------------------------------------------------

def generate_excel(data, output_file):
    """
    Create a new Excel file from the fixed Template.
    """

    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError(
            f"Template file not found: {TEMPLATE_FILE}"
        )

    # Open template
    workbook = load_workbook(
        TEMPLATE_FILE,
        data_only=False
    )

    # -----------------------------------------------------
    # Template sheet
    # -----------------------------------------------------

    SHEET_NAME = "خام (2)"

    if SHEET_NAME not in workbook.sheetnames:
        raise ValueError(
            f"Sheet '{SHEET_NAME}' does not exist. "
            f"Available sheets: {workbook.sheetnames}"
        )

    ws = workbook[SHEET_NAME]

    # -----------------------------------------------------
    # Read structured data
    # -----------------------------------------------------

    seller = data.get("seller", {})
    buyer = data.get("buyer", {})
    invoice = data.get("invoice", {})
    items = data.get("items", [])
    totals = data.get("totals", {})
    notes = data.get("notes", "")

    # -----------------------------------------------------
    # Invoice information
    # -----------------------------------------------------

    ws["E1"] = safe_value(
        invoice.get("title", "پیش فاکتور")
    )

    ws["AB1"] = safe_value(
        invoice.get("number")
    )

    ws["AB3"] = safe_value(
        invoice.get("date")
    )

    # -----------------------------------------------------
    # Seller
    # -----------------------------------------------------

    ws["A5"] = safe_value(
        seller.get("name")
    )

    ws["E5"] = safe_value(
        seller.get("economic_id")
    )

    ws["R5"] = safe_value(
        seller.get("national_id")
    )

    ws["A7"] = safe_value(
        seller.get("address")
    )

    ws["R7"] = safe_value(
        seller.get("registration_id")
    )

    ws["A8"] = safe_value(
        seller.get("phone")
    )

    ws["R8"] = safe_value(
        seller.get("postal_code")
    )

    # -----------------------------------------------------
    # Buyer
    # -----------------------------------------------------

    ws["A10"] = safe_value(
        buyer.get("name")
    )

    ws["E10"] = safe_value(
        buyer.get("economic_id")
    )

    ws["R10"] = safe_value(
        buyer.get("national_id")
    )

    ws["A12"] = safe_value(
        buyer.get("address")
    )

    ws["R12"] = safe_value(
        buyer.get("registration_id")
    )

    ws["A13"] = safe_value(
        buyer.get("phone")
    )

    ws["R13"] = safe_value(
        buyer.get("postal_code")
    )

    # -----------------------------------------------------
    # Items
    # -----------------------------------------------------

    FIRST_ITEM_ROW = 16

    # We use row 16 as the formatting source.
    STYLE_SOURCE_ROW = 16

    # Maximum number of existing template rows to clear.
    MAX_TEMPLATE_ROWS = 100

    # Clear existing item values while preserving formatting/formulas
    for row in range(
        FIRST_ITEM_ROW,
        FIRST_ITEM_ROW + MAX_TEMPLATE_ROWS
    ):

        for col in range(
            1,
            ws.max_column + 1
        ):

            cell = ws.cell(row, col)

            # Don't destroy formulas already present
            if not (
                isinstance(cell.value, str)
                and cell.value.startswith("=")
            ):
                cell.value = None

    # -----------------------------------------------------
    # Write items
    # -----------------------------------------------------

    for index, item in enumerate(items):

        row = FIRST_ITEM_ROW + index

        # If row does not already exist,
        # copy formatting from template row 16.
        if row != STYLE_SOURCE_ROW:
            copy_row(
                ws,
                STYLE_SOURCE_ROW,
                row
            )

        # Row number
        ws.cell(row, 1).value = index + 1

        # Description
        ws.cell(row, 2).value = safe_value(
            item.get("description")
        )

        # Quantity
        ws.cell(row, 5).value = safe_value(
            item.get("quantity")
        )

        # Unit
        ws.cell(row, 7).value = safe_value(
            item.get("unit")
        )

        # Unit price
        ws.cell(row, 8).value = safe_value(
            item.get("unit_price")
        )

        # Total amount
        #
        # Important:
        # We calculate this with Excel,
        # not with the AI.
        ws.cell(row, 10).value = (
            f"=H{row}*E{row}"
        )

        # Discount
        ws.cell(row, 14).value = safe_value(
            item.get("discount")
        )

        # Tax
        ws.cell(row, 18).value = safe_value(
            item.get("tax")
        )

        # Final amount
        ws.cell(row, 24).value = safe_value(
            item.get("final_total")
        )

    # -----------------------------------------------------
    # Totals
    # -----------------------------------------------------

    if len(items) > 0:

        first_row = FIRST_ITEM_ROW

        last_row = (
            FIRST_ITEM_ROW
            + len(items)
            - 1
        )

        total_row = last_row + 1

        # Subtotal
        ws.cell(total_row, 10).value = (
            f"=SUM(J{first_row}:J{last_row})"
        )

    # -----------------------------------------------------
    # Additional totals
    # -----------------------------------------------------

    if totals.get("subtotal") not in (
        None,
        ""
    ):
        ws["J22"] = totals.get("subtotal")

    if totals.get("tax") not in (
        None,
        ""
    ):
        ws["R22"] = totals.get("tax")

    if totals.get("grand_total") not in (
        None,
        ""
    ):
        ws["X22"] = totals.get("grand_total")

    # -----------------------------------------------------
    # Notes
    # -----------------------------------------------------

    if notes:
        ws["A19"] = notes

    # -----------------------------------------------------
    # Excel recalculation
    # -----------------------------------------------------

    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
    except Exception:
        pass

    # -----------------------------------------------------
    # Save
    # -----------------------------------------------------

    workbook.save(output_file)


# ---------------------------------------------------------
# API
# ---------------------------------------------------------

@app.get("/")
def root():
    return {
        "status": "ok",
        "service": "Excel Template Processor"
    }


@app.post("/generate")
async def generate(
    file: UploadFile = File(...),
    data: str = Form(...)
):

    # -----------------------------------------------------
    # Validate template
    # -----------------------------------------------------

    if not TEMPLATE_FILE.exists():
        raise HTTPException(
            status_code=500,
            detail="Template Excel file is missing."
        )

    # -----------------------------------------------------
    # Parse JSON
    # -----------------------------------------------------

    try:
        structured_data = json.loads(data)

    except json.JSONDecodeError as e:

        raise HTTPException(
            status_code=400,
            detail=f"Invalid JSON data: {str(e)}"
        )

    # -----------------------------------------------------
    # Create temporary input file
    # -----------------------------------------------------

    job_id = str(uuid.uuid4())

    input_extension = (
        Path(file.filename or "input.xlsx").suffix
        or ".xlsx"
    )

    input_path = (
        TEMP_DIR
        / f"{job_id}_input{input_extension}"
    )

    output_path = (
        TEMP_DIR
        / f"{job_id}_output.xlsx"
    )

    # -----------------------------------------------------
    # Save uploaded input
    # -----------------------------------------------------

    try:

        with input_path.open("wb") as buffer:

            shutil.copyfileobj(
                file.file,
                buffer
            )

        # -------------------------------------------------
        # Generate Excel
        # -------------------------------------------------

        generate_excel(
            structured_data,
            output_path
        )

        # -------------------------------------------------
        # Return generated file
        # -------------------------------------------------

        return FileResponse(
            path=str(output_path),
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            filename="output.xlsx"
        )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# ---------------------------------------------------------
# Local development
# ---------------------------------------------------------

if __name__ == "__main__":

    import uvicorn

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000
    )