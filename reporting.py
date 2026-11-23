"""Write the results and corrected input to one Excel workbook."""

from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

NOTES = [
    "Annual IRRs use actual days / 365. Fund IRR uses the supplied base-currency amounts.",
    "NAV before payment includes today's cashflow; NAV after payment is the remaining invested exposure used for hedging.",
    "Sell 100% of foreign NAV against the base currency at each three-month reset. Receipts are assumed converted or distributed immediately.",
    "Coverage can drift between resets. Intermediate rolls need a spot leg or an agreed cash-settlement arrangement.",
    "Dates preserve end of month and are unadjusted. Forward rates are blank because quotes were not supplied. The source exit payment contains principal only.",
]


def write_results(path, records, corrections, results):
    book = Workbook()
    book.remove(book.active)
    tables = {
        **results,
        "Cleaned Input": [record.model_dump(by_alias=True) for record in records],
        "Corrections": corrections,
        "Notes": [{"Assumptions": note} for note in NOTES],
    }
    for title, rows in tables.items():
        sheet = book.create_sheet(title)
        if not rows:
            sheet.append(["No records"])
            continue
        headers = list(rows[0])
        sheet.append(headers)
        for row in rows:
            sheet.append(
                [float(value) if isinstance(value, Decimal) else value for value in row.values()]
            )
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            sheet.column_dimensions[cell.column_letter].width = min(
                max(len(str(cell.value)) + 2, 20), 40
            )
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                field = headers[cell.column - 1]
                if isinstance(cell.value, date):
                    cell.number_format = "yyyy-mm-dd"
                elif field == "annual_irr" or field.startswith("coverage_"):
                    cell.number_format = "0.000000%"
                elif isinstance(cell.value, float):
                    cell.number_format = "#,##0.00"
                elif isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.data_type = "s"
        if title == "Notes":
            sheet.column_dimensions["A"].width = 110
            for row in sheet.iter_rows(min_row=2):
                row[0].alignment = Alignment(wrap_text=True, vertical="top")
                sheet.row_dimensions[row[0].row].height = 32
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    book.save(path)
