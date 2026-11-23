"""Excel cashflows, Pydantic validation and input corrections."""

from datetime import date, datetime
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from typing import Literal

from openpyxl import load_workbook
from pydantic import BaseModel, ConfigDict, Field, ValidationError, field_validator, model_validator

Currency = Literal["EUR", "GBP", "USD"]


def parse_date(value):
    if isinstance(value, str):
        value = " ".join(value.strip().removesuffix("`").strip().split())
        for fmt in ("%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y"):
            try:
                value = datetime.strptime(value, fmt)
                break
            except ValueError:
                continue
    if isinstance(value, datetime):
        if value.time() != datetime.min.time():
            raise ValueError("cashflow dates must be at midnight")
        return value.date()
    if isinstance(value, date):
        return value
    raise ValueError("expected an Excel, ISO or day-first date")


class Cashflow(BaseModel):
    model_config = ConfigDict(populate_by_name=True, str_strip_whitespace=True, extra="forbid")

    record_id: int = Field(alias="ID", gt=0, strict=True)
    fund_name: str = Field(alias="Fund Name", min_length=1)
    cashflow_date: date = Field(alias="Date")
    cashflow_type: Literal["Investment", "Interest", "Principal Repayment"] = Field(
        alias="Cashflow Type"
    )
    local_currency: Currency = Field(alias="Local Currency")
    amount_local: Decimal = Field(alias="Cashflow Amount Local", allow_inf_nan=False)
    amount_base: Decimal = Field(alias="Cashflow Amount Base", allow_inf_nan=False)
    base_currency: Currency = Field(alias="Base Currency")

    @field_validator("cashflow_date", mode="before")
    @classmethod
    def validate_date(cls, value):
        return parse_date(value)

    @field_validator("local_currency", "base_currency", mode="before")
    @classmethod
    def clean_currency(cls, value):
        if isinstance(value, str):
            value = value.strip().upper()
            return "GBP" if value == "GPB" else value
        return value

    @field_validator("amount_local", "amount_base", mode="before")
    @classmethod
    def reject_booleans(cls, value):
        if isinstance(value, bool):
            raise ValueError("a boolean is not a cashflow amount")
        return value

    @model_validator(mode="after")
    def validate_amounts(self):
        if self.cashflow_type == "Investment":
            if self.amount_local >= 0 or self.amount_base >= 0:
                raise ValueError("investment amounts must be negative")
        elif self.amount_local <= 0 or self.amount_base <= 0:
            raise ValueError("receipts must be positive")
        if self.local_currency == self.base_currency and abs(
            self.amount_local - self.amount_base
        ) > Decimal("0.01"):
            raise ValueError("local and base amounts must agree for the base currency")
        return self


def header_columns(sheet):
    names = {field.alias: name for name, field in Cashflow.model_fields.items()}
    columns = {}
    for cell in sheet[1]:
        if cell.value is None:
            continue
        name = cell.value.strip() if isinstance(cell.value, str) else cell.value
        if name not in names or names[name] in columns:
            return None
        columns[names[name]] = cell.column
    return columns if len(columns) == len(names) else None


def load_cashflows(path, sheet_name=None):
    snapshot = Path(path).read_bytes()
    values = load_workbook(BytesIO(snapshot), data_only=True)
    formulas = None
    records, corrections, errors = [], [], []
    try:
        formulas = load_workbook(BytesIO(snapshot), data_only=False)
        matching = [sheet.title for sheet in formulas if header_columns(sheet)]
        if sheet_name is None:
            if len(matching) != 1:
                raise ValueError("expected one cashflow sheet; use --sheet to select it")
            sheet_name = matching[0]
        if sheet_name not in formulas.sheetnames:
            raise ValueError(f"sheet {sheet_name!r} does not exist")
        sheet, source = values[sheet_name], formulas[sheet_name]
        columns = header_columns(source)
        if columns is None:
            raise ValueError("expected the eight cashflow headers, each exactly once")
        for field, column in columns.items():
            cell = source.cell(1, column)
            expected = Cashflow.model_fields[field].alias
            if cell.value != expected:
                corrections.append(
                    {"cell": cell.coordinate, "original": cell.value, "corrected": expected}
                )
        for row in source.iter_rows(min_row=2):
            if all(cell.value is None for cell in row):
                continue
            for cell in row:
                cached = sheet[cell.coordinate]
                if cell.value is not None and cell.column not in columns.values():
                    errors.append(f"{sheet_name}!{cell.coordinate}: data without a header")
                if cached.data_type == "e":
                    errors.append(f"{sheet_name}!{cell.coordinate}: Excel error {cached.value}")
                if cell.data_type == "f" and cached.value is None:
                    errors.append(
                        f"{sheet_name}!{cell.coordinate}: formula has no cached result; recalculate and save in Excel"
                    )
            raw = {field: sheet.cell(row[0].row, column).value for field, column in columns.items()}
            try:
                record = Cashflow.model_validate(raw)
            except ValidationError as error:
                errors.append(f"{sheet_name}!row {row[0].row}: {error}")
                continue
            records.append(record)
            for field in ("cashflow_date", "local_currency", "base_currency"):
                if isinstance(raw[field], str) and raw[field] != str(getattr(record, field)):
                    corrections.append(
                        {
                            "cell": sheet.cell(row[0].row, columns[field]).coordinate,
                            "original": raw[field],
                            "corrected": str(getattr(record, field)),
                        }
                    )
        seen, bases = set(), {}
        for record in records:
            if record.record_id in seen:
                errors.append(f"duplicate ID {record.record_id}")
            seen.add(record.record_id)
            if bases.setdefault(record.fund_name, record.base_currency) != record.base_currency:
                errors.append(f"mixed base currencies for {record.fund_name}")
        if not records:
            errors.append("no cashflows found")
        if errors:
            raise ValueError("\n".join(errors))
        return records, corrections
    finally:
        values.close()
        if formulas is not None:
            formulas.close()
