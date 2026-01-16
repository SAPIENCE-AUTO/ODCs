import os
import uuid
from datetime import datetime
from typing import List, Optional

from fastapi import FastAPI
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# -------------------------
# Config
# -------------------------
TEMPLATE_PATH = os.getenv("TEMPLATE_PATH", "templates/odc_template.xlsx")
SHEET_NAME = os.getenv("SHEET_NAME", "")  # si usas template, puedes fijar hoja; si no, se usa activa
OUTPUT_DIR = "/tmp"  # Render permite escribir /tmp


# -------------------------
# Models
# -------------------------
class OdcItem(BaseModel):
    concepto: str
    costo_unitario: float
    unidades: int = Field(ge=1)

    @property
    def subtotal(self) -> float:
        return float(self.costo_unitario) * int(self.unidades)


class OdcPayload(BaseModel):
    odc_num: str
    fecha: str
    proveedor: str
    servicio: str
    proyecto: str
    facturar_a: str
    rfc: str
    direccion: str
    items: List[OdcItem]


# -------------------------
# App
# -------------------------
app = FastAPI(title="ODC XLSX Generator", version="1.0.0")


@app.get("/")
def health():
    return {
        "ok": True,
        "service": "odc-xlsx-generator",
        "template_path": TEMPLATE_PATH,
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "sheet_name": SHEET_NAME or "(active sheet)",
        "endpoints": {
            "generate": "POST /generate-odc-xlsx",
        },
    }


# -------------------------
# Helpers: workbook creation
# -------------------------
def _load_or_create_wb():
    """
    If template exists: load it.
    Else: create a clean workbook with a basic layout.
    """
    if os.path.exists(TEMPLATE_PATH):
        wb = load_workbook(TEMPLATE_PATH)
        if SHEET_NAME and SHEET_NAME in wb.sheetnames:
            ws = wb[SHEET_NAME]
        else:
            ws = wb.active
        return wb, ws

    # No template: build a minimal but decent Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "ODC"

    # Basic column widths
    widths = [6, 55, 14, 10, 14]  # #, concepto, unit, qty, subtotal
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Styles
    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="111827")  # dark
    header_font = Font(bold=True, color="FFFFFF")
    thin = Side(style="thin", color="9CA3AF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title row
    ws["A1"] = "ORDEN DE COMPRA"
    ws["A1"].font = Font(bold=True, size=16)
    ws.merge_cells("A1:E1")
    ws["A1"].alignment = Alignment(horizontal="center")

    # Meta labels
    ws["A3"] = "ODC #"
    ws["A4"] = "Fecha"
    ws["A5"] = "Proveedor"
    ws["A6"] = "Servicio"
    ws["A7"] = "Proyecto"
    ws["A8"] = "Facturar a"
    ws["A9"] = "RFC"
    ws["A10"] = "Dirección"

    for r in range(3, 11):
        ws[f"A{r}"].font = bold

    ws.merge_cells("B10:E10")

    # Table header
    start_row = 12
    headers = ["#", "Concepto", "Costo unitario", "Unidades", "Subtotal"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Freeze
    ws.freeze_panes = "A13"
    return wb, ws


def _find_table_start_row(ws) -> int:
    """
    If using a template, you can control where items begin.
    Convention:
      - If there's a cell with 'Concepto' header, table starts at that row, and items at next row.
      - Else default to row 13.
    """
    # scan a reasonable range
    for r in range(1, 60):
        for c in range(1, 10):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip().lower() == "concepto":
                return r  # header row
    return 12  # default header row


def _write_payload_to_sheet(ws, payload: OdcPayload):
    # If template: try to map to common cells. If not template: our layout expects these cells.
    # Safe writes with fallback.

    # Meta fields: try default cells (for our non-template layout)
    mapping = {
        "odc_num": ("B3", payload.odc_num),
        "fecha": ("B4", payload.fecha),
        "proveedor": ("B5", payload.proveedor),
        "servicio": ("B6", payload.servicio),
        "proyecto": ("B7", payload.proyecto),
        "facturar_a": ("B8", payload.facturar_a),
        "rfc": ("B9", payload.rfc),
        "direccion": ("B10", payload.direccion),
    }
    for _, (cell, value) in mapping.items():
        ws[cell] = value
        ws[cell].alignment = Alignment(wrap_text=True, vertical="top")

    # Find table header row, then write items from next row
    header_row = _find_table_start_row(ws)
    item_row = header_row + 1

    # Ensure table columns: 1..5
    thin = Side(style="thin", color="D1D5DB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    total = 0.0
    for idx, it in enumerate(payload.items, start=1):
        ws.cell(item_row, 1, idx).alignment = Alignment(horizontal="center")
        ws.cell(item_row, 2, it.concepto).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(item_row, 3, it.costo_unitario).number_format = '"$"#,##0.00'
        ws.cell(item_row, 4, it.unidades).alignment = Alignment(horizontal="center")
        ws.cell(item_row, 5, it.subtotal).number_format = '"$"#,##0.00'

        for c in range(1, 6):
            ws.cell(item_row, c).border = border

        total += it.subtotal
        item_row += 1

    # Total row
    ws.cell(item_row, 4, "TOTAL").font = Font(bold=True)
    ws.cell(item_row, 4).alignment = Alignment(horizontal="right")
    ws.cell(item_row, 5, total).font = Font(bold=True)
    ws.cell(item_row, 5).number_format = '"$"#,##0.00'

    for c in range(1, 6):
        ws.cell(item_row, c).border = border


# -------------------------
# Endpoint: Generate XLSX
# -------------------------
@app.post("/generate-odc-xlsx")
def generate_odc_xlsx(payload: OdcPayload):
    try:
        wb, ws = _load_or_create_wb()
        _write_payload_to_sheet(ws, payload)

        safe_odc = "".join(ch for ch in payload.odc_num if ch.isalnum() or ch in ("-", "_"))
        filename = f"ODC-{safe_odc or uuid.uuid4().hex}.xlsx"
        out_path = os.path.join(OUTPUT_DIR, f"{uuid.uuid4().hex}-{filename}")

        wb.save(out_path)

        # ✅ This guarantees a real download in browser/Swagger/n8n
        return FileResponse(
            path=out_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"ok": False, "error": str(e)},
        )
