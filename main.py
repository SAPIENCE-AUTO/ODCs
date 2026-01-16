import os
import io
import re
import uuid
import datetime as dt
from typing import List, Optional

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field
from fastapi.middleware.cors import CORSMiddleware

from starlette.background import BackgroundTask
from fastapi.responses import StreamingResponse

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# -----------------------------
# Helpers
# -----------------------------
def safe_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r"[^\w\-. ]+", "", name, flags=re.UNICODE)
    name = re.sub(r"\s+", "_", name)
    if not name:
        name = "file"
    return name


def pesos(value: float) -> str:
    # $1,400.00
    return f"${value:,.2f}"


def today_es() -> str:
    # "08 ene 2026"
    months = {
        1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
        7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"
    }
    d = dt.datetime.now()
    return f"{d.day:02d} {months[d.month]} {d.year}"


# -----------------------------
# Payload models
# -----------------------------
class OdcItem(BaseModel):
    concepto: str
    costo_unitario: float
    unidades: int
    subtotal: Optional[float] = None


class OdcPayload(BaseModel):
    odc_num: str = Field(..., examples=["RI-02497"])
    fecha: str = Field(default_factory=today_es, examples=["17 nov 2025"])
    proveedor: str
    servicio: str
    proyecto: str

    facturar_a: str
    rfc: str
    direccion: str

    items: List[OdcItem]


# -----------------------------
# Excel generation (Camino B)
# -----------------------------
def generate_excel(payload: OdcPayload) -> str:
    """
    Genera un Excel (xlsx) con formato aproximado a tu referencia.
    Ajusta medidas/colores en esta función según necesites.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "ODC"

    # Page setup (mejor para exportar a PDF desde Excel)
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    ws.sheet_view.showGridLines = False

    # Column widths (ajusta al gusto)
    col_widths = [4, 24, 34, 18, 14, 16, 4]  # A..G
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Row heights
    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 36
    ws.row_dimensions[3].height = 20
    for r in range(4, 40):
        ws.row_dimensions[r].height = 20

    # Brand colors
    DARK_BLUE = "143A4A"   # top bar
    LIGHT_GRAY = "F2F2F2"  # left labels background
    MID_GRAY = "E8ECEF"    # subtle areas
    RED = "E53935"

    white = Font(color="FFFFFF")
    mont_bold_18 = Font(name="Montserrat", bold=True, size=18, color="FFFFFF")
    mont_bold_12 = Font(name="Montserrat", bold=True, size=12, color="143A4A")
    mont_bold_11 = Font(name="Montserrat", bold=True, size=11, color="143A4A")
    mont_reg_11 = Font(name="Montserrat", size=11, color="000000")
    mont_reg_10 = Font(name="Montserrat", size=10, color="000000")
    mont_bold_11_white = Font(name="Montserrat", bold=True, size=11, color="FFFFFF")

    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    right = Alignment(horizontal="right", vertical="center")
    wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

    thin = Side(style="thin", color="6E6E6E")
    border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

    # -----------------------------
    # Header bar
    # -----------------------------
    ws.merge_cells("A1:G2")
    ws["A1"].fill = PatternFill("solid", fgColor=DARK_BLUE)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws["A1"].font = mont_bold_18
    ws["A1"].value = ""  # logo va como imagen

    # Logo (local file in repo)
    # Sube el archivo a /assets/logo sapience blanco.png
    logo_path = os.path.join(os.path.dirname(__file__), "assets", "logo sapience blanco.png")
    if os.path.exists(logo_path):
        img = XLImage(logo_path)
        img.height = 55
        img.width = 240
        ws.add_image(img, "A1")

    # ODC pill (arriba derecha)
    ws.merge_cells("E1:G1")
    ws["E1"].fill = PatternFill("solid", fgColor="F3F3F3")
    ws["E1"].alignment = Alignment(horizontal="right", vertical="center")
    ws["E1"].font = Font(name="Montserrat", bold=True, size=14, color="143A4A")
    ws["E1"].value = "ODC #:"

    ws.merge_cells("F1:G1")  # overlay for number look
    ws["F1"].fill = PatternFill("solid", fgColor="F3F3F3")
    ws["F1"].alignment = Alignment(horizontal="left", vertical="center")
    ws["F1"].font = Font(name="Montserrat", bold=True, size=16, color=RED)
    ws["F1"].value = f" {payload.odc_num}"

    # ensure header fill across
    for cell in ws["A1:G2"][0]:
        cell.fill = PatternFill("solid", fgColor=DARK_BLUE)

    # -----------------------------
    # Left info panel (ODC/FECHA/PROVEEDOR/SERVICIO/PROYECTO)
    # -----------------------------
    start_row = 4
    labels = ["ODC #", "FECHA:", "PROVEEDOR:", "SERVICIO:", "PROYECTO:"]
    values = [payload.odc_num, payload.fecha, payload.proveedor, payload.servicio, payload.proyecto]

    # background block
    ws.merge_cells(f"A{start_row}:C{start_row+4}")
    ws[f"A{start_row}"].fill = PatternFill("solid", fgColor=MID_GRAY)

    for i, (lab, val) in enumerate(zip(labels, values)):
        r = start_row + i
        ws.merge_cells(f"A{r}:B{r}")
        ws[f"A{r}"].value = lab
        ws[f"A{r}"].font = mont_bold_12
        ws[f"A{r}"].alignment = right
        ws[f"A{r}"].fill = PatternFill("solid", fgColor=LIGHT_GRAY)

        ws.merge_cells(f"C{r}:D{r}")
        ws[f"C{r}"].value = val
        ws[f"C{r}"].font = mont_reg_11
        ws[f"C{r}"].alignment = left
        ws[f"C{r}"].fill = PatternFill("solid", fgColor="FFFFFF")

        # divider line (visual)
        ws[f"B{r}"].border = Border(right=thin)

    # -----------------------------
    # Right "Facturar a" block
    # -----------------------------
    ws.merge_cells("E4:G4")
    ws["E4"].value = "FACTURAR A:"
    ws["E4"].font = Font(name="Montserrat", bold=True, size=20, color="143A4A")
    ws["E4"].alignment = left

    ws.merge_cells("E5:G5")
    ws["E5"].value = payload.facturar_a
    ws["E5"].font = Font(name="Montserrat", bold=True, size=13, color="000000")
    ws["E5"].alignment = left

    ws.merge_cells("E6:G6")
    ws["E6"].value = f"RFC: {payload.rfc}"
    ws["E6"].font = Font(name="Montserrat", bold=True, size=12, color="000000")
    ws["E6"].alignment = left

    ws.merge_cells("E7:G8")
    ws["E7"].value = payload.direccion
    ws["E7"].font = mont_reg_11
    ws["E7"].alignment = wrap

    # -----------------------------
    # Items table
    # -----------------------------
    table_top = 11

    # Header row
    ws.merge_cells(f"A{table_top}:C{table_top}")
    ws[f"A{table_top}"].value = "Concepto"
    ws[f"A{table_top}"].font = mont_bold_11_white
    ws[f"A{table_top}"].alignment = center
    ws[f"A{table_top}"].fill = PatternFill("solid", fgColor=DARK_BLUE)

    ws.merge_cells(f"D{table_top}:E{table_top}")
    ws[f"D{table_top}"].value = "Costo unitario"
    ws[f"D{table_top}"].font = mont_bold_11_white
    ws[f"D{table_top}"].alignment = center
    ws[f"D{table_top}"].fill = PatternFill("solid", fgColor=DARK_BLUE)

    ws[f"F{table_top}"].value = "Unidades"
    ws[f"F{table_top}"].font = mont_bold_11_white
    ws[f"F{table_top}"].alignment = center
    ws[f"F{table_top}"].fill = PatternFill("solid", fgColor=DARK_BLUE)

    ws[f"G{table_top}"].value = "Subtotal"
    ws[f"G{table_top}"].font = mont_bold_11_white
    ws[f"G{table_top}"].alignment = center
    ws[f"G{table_top}"].fill = PatternFill("solid", fgColor=DARK_BLUE)

    # Data rows
    row = table_top + 1
    total = 0.0
    for idx, item in enumerate(payload.items):
        subtotal = item.subtotal if item.subtotal is not None else float(item.costo_unitario) * int(item.unidades)
        total += subtotal

        fill = PatternFill("solid", fgColor="FFFFFF" if idx % 2 == 0 else "F7F7F7")

        ws.merge_cells(f"A{row}:C{row}")
        ws[f"A{row}"].value = item.concepto
        ws[f"A{row}"].font = mont_reg_11
        ws[f"A{row}"].alignment = left
        ws[f"A{row}"].fill = fill

        ws.merge_cells(f"D{row}:E{row}")
        ws[f"D{row}"].value = pesos(item.costo_unitario)
        ws[f"D{row}"].font = mont_reg_11
        ws[f"D{row}"].alignment = center
        ws[f"D{row}"].fill = fill

        ws[f"F{row}"].value = item.unidades
        ws[f"F{row}"].font = mont_reg_11
        ws[f"F{row}"].alignment = center
        ws[f"F{row}"].fill = fill

        ws[f"G{row}"].value = pesos(subtotal)
        ws[f"G{row}"].font = mont_reg_11
        ws[f"G{row}"].alignment = center
        ws[f"G{row}"].fill = fill

        # borders
        for c in ["A", "B", "C", "D", "E", "F", "G"]:
            ws[f"{c}{row}"].border = border_thin
        # header borders
        for c in ["A", "B", "C", "D", "E", "F", "G"]:
            ws[f"{c}{table_top}"].border = border_thin

        row += 1

    # Freeze panes under header
    ws.freeze_panes = f"A{table_top+1}"

    # Save to tmp file
    out_dir = "/tmp"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"ODC_{uuid.uuid4().hex}.xlsx")
    wb.save(out_path)
    return out_path


# -----------------------------
# FastAPI app
# -----------------------------
app = FastAPI(title="ODC Generator (Excel)", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
def root():
    return {"ok": True, "endpoints": ["/generate-odc-excel"]}


@app.post("/generate-odc-excel")
def generate_odc_excel(payload: OdcPayload):
    try:
        xlsx_path = generate_excel(payload)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate Excel: {e}")

    filename = safe_filename(f"ODC_{payload.odc_num}.xlsx")

    def cleanup():
        try:
            os.remove(xlsx_path)
        except Exception:
            pass

    file_like = open(xlsx_path, "rb")
    return StreamingResponse(
        file_like,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Cache-Control": "no-store",
        },
        background=BackgroundTask(cleanup),
    )
