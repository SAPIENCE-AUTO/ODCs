# main.py
import os
import re
import tempfile
from datetime import datetime
from typing import List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage


app = FastAPI(title="ODC Generator")


# -----------------------------
# Models
# -----------------------------
class OdcItem(BaseModel):
    concepto: str
    costo_unitario: float
    unidades: int
    subtotal: float


class OdcPayload(BaseModel):
    odc_num: str = Field(..., examples=["RI-02497"])
    fecha: str = Field(..., examples=["17 nov 2025"])
    proveedor: str
    servicio: str
    proyecto: str

    facturar_a: str
    rfc: str
    direccion: str

    items: List[OdcItem]


# -----------------------------
# Helpers
# -----------------------------
def safe_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r"[^\w\-\. ]+", "", name, flags=re.UNICODE)
    name = name.replace(" ", "_")
    return name or "file"


def parse_hex_color(hex_str: str) -> str:
    """Return 6-hex without #. Excel expects 'RRGGBB'."""
    s = hex_str.strip().lstrip("#")
    if len(s) != 6:
        raise ValueError(f"Invalid hex color: {hex_str}")
    return s.upper()


def add_image_if_exists(ws, path: str, anchor: str, *, max_w: Optional[int] = None, max_h: Optional[int] = None):
    """
    Adds image to worksheet if file exists.
    Optional max_w/max_h clamp.
    """
    if not os.path.exists(path):
        return False

    img = XLImage(path)

    # Clamp size if needed (rough; openpyxl uses px)
    if max_w and img.width and img.width > max_w:
        scale = max_w / float(img.width)
        img.width = int(img.width * scale)
        img.height = int(img.height * scale)

    if max_h and img.height and img.height > max_h:
        scale = max_h / float(img.height)
        img.width = int(img.width * scale)
        img.height = int(img.height * scale)

    ws.add_image(img, anchor)
    return True


def build_odc_layout(ws):
    """
    Camino B: layout 100% por código (openpyxl)
    - Barra superior con logo + badge ODC
    - Bloque izquierdo con labels azules y fondos grises
    - Bloque derecho "FACTURAR A"
    - Tabla con header azul + zebra
    """
    # ---- Page basics
    ws.title = "ODC"
    ws.sheet_view.showGridLines = False

    # ---- Column widths
    # Ajusta si necesitas (estos valores están calibrados para verse parecido a tu ejemplo)
    col_widths = {
        "A": 2.0,
        "B": 14.0,
        "C": 46.0,
        "D": 2.0,
        "E": 3.0,
        "F": 30.0,
        "G": 30.0,
        "H": 2.0,
        "I": 2.0,
        "J": 2.0,
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # ---- Row heights
    ws.row_dimensions[1].height = 8
    ws.row_dimensions[2].height = 58  # top bar
    ws.row_dimensions[3].height = 8

    # Left block rows
    for r in range(4, 9):
        ws.row_dimensions[r].height = 26

    ws.row_dimensions[9].height = 14
    ws.row_dimensions[10].height = 8
    ws.row_dimensions[11].height = 30  # table header

    # ---- Colors
    BLUE = parse_hex_color("#143847")       # Sapience-ish deep blue
    LIGHT_GREY = parse_hex_color("#EFEFEF") # left block fill / zebra
    WHITE = parse_hex_color("#FFFFFF")
    BLACK = parse_hex_color("#111111")
    RED = parse_hex_color("#E53935")

    # ---- Styles
    def f(size=12, bold=False, color=BLACK):
        return Font(name="Montserrat", size=size, bold=bold, color=color)

    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    right = Alignment(horizontal="right", vertical="center")

    thin = Side(style="thin", color="8A8A8A")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_none = Border()

    # ---- Top bar (A2:H2)
    ws.merge_cells("A2:H2")
    ws["A2"].fill = PatternFill("solid", fgColor=BLUE)
    ws["A2"].border = border_none

    # ---- Badge on the right (F2:H2)
    # White box for "ODC #:" + red box for number
    # To mimic your layout: make a white rectangle and a red inset rectangle.
    # We'll merge F2:G2 for "ODC #:" right-aligned and H2 for number.
    # But we want the red box to have more width; easiest: merge G2:H2 for red.
    ws.merge_cells("F2:G2")
    ws["F2"].value = "ODC #:"
    ws["F2"].font = f(size=14, bold=True, color=BLUE)
    ws["F2"].alignment = Alignment(horizontal="right", vertical="center")
    ws["F2"].fill = PatternFill("solid", fgColor=WHITE)
    ws["F2"].border = border_none

    ws.merge_cells("G2:H2")
    ws["G2"].value = "RI-00000"
    ws["G2"].font = f(size=14, bold=True, color=WHITE)
    ws["G2"].alignment = center
    ws["G2"].fill = PatternFill("solid", fgColor=RED)
    ws["G2"].border = border_none

    # ---- Left block (B4:C8)
    for r in range(4, 9):
        ws[f"B{r}"].fill = PatternFill("solid", fgColor=LIGHT_GREY)
        ws[f"C{r}"].fill = PatternFill("solid", fgColor=LIGHT_GREY)

        # separator line between label/value
        ws[f"B{r}"].border = Border(right=thin)
        ws[f"C{r}"].border = Border(left=thin)

    labels = [
        ("B4", "ODC #"),
        ("B5", "FECHA:"),
        ("B6", "PROVEEDOR:"),
        ("B7", "SERVICIO:"),
        ("B8", "PROYECTO:"),
    ]
    for cell, txt in labels:
        ws[cell].value = txt
        ws[cell].font = f(size=13, bold=True, color=BLUE)
        ws[cell].alignment = right

    for r in range(4, 9):
        ws[f"C{r}"].font = f(size=13, bold=False, color=BLACK)
        ws[f"C{r}"].alignment = left

    # ---- Right block (F4:G8)
    ws["F4"].value = "FACTURAR A:"
    ws["F4"].font = f(size=18, bold=True, color=BLUE)
    ws["F4"].alignment = left

    ws["F5"].value = ""
    ws["F5"].font = f(size=14, bold=True, color=BLACK)
    ws["F5"].alignment = left

    ws["F6"].value = ""
    ws["F6"].font = f(size=14, bold=True, color=BLACK)
    ws["F6"].alignment = left

    ws.merge_cells("F7:G8")
    ws["F7"].value = ""
    ws["F7"].font = f(size=13, bold=False, color=BLACK)
    ws["F7"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # ---- Table header (B11:G11)
    ws.merge_cells("B11:D11")
    ws["B11"].value = "Concepto"
    ws["E11"].value = "Costo unitario"
    ws["F11"].value = "Unidades"
    ws["G11"].value = "Subtotal"

    for cell in ["B11", "E11", "F11", "G11"]:
        ws[cell].fill = PatternFill("solid", fgColor=BLUE)
        ws[cell].font = f(size=14, bold=True, color=WHITE)
        ws[cell].alignment = center
        ws[cell].border = border_all

    # ensure borders for merged header cells
    for c in range(2, 8):  # B..G
        ws.cell(11, c).border = border_all

    ws.freeze_panes = "B12"


def write_odc_to_layout(ws, payload: OdcPayload):
    BLUE = parse_hex_color("#143847")
    LIGHT_GREY = parse_hex_color("#EFEFEF")
    BLACK = parse_hex_color("#111111")

    def f(size=12, bold=False, color=BLACK):
        return Font(name="Montserrat", size=size, bold=bold, color=color)

    thin = Side(style="thin", color="8A8A8A")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Badge
    ws["G2"].value = payload.odc_num

    # Left block values
    ws["C4"].value = payload.odc_num
    ws["C5"].value = payload.fecha
    ws["C6"].value = payload.proveedor
    ws["C7"].value = payload.servicio
    ws["C8"].value = payload.proyecto

    # Right block
    ws["F5"].value = payload.facturar_a
    ws["F6"].value = f"RFC: {payload.rfc}"
    ws["F7"].value = payload.direccion

    # Items table
    start_row = 12
    zebra = [parse_hex_color("#FFFFFF"), LIGHT_GREY]

    for i, it in enumerate(payload.items):
        r = start_row + i

        ws.merge_cells(f"B{r}:D{r}")
        ws[f"B{r}"].value = it.concepto
        ws[f"E{r}"].value = float(it.costo_unitario)
        ws[f"F{r}"].value = int(it.unidades)
        ws[f"G{r}"].value = float(it.subtotal)

        fill = PatternFill("solid", fgColor=zebra[i % 2])

        # Apply styles to all visible cells in row
        for cell in [f"B{r}", f"C{r}", f"D{r}", f"E{r}", f"F{r}", f"G{r}"]:
            ws[cell].fill = fill
            ws[cell].border = border_all
            ws[cell].font = f(size=13, bold=False)

        # Alignments
        ws[f"B{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws[f"E{r}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"F{r}"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"G{r}"].alignment = Alignment(horizontal="center", vertical="center")

        # Number formats
        ws[f"E{r}"].number_format = '"$"#,##0.00'
        ws[f"G{r}"].number_format = '"$"#,##0.00'

        ws.row_dimensions[r].height = 28


def add_brand_assets(ws):
    """
    Adds the Sapience logo (white) on the top bar.
    Expects the file to exist in repo.
    Adjust the path if your folder differs.
    """
    # Try common paths
    candidates = [
        os.path.join("assets", "logo sapience blanco.png"),
        os.path.join("assets", "logo_sapience_blanco.png"),
        os.path.join("assets", "logo.png"),
        "logo sapience blanco.png",
    ]
    for p in candidates:
        if os.path.exists(p):
            # Place on top bar left
            add_image_if_exists(ws, p, "A2", max_h=46)
            break


def generate_excel(payload: OdcPayload) -> str:
    """
    Returns path to a generated .xlsx in a temp file.
    """
    wb = Workbook()
    ws = wb.active

    build_odc_layout(ws)
    add_brand_assets(ws)
    write_odc_to_layout(ws, payload)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.close()
    wb.save(tmp.name)
    return tmp.name


# -----------------------------
# Routes
# -----------------------------
@app.get("/")
def root():
    return {"ok": True, "endpoints": ["/generate-odc-excel"]}


@app.post("/generate-odc-excel")
def generate_odc_excel(payload: OdcPayload):
    """
    Body: OdcPayload JSON
    Returns: .xlsx download
    """
    try:
        xlsx_path = generate_excel(payload)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate Excel: {e}")

    filename = safe_filename(f"ODC_{payload.odc_num}.xlsx")
    return FileResponse(
        path=xlsx_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# Optional: health check
@app.get("/health")
def health():
    return JSONResponse({"status": "healthy", "ts": datetime.utcnow().isoformat()})
