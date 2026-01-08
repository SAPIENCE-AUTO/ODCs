import os
from io import BytesIO
from datetime import datetime

from fastapi import FastAPI
from fastapi.responses import Response
from pydantic import BaseModel, Field
from typing import List, Optional

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader


# -----------------------------
# App
# -----------------------------
app = FastAPI(title="ODCs Generator", version="1.0.0")


# -----------------------------
# Paths / Assets
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")

LOGO_PATH = os.path.join(ASSETS_DIR, "logo_sapience_blanco.png")
FONTS_DIR = os.path.join(ASSETS_DIR, "fonts")

FONT_REGULAR = os.path.join(FONTS_DIR, "Montserrat-Regular.ttf")
FONT_SEMIBOLD = os.path.join(FONTS_DIR, "Montserrat-SemiBold.ttf")
FONT_BOLD = os.path.join(FONTS_DIR, "Montserrat-Bold.ttf")


def _register_fonts():
    """
    Register Montserrat fonts if present; otherwise fallback to Helvetica.
    """
    try:
        if os.path.exists(FONT_REGULAR):
            pdfmetrics.registerFont(TTFont("Montserrat", FONT_REGULAR))
        if os.path.exists(FONT_SEMIBOLD):
            pdfmetrics.registerFont(TTFont("Montserrat-SemiBold", FONT_SEMIBOLD))
        if os.path.exists(FONT_BOLD):
            pdfmetrics.registerFont(TTFont("Montserrat-Bold", FONT_BOLD))
    except Exception:
        # If something fails, we will fallback silently
        pass


def _font(name: str) -> str:
    """
    Safe font selector.
    """
    available = set(pdfmetrics.getRegisteredFontNames())
    if name in available:
        return name
    # fallback
    return "Helvetica"


# -----------------------------
# Input schema
# -----------------------------
class ODCItem(BaseModel):
    concepto: str
    costo_unitario: float
    unidades: int

    @property
    def subtotal(self) -> float:
        return float(self.costo_unitario) * int(self.unidades)


class ODCRequest(BaseModel):
    odc_num: str = Field(..., example="RI-02497")
    fecha: str = Field(..., example="08 ene 2026")  # puedes mandar texto ya formateado
    proveedor: str = Field(..., example="María Guadalupe Garza Sardaneta")
    servicio: str = Field(..., example="Reclutamiento")
    proyecto: str = Field(..., example="ALONG")

    facturar_a_nombre: str = Field(..., example="ASESORES GLOBALES CORPORATIVOS")
    facturar_a_rfc: str = Field(..., example="AGC051117MX5")
    facturar_a_direccion: str = Field(..., example="Peregrinos 24, Colinas del Sur,\nÁlvaro Obregón, CP. 01430, CDMX")

    items: List[ODCItem]

    mostrar_totales: bool = True


# -----------------------------
# Helpers
# -----------------------------
def money_mx(value: float) -> str:
    # Formato tipo $4,200
    return "${:,.0f}".format(value)


def money_mx_2(value: float) -> str:
    # Formato tipo $1,400.00
    return "${:,.2f}".format(value)


def draw_text(c: canvas.Canvas, x, y, text, font_name, font_size, color=colors.black):
    c.setFillColor(color)
    c.setFont(font_name, font_size)
    c.drawString(x, y, text)


def draw_text_right(c: canvas.Canvas, x, y, text, font_name, font_size, color=colors.black):
    c.setFillColor(color)
    c.setFont(font_name, font_size)
    c.drawRightString(x, y, text)


def draw_text_center(c: canvas.Canvas, x, y, text, font_name, font_size, color=colors.black):
    c.setFillColor(color)
    c.setFont(font_name, font_size)
    c.drawCentredString(x, y, text)


# -----------------------------
# PDF builder
# -----------------------------
def build_odc_pdf(payload: ODCRequest) -> bytes:
    _register_fonts()

    W, H = A4
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    # Colors
    BLUE = colors.HexColor("#123A4B")
    LIGHT_GRAY = colors.HexColor("#EFEFEF")
    LINE_GRAY = colors.HexColor("#8A8A8A")
    RED = colors.HexColor("#E53935")
    TEXT = colors.HexColor("#111111")

    # Fonts (safe)
    F_REG = _font("Montserrat")
    F_SEMI = _font("Montserrat-SemiBold")
    F_BOLD = _font("Montserrat-Bold")

    # Margins / layout
    M = 18 * mm
    header_h = 42 * mm

    # -----------------------------
    # Header bar
    # -----------------------------
    c.setFillColor(BLUE)
    c.rect(0, H - header_h, W, header_h, fill=1, stroke=0)

    # Logo
    if os.path.exists(LOGO_PATH):
        try:
            logo = ImageReader(LOGO_PATH)
            # Caja del logo aprox
            logo_w = 78 * mm
            logo_h = 22 * mm
            logo_x = M
            logo_y = H - header_h + (header_h - logo_h) / 2
            c.drawImage(logo, logo_x, logo_y, width=logo_w, height=logo_h, mask="auto")
        except Exception:
            pass

    # ODC box (right)
    odc_box_w = 72 * mm
    odc_box_h = 14 * mm
    odc_box_x = W - M - odc_box_w
    odc_box_y = H - (header_h / 2) - (odc_box_h / 2)

    c.setFillColor(colors.white)
    c.roundRect(odc_box_x, odc_box_y, odc_box_w, odc_box_h, radius=2.5 * mm, fill=1, stroke=0)

    # Separator inside box
    sep_x = odc_box_x + 32 * mm
    c.setStrokeColor(colors.HexColor("#2A2A2A"))
    c.setLineWidth(0.9)
    c.line(sep_x, odc_box_y + 2, sep_x, odc_box_y + odc_box_h - 2)

    # Text inside box
    draw_text_right(
        c,
        sep_x - 2,
        odc_box_y + 4.2 * mm,
        "ODC #:",
        F_BOLD,
        12,
        color=BLUE
    )
    draw_text(
        c,
        sep_x + 2,
        odc_box_y + 4.2 * mm,
        payload.odc_num,
        F_BOLD,
        12,
        color=RED
    )

    # -----------------------------
    # Top blocks
    # -----------------------------
    top_y = H - header_h - 16 * mm

    # Left block dimensions
    left_x = M
    left_w = 120 * mm
    row_h = 11 * mm
    label_w = 34 * mm
    divider_x = left_x + label_w

    rows = [
        ("ODC #", payload.odc_num),
        ("FECHA", payload.fecha),
        ("PROVEEDOR", payload.proveedor),
        ("SERVICIO", payload.servicio),
        ("PROYECTO", payload.proyecto),
    ]

    # Draw left block rows with light gray backgrounds
    c.setLineWidth(1)
    for i, (lab, val) in enumerate(rows):
        y = top_y - i * row_h

        # background band (full row)
        c.setFillColor(LIGHT_GRAY)
        c.rect(left_x, y - row_h + 1, left_w, row_h, fill=1, stroke=0)

        # label
        draw_text_right(c, divider_x - 2, y - 7.5, f"{lab}:", F_BOLD, 12, color=BLUE)

        # value
        draw_text(c, divider_x + 4, y - 7.5, str(val), F_REG, 12, color=TEXT)

    # Vertical divider line for left block
    c.setStrokeColor(colors.HexColor("#5A5A5A"))
    c.setLineWidth(1.2)
    top_block_y1 = top_y + 1
    top_block_y0 = top_y - len(rows) * row_h + 1
    c.line(divider_x, top_block_y0, divider_x, top_block_y1)

    # Right block (Facturar A)
    right_x = left_x + left_w + 24 * mm
    right_y = top_y

    draw_text(c, right_x, right_y - 7.5, "FACTURAR A:", F_BOLD, 22, color=BLUE)
    draw_text(c, right_x, right_y - 20 * mm, payload.facturar_a_nombre, F_BOLD, 14, color=TEXT)
    draw_text(c, right_x, right_y - 30 * mm, f"RFC: {payload.facturar_a_rfc}", F_BOLD, 13, color=TEXT)

    # Multi-line address
    c.setFont(F_REG, 12)
    c.setFillColor(TEXT)
    addr_lines = str(payload.facturar_a_direccion).split("\n")
    addr_y = right_y - 42 * mm
    for line in addr_lines:
        c.drawString(right_x, addr_y, line.strip())
        addr_y -= 6.5 * mm

    # -----------------------------
    # Table
    # -----------------------------
    table_x = M
    table_w = W - 2 * M
    table_y_top = top_block_y0 - 20 * mm

    th = 14 * mm  # header height
    tr = 12.5 * mm  # row height

    # Column widths (tune to match ref)
    col_concept = table_w * 0.56
    col_unit = table_w * 0.15
    col_units = table_w * 0.13
    col_sub = table_w * 0.16

    col_x = [
        table_x,
        table_x + col_concept,
        table_x + col_concept + col_unit,
        table_x + col_concept + col_unit + col_units,
        table_x + table_w,
    ]

    # Header background
    c.setFillColor(BLUE)
    c.rect(table_x, table_y_top - th, table_w, th, fill=1, stroke=0)

    # Header text
    draw_text_center(c, table_x + col_concept / 2, table_y_top - 10.5 * mm, "Concepto", F_BOLD, 15, color=colors.white)
    draw_text_center(c, col_x[1] + col_unit / 2, table_y_top - 10.5 * mm, "Costo unitario", F_BOLD, 15, color=colors.white)
    draw_text_center(c, col_x[2] + col_units / 2, table_y_top - 10.5 * mm, "Unidades", F_BOLD, 15, color=colors.white)
    draw_text_center(c, col_x[3] + col_sub / 2, table_y_top - 10.5 * mm, "Subtotal", F_BOLD, 15, color=colors.white)

    # Table borders
    c.setStrokeColor(LINE_GRAY)
    c.setLineWidth(1)

    # Header outer border
    c.rect(table_x, table_y_top - th, table_w, th, fill=0, stroke=1)

    # Vertical separators (including header)
    for x in col_x[1:-1]:
        c.line(x, table_y_top - th, x, table_y_top)

    # Rows
    y = table_y_top - th
    for idx, item in enumerate(payload.items):
        row_y_top = y - idx * tr
        row_y_bottom = row_y_top - tr

        # zebra fill
        if idx % 2 == 1:
            c.setFillColor(LIGHT_GRAY)
            c.rect(table_x, row_y_bottom, table_w, tr, fill=1, stroke=0)

        # row border
        c.setStrokeColor(LINE_GRAY)
        c.setLineWidth(1)
        c.rect(table_x, row_y_bottom, table_w, tr, fill=0, stroke=1)

        # vertical separators
        for x in col_x[1:-1]:
            c.line(x, row_y_bottom, x, row_y_top)

        # text
        concept_txt = item.concepto
        unit_txt = money_mx(item.costo_unitario) if abs(item.costo_unitario - round(item.costo_unitario)) < 1e-9 else money_mx_2(item.costo_unitario)
        units_txt = str(item.unidades)
        sub_txt = money_mx(item.subtotal)

        # Concept
        draw_text(c, table_x + 3 * mm, row_y_bottom + 4.2 * mm, concept_txt, F_REG, 14, color=TEXT)

        # Unit
        draw_text_center(c, col_x[1] + col_unit / 2, row_y_bottom + 4.2 * mm, unit_txt, F_REG, 14, color=TEXT)

        # Units
        draw_text_center(c, col_x[2] + col_units / 2, row_y_bottom + 4.2 * mm, units_txt, F_REG, 14, color=TEXT)

        # Subtotal
        draw_text_center(c, col_x[3] + col_sub / 2, row_y_bottom + 4.2 * mm, sub_txt, F_REG, 14, color=TEXT)

    # Totals (optional) — si quieres la versión con totales abajo a la derecha
    if payload.mostrar_totales:
        total = sum(i.subtotal for i in payload.items)
        totals_y = (table_y_top - th) - len(payload.items) * tr - 18 * mm
        # puedes activar si lo necesitas; por ahora lo dejo apagable desde JSON
        # draw_text_right(c, W - M - 40*mm, totals_y, "Total:", F_BOLD, 14, color=TEXT)
        # draw_text_right(c, W - M, totals_y, money_mx(total), F_BOLD, 14, color=TEXT)

    c.showPage()
    c.save()
    return buf.getvalue()


# -----------------------------
# Endpoints
# -----------------------------
@app.get("/")
def health():
    return {"ok": True, "service": "odcs"}


@app.post("/generate-odc")
def generate_odc(payload: ODCRequest):
    pdf_bytes = build_odc_pdf(payload)
    filename = f"ODC-{payload.odc_num}.pdf".replace(" ", "_")

    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )
