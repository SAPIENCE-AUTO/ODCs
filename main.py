from __future__ import annotations

import io
import os
from datetime import date, datetime
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader


# -----------------------------
# App
# -----------------------------
app = FastAPI(title="ODCs Generator", version="1.0.0")

BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"


# -----------------------------
# Models
# -----------------------------
class BillTo(BaseModel):
    name: str
    rfc: str
    address: str


class ODCItem(BaseModel):
    concept: str
    unit_cost: float = Field(..., ge=0)
    units: int = Field(..., ge=1)


class ODCRequest(BaseModel):
    odc_prefix: str = Field(default="RI")
    odc_number: str = Field(..., description="Ej: 02497")
    date: Optional[str] = Field(default=None, description="YYYY-MM-DD. Si no, hoy.")
    provider: str
    service: str
    project: str
    bill_to: BillTo
    items: List[ODCItem]


# -----------------------------
# Helpers
# -----------------------------
def try_register_montserrat() -> dict:
    """
    Tries to register Montserrat fonts if present.
    Returns a dict with font names to use.
    """
    fonts_dir = ASSETS_DIR / "fonts"
    regular = fonts_dir / "Montserrat-Regular.ttf"
    bold = fonts_dir / "Montserrat-Bold.ttf"
    semibold = fonts_dir / "Montserrat-SemiBold.ttf"

    # Defaults (built-in)
    out = {
        "regular": "Helvetica",
        "bold": "Helvetica-Bold",
        "semibold": "Helvetica-Bold",
    }

    try:
        if regular.exists():
            pdfmetrics.registerFont(TTFont("Montserrat", str(regular)))
            out["regular"] = "Montserrat"
        if bold.exists():
            pdfmetrics.registerFont(TTFont("Montserrat-Bold", str(bold)))
            out["bold"] = "Montserrat-Bold"
        if semibold.exists():
            pdfmetrics.registerFont(TTFont("Montserrat-SemiBold", str(semibold)))
            out["semibold"] = "Montserrat-SemiBold"
    except Exception:
        # If anything goes wrong, just fall back silently
        pass

    return out


def money(v: float) -> str:
    return f"${v:,.2f}"


def parse_date(d: Optional[str]) -> date:
    if not d:
        return date.today()
    # Accept YYYY-MM-DD
    return datetime.strptime(d, "%Y-%m-%d").date()


def draw_wrapped_text(c: rl_canvas.Canvas, text: str, x: float, y: float, max_w: float, font: str, size: int, leading: float):
    """
    Basic wrapper: splits by spaces and wraps within max_w.
    Returns final y after drawing.
    """
    c.setFont(font, size)
    words = (text or "").split()
    line = ""
    lines = []
    for w in words:
        test = (line + " " + w).strip()
        if c.stringWidth(test, font, size) <= max_w:
            line = test
        else:
            if line:
                lines.append(line)
            line = w
    if line:
        lines.append(line)

    for ln in lines:
        c.drawString(x, y, ln)
        y -= leading
    return y


def get_logo_path() -> Path:
    # Prefer renamed version, but accept the original name with spaces
    p1 = ASSETS_DIR / "logo_sapience_blanco.png"
    p2 = ASSETS_DIR / "logo sapience blanco.png"
    if p1.exists():
        return p1
    return p2


# -----------------------------
# PDF Generator
# -----------------------------
def generate_odc_pdf(payload: ODCRequest) -> bytes:
    fonts = try_register_montserrat()

    buffer = io.BytesIO()
    c = rl_canvas.Canvas(buffer, pagesize=A4)
    w, h = A4

    # Brand colors (ajústalos si quieres)
    NAVY = colors.HexColor("#143647")   # header + table header
    LIGHT_GREY = colors.HexColor("#EEF2F5")
    MID_GREY = colors.HexColor("#D9E1E7")
    TEXT_DARK = colors.HexColor("#0E1A22")
    ACCENT_RED = colors.HexColor("#E53935")
    WHITE = colors.white

    # Margins
    left = 18 * mm
    right = 18 * mm
    top = 18 * mm
    bottom = 18 * mm
    content_w = w - left - right

    # Header bar
    header_h = 26 * mm
    c.setFillColor(NAVY)
    c.rect(0, h - header_h, w, header_h, fill=1, stroke=0)

    # Logo (left)
    logo_path = get_logo_path()
    if logo_path.exists():
        try:
            logo = ImageReader(str(logo_path))
            logo_w = 55 * mm
            logo_h = 16 * mm
            c.drawImage(logo, left, h - header_h + (header_h - logo_h) / 2, width=logo_w, height=logo_h, mask="auto")
        except Exception:
            # If image fails, just write text
            c.setFillColor(WHITE)
            c.setFont(fonts["bold"], 18)
            c.drawString(left, h - header_h + 8 * mm, "SAPIENCE")
    else:
        c.setFillColor(WHITE)
        c.setFont(fonts["bold"], 18)
        c.drawString(left, h - header_h + 8 * mm, "SAPIENCE")

    # ODC badge (right)
    odc_full = f"{payload.odc_prefix}-{payload.odc_number}"
    badge_h = 10 * mm
    badge_w = 65 * mm
    badge_x = w - right - badge_w
    badge_y = h - header_h + (header_h - badge_h) / 2
    c.setFillColor(WHITE)
    c.roundRect(badge_x, badge_y, badge_w, badge_h, 3, fill=1, stroke=0)

    # "ODC #:" label
    c.setFillColor(TEXT_DARK)
    c.setFont(fonts["bold"], 10)
    c.drawString(badge_x + 6, badge_y + 3.2, "ODC #:")

    # red pill for number
    pill_x = badge_x + 27
    pill_w = badge_w - 33
    c.setFillColor(ACCENT_RED)
    c.roundRect(pill_x, badge_y + 1.3, pill_w, badge_h - 2.6, 3, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont(fonts["bold"], 10)
    c.drawCentredString(pill_x + pill_w / 2, badge_y + 3.2, odc_full)

    # Body start
    y = h - header_h - 14 * mm

    # Left info box background
    box_h = 30 * mm
    box_w = (content_w * 0.58)
    box_x = left
    box_y = y - box_h
    c.setFillColor(LIGHT_GREY)
    c.rect(box_x, box_y, box_w, box_h, fill=1, stroke=0)

    # Right info area (Factura)
    r_x = left + box_w + 10 * mm
    r_w = w - right - r_x

    # Left labels + values
    labels = [
        ("ODC #", odc_full),
        ("FECHA", parse_date(payload.date).strftime("%d %b %Y").lower()),
        ("PROVEEDOR", payload.provider),
        ("SERVICIO", payload.service),
        ("PROYECTO", payload.project),
    ]

    label_x = box_x + 6 * mm
    value_x = box_x + 35 * mm
    row_y = y - 6 * mm
    row_gap = 6 * mm

    c.setFillColor(NAVY)
    c.setFont(fonts["bold"], 11)
    for i, (lab, val) in enumerate(labels):
        yy = row_y - i * row_gap
        c.drawRightString(value_x - 3 * mm, yy, f"{lab}:")
        c.setFillColor(TEXT_DARK)
        c.setFont(fonts["regular"], 11)
        c.drawString(value_x, yy, str(val))
        c.setFillColor(NAVY)
        c.setFont(fonts["bold"], 11)

    # Right: FACTURAR A
    c.setFillColor(NAVY)
    c.setFont(fonts["bold"], 14)
    c.drawString(r_x, y - 6 * mm, "FACTURAR A:")

    c.setFillColor(TEXT_DARK)
    c.setFont(fonts["bold"], 12)
    c.drawString(r_x, y - 13 * mm, payload.bill_to.name.upper())

    c.setFont(fonts["bold"], 11)
    c.drawString(r_x, y - 19 * mm, f"RFC: {payload.bill_to.rfc}")

    c.setFont(fonts["regular"], 11)
    draw_wrapped_text(
        c,
        payload.bill_to.address,
        r_x,
        y - 26 * mm,
        r_w,
        fonts["regular"],
        11,
        leading=4.8 * mm,
    )

    # Table
    y = box_y - 12 * mm

    table_x = left
    table_w = content_w
    table_top = y
    header_row_h = 10 * mm
    row_h = 9 * mm

    # Columns (like reference)
    col_concept = table_w * 0.52
    col_unit = table_w * 0.18
    col_units = table_w * 0.12
    col_sub = table_w * 0.18

    # Table header background
    c.setFillColor(NAVY)
    c.rect(table_x, table_top - header_row_h, table_w, header_row_h, fill=1, stroke=0)

    c.setFillColor(WHITE)
    c.setFont(fonts["bold"], 12)
    c.drawCentredString(table_x + col_concept / 2, table_top - 6.8 * mm, "Concepto")
    c.drawCentredString(table_x + col_concept + col_unit / 2, table_top - 6.8 * mm, "Costo unitario")
    c.drawCentredString(table_x + col_concept + col_unit + col_units / 2, table_top - 6.8 * mm, "Unidades")
    c.drawCentredString(table_x + col_concept + col_unit + col_units + col_sub / 2, table_top - 6.8 * mm, "Subtotal")

    # Rows
    y_row = table_top - header_row_h
    c.setStrokeColor(MID_GREY)
    c.setLineWidth(1)

    total = 0.0
    for idx, it in enumerate(payload.items):
        subtotal = float(it.unit_cost) * int(it.units)
        total += subtotal

        # zebra background
        if idx % 2 == 0:
            c.setFillColor(colors.white)
        else:
            c.setFillColor(colors.HexColor("#F6F8FA"))
        c.rect(table_x, y_row - row_h, table_w, row_h, fill=1, stroke=0)

        # Borders
        c.setStrokeColor(MID_GREY)
        c.rect(table_x, y_row - row_h, table_w, row_h, fill=0, stroke=1)

        # Vertical lines
        x1 = table_x + col_concept
        x2 = x1 + col_unit
        x3 = x2 + col_units
        c.line(x1, y_row - row_h, x1, y_row)
        c.line(x2, y_row - row_h, x2, y_row)
        c.line(x3, y_row - row_h, x3, y_row)

        # Text
        c.setFillColor(TEXT_DARK)
        c.setFont(fonts["regular"], 11)

        # Concept (wrap)
        concept_x = table_x + 3 * mm
        concept_y = y_row - 6.2 * mm
        draw_wrapped_text(c, it.concept, concept_x, concept_y, col_concept - 6 * mm, fonts["regular"], 11, leading=4.5 * mm)

        # Unit cost
        c.setFont(fonts["regular"], 11)
        c.drawCentredString(x1 + col_unit / 2, y_row - 6.2 * mm, money(it.unit_cost))

        # Units
        c.drawCentredString(x2 + col_units / 2, y_row - 6.2 * mm, str(it.units))

        # Subtotal
        c.drawCentredString(x3 + col_sub / 2, y_row - 6.2 * mm, money(subtotal))

        y_row -= row_h

        # Simple page break guard (if many rows)
        if y_row < bottom + 40 * mm:
            c.showPage()
            c.setFont(fonts["regular"], 11)
            y_row = h - top

    # Footer note
    footer_text = "Esta Orden de Compra constituye el acuerdo formal para la prestación del servicio descrito."
    c.setFillColor(colors.HexColor("#6B7780"))
    c.setFont(fonts["regular"], 10)
    c.drawCentredString(w / 2, bottom, footer_text)

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------
# Routes
# -----------------------------
@app.get("/")
def root():
    return {
        "ok": True,
        "service": "ODCs Generator",
        "endpoints": {
            "health": "/health",
            "generate_pdf": "/generate-odc",
        },
    }


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/generate-odc")
def generate_odc(req: ODCRequest):
    pdf_bytes = generate_odc_pdf(req)
    filename = f"ODC_{req.odc_prefix}-{req.odc_number}.pdf"
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"'
    }
    return StreamingResponse(io.BytesIO(pdf_bytes), media_type="application/pdf", headers=headers)
