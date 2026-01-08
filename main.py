from __future__ import annotations

import os
from io import BytesIO
from typing import List, Optional

import requests
from fastapi import FastAPI, Response
from pydantic import BaseModel, Field
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle


PAGE_W, PAGE_H = A4

SAP_BLUE = HexColor("#153646")
SAP_BLUE_2 = HexColor("#23495A")
LIGHT_GRAY = HexColor("#F2F2F2")
MID_GRAY = HexColor("#D9D9D9")
TEXT_DARK = HexColor("#111111")
ACCENT_RED = HexColor("#FF3B30")

DEFAULT_LOGO_URL = "https://i.imghippo.com/files/qW2090cp.png"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def money_fmt(value: float) -> str:
    return "${:,.2f}".format(float(value))


def safe_int(value) -> int:
    try:
        return int(value)
    except Exception:
        return 0


def safe_float(value) -> float:
    try:
        return float(value)
    except Exception:
        return 0.0


def register_fonts_if_present() -> None:
    fonts_dir = os.path.join(BASE_DIR, "assets", "fonts")
    files = {
        "Montserrat": "Montserrat-Regular.ttf",
        "Montserrat-Medium": "Montserrat-Medium.ttf",
        "Montserrat-SemiBold": "Montserrat-SemiBold.ttf",
        "Montserrat-Bold": "Montserrat-Bold.ttf",
    }
    for font_name, filename in files.items():
        path = os.path.join(fonts_dir, filename)
        if os.path.isfile(path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, path))
            except Exception:
                pass


def pick_font(preferred: str, fallback: str = "Helvetica") -> str:
    try:
        pdfmetrics.getFont(preferred)
        return preferred
    except Exception:
        return fallback


def load_logo_from_url(url: str) -> ImageReader:
    r = requests.get(url, timeout=20)
    r.raise_for_status()
    return ImageReader(BytesIO(r.content))


class FacturarA(BaseModel):
    razon_social: str
    rfc: str
    direccion_linea1: str
    direccion_linea2: Optional[str] = ""


class Item(BaseModel):
    concepto: str
    precio_unitario: float
    unidades: int

    @property
    def subtotal(self) -> float:
        return safe_float(self.precio_unitario) * safe_int(self.unidades)


class ODCRequest(BaseModel):
    odc_num: str = Field(..., examples=["RI-02497"])
    fecha: str = Field(..., examples=["08 enero 2026"])
    proveedor: str
    servicio: str
    proyecto: str

    facturar_a: FacturarA
    items: List[Item]
    anticipo: float = 0.0

    condiciones_titulo: str = "NOTAS Y CONDICIONES DE LA ORDEN DE COMPRA"
    condiciones: Optional[List[dict]] = None


def draw_header(c: canvas.Canvas, odc_num: str, logo: ImageReader) -> None:
    header_h = 90
    y0 = PAGE_H - header_h

    c.setFillColor(SAP_BLUE)
    c.rect(0, y0, PAGE_W, header_h, stroke=0, fill=1)

    c.drawImage(logo, 22, y0 + 22, width=190, height=48, mask="auto")

    box_w, box_h = 210, 34
    box_x = PAGE_W - box_w - 24
    box_y = y0 + 34

    c.setFillColor(LIGHT_GRAY)
    c.rect(box_x, box_y, box_w, box_h, stroke=0, fill=1)

    c.setStrokeColor(MID_GRAY)
    c.setLineWidth(1)
    c.line(box_x + 105, box_y, box_x + 105, box_y + box_h)

    font_sb = pick_font("Montserrat-SemiBold", "Helvetica-Bold")
    font_b = pick_font("Montserrat-Bold", "Helvetica-Bold")

    c.setFillColor(SAP_BLUE)
    c.setFont(font_sb, 12)
    c.drawRightString(box_x + 95, box_y + 11, "ODC #:")

    c.setFillColor(ACCENT_RED)
    c.setFont(font_b, 13)
    c.drawString(box_x + 115, box_y + 11, odc_num)


def draw_kv_block_left(c: canvas.Canvas, x: float, y_top: float, label: str, value: str) -> float:
    font_sb = pick_font("Montserrat-SemiBold", "Helvetica-Bold")
    font_r = pick_font("Montserrat", "Helvetica")

    row_h = 16
    label_w = 33 * mm
    value_w = 95 * mm
    block_h = 26

    c.setFillColor(LIGHT_GRAY)
    c.rect(x + label_w, y_top - block_h + 4, value_w, block_h, stroke=0, fill=1)

    c.setFillColor(SAP_BLUE)
    c.setFont(font_sb, 12)
    c.drawRightString(x + label_w - 6, y_top - row_h + 2, f"{label}:")

    c.setFillColor(TEXT_DARK)
    c.setFont(font_r, 12)
    c.drawString(x + label_w + 8, y_top - row_h + 2, value)

    c.setStrokeColor(MID_GRAY)
    c.setLineWidth(1)
    c.line(x + label_w, y_top - block_h + 4, x + label_w, y_top + 4)

    return y_top - (block_h + 8)


def draw_facturar_a(c: canvas.Canvas, x: float, y_top: float, data: FacturarA) -> None:
    font_b = pick_font("Montserrat-Bold", "Helvetica-Bold")
    font_sb = pick_font("Montserrat-SemiBold", "Helvetica-Bold")
    font_r = pick_font("Montserrat", "Helvetica")

    c.setFillColor(SAP_BLUE)
    c.setFont(font_b, 20)
    c.drawString(x, y_top, "FACTURAR A:")

    y = y_top - 28

    c.setFillColor(TEXT_DARK)
    c.setFont(font_b, 14)
    c.drawString(x, y, data.razon_social)

    y -= 22
    c.setFont(font_sb, 12)
    c.drawString(x, y, f"RFC: {data.rfc}")

    y -= 22
    c.setFont(font_r, 12)
    c.drawString(x, y, data.direccion_linea1)

    if (data.direccion_linea2 or "").strip():
        y -= 18
        c.drawString(x, y, data.direccion_linea2)


def build_items_table(items: List[Item]) -> Table:
    font_b = pick_font("Montserrat-Bold", "Helvetica-Bold")
    font_r = pick_font("Montserrat", "Helvetica")

    data = [["Concepto", "Precio unitario", "Unidades", "Subtotal"]]
    for it in items:
        data.append([it.concepto, money_fmt(it.precio_unitario), str(it.unidades), money_fmt(it.subtotal)])

    col_widths = [110 * mm, 40 * mm, 30 * mm, 40 * mm]
    t = Table(data, colWidths=col_widths)

    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), SAP_BLUE_2),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), font_b),
                ("FONTSIZE", (0, 0), (-1, 0), 11),
                ("ALIGN", (1, 0), (-1, 0), "CENTER"),
                ("FONTNAME", (0, 1), (-1, -1), font_r),
                ("FONTSIZE", (0, 1), (-1, -1), 10.5),
                ("GRID", (0, 0), (-1, -1), 0.6, MID_GRAY),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT_GRAY]),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("ALIGN", (1, 1), (-1, -1), "CENTER"),
                ("ALIGN", (0, 1), (0, -1), "LEFT"),
            ]
        )
    )
    return t


def draw_table(c: canvas.Canvas, table: Table, x: float, y_top: float) -> float:
    w, h = table.wrap(0, 0)
    table.drawOn(c, x, y_top - h)
    return h


def draw_totals_box(c: canvas.Canvas, x: float, y_top: float, subtotal: float, anticipo: float) -> None:
    font_b = pick_font("Montserrat-Bold", "Helvetica-Bold")
    font_sb = pick_font("Montserrat-SemiBold", "Helvetica-Bold")

    total = subtotal - safe_float(anticipo)

    box_w = 95 * mm
    row_h = 14 * mm
    box_h = row_h * 3

    c.setFillColor(SAP_BLUE_2)
    c.rect(x, y_top - box_h, box_w * 0.55, box_h, stroke=0, fill=1)

    c.setFillColor(colors.white)
    c.rect(x + box_w * 0.55, y_top - box_h, box_w * 0.45, box_h, stroke=1, fill=1)

    c.setStrokeColor(MID_GRAY)
    c.setLineWidth(1)
    c.line(x, y_top - row_h, x + box_w, y_top - row_h)
    c.line(x, y_top - row_h * 2, x + box_w, y_top - row_h * 2)

    labels = ["Subtotal", "Anticipo", "Total"]
    values = [
        money_fmt(subtotal),
        f"-{money_fmt(anticipo)}" if anticipo else money_fmt(0),
        money_fmt(total),
    ]

    for i, (lab, val) in enumerate(zip(labels, values)):
        y = y_top - row_h * i - 10 * mm

        c.setFillColor(colors.white)
        c.setFont(font_b, 13)
        c.drawRightString(x + box_w * 0.53, y, lab)

        if lab == "Anticipo" and safe_float(anticipo) > 0:
            c.setFillColor(ACCENT_RED)
        else:
            c.setFillColor(TEXT_DARK)

        c.setFont(font_sb, 13)
        c.drawString(x + box_w * 0.58, y, val)


def draw_footer_note(c: canvas.Canvas, text: str) -> None:
    font_r = pick_font("Montserrat", "Helvetica")
    c.setFillColor(TEXT_DARK)
    c.setFont(font_r, 9.5)
    c.drawCentredString(PAGE_W / 2, 18 * mm, text)


def draw_conditions_page(c: canvas.Canvas, title: str, sections: Optional[List[dict]]) -> None:
    c.showPage()

    font_b = pick_font("Montserrat-Bold", "Helvetica-Bold")
    font_sb = pick_font("Montserrat-SemiBold", "Helvetica-Bold")
    font_r = pick_font("Montserrat", "Helvetica")

    c.setFillColor(TEXT_DARK)
    c.setFont(font_b, 14)
    c.drawCentredString(PAGE_W / 2, PAGE_H - 25 * mm, title)

    margin_x = 18 * mm
    top = PAGE_H - 40 * mm
    col_gap = 14 * mm
    col_w = (PAGE_W - margin_x * 2 - col_gap) / 2

    x_left = margin_x
    x_right = margin_x + col_w + col_gap
    y_left = top
    y_right = top

    if not sections:
        c.setFont(font_r, 11)
        c.drawString(x_left, y_left, "Sin condiciones configuradas.")
        return

    def draw_section(x: float, y: float, sec: dict) -> float:
        c.setFont(font_sb, 11)
        c.setFillColor(TEXT_DARK)
        c.drawString(x, y, (sec.get("titulo", "") or "").strip())
        y -= 8 * mm

        bullets = sec.get("bullets", []) or []
        c.setFont(font_r, 10.2)

        for b in bullets:
            text = (b or "").strip()
            if not text:
                continue

            max_chars = 78
            lines = []
            while len(text) > max_chars:
                cut = text.rfind(" ", 0, max_chars)
                if cut <= 0:
                    cut = max_chars
                lines.append(text[:cut].strip())
                text = text[cut:].strip()
            if text:
                lines.append(text)

            for j, line in enumerate(lines):
                prefix = "— " if j == 0 else "  "
                c.drawString(x, y, f"{prefix}{line}")
                y -= 5.2 * mm

            y -= 2.5 * mm

        y -= 3 * mm
        return y

    for idx, sec in enumerate(sections):
        if idx % 2 == 0:
            y_left = draw_section(x_left, y_left, sec)
            if y_left < 25 * mm:
                c.showPage()
                c.setFont(font_b, 14)
                c.drawCentredString(PAGE_W / 2, PAGE_H - 25 * mm, title)
                y_left = top
                y_right = top
        else:
            y_right = draw_section(x_right, y_right, sec)
            if y_right < 25 * mm:
                c.showPage()
                c.setFont(font_b, 14)
                c.drawCentredString(PAGE_W / 2, PAGE_H - 25 * mm, title)
                y_left = top
                y_right = top


def generate_pdf(payload: ODCRequest) -> bytes:
    register_fonts_if_present()
    logo = load_logo_from_url(DEFAULT_LOGO_URL)

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    draw_header(c, payload.odc_num, logo)

    left_x = 18 * mm
    right_x = 115 * mm + 18 * mm
    y_top = PAGE_H - 105

    y = y_top
    y = draw_kv_block_left(c, left_x, y, "ODC #", payload.odc_num)
    y = draw_kv_block_left(c, left_x, y, "FECHA", payload.fecha)
    y = draw_kv_block_left(c, left_x, y, "PROVEEDOR", payload.proveedor)
    y = draw_kv_block_left(c, left_x, y, "SERVICIO", payload.servicio)
    y = draw_kv_block_left(c, left_x, y, "PROYECTO", payload.proyecto)

    draw_facturar_a(c, right_x, y_top - 18, payload.facturar_a)

    table = build_items_table(payload.items)
    table_x = 18 * mm
    table_y_top = y - 10 * mm
    table_h = draw_table(c, table, table_x, table_y_top)

    subtotal = sum([it.subtotal for it in payload.items])
    anticipo = safe_float(payload.anticipo)

    totals_x = PAGE_W - (18 * mm) - (95 * mm)
    totals_y_top = table_y_top - table_h - 10 * mm
    draw_totals_box(c, totals_x, totals_y_top, subtotal, anticipo)

    draw_footer_note(c, "Esta Orden de Compra constituye el acuerdo formal para la prestación del servicio descrito.")

    draw_conditions_page(c, payload.condiciones_titulo, payload.condiciones)

    c.save()
    return buf.getvalue()


app = FastAPI(title="ODCs PDF Generator", version="1.0.0")


@app.get("/")
def health():
    return {"ok": True, "service": "ODCs PDF Generator"}


@app.post("/generate-odc")
def generate_odc(payload: ODCRequest):
    pdf_bytes = generate_pdf(payload)
    filename = f"ODC-{payload.odc_num}.pdf".replace(" ", "_")
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(content=pdf_bytes, media_type="application/pdf", headers=headers)
