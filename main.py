# main.py
# Sapience ODCs — Excel generator
# FastAPI + XlsxWriter (Render-ready)

from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import List, Optional
from datetime import datetime
from io import BytesIO
import math
import textwrap
import requests
import struct

import xlsxwriter
from xlsxwriter.utility import xl_range

app = FastAPI(title="Sapience ODCs (Excel)", version="1.1.0")


# -----------------------------
# Models
# -----------------------------
class ODCItem(BaseModel):
    concept: str
    unit_cost: float
    units: float
    subtotal: Optional[float] = None


class ODCPayload(BaseModel):
    odc_number: str
    date_str: str
    provider: str
    service: str
    project: str

    bill_to_title: str = "FACTURAR A:"
    bill_to_name: str
    bill_to_rfc: str
    bill_to_address_1: str
    bill_to_address_2: str

    items: List[ODCItem]

    subtotal_amount: Optional[float] = None
    advance_amount: float = 0.0
    total: float

    currency_symbol: str = "$"
    logo_url: Optional[str] = "https://i.postimg.cc/Pf8KhptD/logo-sapience-blanco-15-ene-26.png"


# -----------------------------
# Helpers
# -----------------------------
def range_a1(r1, c1, r2, c2):
    return xl_range(r1 - 1, c1 - 1, r2 - 1, c2 - 1)


def safe_float(x):
    try:
        return float(x)
    except Exception:
        return 0.0


def fill_range(ws, r1, c1, r2, c2, fmt):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.write_blank(r - 1, c - 1, "", fmt)


def png_size(b):
    if b[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    w = struct.unpack(">I", b[16:20])[0]
    h = struct.unpack(">I", b[20:24])[0]
    return w, h


# -----------------------------
# Routes
# -----------------------------
@app.post("/generate-odc-excel")
def generate_excel(payload: ODCPayload):
    try:
        data = build_excel(payload)
        filename = f"ODC_{payload.odc_number}.xlsx"
        return StreamingResponse(
            BytesIO(data),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


# -----------------------------
# Excel builder
# -----------------------------
def build_excel(p: ODCPayload) -> bytes:
    out = BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = wb.add_worksheet("ODC")

    FONT = "Montserrat"
    TEAL = "#0F3D4C"
    TEAL_D = "#0E4A5A"
    GRAY = "#EFEFEF"
    GRID = "#8A8A8A"
    RED = "#E10600"
    WHITE = "#FFFFFF"
    BLACK = "#111111"

    # Column grid
    for c in range(26):
        ws.set_column(c, c, 2.6)

    white_bg = wb.add_format({"bg_color": WHITE})
    fill_range(ws, 1, 1, 200, 26, white_bg)

    # ---------------- Header (3 rows x 16 pts)
    ws.set_row(1, 16)
    ws.set_row(2, 16)
    ws.set_row(3, 16)

    banner = wb.add_format({"bg_color": TEAL})
    fill_range(ws, 2, 2, 4, 26, banner)

    # Logo (smaller)
    if p.logo_url:
        try:
            img = requests.get(p.logo_url, timeout=10).content
            size = png_size(img)
            scale = 0.16
            ws.insert_image(
                2, 1, "logo.png",
                {
                    "image_data": BytesIO(img),
                    "x_scale": scale,
                    "y_scale": scale,
                    "x_offset": 6,
                    "y_offset": 6,
                },
            )
        except Exception:
            pass

    # ODC box (compact)
    odc_lbl = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": TEAL_D, "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })
    odc_val = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": RED, "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })

    ws.merge_range(1, 19, 2, 22, "ODC #:", odc_lbl)
    ws.merge_range(1, 23, 2, 25, p.odc_number, odc_val)

    # ---------------- Left meta
    label = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_D, "bg_color": GRAY,
        "right": 1, "right_color": GRID,
    })
    value = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": GRAY,
        "text_wrap": True,
    })

    rows = [
        ("ODC #", p.odc_number),
        ("FECHA:", p.date_str),
        ("PROVEEDOR:", p.provider),
        ("SERVICIO:", p.service),
        ("PROYECTO:", p.project),
    ]

    for i, (k, v) in enumerate(rows):
        r = 5 + i
        ws.set_row(r - 1, 20)
        ws.merge_range(r - 1, 1, r - 1, 4, k, label)
        ws.merge_range(r - 1, 5, r - 1, 13, v, value)

    # ---------------- Bill to
    bt_title = wb.add_format({
        "font_name": FONT, "font_size": 10, "bold": True,
        "align": "left", "font_color": TEAL_D,
    })
    bt_bold = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
    })
    bt_norm = wb.add_format({
        "font_name": FONT, "font_size": 8,
    })

    ws.merge_range(4, 14, 4, 25, p.bill_to_title, bt_title)
    ws.merge_range(5, 14, 5, 25, p.bill_to_name, bt_bold)
    ws.merge_range(6, 14, 6, 25, f"RFC: {p.bill_to_rfc}", bt_bold)
    ws.merge_range(7, 14, 7, 25, p.bill_to_address_1, bt_norm)
    ws.merge_range(8, 14, 8, 25, p.bill_to_address_2, bt_norm)

    # ---------------- Table header
    th = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": WHITE, "bg_color": TEAL_D,
        "border": 1, "border_color": GRID,
    })

    header_row = 11
    ws.set_row(header_row - 1, 26)
    ws.merge_range(header_row - 1, 1,  header_row - 1, 13, "Concepto", th)
    ws.merge_range(header_row - 1, 14, header_row - 1, 18, "Costo unitario", th)
    ws.merge_range(header_row - 1, 19, header_row - 1, 21, "Unidades", th)
    ws.merge_range(header_row - 1, 22, header_row - 1, 25, "Subtotal", th)

    # ---------------- Table rows
    cell = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "left", "valign": "vcenter",
        "text_wrap": True,
        "border": 1, "border_color": GRID,
    })
    money = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": GRID,
        "num_format": f'"{p.currency_symbol}"#,##0.00',
    })
    units = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": GRID,
    })

    start = 12
    last = start

    for i, it in enumerate(p.items):
        r = start + i
        ws.set_row(r - 1, 30)
        subtotal = it.subtotal if it.subtotal is not None else it.unit_cost * it.units

        ws.merge_range(r - 1, 1,  r - 1, 13, it.concept, cell)
        ws.merge_range(r - 1, 14, r - 1, 18, it.unit_cost, money)
        ws.merge_range(r - 1, 19, r - 1, 21, it.units, units)
        ws.merge_range(r - 1, 22, r - 1, 25, subtotal, money)

        last = r

    # ---------------- Totals block
    suma = p.subtotal_amount if p.subtotal_amount is not None else sum(
        (it.subtotal if it.subtotal is not None else it.unit_cost * it.units)
        for it in p.items
    )

    r0 = last + 2
    ws.set_row(r0 - 1, 20)
    ws.set_row(r0, 20)
    ws.set_row(r0 + 1, 22)

    lbl = wb.add_format({
        "font_name": FONT, "font_size": 10, "bold": True,
        "align": "right",
        "font_color": TEAL_D,
    })
    val_gray = wb.add_format({
        "font_name": FONT, "font_size": 10,
        "align": "right",
        "font_color": "#6B6B6B",
        "num_format": f'"{p.currency_symbol}"#,##0.00',
    })
    val_red = wb.add_format({
        "font_name": FONT, "font_size": 10, "bold": True,
        "align": "right",
        "font_color": RED,
        "num_format": f'"{p.currency_symbol}"#,##0.00',
    })
    val_total = wb.add_format({
        "font_name": FONT, "font_size": 11, "bold": True,
        "align": "right",
        "font_color": BLACK,
        "num_format": f'"{p.currency_symbol}"#,##0.00',
    })

    ws.merge_range(r0 - 1, 19, r0 - 1, 21, "SUMA:", lbl)
    ws.merge_range(r0 - 1, 22, r0 - 1, 25, suma, val_gray)

    ws.merge_range(r0, 19, r0, 21, "ANTICIPO:", lbl)
    ws.merge_range(r0, 22, r0, 25, p.advance_amount, val_red)

    ws.merge_range(r0 + 1, 19, r0 + 1, 21, "TOTAL:", lbl)
    ws.merge_range(r0 + 1, 22, r0 + 1, 25, p.total, val_total)

    # ---------------- Print
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.set_margins(0.25, 0.25, 0.35, 0.35)
    ws.print_area(range_a1(1, 1, r0 + 4, 26))

    wb.close()
    return out.getvalue()
