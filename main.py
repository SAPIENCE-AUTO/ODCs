# main.py
# Sapience ODCs — Render + FastAPI — Excel only (XlsxWriter)
# Key rules:
# - White background is explicit (no "transparent" cells).
# - Small/controlled merges only (for layout + to avoid ####).
# - TOTAL font size = 11 and always readable.
# - "FACTURAR A:" aligned LEFT.
# - Concept rows auto-height with extra vertical "air".

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

app = FastAPI(title="Sapience ODCs (Excel)", version="1.0.4")


# -----------------------------
# Models
# -----------------------------
class ODCItem(BaseModel):
    concept: str = ""
    unit_cost: float = 0
    units: float = 0
    subtotal: Optional[float] = None  # if omitted, we compute = unit_cost * units


class ODCPayload(BaseModel):
    # Header / left meta
    odc_number: str
    date_str: str
    provider: str
    service: str
    project: str

    # Bill-to
    bill_to_title: str = "FACTURAR A:"
    bill_to_name: str
    bill_to_rfc: str
    bill_to_address_1: str
    bill_to_address_2: str

    # Items
    items: List[ODCItem] = Field(default_factory=list)

    # Total comes from Monday (number)
    total: float
    currency_symbol: str = "$"

    # Logo
    logo_url: Optional[str] = "https://i.postimg.cc/Pf8KhptD/logo-sapience-blanco-15-ene-26.png"


# -----------------------------
# Helpers
# -----------------------------
def _png_size(img_bytes: bytes):
    """Return (w,h) for PNG bytes or None."""
    if len(img_bytes) < 24 or img_bytes[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    try:
        w = struct.unpack(">I", img_bytes[16:20])[0]
        h = struct.unpack(">I", img_bytes[20:24])[0]
        return int(w), int(h)
    except Exception:
        return None


def range_a1(row1: int, col1: int, row2: int, col2: int) -> str:
    """1-based row/col -> Excel range like A1:Z20"""
    return xl_range(row1 - 1, col1 - 1, row2 - 1, col2 - 1)


def safe_float(x) -> float:
    try:
        return float(x)
    except Exception:
        return 0.0


def fill_range(ws, row1: int, col1: int, row2: int, col2: int, fmt):
    """Paint a rectangle explicitly with blank cells (avoid transparent background)."""
    for rr in range(row1, row2 + 1):
        for cc in range(col1, col2 + 1):
            ws.write_blank(rr - 1, cc - 1, "", fmt)


def row_height_for_wrapped_text(
    text: str,
    wrap_width_chars: int,
    base_line_height: float = 12.5,
    extra_lines: float = 1.0,
) -> float:
    """
    No real padding in XlsxWriter; we simulate "air" by row height.
    - estimate wrapped lines
    - add extra_lines for breathing room
    """
    if not text:
        return base_line_height * (1 + extra_lines)

    paragraphs = str(text).splitlines() or [""]
    total_lines = 0
    for p in paragraphs:
        wrapped = textwrap.wrap(
            p,
            width=max(1, wrap_width_chars),
            break_long_words=True,
            replace_whitespace=False,
        ) or [""]
        total_lines += len(wrapped)

    return base_line_height * (total_lines + extra_lines)


# -----------------------------
# Routes
# -----------------------------
@app.get("/health")
def health():
    return {"ok": True, "time": datetime.now().isoformat()}


@app.post("/generate-odc-excel")
def generate_odc_excel(payload: ODCPayload):
    try:
        xlsx_bytes = build_odc_excel(payload)
        filename = f"ODC_{payload.odc_number}_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
        return StreamingResponse(
            BytesIO(xlsx_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})


# -----------------------------
# Excel Builder (XlsxWriter)
# -----------------------------
def build_odc_excel(payload: ODCPayload) -> bytes:
    out = BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = wb.add_worksheet("ODC")

    # -------- Palette / Font --------
    FONT = "Montserrat"
    TEAL = "#0F3D4C"
    TEAL_2 = "#0E4A5A"
    WHITE = "#FFFFFF"
    LIGHT_GRAY = "#EFEFEF"
    GRID = "#7C7C7C"
    RED = "#E10600"
    BLACK = "#111111"

    # -------- Grid columns (uniform 2.5) --------
    # Columns A..Z (26)
    for c in range(0, 26):
        ws.set_column(c, c, 2.5)

    # -------- Explicit white canvas (no transparency) --------
    white_bg = wb.add_format({"bg_color": WHITE})
    # paint a safe canvas; print_area will cut it later
    fill_range(ws, 1, 1, 200, 26, white_bg)

    # -------- Base fills --------
    banner_fill = wb.add_format({"bg_color": TEAL})
    gray_fill = wb.add_format({"bg_color": LIGHT_GRAY})

    # -------- Formats --------
    # ODC box (top-right) - WHITE background (not banner)
    odc_box_lbl = wb.add_format({
        "font_name": FONT, "font_size": 14, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })
    odc_box_val = wb.add_format({
        "font_name": FONT, "font_size": 14, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": RED, "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })

    # Left meta labels/values (striped gray/white)
    label_gray = wb.add_format({
        "font_name": FONT, "font_size": 11, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": LIGHT_GRAY
    })
    value_gray = wb.add_format({
        "font_name": FONT, "font_size": 11,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": LIGHT_GRAY,
        "text_wrap": True
    })
    label_white = wb.add_format({
        "font_name": FONT, "font_size": 11, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })
    value_white = wb.add_format({
        "font_name": FONT, "font_size": 11,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE,
        "text_wrap": True
    })

    # Bill-to block (FACTURAR A aligned left)
    bill_title_fmt = wb.add_format({
        "font_name": FONT, "font_size": 18, "bold": True,
        "align": "left",          # ✅ left aligned
        "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })
    bill_bold_fmt = wb.add_format({
        "font_name": FONT, "font_size": 12, "bold": True,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE
    })
    bill_norm_fmt = wb.add_format({
        "font_name": FONT, "font_size": 12,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE
    })

    # Table header
    th_fmt = wb.add_format({
        "font_name": FONT, "font_size": 12, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": WHITE, "bg_color": TEAL_2,
        "border": 1, "border_color": GRID,
    })

    # Table cells (concept uses top valign + wrap)
    concept_w = wb.add_format({
        "font_name": FONT, "font_size": 10,
        "align": "left", "valign": "top",
        "font_color": BLACK,
        "text_wrap": True,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })
    concept_g = wb.add_format({
        "font_name": FONT, "font_size": 10,
        "align": "left", "valign": "top",
        "font_color": BLACK,
        "text_wrap": True,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
    })

    money_w = wb.add_format({
        "font_name": FONT, "font_size": 11,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
        "num_format": f'"{payload.currency_symbol}"#,##0'
    })
    money_g = wb.add_format({
        "font_name": FONT, "font_size": 11,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
        "num_format": f'"{payload.currency_symbol}"#,##0'
    })

    units_w = wb.add_format({
        "font_name": FONT, "font_size": 11,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })
    units_g = wb.add_format({
        "font_name": FONT, "font_size": 11,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
    })

    # Total formats (font 11)
    total_label_fmt = wb.add_format({
        "font_name": FONT, "font_size": 11, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })
    total_money_fmt = wb.add_format({
        "font_name": FONT, "font_size": 11, "bold": True,  # ✅ 11
        "align": "right", "valign": "vcenter",
        "font_color": RED, "bg_color": WHITE,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })

    # -------- Layout coordinates (0-based col indices, but we use 0-based in merge_range) --------
    # A..Z => 0..25
    LEFT = 1     # B
    RIGHT = 25   # Z

    # Banner: rows 1..4 (1-based) => we fill B2:Z4 (as per your layout)
    # We'll implement with 1-based in our fill helper:
    # Make banner height with a "thin" first row, then two taller, then medium.
    ws.set_row(0, 10)   # row 1
    ws.set_row(1, 38)   # row 2
    ws.set_row(2, 38)   # row 3
    ws.set_row(3, 26)   # row 4

    fill_range(ws, 2, 2, 4, 26, banner_fill)  # B2:Z4

    # Insert logo (smaller + centered-ish)
    if payload.logo_url:
        try:
            resp = requests.get(payload.logo_url, timeout=15)
            resp.raise_for_status()
            img = resp.content
            wh = _png_size(img)

            # place at B2 (row=2,col=B)
            x_scale = y_scale = 0.20
            y_off = 4

            if wh:
                w_px, h_px = wh
                # Soft cap to keep it small
                scale = min(460 / max(1, w_px), 105 / max(1, h_px))
                scale = max(0.10, min(scale, 0.26))
                x_scale = y_scale = scale

                # center vertically across rows 2..4 (approx height in px)
                target_px_h = (38 + 38 + 26) * 4 / 3
                scaled_h = h_px * y_scale
                y_off = max(0, int((target_px_h - scaled_h) / 2))

            ws.insert_image(
                1,  # row index 1 => Excel row 2
                LEFT,  # col index 1 => B
                "logo.png",
                {
                    "image_data": BytesIO(img),
                    "x_scale": x_scale,
                    "y_scale": y_scale,
                    "x_offset": 6,
                    "y_offset": y_off,
                    "object_position": 1,
                },
            )
        except Exception:
            pass

    # ODC box top-right inside banner: rows 2..3 (Excel) and cols T..Z
    # T=20th col (1-based) => index 19 (0-based). Z=26th => index 25.
    # We'll use merges:
    odc_r1, odc_r2 = 2, 3  # 1-based rows
    ws.merge_range(odc_r1 - 1, 19, odc_r2 - 1, 22, "ODC #:", odc_box_lbl)          # T..W
    ws.merge_range(odc_r1 - 1, 23, odc_r2 - 1, 25, payload.odc_number, odc_box_val) # X..Z

    # Left meta block rows 5..9 (Excel)
    meta_rows = [
        ("ODC #", payload.odc_number),
        ("FECHA:", payload.date_str),
        ("PROVEEDOR:", payload.provider),
        ("SERVICIO:", payload.service),
        ("PROYECTO:", payload.project),
    ]
    # Left block columns B..N (B=2, N=14 1-based) => indices 1..13
    for i, (lab, val) in enumerate(meta_rows):
        rr = 5 + i  # 1-based row
        ws.set_row(rr - 1, 22)

        is_gray = (i % 2 == 0)  # rows 5,7,9 gray stripes
        fill_range(ws, rr, 2, rr, 14, gray_fill if is_gray else white_bg)

        # label B..E (2..5), value F..N (6..14)
        ws.merge_range(rr - 1, 1, rr - 1, 4, lab, label_gray if is_gray else label_white)
        ws.merge_range(rr - 1, 5, rr - 1, 13, val, value_gray if is_gray else value_white)

    # Bill-to block right side rows 5..9 and cols O..Z (O=15 1-based => idx 14)
    # Title row 5
    fill_range(ws, 5, 15, 9, 26, white_bg)
    ws.merge_range(4, 14, 4, 25, payload.bill_to_title, bill_title_fmt)  # row 5 => idx 4
    ws.merge_range(5, 14, 5, 25, payload.bill_to_name, bill_bold_fmt)
    ws.merge_range(6, 14, 6, 25, f"RFC: {payload.bill_to_rfc}", bill_bold_fmt)
    ws.merge_range(7, 14, 7, 25, payload.bill_to_address_1, bill_norm_fmt)
    ws.merge_range(8, 14, 8, 25, payload.bill_to_address_2, bill_norm_fmt)

    # Spacer row 10 (Excel)
    ws.set_row(9, 14)

    # Table header row 11 (Excel)
    header_row = 11
    ws.set_row(header_row - 1, 30)

    # Column groups:
    # Concepto B..N (2..14) => idx 1..13
    # Costo O..S (15..19) => idx 14..18
    # Unidades T..V (20..22) => idx 19..21
    # Subtotal W..Z (23..26) => idx 22..25
    ws.merge_range(header_row - 1, 1, header_row - 1, 13, "Concepto", th_fmt)
    ws.merge_range(header_row - 1, 14, header_row - 1, 18, "Costo unitario", th_fmt)
    ws.merge_range(header_row - 1, 19, header_row - 1, 21, "Unidades", th_fmt)
    ws.merge_range(header_row - 1, 22, header_row - 1, 25, "Subtotal", th_fmt)

    # Items start row 12 (Excel)
    start_items = 12
    items = payload.items or [ODCItem(concept="", unit_cost=0, units=0)]

    max_items = min(len(items), 18)
    wrap_chars = 55  # approximate for B..N merge width
    min_row_h = 30

    last_item_row = start_items - 1
    for idx in range(max_items):
        rr = start_items + idx  # 1-based row
        it = items[idx]

        zebra = (idx % 2 == 1)
        row_fill = gray_fill if zebra else white_bg
        fill_range(ws, rr, 2, rr, 26, row_fill)

        # Auto-height for concept (adds "air")
        needed = row_height_for_wrapped_text(it.concept, wrap_chars, base_line_height=12.5, extra_lines=1.0)
        ws.set_row(rr - 1, int(max(min_row_h, math.ceil(needed))))

        unit_cost = safe_float(it.unit_cost)
        units = safe_float(it.units)
        subtotal = safe_float(it.subtotal) if it.subtotal is not None else (unit_cost * units)

        ws.merge_range(rr - 1, 1, rr - 1, 13, it.concept, concept_g if zebra else concept_w)
        ws.merge_range(rr - 1, 14, rr - 1, 18, unit_cost, money_g if zebra else money_w)
        ws.merge_range(rr - 1, 19, rr - 1, 21, units, units_g if zebra else units_w)
        ws.merge_range(rr - 1, 22, rr - 1, 25, subtotal, money_g if zebra else money_w)

        last_item_row = rr

    # TOTAL row: 2 rows below last item
    total_row = last_item_row + 2
    ws.set_row(total_row - 1, 26)
    fill_range(ws, total_row, 2, total_row, 26, white_bg)

    # Make amount area wider to avoid ####:
    # TOTAL label at T..V (20..22 => idx 19..21)
    # Amount at W..Z (23..26 => idx 22..25)
    ws.merge_range(total_row - 1, 19, total_row - 1, 21, "TOTAL:", total_label_fmt)
    ws.merge_range(total_row - 1, 22, total_row - 1, 25, safe_float(payload.total), total_money_fmt)

    # -------------------- Print settings --------------------
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    ws.fit_to_pages(1, 0)

    # Print area (XlsxWriter)
    ws.print_area(range_a1(1, 1, total_row + 3, 26))  # A1:Z(...)

    wb.close()
    return out.getvalue()
