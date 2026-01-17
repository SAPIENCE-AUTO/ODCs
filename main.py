# main.py
# Sapience ODCs — Render + FastAPI — Excel only (XlsxWriter)
# v1.1.0
#
# Key rules (requested):
# - Banner azul: 3 filas de 16 pts (48 total)
# - Logo más chico y bien ubicado (no se sale del banner)
# - ODC box arriba derecha: menos alto (1 fila)
# - Meta izquierda: blanco/gris alternado empezando en blanco
# - Línea divisoria gris vertical junto a labels (ODC/FECHA/PROVEEDOR/etc.)
# - FACTURAR A y datos con fondo blanco explícito
# - Headers tabla: 9 pt
# - Celdas tabla: 8 pt y centrado vertical
# - Abajo: SUMA / ANTICIPO / TOTAL con números centrados horizontalmente
#
# Notes:
# - Pydantic acepta: sum_amount, advance_amount, total_due
# - Compatibilidad: si solo viene total, se usa como total_due (y suma/anticipo quedan en 0)

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

    # Bottom summary
    sum_amount: float = 0
    advance_amount: float = 0
    total_due: float = 0

    # Backward compatibility (older payloads)
    total: Optional[float] = None  # if provided and total_due==0, we'll use it

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
    base_line_height: float = 11.5,
    extra_lines: float = 0.9,
) -> float:
    """
    Simulate padding via row height.
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

    # -------- Data normalization --------
    # If old payloads only send `total`, map it to total_due (unless already set)
    if safe_float(payload.total_due) == 0 and payload.total is not None:
        payload.total_due = safe_float(payload.total)

    # If sum_amount not provided but we can infer from items:
    if safe_float(payload.sum_amount) == 0 and payload.items:
        payload.sum_amount = sum(
            safe_float(it.subtotal) if it.subtotal is not None else safe_float(it.unit_cost) * safe_float(it.units)
            for it in payload.items
        )

    # If total_due not provided but sum and advance are, infer:
    if safe_float(payload.total_due) == 0 and safe_float(payload.sum_amount) != 0:
        payload.total_due = safe_float(payload.sum_amount) - safe_float(payload.advance_amount)

    # -------- Palette / Font --------
    FONT = "Montserrat"
    TEAL = "#0F3D4C"
    TEAL_2 = "#0E4A5A"
    WHITE = "#FFFFFF"
    LIGHT_GRAY = "#EFEFEF"
    GRID = "#7C7C7C"
    RED = "#E10600"
    BLACK = "#111111"
    MID_GRAY = "#C9C9C9"

    # -------- Grid columns (uniform 2.5) --------
    for c in range(0, 26):  # A..Z
        ws.set_column(c, c, 2.5)

    # -------- Explicit white canvas --------
    white_bg = wb.add_format({"bg_color": WHITE})
    fill_range(ws, 1, 1, 200, 26, white_bg)

    banner_fill = wb.add_format({"bg_color": TEAL})
    gray_fill = wb.add_format({"bg_color": LIGHT_GRAY})

    # -------- Formats --------
    # Top-right ODC box (smaller height: 1 row) — label 9 pt + value 9 pt
    odc_box_lbl = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })
    odc_box_val = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": RED, "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })

    # Left meta labels/values (8 pt labels / 8 pt values)
    label_white = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE,
        "right": 1, "right_color": MID_GRAY,  # vertical divider line
    })
    value_white = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE,
        "text_wrap": True,
    })
    label_gray = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": LIGHT_GRAY,
        "right": 1, "right_color": MID_GRAY,  # vertical divider line
    })
    value_gray = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": LIGHT_GRAY,
        "text_wrap": True,
    })

    # Bill-to block (white bg explicit)
    bill_title_fmt = wb.add_format({
        "font_name": FONT, "font_size": 10, "bold": True,
        "align": "left", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE,
    })
    bill_bold_fmt = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE,
    })
    bill_norm_fmt = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE,
    })

    # Table header (9 pt)
    th_fmt = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": WHITE, "bg_color": TEAL_2,
        "border": 1, "border_color": GRID,
    })

    # Table cells (8 pt, vertical centered)
    concept_w = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK,
        "text_wrap": True,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })
    concept_g = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK,
        "text_wrap": True,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
    })

    money_w = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })
    money_g = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })

    units_w = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
        "num_format": "0"
    })
    units_g = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
        "num_format": "0"
    })

    # Bottom summary formats (labels teal, values centered horizontally)
    sum_label_fmt = wb.add_format({
        "font_name": FONT, "font_size": 10, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE,
    })
    sum_value_fmt = wb.add_format({
        "font_name": FONT, "font_size": 10,
        "align": "center", "valign": "vcenter",  # centered horizontally
        "font_color": "#6F6F6F", "bg_color": WHITE,
        "num_format": f'"{payload.currency_symbol}"#,##0.00',
    })

    adv_label_fmt = wb.add_format({
        "font_name": FONT, "font_size": 10, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE,
    })
    adv_value_fmt = wb.add_format({
        "font_name": FONT, "font_size": 10, "bold": True,
        "align": "center", "valign": "vcenter",  # centered horizontally
        "font_color": RED, "bg_color": WHITE,
        "num_format": f'"{payload.currency_symbol}"#,##0.00',
    })

    total_label_fmt = wb.add_format({
        "font_name": FONT, "font_size": 10, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE,
    })
    total_value_fmt = wb.add_format({
        "font_name": FONT, "font_size": 11, "bold": True,  # prominent
        "align": "center", "valign": "vcenter",  # centered horizontally
        "font_color": "#3A3A3A", "bg_color": WHITE,
        "num_format": f'"{payload.currency_symbol}"#,##0.00',
    })

    # Thin gray line format (for separators under summary rows)
    hline_fmt = wb.add_format({"bottom": 1, "bottom_color": MID_GRAY, "bg_color": WHITE})

    # -------- Layout --------
    # Banner rows: Excel rows 1..3 (each 16 pts)
    ws.set_row(0, 16)
    ws.set_row(1, 16)
    ws.set_row(2, 16)

    # Paint banner area across B..Z in rows 1..3
    fill_range(ws, 1, 2, 3, 26, banner_fill)  # B1:Z3

    # Insert logo (smaller, anchored within banner, no overflow)
    # Put it around B1 with conservative scaling and offsets.
    if payload.logo_url:
        try:
            resp = requests.get(payload.logo_url, timeout=15)
            resp.raise_for_status()
            img = resp.content
            wh = _png_size(img)

            # default conservative
            x_scale = y_scale = 0.18
            x_off = 6
            y_off = 2

            if wh:
                w_px, h_px = wh
                # Banner height ~ 48 pts -> approx 64 px (varies), we cap by 52 px to be safe
                target_h_px = 52
                target_w_px = 260  # keep modest width
                scale = min(target_w_px / max(1, w_px), target_h_px / max(1, h_px))
                # hard caps to avoid overflow
                scale = max(0.10, min(scale, 0.20))
                x_scale = y_scale = scale

                # vertical centering inside 3 rows
                scaled_h = h_px * y_scale
                y_off = max(0, int((target_h_px - scaled_h) / 2))

            ws.insert_image(
                0,      # row 1 (0-based)
                1,      # col B (0-based)
                "logo.png",
                {
                    "image_data": BytesIO(img),
                    "x_scale": x_scale,
                    "y_scale": y_scale,
                    "x_offset": x_off,
                    "y_offset": y_off,
                    "object_position": 1,
                },
            )
        except Exception:
            pass

    # ODC box (top-right) — 1 row height (row 1), cols U..Z (just tighter visually)
    # We'll use row 1 (Excel) => row index 0
    # Label at U..W, Value at X..Z
    ws.merge_range(0, 20, 0, 22, "ODC #:", odc_box_lbl)                 # U..W
    ws.merge_range(0, 23, 0, 25, payload.odc_number, odc_box_val)       # X..Z

    # Meta + Bill-to block starts at Excel row 4
    meta_start_row = 4  # Excel row 4
    meta_rows = [
        ("ODC #", payload.odc_number),
        ("FECHA:", payload.date_str),
        ("PROVEEDOR:", payload.provider),
        ("SERVICIO:", payload.service),
        ("PROYECTO:", payload.project),
    ]

    # Left block cols B..N (B=2..N=14 => 1..13 0-based)
    # Alternating starts WHITE then GRAY...
    for i, (lab, val) in enumerate(meta_rows):
        rr = meta_start_row + i  # Excel row number
        ws.set_row(rr - 1, 18)

        is_gray = (i % 2 == 1)  # start with white, then gray
        fill_range(ws, rr, 2, rr, 14, gray_fill if is_gray else white_bg)

        ws.merge_range(rr - 1, 1, rr - 1, 4, lab, label_gray if is_gray else label_white)  # B..E
        ws.merge_range(rr - 1, 5, rr - 1, 13, val, value_gray if is_gray else value_white) # F..N

    # Bill-to block right side rows 4..8 and cols O..Z (O=15..Z=26)
    # Ensure white fill explicit
    fill_range(ws, meta_start_row, 15, meta_start_row + 4, 26, white_bg)

    ws.merge_range(meta_start_row - 1, 14, meta_start_row - 1, 25, payload.bill_to_title, bill_title_fmt)   # O..Z
    ws.merge_range(meta_start_row,     14, meta_start_row,     25, payload.bill_to_name,  bill_bold_fmt)
    ws.merge_range(meta_start_row + 1, 14, meta_start_row + 1, 25, f"RFC: {payload.bill_to_rfc}", bill_bold_fmt)
    ws.merge_range(meta_start_row + 2, 14, meta_start_row + 2, 25, payload.bill_to_address_1, bill_norm_fmt)
    ws.merge_range(meta_start_row + 3, 14, meta_start_row + 3, 25, payload.bill_to_address_2, bill_norm_fmt)

    # Spacer row (after meta blocks)
    spacer_row = meta_start_row + 5  # Excel row 9
    ws.set_row(spacer_row - 1, 12)

    # Table header row
    header_row = spacer_row + 1  # Excel row 10
    ws.set_row(header_row - 1, 22)

    # Column groups:
    # Concepto B..N (2..14) => idx 1..13
    # Costo O..S (15..19) => idx 14..18
    # Unidades T..V (20..22) => idx 19..21
    # Subtotal W..Z (23..26) => idx 22..25
    ws.merge_range(header_row - 1, 1,  header_row - 1, 13, "Concepto",      th_fmt)
    ws.merge_range(header_row - 1, 14, header_row - 1, 18, "Costo unitario", th_fmt)
    ws.merge_range(header_row - 1, 19, header_row - 1, 21, "Unidades",      th_fmt)
    ws.merge_range(header_row - 1, 22, header_row - 1, 25, "Subtotal",      th_fmt)

    # Items start row
    start_items = header_row + 1  # Excel row 11
    items = payload.items or [ODCItem(concept="", unit_cost=0, units=0)]

    max_items = min(len(items), 18)
    wrap_chars = 62
    min_row_h = 22

    last_item_row = start_items - 1
    for idx, it in enumerate(items[:max_items]):
        rr = start_items + idx
        zebra = (idx % 2 == 1)
        row_fill = gray_fill if zebra else white_bg
        fill_range(ws, rr, 2, rr, 26, row_fill)

        needed = row_height_for_wrapped_text(it.concept, wrap_chars, base_line_height=11.5, extra_lines=0.9)
        ws.set_row(rr - 1, int(max(min_row_h, math.ceil(needed))))

        unit_cost = safe_float(it.unit_cost)
        units = safe_float(it.units)
        subtotal = safe_float(it.subtotal) if it.subtotal is not None else (unit_cost * units)

        ws.merge_range(rr - 1, 1,  rr - 1, 13, it.concept, concept_g if zebra else concept_w)
        ws.merge_range(rr - 1, 14, rr - 1, 18, unit_cost,  money_g if zebra else money_w)
        ws.merge_range(rr - 1, 19, rr - 1, 21, units,      units_g if zebra else units_w)
        ws.merge_range(rr - 1, 22, rr - 1, 25, subtotal,   money_g if zebra else money_w)

        last_item_row = rr

    # Summary block (SUMA / ANTICIPO / TOTAL)
    summary_start = last_item_row + 2  # one blank row then summary

    # Give some air
    ws.set_row(summary_start - 1, 10)

    # We'll use rows: summary_start+1, +2, +3 for the three lines
    sum_row = summary_start + 1
    adv_row = summary_start + 2
    tot_row = summary_start + 3

    for r in [sum_row, adv_row, tot_row]:
        ws.set_row(r - 1, 22)
        fill_range(ws, r, 2, r, 26, white_bg)

    # Place summary at right side similar to reference:
    # Labels in V..X (22..24 1-based => idx 21..23)
    # Values in Y..Z (25..26 1-based => idx 24..25)
    # And a subtle vertical divider line between label/value (we use borders on the value cell)
    # We'll also add a thin bottom line under each row using hline_fmt over the area.
    # Underlines (light gray) across V..Z
    fill_range(ws, sum_row, 22, sum_row, 26, hline_fmt)
    fill_range(ws, adv_row, 22, adv_row, 26, hline_fmt)
    fill_range(ws, tot_row, 22, tot_row, 26, white_bg)

    # Label merges: V..X (idx 21..23)
    ws.merge_range(sum_row - 1, 21, sum_row - 1, 23, "SUMA:",     sum_label_fmt)
    ws.merge_range(adv_row - 1, 21, adv_row - 1, 23, "ANTICIPO:", adv_label_fmt)
    ws.merge_range(tot_row - 1, 21, tot_row - 1, 23, "TOTAL:",    total_label_fmt)

    # Value merges: Y..Z (idx 24..25)
    ws.merge_range(sum_row - 1, 24, sum_row - 1, 25, safe_float(payload.sum_amount),     sum_value_fmt)
    ws.merge_range(adv_row - 1, 24, adv_row - 1, 25, safe_float(payload.advance_amount), adv_value_fmt)
    ws.merge_range(tot_row - 1, 24, tot_row - 1, 25, safe_float(payload.total_due),      total_value_fmt)

    # -------------------- Print settings --------------------
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    ws.fit_to_pages(1, 0)

    # Print area: until a bit after total
    ws.print_area(range_a1(1, 1, tot_row + 2, 26))  # A1:Z...

    wb.close()
    return out.getvalue()
