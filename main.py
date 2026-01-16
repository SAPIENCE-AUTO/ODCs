# main.py
# ODCs - Render + FastAPI (Excel only)
# Engine: XlsxWriter (stable merges + precise layout)
#
# IMPORTANT:
# - TOTAL comes from Monday -> it MUST be present in JSON as a number (float/int).
# - Logo comes from URL (PNG recommended). Default is the Sapience white logo you provided.
#
# requirements.txt must include:
#   fastapi==0.110.3
#   uvicorn==0.29.0
#   pydantic==2.7.4
#   XlsxWriter==3.2.0
#
# Endpoints:
#   GET  /health
#   POST /generate-odc-excel

from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import List, Optional
from datetime import datetime
from io import BytesIO
import struct
import urllib.request

import xlsxwriter
from xlsxwriter.utility import xl_range


app = FastAPI(title="Sapience ODCs (Excel)", version="1.0.1")


# -------------------- MODELOS --------------------
class ODCItem(BaseModel):
    concept: str = ""
    unit_cost: float = 0
    units: float = 0


class ODCPayload(BaseModel):
    odc_number: str
    issue_date: str  # keep string (e.g., "15 Ene 2026")
    supplier: str
    service: str
    project: str

    bill_to_name: str
    bill_to_rfc: str
    bill_to_address_1: str
    bill_to_address_2: str

    # Comes from Monday (number)
    total: float

    # Optional
    currency_symbol: str = "$"
    logo_url: Optional[str] = "https://i.postimg.cc/Pf8KhptD/logo-sapience-blanco-15-ene-26.png"
    items: List[ODCItem] = Field(default_factory=list)


# -------------------- UTILIDADES --------------------
def _png_size(img_bytes: bytes):
    # Returns (w,h) if valid PNG
    if len(img_bytes) < 24 or img_bytes[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    try:
        w = struct.unpack(">I", img_bytes[16:20])[0]
        h = struct.unpack(">I", img_bytes[20:24])[0]
        return int(w), int(h)
    except Exception:
        return None


def excel_col_width_to_pixels(w: float) -> int:
    # Standard approximation
    if w < 1.0:
        return int(w * 12 + 0.5)
    return int(w * 7 + 5 + 0.5)


def points_to_pixels(pt: float) -> float:
    return pt * 4 / 3


def r0(row_1based: int) -> int:
    return row_1based - 1


def range_a1(row1: int, col1: int, row2: int, col2: int) -> str:
    # row/col are 1-based, col1=1 => A
    return xl_range(r0(row1), col1 - 1, r0(row2), col2 - 1)


def safe_float(x) -> float:
    try:
        return float(x)
    except Exception:
        return 0.0


# -------------------- RUTAS --------------------
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


# -------------------- GENERADOR --------------------
def build_odc_excel(payload: ODCPayload) -> bytes:
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = wb.add_worksheet("ODC")

    # -------------------- PALETA + TIPOGRAFIA --------------------
    TEAL = "#0F3C4B"
    TEAL_2 = "#0F4B5A"
    LIGHT = "#EFEFEF"
    WHITE = "#FFFFFF"
    BLACK = "#111111"
    BORDER = "#6F6F6F"
    RED = "#E00000"

    # NOTE: Excel will fall back if Montserrat isn't installed on the viewer machine
    FONT = "Montserrat"

    # -------------------- GRID (B..Z width = 2.5) --------------------
    ws.set_column(0, 0, 1.2)  # A margin
    for c in range(1, 26):    # B..Z (0-based)
        ws.set_column(c, c, 2.5)

    # Semantic splits (1-based col indices)
    # B=2 ... Z=26
    SPLITS = {
        "concept": (2, 14),    # B:N
        "unit": (15, 19),      # O:S
        "units": (20, 22),     # T:V
        "subtotal": (23, 26),  # W:Z
    }

    # -------------------- FORMATOS (font -1pt vs previous) --------------------
    teal_fill = wb.add_format({"bg_color": TEAL})
    white_fill = wb.add_format({"bg_color": WHITE})
    light_fill = wb.add_format({"bg_color": LIGHT})

    # Left meta rows (zebra)
    label_g = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
        "font_color": TEAL_2, "align": "right", "valign": "vcenter",
        "bg_color": LIGHT
    })
    value_g = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "font_color": BLACK, "align": "left", "valign": "vcenter",
        "bg_color": LIGHT, "text_wrap": True
    })
    label_w = wb.add_format({
        "font_name": FONT, "font_size": 8, "bold": True,
        "font_color": TEAL_2, "align": "right", "valign": "vcenter",
        "bg_color": WHITE
    })
    value_w = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "font_color": BLACK, "align": "left", "valign": "vcenter",
        "bg_color": WHITE, "text_wrap": True
    })

    divider_g = wb.add_format({"right": 1, "right_color": BORDER, "bg_color": LIGHT})
    divider_w = wb.add_format({"right": 1, "right_color": BORDER, "bg_color": WHITE})

    # Bill to block (ensure REAL white background)
    facturar_fmt = wb.add_format({
        "font_name": FONT, "font_size": 13, "bold": True,
        "font_color": TEAL_2, "align": "center", "valign": "vcenter",
        "bg_color": WHITE
    })
    bill_bold = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "font_color": BLACK, "align": "left", "valign": "vcenter",
        "bg_color": WHITE
    })
    bill = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "font_color": BLACK, "align": "left", "valign": "vcenter",
        "bg_color": WHITE
    })

    # ODC box in banner
    odc_box = wb.add_format({
        "font_name": FONT, "font_size": 11, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": TEAL_2,
        "bg_color": LIGHT,
        "border": 1, "border_color": BORDER
    })
    odc_num = wb.add_format({
        "font_name": FONT, "font_size": 11, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": RED,
        "bg_color": LIGHT,
        "border": 1, "border_color": BORDER
    })

    # Table header
    th = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "font_color": WHITE,
        "align": "center", "valign": "vcenter",
        "bg_color": TEAL,
        "border": 1, "border_color": BORDER
    })

    # Table body (force bg_color WHITE / LIGHT)
    td_left_w = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "font_color": BLACK,
        "align": "left", "valign": "vcenter",
        "text_wrap": True,
        "bg_color": WHITE,
        "border": 1, "border_color": BORDER
    })
    td_left_g = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "font_color": BLACK,
        "align": "left", "valign": "vcenter",
        "text_wrap": True,
        "bg_color": LIGHT,
        "border": 1, "border_color": BORDER
    })
    td_c_w = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "font_color": BLACK,
        "align": "center", "valign": "vcenter",
        "bg_color": WHITE,
        "border": 1, "border_color": BORDER
    })
    td_c_g = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "font_color": BLACK,
        "align": "center", "valign": "vcenter",
        "bg_color": LIGHT,
        "border": 1, "border_color": BORDER
    })

    # Currency formats use symbol from payload (must be created after payload is available)
    money_fmt_w = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "font_color": BLACK,
        "align": "center", "valign": "vcenter",
        "bg_color": WHITE,
        "border": 1, "border_color": BORDER,
        "num_format": f'"{payload.currency_symbol}"#,##0'
    })
    money_fmt_g = wb.add_format({
        "font_name": FONT, "font_size": 8,
        "font_color": BLACK,
        "align": "center", "valign": "vcenter",
        "bg_color": LIGHT,
        "border": 1, "border_color": BORDER,
        "num_format": f'"{payload.currency_symbol}"#,##0'
    })

    # TOTAL (suelto, como mockup)
    total_label_fmt = wb.add_format({
        "font_name": FONT, "font_size": 16, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2,
        "bg_color": WHITE
    })
    total_money_fmt = wb.add_format({
        "font_name": FONT, "font_size": 18, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": RED,
        "bg_color": WHITE,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })

    # -------------------- ALTURAS BASE --------------------
    # Banner rows (2..4)
    ws.set_row(r0(1), 6)
    ws.set_row(r0(2), 38)
    ws.set_row(r0(3), 38)
    ws.set_row(r0(4), 22)

    # Meta block rows 5..9
    for rr in range(5, 10):
        ws.set_row(r0(rr), 24)

    # Spacer + header
    ws.set_row(r0(10), 14)
    ws.set_row(r0(11), 34)

    # Items (a little taller so text doesn't feel tight)
    row_h_items = 27  # tweak here

    # -------------------- BANNER (fill only, no big merges) --------------------
    ws.conditional_format(range_a1(2, 2, 4, 26), {"type": "no_blanks", "format": teal_fill})
    ws.conditional_format(range_a1(2, 2, 4, 26), {"type": "blanks", "format": teal_fill})

    # ODC mini box inside banner
    ws.merge_range(range_a1(3, 20, 3, 23), "ODC #:", odc_box)                # T3:W3
    ws.merge_range(range_a1(3, 24, 3, 26), payload.odc_number, odc_num)      # X3:Z3

    # -------------------- LOGO (smaller + centered vertically) --------------------
    if payload.logo_url:
        try:
            with urllib.request.urlopen(payload.logo_url, timeout=12) as resp:
                img = resp.read()

            wh = _png_size(img)
            # Target area for logo: B2:N4
            col_w = 2.5
            target_px_w = sum(excel_col_width_to_pixels(col_w) for _ in range(2, 15))  # B..N
            banner_h_pt = float(38 + 38 + 22)
            target_px_h = points_to_pixels(banner_h_pt)

            x_scale = y_scale = 1.0
            y_off = 0
            if wh:
                w_px, h_px = wh
                scale = min(target_px_w / w_px, target_px_h / h_px)
                scale = max(0.05, min(scale, 3.0))
                scale *= 0.86  # smaller per your note
                x_scale = y_scale = scale
                scaled_h = h_px * y_scale
                y_off = max(0, int((target_px_h - scaled_h) / 2))

            # Insert at B2
            ws.insert_image(
                r0(2), 1, "sapience_logo.png",
                {
                    "image_data": BytesIO(img),
                    "x_scale": x_scale,
                    "y_scale": y_scale,
                    "x_offset": 4,
                    "y_offset": y_off,
                    "object_position": 1,
                },
            )
        except Exception:
            # Keep going even if logo fails
            pass

    # -------------------- BLOQUE IZQUIERDO (B..O) --------------------
    # label B:E, divider in E, value F:O
    left_rows = [
        ("ODC #", payload.odc_number),
        ("FECHA:", payload.issue_date),
        ("PROVEEDOR:", payload.supplier),
        ("SERVICIO:", payload.service),
        ("PROYECTO:", payload.project),
    ]

    for i, (lab, val) in enumerate(left_rows):
        rr = 5 + i
        is_grey = (i % 2 == 0)  # rows 5,7,9 grey
        base_fill = light_fill if is_grey else white_fill
        lab_fmt = label_g if is_grey else label_w
        val_fmt = value_g if is_grey else value_w
        div_fmt = divider_g if is_grey else divider_w

        # Force background (avoid "transparent")
        ws.conditional_format(range_a1(rr, 2, rr, 15), {"type": "blanks", "format": base_fill})
        ws.conditional_format(range_a1(rr, 2, rr, 15), {"type": "no_blanks", "format": base_fill})

        ws.merge_range(range_a1(rr, 2, rr, 5), lab, lab_fmt)     # B:E
        ws.write(r0(rr), 4, "", div_fmt)                          # E divider overlay
        ws.merge_range(range_a1(rr, 6, rr, 15), val, val_fmt)     # F:O

    # -------------------- BLOQUE DERECHO (FACTURAR A) Q..Z --------------------
    # Q=17 ... Z=26
    ws.conditional_format(range_a1(5, 17, 9, 26), {"type": "blanks", "format": white_fill})
    ws.conditional_format(range_a1(5, 17, 9, 26), {"type": "no_blanks", "format": white_fill})

    ws.merge_range(range_a1(5, 17, 5, 26), "FACTURAR A:", facturar_fmt)
    ws.merge_range(range_a1(6, 17, 6, 26), payload.bill_to_name, bill_bold)
    ws.merge_range(range_a1(7, 17, 7, 26), f"RFC: {payload.bill_to_rfc}", bill_bold)
    ws.merge_range(range_a1(8, 17, 8, 26), payload.bill_to_address_1, bill)
    ws.merge_range(range_a1(9, 17, 9, 26), payload.bill_to_address_2, bill)

    # -------------------- HEADER TABLA (row 11) --------------------
    ws.merge_range(range_a1(11, SPLITS["concept"][0], 11, SPLITS["concept"][1]), "Concepto", th)
    ws.merge_range(range_a1(11, SPLITS["unit"][0], 11, SPLITS["unit"][1]), "Costo unitario", th)
    ws.merge_range(range_a1(11, SPLITS["units"][0], 11, SPLITS["units"][1]), "Unidades", th)
    ws.merge_range(range_a1(11, SPLITS["subtotal"][0], 11, SPLITS["subtotal"][1]), "Subtotal", th)

    # -------------------- ITEMS (row 12+) --------------------
    items = payload.items or []
    start_row = 12
    max_rows = 12  # safety
    last_item_row = start_row - 1

    # If no items, still render one blank row
    if not items:
        items = [ODCItem(concept="", unit_cost=0, units=0)]

    for idx, it in enumerate(items[:max_rows]):
        rr = start_row + idx
        ws.set_row(r0(rr), row_h_items)

        zebra_grey = (idx % 2 == 1)
        row_fill = light_fill if zebra_grey else white_fill
        concept_fmt = td_left_g if zebra_grey else td_left_w
        center_fmt = td_c_g if zebra_grey else td_c_w
        money_fmt = money_fmt_g if zebra_grey else money_fmt_w

        # Force row background (prevents "transparent white")
        ws.conditional_format(range_a1(rr, 2, rr, 26), {"type": "blanks", "format": row_fill})
        ws.conditional_format(range_a1(rr, 2, rr, 26), {"type": "no_blanks", "format": row_fill})

        unit_cost = safe_float(it.unit_cost)
        units = safe_float(it.units)
        subtotal = unit_cost * units

        ws.merge_range(range_a1(rr, SPLITS["concept"][0], rr, SPLITS["concept"][1]), it.concept, concept_fmt)
        ws.merge_range(range_a1(rr, SPLITS["unit"][0], rr, SPLITS["unit"][1]), unit_cost, money_fmt)
        ws.merge_range(range_a1(rr, SPLITS["units"][0], rr, SPLITS["units"][1]), units, center_fmt)
        ws.merge_range(range_a1(rr, SPLITS["subtotal"][0], rr, SPLITS["subtotal"][1]), subtotal, money_fmt)

        last_item_row = rr

    # -------------------- TOTAL (FROM PAYLOAD) --------------------
    gap_rows = 2
    total_row = last_item_row + gap_rows
    ws.set_row(r0(total_row), 34)

    # Force white background in total row
    ws.conditional_format(range_a1(total_row, 2, total_row, 26), {"type": "blanks", "format": white_fill})
    ws.conditional_format(range_a1(total_row, 2, total_row, 26), {"type": "no_blanks", "format": white_fill})

    ws.merge_range(
        range_a1(total_row, SPLITS["units"][0], total_row, SPLITS["units"][1]),
        "TOTAL:",
        total_label_fmt
    )
    ws.merge_range(
        range_a1(total_row, SPLITS["subtotal"][0], total_row, SPLITS["subtotal"][1]),
        safe_float(payload.total),
        total_money_fmt
    )

    last_row = total_row

    # -------------------- PRINT SETTINGS --------------------
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    ws.fit_to_pages(1, 0)
    ws.set_print_area(r0(1), 0, r0(last_row), 25)  # A1..Z(last_row)

    wb.close()
    return output.getvalue()
