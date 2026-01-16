# main.py
# NOTE: This file uses XlsxWriter (NOT openpyxl) because merges + layout control are far more stable.
# Make sure your requirements.txt includes:
#   xlsxwriter==3.2.0
# (FastAPI + uvicorn + pydantic already in your repo)
#
# Endpoint:
#   POST /generate-odc-excel
#
# Payload:
#   {
#     "odc_number": "RI-02497",
#     "issue_date": "17 nov 2025",
#     "supplier": "María Guadalupe Garza Sardaneta",
#     "service": "Reclutamiento",
#     "project": "ALONG",
#     "bill_to_name": "ASESORES GLOBALES CORPORATIVOS",
#     "bill_to_rfc": "AGC051117MX5",
#     "bill_to_address_1": "Peregrinos 24, Colinas del Sur,",
#     "bill_to_address_2": "Álvaro Obregón, CP. 01430, CDMX",
#     "logo_url": "https://....png",
#     "items": [{"concept":"...", "unit_cost":1400, "units":3}, ...]
#   }

from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import List, Optional
from datetime import datetime
from io import BytesIO
import urllib.request
import struct
import math

import xlsxwriter


app = FastAPI(title="Sapience ODCs (Excel)", version="1.0.0")


# -------------------- MODELOS --------------------
class ODCItem(BaseModel):
    concept: str
    unit_cost: float = 0
    units: float = 0


class ODCPayload(BaseModel):
    odc_number: str
    issue_date: str  # mantener como string (ej. "17 nov 2025")
    supplier: str
    service: str
    project: str

    bill_to_name: str
    bill_to_rfc: str
    bill_to_address_1: str
    bill_to_address_2: str

    logo_url: Optional[str] = None
    items: List[ODCItem] = Field(default_factory=list)


# -------------------- UTILIDADES --------------------
def _png_size(img_bytes: bytes):
    # Devuelve (w,h) si es PNG válido
    if len(img_bytes) < 24 or img_bytes[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    try:
        w = struct.unpack(">I", img_bytes[16:20])[0]
        h = struct.unpack(">I", img_bytes[20:24])[0]
        return int(w), int(h)
    except Exception:
        return None


def excel_col_width_to_pixels(w: float) -> int:
    # Aproximación estándar (misma idea que tu ejemplo)
    if w < 1.0:
        return int(w * 12 + 0.5)
    return int(w * 7 + 5 + 0.5)


def points_to_pixels(pt: float) -> float:
    return pt * 4 / 3


def r0(row_1based: int) -> int:
    return row_1based - 1


def range_a1(row1: int, col1: int, row2: int, col2: int) -> str:
    # row/col en 1-based; col 1=A
    return xlsxwriter.utility.xl_range(r0(row1), col1 - 1, r0(row2), col2 - 1)


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
    # -------- CONFIG BASE --------
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})
    ws = wb.add_worksheet("ODC")

    # Paleta (ajusta si quieres exactitud absoluta)
    TEAL = "#0F3C4B"
    TEAL_2 = "#0F4B5A"   # para "FACTURAR A:" y "TOTAL:"
    LIGHT = "#EFEFEF"
    WHITE = "#FFFFFF"
    BLACK = "#111111"
    BORDER = "#6F6F6F"
    RED = "#E00000"

    # Fuente (Excel hará fallback si no existe en el servidor)
    FONT = "Montserrat"

    # Grid: columnas B..Z a 2.5 (como tu diseño); A como margen
    ws.set_column(0, 0, 1.2)  # A
    for c in range(1, 26):    # B..Z
        ws.set_column(c, c, 2.5)

    # Anchos semánticos (splits)
    # B=2, ..., Z=26 (1-based col indices)
    SPLITS = {
        "concept": (2, 14),   # B:N
        "unit": (15, 19),     # O:S
        "units": (20, 22),    # T:V
        "subtotal": (23, 26), # W:Z
    }

    # -------- FORMATOS --------
    # Banner / fondos
    teal_fill = wb.add_format({"bg_color": TEAL})
    white_fill = wb.add_format({"bg_color": WHITE})
    light_fill = wb.add_format({"bg_color": LIGHT})

    # Labels izquierda
    label_fmt = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "font_color": TEAL_2,
        "align": "right", "valign": "vcenter",
        "bg_color": LIGHT
    })
    value_fmt = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "font_color": BLACK,
        "align": "left", "valign": "vcenter",
        "bg_color": LIGHT
    })
    # Variante blanco (para filas sin gris)
    label_w = wb.add_format({
        "font_name": FONT, "font_size": 9, "bold": True,
        "font_color": TEAL_2,
        "align": "right", "valign": "vcenter",
        "bg_color": WHITE
    })
    value_w = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "font_color": BLACK,
        "align": "left", "valign": "vcenter",
        "bg_color": WHITE
    })

    # Separador vertical (col E)
    divider_fmt = wb.add_format({"right": 1, "right_color": BORDER, "bg_color": LIGHT})
    divider_w = wb.add_format({"right": 1, "right_color": BORDER, "bg_color": WHITE})

    # Facturar a (derecha)
    facturar_fmt = wb.add_format({
        "font_name": FONT, "font_size": 18, "bold": True,
        "font_color": TEAL_2,
        "align": "center", "valign": "vcenter",
        "bg_color": WHITE
    })
    bill_bold = wb.add_format({
        "font_name": FONT, "font_size": 11, "bold": True,
        "font_color": BLACK,
        "align": "left", "valign": "vcenter",
        "bg_color": WHITE
    })
    bill = wb.add_format({
        "font_name": FONT, "font_size": 11,
        "font_color": BLACK,
        "align": "left", "valign": "vcenter",
        "bg_color": WHITE
    })

    # Caja ODC en banner
    odc_box = wb.add_format({
        "font_name": FONT, "font_size": 14, "bold": True,
        "align": "center", "valign": "vcenter",
        "bg_color": LIGHT,
        "border": 1, "border_color": BORDER
    })
    odc_num = wb.add_format({
        "font_name": FONT, "font_size": 14, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": RED,
        "bg_color": LIGHT,
        "border": 1, "border_color": BORDER
    })

    # Tabla header
    th = wb.add_format({
        "font_name": FONT, "font_size": 12, "bold": True,
        "font_color": WHITE,
        "align": "center", "valign": "vcenter",
        "bg_color": TEAL,
        "border": 1, "border_color": BORDER
    })

    # Tabla body
    td_left_w = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "font_color": BLACK,
        "align": "left", "valign": "vcenter",
        "text_wrap": True,
        "bg_color": WHITE,
        "border": 1, "border_color": BORDER
    })
    td_left_g = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "font_color": BLACK,
        "align": "left", "valign": "vcenter",
        "text_wrap": True,
        "bg_color": LIGHT,
        "border": 1, "border_color": BORDER
    })
    td_c_w = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "font_color": BLACK,
        "align": "center", "valign": "vcenter",
        "bg_color": WHITE,
        "border": 1, "border_color": BORDER
    })
    td_c_g = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "font_color": BLACK,
        "align": "center", "valign": "vcenter",
        "bg_color": LIGHT,
        "border": 1, "border_color": BORDER
    })
    money_w = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "font_color": BLACK,
        "align": "center", "valign": "vcenter",
        "bg_color": WHITE,
        "border": 1, "border_color": BORDER,
        "num_format": '"$"#,##0'
    })
    money_g = wb.add_format({
        "font_name": FONT, "font_size": 9,
        "font_color": BLACK,
        "align": "center", "valign": "vcenter",
        "bg_color": LIGHT,
        "border": 1, "border_color": BORDER,
        "num_format": '"$"#,##0'
    })

    # TOTAL (estilo suelto)
    total_label_fmt = wb.add_format({
        "font_name": FONT, "font_size": 18, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2,
        "bg_color": WHITE
    })
    total_money_fmt = wb.add_format({
        "font_name": FONT, "font_size": 22, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": RED,
        "bg_color": WHITE,
        "num_format": '#,##0.00'
    })

    # -------- ROW HEIGHTS (1-based) --------
    # Banner (3 filas)
    ws.set_row(r0(1), 6)
    ws.set_row(r0(2), 38)
    ws.set_row(r0(3), 38)
    ws.set_row(r0(4), 22)

    # Info block
    for rr in range(5, 10):
        ws.set_row(r0(rr), 24)

    # Espacio
    ws.set_row(r0(10), 14)

    # Header tabla
    ws.set_row(r0(11), 34)

    # -------- BANNER (sin merge grande; solo fill de rango) --------
    ws.conditional_format(range_a1(2, 2, 4, 26), {"type": "no_blanks", "format": teal_fill})
    ws.conditional_format(range_a1(2, 2, 4, 26), {"type": "blanks", "format": teal_fill})

    # Caja ODC dentro del banner (merge chico)
    # T3:W3 y X3:Z3 (fila 3)
    ws.merge_range(range_a1(3, 20, 3, 23), "ODC #:", odc_box)          # T:W
    ws.merge_range(range_a1(3, 24, 3, 26), payload.odc_number, odc_num) # X:Z

    # Logo (insertado; centrado verticalmente dentro de filas 2-4)
    if payload.logo_url:
        try:
            with urllib.request.urlopen(payload.logo_url, timeout=12) as resp:
                img = resp.read()
            wh = _png_size(img)

            # área objetivo para el logo: B2:N4 (B..N)
            col_w = 2.5
            target_px_w = 0
            for _ in range(2, 15):  # B..N (2..14)
                target_px_w += excel_col_width_to_pixels(col_w)
            banner_h_pt = 38 + 38 + 22
            target_px_h = points_to_pixels(banner_h_pt)

            x_scale = y_scale = 1.0
            y_off = 0
            if wh:
                w_px, h_px = wh
                scale = min(target_px_w / w_px, target_px_h / h_px)
                scale = max(0.05, min(scale, 3.0))
                # un poco más chico (como pediste)
                scale *= 0.88
                x_scale = y_scale = scale
                scaled_h = h_px * y_scale
                y_off = max(0, int((target_px_h - scaled_h) / 2))

            ws.insert_image(
                r0(2), 1, "sapience_logo.png",  # B2
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
            pass

    # -------- BLOQUE IZQUIERDO (ODC/FECHA/PROVEEDOR/SERVICIO/PROYECTO) --------
    # Cols: label B:E, value F:O, divider en E (borde derecho)
    left_rows = [
        ("ODC #", payload.odc_number),
        ("FECHA:", payload.issue_date),
        ("PROVEEDOR:", payload.supplier),
        ("SERVICIO:", payload.service),
        ("PROYECTO:", payload.project),
    ]

    for i, (lab, val) in enumerate(left_rows):
        rr = 5 + i
        is_grey = (i % 2 == 0)  # 5,7,9 gris como en mockup (ajusta si quieres)
        lab_fmt = label_fmt if is_grey else label_w
        val_fmt = value_fmt if is_grey else value_w
        div_fmt = divider_fmt if is_grey else divider_w

        # Relleno base (para que NO sea transparente)
        ws.conditional_format(range_a1(rr, 2, rr, 15), {"type": "blanks", "format": light_fill if is_grey else white_fill})
        ws.conditional_format(range_a1(rr, 2, rr, 15), {"type": "no_blanks", "format": light_fill if is_grey else white_fill})

        ws.merge_range(range_a1(rr, 2, rr, 5), lab, lab_fmt)   # B:E
        # divisor en E (lo pintamos encima)
        ws.write(r0(rr), 4, "", div_fmt)  # E (0-based col 4)

        ws.merge_range(range_a1(rr, 6, rr, 15), val, val_fmt)  # F:O

    # -------- BLOQUE DERECHO (FACTURAR A) --------
    # zona derecha: Q:Z (17..26)
    ws.merge_range(range_a1(5, 17, 5, 26), "FACTURAR A:", facturar_fmt)
    ws.merge_range(range_a1(6, 17, 6, 26), payload.bill_to_name, bill_bold)
    ws.merge_range(range_a1(7, 17, 7, 26), f"RFC: {payload.bill_to_rfc}", bill_bold)
    ws.merge_range(range_a1(8, 17, 8, 26), payload.bill_to_address_1, bill)
    ws.merge_range(range_a1(9, 17, 9, 26), payload.bill_to_address_2, bill)

    # Asegura blancos reales (Q:Z filas 5..9)
    ws.conditional_format(range_a1(5, 17, 9, 26), {"type": "blanks", "format": white_fill})
    ws.conditional_format(range_a1(5, 17, 9, 26), {"type": "no_blanks", "format": white_fill})

    # -------- HEADER TABLA --------
    ws.merge_range(range_a1(11, *SPLITS["concept"]), "Concepto", th)
    ws.merge_range(range_a1(11, *SPLITS["unit"]), "Costo unitario", th)
    ws.merge_range(range_a1(11, *SPLITS["units"]), "Unidades", th)
    ws.merge_range(range_a1(11, *SPLITS["subtotal"]), "Subtotal", th)

    # -------- ITEMS --------
    items = payload.items or []
    start_row = 12
    row_h_items = 26  # “poquito más grande” para que respire
    max_rows = 12     # por si te mandan un chorro

    grand_total = 0.0
    last_item_row = start_row - 1

    for idx, it in enumerate(items[:max_rows]):
        rr = start_row + idx
        ws.set_row(r0(rr), row_h_items)

        zebra_grey = (idx % 2 == 1)
        concept_fmt = td_left_g if zebra_grey else td_left_w
        center_fmt = td_c_g if zebra_grey else td_c_w
        money_fmt = money_g if zebra_grey else money_w

        # fill del renglón completo (para evitar transparencia)
        ws.conditional_format(range_a1(rr, 2, rr, 26), {"type": "blanks", "format": light_fill if zebra_grey else white_fill})
        ws.conditional_format(range_a1(rr, 2, rr, 26), {"type": "no_blanks", "format": light_fill if zebra_grey else white_fill})

        unit_cost = safe_float(it.unit_cost)
        units = safe_float(it.units)
        subtotal = unit_cost * units
        grand_total += subtotal

        ws.merge_range(range_a1(rr, *SPLITS["concept"]), it.concept, concept_fmt)
        ws.merge_range(range_a1(rr, *SPLITS["unit"]), unit_cost, money_fmt)
        ws.merge_range(range_a1(rr, *SPLITS["units"]), units, center_fmt)
        ws.merge_range(range_a1(rr, *SPLITS["subtotal"]), subtotal, money_fmt)

        last_item_row = rr

    # Si no hay items, al menos un renglón vacío
    if last_item_row < start_row:
        rr = start_row
        ws.set_row(r0(rr), row_h_items)
        ws.conditional_format(range_a1(rr, 2, rr, 26), {"type": "blanks", "format": white_fill})
        ws.conditional_format(range_a1(rr, 2, rr, 26), {"type": "no_blanks", "format": white_fill})
        ws.merge_range(range_a1(rr, *SPLITS["concept"]), "", td_left_w)
        ws.merge_range(range_a1(rr, *SPLITS["unit"]), "", money_w)
        ws.merge_range(range_a1(rr, *SPLITS["units"]), "", td_c_w)
        ws.merge_range(range_a1(rr, *SPLITS["subtotal"]), "", money_w)
        last_item_row = rr

    # -------- TOTAL (como tu mockup: suelto, grande, rojo) --------
    gap_rows = 2
    total_row = last_item_row + gap_rows
    ws.set_row(r0(total_row), 34)

    # Fondo blanco real en la zona del total
    ws.conditional_format(range_a1(total_row, 2, total_row, 26), {"type": "blanks", "format": white_fill})
    ws.conditional_format(range_a1(total_row, 2, total_row, 26), {"type": "no_blanks", "format": white_fill})

    ws.merge_range(range_a1(total_row, *SPLITS["units"]), "TOTAL:", total_label_fmt)
    ws.merge_range(range_a1(total_row, *SPLITS["subtotal"]), grand_total, total_money_fmt)

    last_row = total_row

    # -------- PRINT SETTINGS --------
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    ws.fit_to_pages(1, 0)
    ws.set_print_area(r0(1), 0, r0(last_row), 25)  # A1..Z(last_row)

    wb.close()
    return output.getvalue()

