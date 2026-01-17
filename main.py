# main.py
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

app = FastAPI(title="Sapience ODCs (Excel)", version="1.0.0")


# -----------------------------
# Models
# -----------------------------
class ODCItem(BaseModel):
    concept: str
    unit_cost: float
    units: int
    subtotal: float


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

    total: float  # viene del Monday (número)
    currency_symbol: str = "$"

    logo_url: str = "https://i.postimg.cc/Pf8KhptD/logo-sapience-blanco-15-ene-26.png"


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
    """1-based row/col -> Excel A1 notation e.g. A1:Z20"""
    return xlsxwriter.utility.xl_range(row1 - 1, col1 - 1, row2 - 1, col2 - 1)


def _row_height_for_wrapped_text(
    text: str,
    wrap_width_chars: int,
    base_line_height: float = 12.5,
    extra_lines: float = 1.0,
) -> float:
    """
    XlsxWriter no tiene padding real. Para que el texto no se sienta apretado:
    - estimamos líneas por wrap
    - damos "extra_lines" de aire vertical
    """
    if not text:
        return base_line_height * (1 + extra_lines)

    paragraphs = str(text).splitlines() or [""]
    total_lines = 0
    for p in paragraphs:
        wrapped = (
            textwrap.wrap(
                p,
                width=max(1, wrap_width_chars),
                break_long_words=True,
                replace_whitespace=False,
            )
            or [""]
        )
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
    # (Montserrat solo se verá si está instalada en la máquina donde se abre el XLSX)
    FONT = "Montserrat"
    TEAL = "#0F3D4C"
    TEAL_2 = "#0E4A5A"
    WHITE = "#FFFFFF"
    LIGHT_GRAY = "#EFEFEF"
    GRID = "#7C7C7C"
    RED = "#E10600"

    # -------- Global grid canvas --------
    # Usamos un grid fijo de 2.5 para "dibujar" el layout
    # Columnas: A..Z (26 columnas)
    for c in range(0, 26):
        ws.set_column(c, c, 2.5)

    # Fondo blanco explícito (evitar "sin fondo")
    white_bg = wb.add_format({"bg_color": WHITE})
    for r in range(0, 120):
        ws.set_row(r, None, white_bg)

    # -------- Formats --------
    # Banner teal
    banner_fmt = wb.add_format({"bg_color": TEAL, "border": 0})

    # ODC box (top-right)
    odc_box_lbl = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 14,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "font_color": TEAL_2,
            "bg_color": WHITE,
            "border": 1,
            "border_color": GRID,
        }
    )
    odc_box_val = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 14,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "font_color": RED,
            "bg_color": WHITE,
            "border": 1,
            "border_color": GRID,
        }
    )

    # Left labels (ODC/FECHA/PROVEEDOR/SERVICIO/PROYECTO)
    left_label_fmt = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 11,
            "bold": True,
            "align": "right",
            "valign": "vcenter",
            "font_color": TEAL_2,
            "bg_color": LIGHT_GRAY,
        }
    )
    left_value_fmt = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 11,
            "bold": False,
            "align": "left",
            "valign": "vcenter",
            "bg_color": LIGHT_GRAY,
        }
    )

    # Bill-to block
    # ✅ “FACTURAR A:” alineado a la izquierda
    bill_title_fmt = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 18,
            "bold": True,
            "align": "left",  # <-- ajuste clave
            "valign": "vcenter",
            "font_color": TEAL_2,
            "bg_color": WHITE,
        }
    )
    bill_bold_fmt = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 12,
            "bold": True,
            "align": "left",
            "valign": "vcenter",
            "font_color": "#000000",
            "bg_color": WHITE,
        }
    )
    bill_norm_fmt = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 12,
            "bold": False,
            "align": "left",
            "valign": "vcenter",
            "font_color": "#000000",
            "bg_color": WHITE,
        }
    )

    # Table header
    th_fmt = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 12,
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "font_color": WHITE,
            "bg_color": TEAL_2,
            "border": 1,
            "border_color": GRID,
        }
    )

    # Table cells
    concept_fmt = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 10,
            "align": "left",
            "valign": "top",
            "text_wrap": True,
            "bg_color": WHITE,
            "border": 1,
            "border_color": GRID,
        }
    )
    money_fmt = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 11,
            "align": "center",
            "valign": "vcenter",
            "bg_color": WHITE,
            "border": 1,
            "border_color": GRID,
            "num_format": f'"{payload.currency_symbol}"#,##0',
        }
    )
    units_fmt = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 11,
            "align": "center",
            "valign": "vcenter",
            "bg_color": WHITE,
            "border": 1,
            "border_color": GRID,
        }
    )
    subtotal_fmt = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 11,
            "align": "center",
            "valign": "vcenter",
            "bg_color": WHITE,
            "border": 1,
            "border_color": GRID,
            "num_format": f'"{payload.currency_symbol}"#,##0',
        }
    )

    # Zebra alt row (solo para concepto; el resto puede quedar blanco)
    concept_fmt_z = wb.add_format({**concept_fmt.properties, "bg_color": LIGHT_GRAY})
    money_fmt_z = wb.add_format({**money_fmt.properties, "bg_color": LIGHT_GRAY})
    units_fmt_z = wb.add_format({**units_fmt.properties, "bg_color": LIGHT_GRAY})
    subtotal_fmt_z = wb.add_format({**subtotal_fmt.properties, "bg_color": LIGHT_GRAY})

    # Total formats
    total_label_fmt = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 11,  # ✅ ajuste: 11
            "bold": True,
            "align": "right",
            "valign": "vcenter",
            "font_color": TEAL_2,
            "bg_color": WHITE,
        }
    )
    total_money_fmt = wb.add_format(
        {
            "font_name": FONT,
            "font_size": 11,  # ✅ ajuste: 11
            "bold": True,
            "align": "right",
            "valign": "vcenter",
            "font_color": RED,
            "bg_color": WHITE,
            "num_format": f'"{payload.currency_symbol}"#,##0.00',
        }
    )

    # -------- Layout coordinates (0-based) --------
    # Canvas: columnas A..Z (0..25)
    # Dejamos un margen: arrancar en B (col=1)
    LEFT = 1
    RIGHT = 25

    # Banner rows
    banner_r1, banner_r2 = 0, 3  # 4 filas altas
    for r in range(banner_r1, banner_r2 + 1):
        for c in range(LEFT, RIGHT + 1):
            ws.write_blank(r, c, None, banner_fmt)

    # Ajuste de altura del banner (logo más chico y centrado)
    ws.set_row(0, 10, white_bg)         # primera fila más angosta
    ws.set_row(1, 38, white_bg)
    ws.set_row(2, 38, white_bg)
    ws.set_row(3, 26, white_bg)

    # Insert logo (más chico)
    try:
        r = requests.get(payload.logo_url, timeout=15)
        img = r.content
        wh = _png_size(img)

        # Lo insertamos en B2 (row=1,col=1), centrado visualmente en el banner
        # Escala target (más chico que antes)
        x_scale = y_scale = 0.22
        if wh:
            w_px, h_px = wh
            # cap suave (logo más chico)
            scale = min(520 / max(1, w_px), 110 / max(1, h_px))
            scale = max(0.10, min(scale, 0.30))
            x_scale = y_scale = scale

        ws.insert_image(
            1,
            LEFT,
            "logo.png",
            {
                "image_data": BytesIO(img),
                "x_scale": x_scale,
                "y_scale": y_scale,
                "x_offset": 6,
                "y_offset": 6,
                "object_position": 1,
            },
        )
    except Exception:
        pass

    # Top-right ODC box (con fondo blanco, no “transparente”)
    # Ocupa aprox columnas S..Z (18..25)
    odc_box_row1, odc_box_row2 = 1, 2
    odc_lbl_c1, odc_lbl_c2 = 19, 22
    odc_val_c1, odc_val_c2 = 23, 25

    ws.merge_range(odc_box_row1, odc_lbl_c1, odc_box_row2, odc_lbl_c2, "ODC #:", odc_box_lbl)
    ws.merge_range(odc_box_row1, odc_val_c1, odc_box_row2, odc_val_c2, payload.odc_number, odc_box_val)

    # Left meta block (rows 4..8)
    meta_start = 4
    meta_rows = [
        ("ODC #", payload.odc_number),
        ("FECHA:", payload.date_str),
        ("PROVEEDOR:", payload.provider),
        ("SERVICIO:", payload.service),
        ("PROYECTO:", payload.project),
    ]

    # Bloque gris claro a la izquierda (B..N aprox)
    left_c1, left_c2 = LEFT, 13  # B..N
    label_c1, label_c2 = left_c1, 4  # B..E
    value_c1, value_c2 = 5, left_c2  # F..N

    for i, (lab, val) in enumerate(meta_rows):
        rr = meta_start + i
        ws.set_row(rr, 22, white_bg)
        ws.merge_range(rr, label_c1, rr, label_c2, lab, left_label_fmt)
        ws.merge_range(rr, value_c1, rr, value_c2, val, left_value_fmt)

    # Bill-to block (right side)
    # ✅ FACTURAR A alineado a la izquierda
    bill_c1, bill_c2 = 14, RIGHT  # O..Z
    bill_r = 4
    ws.merge_range(bill_r, bill_c1, bill_r, bill_c2, payload.bill_to_title, bill_title_fmt)

    ws.merge_range(bill_r + 1, bill_c1, bill_r + 1, bill_c2, payload.bill_to_name, bill_bold_fmt)
    ws.merge_range(bill_r + 2, bill_c1, bill_r + 2, bill_c2, f"RFC: {payload.bill_to_rfc}", bill_bold_fmt)
    ws.merge_range(bill_r + 3, bill_c1, bill_r + 3, bill_c2, payload.bill_to_address_1, bill_norm_fmt)
    ws.merge_range(bill_r + 4, bill_c1, bill_r + 4, bill_c2, payload.bill_to_address_2, bill_norm_fmt)

    # Ensure rows under bill-to are white (explicit)
    for rr in range(bill_r, bill_r + 5):
        for cc in range(bill_c1, bill_c2 + 1):
            # si no hay valor, pinto blanco
            pass

    # Table header row
    header_row = 10
    ws.set_row(header_row, 30, white_bg)

    # Column groups (Concepto | Costo unitario | Unidades | Subtotal)
    # B..N, O..S, T..V, W..Z
    concept_c1, concept_c2 = LEFT, 13       # B..N
    cost_c1, cost_c2 = 14, 18               # O..S
    units_c1, units_c2 = 19, 21             # T..V
    sub_c1, sub_c2 = 22, RIGHT              # W..Z

    ws.merge_range(header_row, concept_c1, header_row, concept_c2, "Concepto", th_fmt)
    ws.merge_range(header_row, cost_c1, cost_c2, header_row, "Costo unitario", th_fmt)
    ws.merge_range(header_row, units_c1, units_c2, header_row, "Unidades", th_fmt)
    ws.merge_range(header_row, sub_c1, sub_c2, header_row, "Subtotal", th_fmt)

    # Items rows
    start_items_row = header_row + 1
    max_rows = max(1, min(len(payload.items), 18))  # tope conservador
    wrap_chars = 55  # ancho estimado para B..N (merge)

    for idx in range(max_rows):
        rr = start_items_row + idx
        item = payload.items[idx]

        # Zebra
        zebra = (idx % 2 == 1)
        cfmt = concept_fmt_z if zebra else concept_fmt
        mfmt = money_fmt_z if zebra else money_fmt
        ufmt = units_fmt_z if zebra else units_fmt
        sfmt = subtotal_fmt_z if zebra else subtotal_fmt

        # Row height “aireada” en Concepto
        min_h = 30
        needed_h = _row_height_for_wrapped_text(item.concept, wrap_chars, base_line_height=12.5, extra_lines=1.0)
        ws.set_row(rr, int(max(min_h, math.ceil(needed_h))), white_bg)

        # Write cells
        ws.merge_range(rr, concept_c1, rr, concept_c2, item.concept, cfmt)
        ws.merge_range(rr, cost_c1, rr, cost_c2, item.unit_cost, mfmt)
        ws.merge_range(rr, units_c1, rr, units_c2, item.units, ufmt)
        ws.merge_range(rr, sub_c1, rr, sub_c2, item.subtotal, sfmt)

    last_row = start_items_row + max_rows - 1

    # Total row area (white background explicit)
    total_row = last_row + 2
    ws.set_row(total_row, 26, white_bg)

    # ✅ Evitar ###: damos suficiente ancho al área del total
    # En grid 2.5, el área W..Z (4 cols) puede quedar justa con font grande.
    # Aquí colocamos TOTAL en T..V y el número en W..Z (ya ampliado por merge),
    # y además bajamos a font 11 (ya aplicado).
    ws.merge_range(total_row, units_c1, total_row, units_c2, "TOTAL:", total_label_fmt)
    ws.merge_range(total_row, sub_c1, total_row, sub_c2, payload.total, total_money_fmt)

    # White fill for all canvas outside shapes (ya está con white_bg global),
    # pero reforzamos que todo lo no azul/gris sea blanco:
    # (no-op extra; lo importante es usar bg_color en formatos anteriores)

    # -------------------- Print settings --------------------
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    ws.fit_to_pages(1, 0)

    # Print area (XlsxWriter usa print_area)
    ws.print_area(range_a1(1, 1, total_row + 3, 26))  # A1:Z( ... )

    wb.close()
    return out.getvalue()

