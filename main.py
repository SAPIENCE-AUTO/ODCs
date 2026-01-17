# main.py
# Sapience ODCs — Render + FastAPI — Excel only (XlsxWriter)

from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field, model_validator
from typing import List, Optional, Any, Dict
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
    subtotal: Optional[float] = None

    @model_validator(mode="before")
    @classmethod
    def _coerce_aliases(cls, data: Any):
        if not isinstance(data, dict):
            return data
        d = dict(data)

        if "concepto" in d and "concept" not in d:
            d["concept"] = d.get("concepto")
        if "costo_unitario" in d and "unit_cost" not in d:
            d["unit_cost"] = d.get("costo_unitario")
        if "unidades" in d and "units" not in d:
            d["units"] = d.get("unidades")

        return d


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

    items: List[ODCItem] = Field(default_factory=list)

    total: float
    currency_symbol: str = "$"

    logo_url: Optional[str] = "https://i.postimg.cc/Pf8KhptD/logo-sapience-blanco-15-ene-26.png"

    @model_validator(mode="before")
    @classmethod
    def _coerce_payload_aliases(cls, data: Any):
        if not isinstance(data, dict):
            return data
        d: Dict[str, Any] = dict(data)

        # date_str from issue_date / fecha
        if "date_str" not in d:
            iso = d.get("issue_date") or d.get("fecha")
            if iso:
                try:
                    dt = datetime.fromisoformat(str(iso)).date()
                    meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
                    d["date_str"] = f"{dt.day} {meses[dt.month - 1]} {dt.year}"
                except Exception:
                    d["date_str"] = str(iso)

        if "provider" not in d:
            d["provider"] = d.get("supplier") or d.get("proveedor") or ""
        if "service" not in d:
            d["service"] = d.get("servicio") or ""
        if "project" not in d:
            d["project"] = d.get("proyecto") or ""

        fact = d.get("facturar_a")
        if isinstance(fact, dict):
            d.setdefault("bill_to_title", "FACTURAR A:")
            d.setdefault("bill_to_name", fact.get("razon_social") or "")
            d.setdefault("bill_to_rfc", fact.get("rfc") or "")
            d.setdefault("bill_to_address_1", fact.get("direccion_linea1") or "")
            d.setdefault("bill_to_address_2", fact.get("direccion_linea2") or "")

        return d


# -----------------------------
# Helpers
# -----------------------------
def _png_size(img_bytes: bytes):
    if len(img_bytes) < 24 or img_bytes[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    try:
        w = struct.unpack(">I", img_bytes[16:20])[0]
        h = struct.unpack(">I", img_bytes[20:24])[0]
        return int(w), int(h)
    except Exception:
        return None


def range_a1(row1: int, col1: int, row2: int, col2: int) -> str:
    return xl_range(row1 - 1, col1 - 1, row2 - 1, col2 - 1)


def safe_float(x) -> float:
    try:
        return float(x)
    except Exception:
        return 0.0


def fill_range(ws, row1: int, col1: int, row2: int, col2: int, fmt):
    for rr in range(row1, row2 + 1):
        for cc in range(col1, col2 + 1):
            ws.write_blank(rr - 1, cc - 1, "", fmt)


def row_height_for_wrapped_text(
    text: str,
    wrap_width_chars: int,
    base_line_height: float = 11.5,
    extra_lines: float = 1.3,
) -> float:
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
# Excel Builder
# -----------------------------
def build_odc_excel(payload: ODCPayload) -> bytes:
    out = BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = wb.add_worksheet("ODC")

    FONT = "Montserrat"
    TEAL = "#0F3D4C"
    TEAL_2 = "#0E4A5A"
    WHITE = "#FFFFFF"
    LIGHT_GRAY = "#EFEFEF"
    GRID = "#7C7C7C"
    RED = "#E10600"
    BLACK = "#111111"

    # Font sizes
    FS_ODC_TOPRIGHT = 9
    FS_LEFT_TITLES = 8
    FS_LEFT_VALUES = 8
    FS_BILL_TITLE = 10
    FS_BILL_TEXT = 8
    FS_TABLE_HDR = 9
    FS_TABLE_BODY = 8
    FS_TOTAL_LABEL = 10
    FS_TOTAL_VALUE = 11

    # Columns A..Z
    for c in range(0, 26):
        ws.set_column(c, c, 2.5)

    white_bg = wb.add_format({"bg_color": WHITE})
    fill_range(ws, 1, 1, 220, 26, white_bg)

    banner_fill = wb.add_format({"bg_color": TEAL})
    gray_fill = wb.add_format({"bg_color": LIGHT_GRAY})

    odc_box_lbl = wb.add_format({
        "font_name": FONT, "font_size": FS_ODC_TOPRIGHT, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })
    odc_box_val = wb.add_format({
        "font_name": FONT, "font_size": FS_ODC_TOPRIGHT, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": RED, "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })

    label_gray = wb.add_format({
        "font_name": FONT, "font_size": FS_LEFT_TITLES, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": LIGHT_GRAY,
        "right": 1, "right_color": GRID,
    })
    value_gray = wb.add_format({
        "font_name": FONT, "font_size": FS_LEFT_VALUES,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": LIGHT_GRAY,
        "text_wrap": True
    })
    label_white = wb.add_format({
        "font_name": FONT, "font_size": FS_LEFT_TITLES, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE,
        "right": 1, "right_color": GRID,
    })
    value_white = wb.add_format({
        "font_name": FONT, "font_size": FS_LEFT_VALUES,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE,
        "text_wrap": True
    })

    bill_title_fmt = wb.add_format({
        "font_name": FONT, "font_size": FS_BILL_TITLE, "bold": True,
        "align": "left", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })
    bill_bold_fmt = wb.add_format({
        "font_name": FONT, "font_size": FS_BILL_TEXT, "bold": True,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE
    })
    bill_norm_fmt = wb.add_format({
        "font_name": FONT, "font_size": FS_BILL_TEXT,
        "align": "left", "valign": "vcenter",
        "font_color": BLACK, "bg_color": WHITE
    })

    th_fmt = wb.add_format({
        "font_name": FONT, "font_size": FS_TABLE_HDR, "bold": True,
        "align": "center", "valign": "vcenter",
        "font_color": WHITE, "bg_color": TEAL_2,
        "border": 1, "border_color": GRID,
    })

    # ✅ Change valign to vcenter for all table cells
    concept_w = wb.add_format({
        "font_name": FONT, "font_size": FS_TABLE_BODY,
        "align": "left", "valign": "vcenter",     # ✅
        "font_color": BLACK,
        "text_wrap": True,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })
    concept_g = wb.add_format({
        "font_name": FONT, "font_size": FS_TABLE_BODY,
        "align": "left", "valign": "vcenter",     # ✅
        "font_color": BLACK,
        "text_wrap": True,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
    })

    money_w = wb.add_format({
        "font_name": FONT, "font_size": FS_TABLE_BODY,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })
    money_g = wb.add_format({
        "font_name": FONT, "font_size": FS_TABLE_BODY,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })

    units_w = wb.add_format({
        "font_name": FONT, "font_size": FS_TABLE_BODY,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": WHITE,
        "border": 1, "border_color": GRID,
    })
    units_g = wb.add_format({
        "font_name": FONT, "font_size": FS_TABLE_BODY,
        "align": "center", "valign": "vcenter",
        "font_color": BLACK,
        "bg_color": LIGHT_GRAY,
        "border": 1, "border_color": GRID,
    })

    total_label_fmt = wb.add_format({
        "font_name": FONT, "font_size": FS_TOTAL_LABEL, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": TEAL_2, "bg_color": WHITE
    })
    total_money_fmt = wb.add_format({
        "font_name": FONT, "font_size": FS_TOTAL_VALUE, "bold": True,
        "align": "right", "valign": "vcenter",
        "font_color": RED, "bg_color": WHITE,
        "num_format": f'"{payload.currency_symbol}"#,##0.00'
    })

    # Layout
    ws.set_row(0, 10)

    # Banner 3x16
    ws.set_row(1, 16)
    ws.set_row(2, 16)
    ws.set_row(3, 16)
    fill_range(ws, 2, 2, 4, 26, banner_fill)

    # ✅ Smaller logo: tighter caps
    if payload.logo_url:
        try:
            resp = requests.get(payload.logo_url, timeout=15)
            resp.raise_for_status()
            img = resp.content
            wh = _png_size(img)

            x_scale = y_scale = 0.10
            y_off = 0

            if wh:
                w_px, h_px = wh
                # tighter than before
                scale = min(280 / max(1, w_px), 40 / max(1, h_px))
                scale = max(0.06, min(scale, 0.11))
                x_scale = y_scale = scale

                target_px_h = 16 * 3 * 4 / 3
                scaled_h = h_px * y_scale
                y_off = max(0, int((target_px_h - scaled_h) / 2))

            ws.insert_image(
                1, 1, "logo.png",
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

    # ODC box (1 row)
    ws.merge_range(1, 19, 1, 22, "ODC #:", odc_box_lbl)
    ws.merge_range(1, 23, 1, 25, payload.odc_number, odc_box_val)

    # Left meta
    meta_rows = [
        ("ODC #", payload.odc_number),
        ("FECHA:", payload.date_str),
        ("PROVEEDOR:", payload.provider),
        ("SERVICIO:", payload.service),
        ("PROYECTO:", payload.project),
    ]
    for i, (lab, val) in enumerate(meta_rows):
        rr = 5 + i
        ws.set_row(rr - 1, 20)
        is_gray = (i % 2 == 0)
        fill_range(ws, rr, 2, rr, 14, gray_fill if is_gray else white_bg)
        ws.merge_range(rr - 1, 1, rr - 1, 4, lab, label_gray if is_gray else label_white)
        ws.merge_range(rr - 1, 5, rr - 1, 13, val, value_gray if is_gray else value_white)

    # Bill-to
    fill_range(ws, 5, 15, 9, 26, white_bg)
    ws.merge_range(4, 14, 4, 25, payload.bill_to_title, bill_title_fmt)
    ws.merge_range(5, 14, 5, 25, payload.bill_to_name, bill_bold_fmt)
    ws.merge_range(6, 14, 6, 25, f"RFC: {payload.bill_to_rfc}", bill_bold_fmt)
    ws.merge_range(7, 14, 7, 25, payload.bill_to_address_1, bill_norm_fmt)
    ws.merge_range(8, 14, 8, 25, payload.bill_to_address_2, bill_norm_fmt)

    # Spacer
    ws.set_row(9, 12)

    # Table header
    header_row = 11
    ws.set_row(header_row - 1, 26)

    ws.merge_range(header_row - 1, 1, header_row - 1, 13, "Concepto", th_fmt)
    ws.merge_range(header_row - 1, 14, header_row - 1, 18, "Costo unitario", th_fmt)
    ws.merge_range(header_row - 1, 19, header_row - 1, 21, "Unidades", th_fmt)
    ws.merge_range(header_row - 1, 22, header_row - 1, 25, "Subtotal", th_fmt)

    # Items
    start_items = 12
    items = payload.items or [ODCItem(concept="", unit_cost=0, units=0)]
    max_items = min(len(items), 18)

    wrap_chars = 48
    min_row_h = 28

    last_item_row = start_items - 1
    for idx in range(max_items):
        rr = start_items + idx
        it = items[idx]

        zebra = (idx % 2 == 1)
        row_fill = gray_fill if zebra else white_bg
        fill_range(ws, rr, 2, rr, 26, row_fill)

        needed = row_height_for_wrapped_text(it.concept, wrap_chars, base_line_height=11.5, extra_lines=1.3)
        ws.set_row(rr - 1, int(max(min_row_h, math.ceil(needed))))

        unit_cost = safe_float(it.unit_cost)
        units = safe_float(it.units)
        subtotal = safe_float(it.subtotal) if it.subtotal is not None else (unit_cost * units)

        ws.merge_range(rr - 1, 1, rr - 1, 13, it.concept, concept_g if zebra else concept_w)
        ws.merge_range(rr - 1, 14, rr - 1, 18, unit_cost, money_g if zebra else money_w)
        ws.merge_range(rr - 1, 19, rr - 1, 21, units, units_g if zebra else units_w)
        ws.merge_range(rr - 1, 22, rr - 1, 25, subtotal, money_g if zebra else money_w)

        last_item_row = rr

    # TOTAL
    total_row = last_item_row + 2
    ws.set_row(total_row - 1, 22)
    fill_range(ws, total_row, 2, total_row, 26, white_bg)

    ws.merge_range(total_row - 1, 19, total_row - 1, 21, "TOTAL:", total_label_fmt)
    ws.merge_range(total_row - 1, 22, total_row - 1, 25, safe_float(payload.total), total_money_fmt)

    # Print settings
    ws.set_landscape()
    ws.set_paper(9)
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    ws.fit_to_pages(1, 0)
    ws.print_area(range_a1(1, 1, total_row + 3, 26))

    wb.close()
    return out.getvalue()
