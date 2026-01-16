# main.py
# ODCs - Excel only (NO big merges; banner is just fill; small functional merges allowed)
# Render + FastAPI + xlsxwriter
#
# Endpoints:
#   GET  /health
#   POST /generate-odc-excel
#
# Key ideas:
# - Grid of equal-width columns (2.5) and controlled row heights.
# - Banner = fill only + logo image from URL (no merge).
# - ODC small box = pseudo-merge via "center across selection" (xlsxwriter) + fill.
# - Table headers/body use the same "span" helper (write in left cell + blank fill across range),
#   but we DO NOT write values into the other cells (only blanks), which is safe in xlsxwriter.
#
# Dependencies (requirements.txt):
#   fastapi==0.110.3
#   uvicorn==0.29.0
#   pydantic==2.7.4
#   XlsxWriter==3.2.0
#   requests==2.32.3

from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import List, Optional
from io import BytesIO
from datetime import datetime
import struct

import requests
import xlsxwriter

app = FastAPI(title="ODCs Render (Excel Only)", version="1.1.0")


# ----------------- MODELOS -----------------
class ODCItem(BaseModel):
    concept: str = ""
    unit_cost: float = 0.0
    units: float = 0.0


class ODCRequest(BaseModel):
    odc_number: str = Field(..., examples=["RI-02497"])
    issue_date: Optional[str] = Field(default_factory=lambda: datetime.now().strftime("%d %b %Y"))

    supplier: str
    service: str
    project: str

    bill_to_name: str
    bill_to_rfc: str
    bill_to_address_1: str
    bill_to_address_2: str

    items: List[ODCItem] = Field(default_factory=list)

    # Logo for banner (PNG recommended)
    logo_url: Optional[str] = None


# ----------------- UTILIDADES IMAGEN -----------------
def _png_size(img_bytes: bytes):
    if len(img_bytes) < 24 or img_bytes[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    try:
        w = struct.unpack(">I", img_bytes[16:20])[0]
        h = struct.unpack(">I", img_bytes[20:24])[0]
        return int(w), int(h)
    except Exception:
        return None


def excel_col_width_to_pixels(w: float) -> int:
    # same as your calendar best practice
    if w < 1.0:
        return int(w * 12 + 0.5)
    return int(w * 7 + 5 + 0.5)


def points_to_pixels(pt: float) -> float:
    return pt * 4 / 3


# ----------------- LAYOUT SPEC (ajusta aquí) -----------------
LAYOUT = {
    "grid": {"col_start": "B", "col_end": "Z", "col_width": 2.5, "rows": 120, "base_row_h": 18},

    # Banner rows 2..4
    "banner": {"r1": 2, "c1": "B", "r2": 4, "c2": "Z", "row_heights": {2: 26, 3: 26, 4: 22}},

    # ODC box (top-right) at row 3
    # label block T..W, value block X..Z
    "odc_box": {"row": 3, "label": ("T", "W"), "value": ("X", "Z")},

    # Left meta block rows 5..9: labels B..E, values F..O, with divider at E
    "left": {
        "label_cols": ("B", "E"),
        "value_cols": ("F", "O"),
        "rows": {"odc": 5, "fecha": 6, "proveedor": 7, "servicio": 8, "proyecto": 9},
        "zebra_rows": [5, 7, 9],
        "divider_col": "E",
    },

    # Right "FACTURAR A" block starts at Q
    "bill_to": {"title_row": 5, "cols": ("Q", "Z"), "lines": {6: "name", 7: "rfc", 8: "a1", 9: "a2"}},

    # Items table
    "items": {
        "header_row": 11,
        "header_height": 28,
        "start_row": 12,
        "zebra_gray_on_odd_rows": True,
        # Concept B..N | Unit O..S | Units T..V | Subtotal W..Z
        "splits": {
            "concept": ("B", "N"),
            "unit_cost": ("O", "S"),
            "units": ("T", "V"),
            "subtotal": ("W", "Z"),
        },
    },

    # Logo placement over banner (anchor cell + how many grid columns to approximate max width)
    "logo": {"anchor": ("B", 2), "cols_for_logo": 12, "x_offset": 10, "y_offset": 8},
}


# ----------------- RUTAS -----------------
@app.get("/health")
def health():
    return {"ok": True, "time": datetime.now().isoformat()}


@app.post("/generate-odc-excel")
def generate_odc_excel(payload: ODCRequest):
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


# ----------------- GENERADOR -----------------
def build_odc_excel(payload: ODCRequest) -> bytes:
    out = BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = wb.add_worksheet("ODC")

    # ---- Paleta + fuente ----
    TEAL = "#0F3B4A"
    TEAL_2 = "#0D4A60"
    WHITE = "#FFFFFF"
    GRAY = "#EFEFEF"
    BLACK = "#000000"
    RED = "#D50000"
    BORDER = "#6E6E6E"
    FONT = "Calibri"

    # ---- Formats (explícitos, sin .properties) ----
    white_fmt = wb.add_format({"bg_color": WHITE})

    banner_bg = wb.add_format({"bg_color": TEAL})

    # Box
    box_bg = wb.add_format({"bg_color": GRAY, "border": 1, "border_color": BORDER})
    box_label = wb.add_format({
        "font_name": FONT, "font_size": 16, "bold": True,
        "font_color": TEAL_2, "bg_color": GRAY,
        "border": 1, "border_color": BORDER,
        "align": "center", "valign": "vcenter",
        "text_h_align": 6,  # center_across (xlsxwriter internal)
    })
    box_value = wb.add_format({
        "font_name": FONT, "font_size": 18, "bold": True,
        "font_color": RED, "bg_color": GRAY,
        "border": 1, "border_color": BORDER,
        "align": "center", "valign": "vcenter",
        "text_h_align": 6,  # center_across
    })

    label = wb.add_format({
        "font_name": FONT, "font_size": 14, "bold": True,
        "font_color": TEAL_2, "align": "right", "valign": "vcenter"
    })
    value = wb.add_format({
        "font_name": FONT, "font_size": 16,
        "font_color": BLACK, "align": "left", "valign": "vcenter"
    })
    label_g = wb.add_format({
        "font_name": FONT, "font_size": 14, "bold": True,
        "font_color": TEAL_2, "align": "right", "valign": "vcenter",
        "bg_color": GRAY
    })
    value_g = wb.add_format({
        "font_name": FONT, "font_size": 16,
        "font_color": BLACK, "align": "left", "valign": "vcenter",
        "bg_color": GRAY
    })

    # Divider (right border only) on E
    div_w = wb.add_format({"bg_color": WHITE, "right": 1, "right_color": BORDER})
    div_g = wb.add_format({"bg_color": GRAY, "right": 1, "right_color": BORDER})

    facturar = wb.add_format({
        "font_name": FONT, "font_size": 22, "bold": True,
        "font_color": TEAL_2, "align": "center", "valign": "vcenter"
    })
    bill_bold = wb.add_format({
        "font_name": FONT, "font_size": 16, "bold": True,
        "font_color": BLACK, "align": "left", "valign": "vcenter"
    })
    bill = wb.add_format({
        "font_name": FONT, "font_size": 16,
        "font_color": BLACK, "align": "left", "valign": "vcenter"
    })

    th = wb.add_format({
        "font_name": FONT, "font_size": 16, "bold": True,
        "font_color": WHITE, "bg_color": TEAL_2,
        "border": 1, "border_color": BORDER,
        "align": "center", "valign": "vcenter",
        "text_h_align": 6,  # center_across (makes header behave like merged)
    })

    td_left = wb.add_format({
        "font_name": FONT, "font_size": 16, "font_color": BLACK,
        "border": 1, "border_color": BORDER,
        "align": "left", "valign": "vcenter",
        "text_wrap": True
    })
    td_left_g = wb.add_format({
        "font_name": FONT, "font_size": 16, "font_color": BLACK,
        "border": 1, "border_color": BORDER,
        "align": "left", "valign": "vcenter",
        "text_wrap": True, "bg_color": GRAY
    })
    td_c = wb.add_format({
        "font_name": FONT, "font_size": 16, "font_color": BLACK,
        "border": 1, "border_color": BORDER,
        "align": "center", "valign": "vcenter"
    })
    td_c_g = wb.add_format({
        "font_name": FONT, "font_size": 16, "font_color": BLACK,
        "border": 1, "border_color": BORDER,
        "align": "center", "valign": "vcenter",
        "bg_color": GRAY
    })
    money = wb.add_format({
        "font_name": FONT, "font_size": 16, "font_color": BLACK,
        "border": 1, "border_color": BORDER,
        "align": "center", "valign": "vcenter",
        "num_format": '"$"#,##0'
    })
    money_g = wb.add_format({
        "font_name": FONT, "font_size": 16, "font_color": BLACK,
        "border": 1, "border_color": BORDER,
        "align": "center", "valign": "vcenter",
        "bg_color": GRAY,
        "num_format": '"$"#,##0'
    })

    # ----------------- Helpers -----------------
    def r0(r: int) -> int:
        return r - 1

    def c0(col_letter: str) -> int:
        return ord(col_letter.upper()) - ord("A")

    def fill_range(row1: int, col1: str, row2: int, col2: str, fmt) -> None:
        """Paint range with blanks (safe in xlsxwriter)."""
        for rr in range(r0(row1), r0(row2) + 1):
            for cc in range(c0(col1), c0(col2) + 1):
                ws.write_blank(rr, cc, None, fmt)

    def span_write(row: int, col_start: str, col_end: str, text, fmt) -> None:
        """Write text in left cell + paint blanks to the right (visual span, no merges)."""
        rr = r0(row)
        cs, ce = c0(col_start), c0(col_end)
        ws.write(rr, cs, text, fmt)
        for cc in range(cs + 1, ce + 1):
            ws.write_blank(rr, cc, None, fmt)

    # ----------------- Grid setup -----------------
    # global white background
    for c in range(0, 70):
        ws.set_column(c, c, None, white_fmt)

    gs = LAYOUT["grid"]
    for c in range(c0(gs["col_start"]), c0(gs["col_end"]) + 1):
        ws.set_column(c, c, gs["col_width"], white_fmt)
    for r in range(0, gs["rows"]):
        ws.set_row(r, gs["base_row_h"])

    # ----------------- Banner (fill only) -----------------
    b = LAYOUT["banner"]
    for rr, hh in b["row_heights"].items():
        ws.set_row(r0(rr), hh)

    fill_range(b["r1"], b["c1"], b["r2"], b["c2"], banner_bg)

    # ----------------- Logo over banner -----------------
    if payload.logo_url:
        try:
            r = requests.get(payload.logo_url, timeout=12)
            img = r.content
            wh = _png_size(img)

            # anchor
            anchor_col, anchor_row_1based = LAYOUT["logo"]["anchor"][0], LAYOUT["logo"]["anchor"][1]
            anchor_row = r0(anchor_row_1based)
            anchor_col0 = c0(anchor_col)

            cols_for_logo = int(LAYOUT["logo"]["cols_for_logo"])
            target_w_px = excel_col_width_to_pixels(cols_for_logo * gs["col_width"])
            target_h_pt = float(b["row_heights"][2] + b["row_heights"][3])
            target_h_px = points_to_pixels(target_h_pt)

            x_scale = y_scale = 1.0
            if wh:
                w_px, h_px = wh
                scale = min(target_w_px / w_px, target_h_px / h_px)
                scale = max(0.05, min(scale, 3.0))
                x_scale = y_scale = scale

            ws.insert_image(
                anchor_row, anchor_col0, "sapience_logo.png",
                {
                    "image_data": BytesIO(img),
                    "x_scale": x_scale,
                    "y_scale": y_scale,
                    "x_offset": int(LAYOUT["logo"]["x_offset"]),
                    "y_offset": int(LAYOUT["logo"]["y_offset"]),
                    "object_position": 1
                }
            )
        except Exception:
            pass

    # ----------------- ODC small box (top-right) -----------------
    ob = LAYOUT["odc_box"]
    fill_range(ob["row"], ob["label"][0], ob["row"], ob["value"][1], box_bg)
    span_write(ob["row"], ob["label"][0], ob["label"][1], "ODC #:", box_label)
    span_write(ob["row"], ob["value"][0], ob["value"][1], payload.odc_number, box_value)

    # ----------------- Left meta block -----------------
    lb = LAYOUT["left"]
    rmap = lb["rows"]
    for rr in rmap.values():
        is_gray = rr in lb["zebra_rows"]
        bg = wb.add_format({"bg_color": GRAY if is_gray else WHITE})
        fill_range(rr, lb["label_cols"][0], rr, lb["value_cols"][1], bg)
        # divider at E
        ws.write_blank(r0(rr), c0(lb["divider_col"]), None, div_g if is_gray else div_w)

    pairs = [
        ("ODC #", payload.odc_number, rmap["odc"]),
        ("FECHA:", payload.issue_date or "", rmap["fecha"]),
        ("PROVEEDOR:", payload.supplier, rmap["proveedor"]),
        ("SERVICIO:", payload.service, rmap["servicio"]),
        ("PROYECTO:", payload.project, rmap["proyecto"]),
    ]
    for lab, val, rr in pairs:
        span_write(rr, lb["label_cols"][0], lb["label_cols"][1], lab, label_g if rr in lb["zebra_rows"] else label)
        span_write(rr, lb["value_cols"][0], lb["value_cols"][1], val, value_g if rr in lb["zebra_rows"] else value)

    # ----------------- Bill to block -----------------
    bt = LAYOUT["bill_to"]
    span_write(bt["title_row"], bt["cols"][0], bt["cols"][1], "FACTURAR A:", facturar)
    ws.write(r0(6), c0(bt["cols"][0]), payload.bill_to_name, bill_bold)
    ws.write(r0(7), c0(bt["cols"][0]), f"RFC: {payload.bill_to_rfc}", bill_bold)
    ws.write(r0(8), c0(bt["cols"][0]), payload.bill_to_address_1, bill)
    ws.write(r0(9), c0(bt["cols"][0]), payload.bill_to_address_2, bill)

    # ----------------- Items table -----------------
    it = LAYOUT["items"]
    ws.set_row(r0(it["header_row"]), it["header_height"])
    fill_range(it["header_row"], gs["col_start"], it["header_row"], gs["col_end"], th)

    s = it["splits"]
    span_write(it["header_row"], s["concept"][0], s["concept"][1], "Concepto", th)
    span_write(it["header_row"], s["unit_cost"][0], s["unit_cost"][1], "Costo unitario", th)
    span_write(it["header_row"], s["units"][0], s["units"][1], "Unidades", th)
    span_write(it["header_row"], s["subtotal"][0], s["subtotal"][1], "Subtotal", th)

    start = it["start_row"]
    items = payload.items or [ODCItem()]
    row = start

    for item in items:
        gray = (row % 2 == 1) if it["zebra_gray_on_odd_rows"] else (row % 2 == 0)
        lf = td_left_g if gray else td_left
        cf = td_c_g if gray else td_c
        mf = money_g if gray else money

        span_write(row, s["concept"][0], s["concept"][1], item.concept, lf)
        span_write(row, s["unit_cost"][0], s["unit_cost"][1], item.unit_cost or 0, mf)
        span_write(row, s["units"][0], s["units"][1], item.units or 0, cf)

        subtotal = float(item.unit_cost or 0) * float(item.units or 0)
        span_write(row, s["subtotal"][0], s["subtotal"][1], subtotal, mf)

        row += 1

    last_row = max(start, row - 1)

    # ----------------- Print setup -----------------
    ws.set_paper(1)        # Letter
    ws.set_portrait()      # portrait
    ws.set_margins(0.3, 0.3, 0.35, 0.35)
    ws.fit_to_pages(1, 1)
    ws.print_area(r0(2), c0(gs["col_start"]), r0(last_row), c0(gs["col_end"]))

    wb.close()
    return out.getvalue()
