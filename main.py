# main.py
# ODCs - Excel only (NO merged cells), grid-painted layout.
# Render + FastAPI + xlsxwriter
#
# Endpoints:
#   GET  /health
#   POST /generate-odc-excel
#
# Notes:
# - Uses xlsxwriter (recommended for this "painted grid" approach).
# - No merges; "wide cells" are simulated by writing text in the left cell and painting blanks across the range.
# - All formats are explicit (no Format.properties cloning).

from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import List, Optional
from io import BytesIO
from datetime import datetime

import xlsxwriter

app = FastAPI(title="ODCs Render (Excel Only - No Merges)", version="1.0.2")


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


# ----------------- LAYOUT SPEC (único lugar para mover rangos) -----------------
LAYOUT = {
    "grid": {"col_start": "B", "col_end": "Z", "col_width": 2.5, "rows": 200, "base_row_h": 18},

    "banner": {"r1": 2, "c1": "B", "r2": 4, "c2": "Z", "row_heights": {2: 26, 3: 26, 4: 22}},

    # ODC # box
    "odc_box": {"row": 3, "label": ("T", "W"), "value": ("X", "Z")},

    # Left block (B:O rows 5-9) — labels B:E, values F:O
    "left": {
        "label_cols": ("B", "E"),
        "value_cols": ("F", "O"),
        "rows": {"odc": 5, "fecha": 6, "proveedor": 7, "servicio": 8, "proyecto": 9},
        "zebra_rows": [5, 7, 9],
        "divider_col": "E",  # draw right border
    },

    # Right block (Q:Z rows 5-9)
    "bill_to": {"title_row": 5, "cols": ("Q", "Z"), "lines": {6: "name", 7: "rfc", 8: "a1", 9: "a2"}},

    # Items table
    "items": {
        "header_row": 11,
        "header_height": 28,
        "start_row": 12,
        "zebra_gray_on_odd_rows": True,  # 13,15...
        "splits": {  # Concept B:N | Unit O:S | Units T:V | Subtotal W:Z
            "concept": ("B", "N"),
            "unit_cost": ("O", "S"),
            "units": ("T", "V"),
            "subtotal": ("W", "Z"),
        },
    },
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
    FONT = "Calibri"

    # ---- Formats (explícitos) ----
    white_fmt = wb.add_format({"bg_color": WHITE})

    banner_bg = wb.add_format({"bg_color": TEAL})
    banner_title = wb.add_format({
        "font_name": FONT, "font_size": 34, "bold": True,
        "font_color": WHITE, "bg_color": TEAL,
        "align": "left", "valign": "vcenter"
    })
    banner_sub = wb.add_format({
        "font_name": FONT, "font_size": 12,
        "font_color": "#CFE3EA", "bg_color": TEAL,
        "align": "left", "valign": "vcenter"
    })

    box_bg = wb.add_format({"bg_color": GRAY, "border": 1})
    box_label = wb.add_format({
        "font_name": FONT, "font_size": 16, "bold": True,
        "font_color": TEAL_2, "bg_color": GRAY,
        "border": 1, "align": "center", "valign": "vcenter"
    })
    box_value = wb.add_format({
        "font_name": FONT, "font_size": 18, "bold": True,
        "font_color": RED, "bg_color": GRAY,
        "border": 1, "align": "center", "valign": "vcenter"
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

    # Divider cell formats (same bg, right border)
    div_w = wb.add_format({"bg_color": WHITE, "right": 1})
    div_g = wb.add_format({"bg_color": GRAY, "right": 1})

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
        "border": 1, "align": "center", "valign": "vcenter"
    })

    td_left = wb.add_format({
        "font_name": FONT, "font_size": 16, "font_color": BLACK,
        "border": 1, "align": "left", "valign": "vcenter",
        "text_wrap": True
    })
    td_left_g = wb.add_format({
        "font_name": FONT, "font_size": 16, "font_color": BLACK,
        "border": 1, "align": "left", "valign": "vcenter",
        "text_wrap": True, "bg_color": GRAY
    })
    td_c = wb.add_format({
        "font_name": FONT, "font_size": 16, "font_color": BLACK,
        "border": 1, "align": "center", "valign": "vcenter"
    })
    td_c_g = wb.add_format({
        "font_name": FONT, "font_size": 16, "font_color": BLACK,
        "border": 1, "align": "center", "valign": "vcenter",
        "bg_color": GRAY
    })
    money = wb.add_format({
        "font_name": FONT, "font_size": 16, "font_color": BLACK,
        "border": 1, "align": "center", "valign": "vcenter",
        "num_format": '"$"#,##0'
    })
    money_g = wb.add_format({
        "font_name": FONT, "font_size": 16, "font_color": BLACK,
        "border": 1, "align": "center", "valign": "vcenter",
        "bg_color": GRAY,
        "num_format": '"$"#,##0'
    })

    # ---- Helpers ----
    def r0(r: int) -> int:
        return r - 1

    def c0(col_letter: str) -> int:
        return ord(col_letter.upper()) - ord("A")

    def rect(r1: int, c1: str, r2: int, c2: str, fmt):
        """Paint a rectangle by writing blanks in each cell."""
        for rr in range(r0(r1), r0(r2) + 1):
            for cc in range(c0(c1), c0(c2) + 1):
                ws.write_blank(rr, cc, None, fmt)

    def across(row: int, c_start: str, c_end: str, text, fmt):
        """Write text in the leftmost cell and paint blanks across (no merges)."""
        rr = r0(row)
        cs = c0(c_start)
        ce = c0(c_end)
        ws.write(rr, cs, text, fmt)
        for cc in range(cs + 1, ce + 1):
            ws.write_blank(rr, cc, None, fmt)

    # ---- Grid ----
    # global white background
    for c in range(0, 60):
        ws.set_column(c, c, None, white_fmt)

    gs = LAYOUT["grid"]
    for c in range(c0(gs["col_start"]), c0(gs["col_end"]) + 1):
        ws.set_column(c, c, gs["col_width"], white_fmt)

    for r in range(0, gs["rows"]):
        ws.set_row(r, gs["base_row_h"])

    # ---- Banner ----
    b = LAYOUT["banner"]
    for rr, hh in b["row_heights"].items():
        ws.set_row(r0(rr), hh)
    rect(b["r1"], b["c1"], b["r2"], b["c2"], banner_bg)
    ws.write(r0(3), c0("B"), "SAPIENCE", banner_title)
    ws.write(r0(4), c0("B"), "Human Insights Strategy", banner_sub)

    # ---- ODC box ----
    ob = LAYOUT["odc_box"]
    rect(ob["row"], ob["label"][0], ob["row"], ob["value"][1], box_bg)
    across(ob["row"], ob["label"][0], ob["label"][1], "ODC #:", box_label)
    across(ob["row"], ob["value"][0], ob["value"][1], payload.odc_number, box_value)

    # ---- Left block ----
    lb = LAYOUT["left"]
    rmap = lb["rows"]

    for rr in rmap.values():
        is_gray = rr in lb["zebra_rows"]
        bg_fmt = wb.add_format({"bg_color": GRAY if is_gray else WHITE})
        rect(rr, lb["label_cols"][0], rr, lb["value_cols"][1], bg_fmt)
        # divider cell at E with right border
        ws.write_blank(r0(rr), c0(lb["divider_col"]), None, div_g if is_gray else div_w)

    labels = [
        ("ODC #", rmap["odc"]),
        ("FECHA:", rmap["fecha"]),
        ("PROVEEDOR:", rmap["proveedor"]),
        ("SERVICIO:", rmap["servicio"]),
        ("PROYECTO:", rmap["proyecto"]),
    ]
    vals = [
        (payload.odc_number, rmap["odc"]),
        (payload.issue_date or "", rmap["fecha"]),
        (payload.supplier, rmap["proveedor"]),
        (payload.service, rmap["servicio"]),
        (payload.project, rmap["proyecto"]),
    ]

    for txt, rr in labels:
        across(rr, lb["label_cols"][0], lb["label_cols"][1], txt, label_g if rr in lb["zebra_rows"] else label)
    for txt, rr in vals:
        across(rr, lb["value_cols"][0], lb["value_cols"][1], txt, value_g if rr in lb["zebra_rows"] else value)

    # ---- Bill to ----
    bt = LAYOUT["bill_to"]
    across(bt["title_row"], bt["cols"][0], bt["cols"][1], "FACTURAR A:", facturar)

    start_col = bt["cols"][0]
    ws.write(r0(6), c0(start_col), payload.bill_to_name, bill_bold)
    ws.write(r0(7), c0(start_col), f"RFC: {payload.bill_to_rfc}", bill_bold)
    ws.write(r0(8), c0(start_col), payload.bill_to_address_1, bill)
    ws.write(r0(9), c0(start_col), payload.bill_to_address_2, bill)

    # ---- Items ----
    it = LAYOUT["items"]
    ws.set_row(r0(it["header_row"]), it["header_height"])
    rect(it["header_row"], gs["col_start"], it["header_row"], gs["col_end"], th)

    s = it["splits"]
    across(it["header_row"], s["concept"][0], s["concept"][1], "Concepto", th)
    across(it["header_row"], s["unit_cost"][0], s["unit_cost"][1], "Costo unitario", th)
    across(it["header_row"], s["units"][0], s["units"][1], "Unidades", th)
    across(it["header_row"], s["subtotal"][0], s["subtotal"][1], "Subtotal", th)

    start = it["start_row"]
    items = payload.items or [ODCItem()]
    row = start

    for item in items:
        gray = (row % 2 == 1) if it["zebra_gray_on_odd_rows"] else (row % 2 == 0)
        lf = td_left_g if gray else td_left
        cf = td_c_g if gray else td_c
        mf = money_g if gray else money

        across(row, s["concept"][0], s["concept"][1], item.concept, lf)
        across(row, s["unit_cost"][0], s["unit_cost"][1], (item.unit_cost or ""), mf)
        across(row, s["units"][0], s["units"][1], (item.units or ""), cf)

        subtotal = (item.unit_cost or 0.0) * (item.units or 0.0)
        across(row, s["subtotal"][0], s["subtotal"][1], (subtotal or ""), mf)

        row += 1

    last_row = row - 1

    # ---- Print setup (xlsxwriter API) ----
    ws.set_paper(1)        # Letter
    ws.set_portrait()      # portrait (NO args)
    ws.set_margins(0.3, 0.3, 0.35, 0.35)
    ws.fit_to_pages(1, 1)
    ws.print_area(r0(2), c0(gs["col_start"]), r0(last_row), c0(gs["col_end"]))

    wb.close()
    return out.getvalue()
