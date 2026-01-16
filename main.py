from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import List, Optional
from io import BytesIO
from datetime import datetime

import xlsxwriter

app = FastAPI(title="ODCs Render (Excel Only - No Merges)", version="1.0.0")


# ----------------- MODELOS -----------------
class ODCItem(BaseModel):
    concept: str
    unit_cost: float
    units: float


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
        # devuelve el error real (como tu main que funciona)
        return JSONResponse(status_code=500, content={"error": str(e)})


# ----------------- GENERADOR -----------------
def build_odc_excel(payload: ODCRequest) -> bytes:
    out = BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = wb.add_worksheet("ODC")

    # ----- Paleta (ajústala si quieres calcar exacto) -----
    TEAL = "#0F3B4A"
    TEAL_2 = "#0D4A60"
    WHITE = "#FFFFFF"
    GRAY = "#EFEFEF"
    BLACK = "#000000"
    RED = "#D50000"

    FONT = "Calibri"  # usa Calibri para máxima compatibilidad Excel/PDF

    # ----- Formats -----
    fmt_white = wb.add_format({"bg_color": WHITE})
    fmt_banner_bg = wb.add_format({"bg_color": TEAL})
    fmt_banner_title = wb.add_format({"font_name": FONT, "font_size": 34, "bold": True, "font_color": WHITE, "bg_color": TEAL, "align": "left", "valign": "vcenter"})
    fmt_banner_sub = wb.add_format({"font_name": FONT, "font_size": 12, "font_color": "#CFE3EA", "bg_color": TEAL, "align": "left", "valign": "vcenter"})

    fmt_box = wb.add_format({"bg_color": GRAY, "border": 1, "align": "center", "valign": "vcenter", "font_name": FONT})
    fmt_box_label = wb.add_format({"bg_color": GRAY, "border": 1, "align": "center", "valign": "vcenter", "font_name": FONT, "font_size": 16, "bold": True, "font_color": TEAL_2})
    fmt_box_value = wb.add_format({"bg_color": GRAY, "border": 1, "align": "center", "valign": "vcenter", "font_name": FONT, "font_size": 18, "bold": True, "font_color": RED})

    fmt_label = wb.add_format({"font_name": FONT, "font_size": 14, "bold": True, "font_color": TEAL_2, "align": "right", "valign": "vcenter"})
    fmt_value = wb.add_format({"font_name": FONT, "font_size": 16, "font_color": BLACK, "align": "left", "valign": "vcenter"})

    fmt_label_gray = wb.add_format({**fmt_label.properties, "bg_color": GRAY})
    fmt_value_gray = wb.add_format({**fmt_value.properties, "bg_color": GRAY})

    fmt_facturar = wb.add_format({"font_name": FONT, "font_size": 22, "bold": True, "font_color": TEAL_2, "align": "center", "valign": "vcenter"})
    fmt_bill_bold = wb.add_format({"font_name": FONT, "font_size": 16, "bold": True, "font_color": BLACK, "align": "left", "valign": "vcenter"})
    fmt_bill = wb.add_format({"font_name": FONT, "font_size": 16, "font_color": BLACK, "align": "left", "valign": "vcenter"})

    fmt_th = wb.add_format({"font_name": FONT, "font_size": 16, "bold": True, "font_color": WHITE, "bg_color": TEAL_2, "border": 1, "align": "center", "valign": "vcenter"})
    fmt_td_left = wb.add_format({"font_name": FONT, "font_size": 16, "font_color": BLACK, "border": 1, "align": "left", "valign": "vcenter", "text_wrap": True})
    fmt_td_center = wb.add_format({"font_name": FONT, "font_size": 16, "font_color": BLACK, "border": 1, "align": "center", "valign": "vcenter"})
    fmt_td_center_gray = wb.add_format({**fmt_td_center.properties, "bg_color": GRAY})
    fmt_td_left_gray = wb.add_format({**fmt_td_left.properties, "bg_color": GRAY})

    money_fmt = wb.add_format({"font_name": FONT, "font_size": 16, "font_color": BLACK, "border": 1, "align": "center", "valign": "vcenter", "num_format": '"$"#,##0'})
    money_fmt_gray = wb.add_format({**money_fmt.properties, "bg_color": GRAY})

    # ----- Grid uniform B:Z width=2.5 -----
    # xlsxwriter uses 0-indexed columns: A=0, B=1 ... Z=25
    # Set a big white background format so unpainted areas are clean:
    for c in range(0, 60):
        ws.set_column(c, c, None, fmt_white)

    # Columns B..Z all 2.5
    for c in range(1, 26):  # B(1) .. Z(25)
        ws.set_column(c, c, 2.5, fmt_white)

    # Uniform row heights baseline + special rows
    for r in range(0, 200):
        ws.set_row(r, 18)
    # Banner (rows 2-4 in Excel = index 1-3 if you start at 1? careful)
    # We'll work in Excel-like numbers below and convert to 0-index.
    def r0(excel_row: int) -> int:
        return excel_row - 1

    # helper to get col index from Excel letter
    def c0(letter: str) -> int:
        return ord(letter.upper()) - ord("A")

    # Helpers: rect + across writing (no merges)
    def rect(r1, c1, r2, c2, fmt):
        # r/c passed as 1-based Excel coordinates; uses write_blank to paint
        for rr in range(r0(r1), r0(r2) + 1):
            for cc in range(c0(c1), c0(c2) + 1):
                ws.write_blank(rr, cc, None, fmt)

    def across(row, c_start, c_end, text, fmt):
        # write text in leftmost cell, paint the rest as blanks
        rr = r0(row)
        cs = c0(c_start)
        ce = c0(c_end)
        ws.write(rr, cs, text, fmt)
        for cc in range(cs + 1, ce + 1):
            ws.write_blank(rr, cc, None, fmt)

    # ----- Layout (based on your screenshot) -----
    # Banner B2:Z4
    ws.set_row(r0(2), 26)
    ws.set_row(r0(3), 26)
    ws.set_row(r0(4), 22)
    rect(2, "B", 4, "Z", fmt_banner_bg)
    ws.write(r0(3), c0("B"), "SAPIENCE", fmt_banner_title)
    ws.write(r0(4), c0("B"), "Human Insights Strategy", fmt_banner_sub)

    # ODC box T3:Z3 with divider at W
    rect(3, "T", 3, "Z", fmt_box)
    across(3, "T", "W", "ODC #:", fmt_box_label)
    across(3, "X", "Z", payload.odc_number, fmt_box_value)

    # Left table B5:O9 with zebra rows 5,7,9 gray
    # Labels B:E, Values F:O, divider after E
    def zebra_fmt(row_num):
        return GRAY if row_num in (5, 7, 9) else WHITE

    for row in range(5, 10):
        bg = zebra_fmt(row)
        if bg == GRAY:
            rect(row, "B", row, "O", wb.add_format({"bg_color": GRAY}))
        else:
            rect(row, "B", row, "O", wb.add_format({"bg_color": WHITE}))

    # Divider line: simulate by applying border to column E cells (right border)
    divider_fmt_gray = wb.add_format({"bg_color": GRAY, "right": 1})
    divider_fmt_white = wb.add_format({"bg_color": WHITE, "right": 1})
    for row in range(5, 10):
        ws.write_blank(r0(row), c0("E"), None, divider_fmt_gray if zebra_fmt(row) == GRAY else divider_fmt_white)

    # Labels
    labels = [
        ("ODC #", 5),
        ("FECHA:", 6),
        ("PROVEEDOR:", 7),
        ("SERVICIO:", 8),
        ("PROYECTO:", 9),
    ]
    for txt, row in labels:
        bg = zebra_fmt(row)
        across(row, "B", "E", txt, fmt_label_gray if bg == GRAY else fmt_label)

    # Values
    values = [
        (payload.odc_number, 5),
        (payload.issue_date or "", 6),
        (payload.supplier, 7),
        (payload.service, 8),
        (payload.project, 9),
    ]
    for txt, row in values:
        bg = zebra_fmt(row)
        across(row, "F", "O", txt, fmt_value_gray if bg == GRAY else fmt_value)

    # Right block Q5:Z9 (Facturar a)
    # Title centered across Q:Z row 5
    across(5, "Q", "Z", "FACTURAR A:", fmt_facturar)
    ws.write(r0(6), c0("Q"), payload.bill_to_name, fmt_bill_bold)
    ws.write(r0(7), c0("Q"), f"RFC: {payload.bill_to_rfc}", fmt_bill_bold)
    ws.write(r0(8), c0("Q"), payload.bill_to_address_1, fmt_bill)
    ws.write(r0(9), c0("Q"), payload.bill_to_address_2, fmt_bill)

    # Items header row 11: B..Z teal
    ws.set_row(r0(11), 28)
    rect(11, "B", 11, "Z", fmt_th)

    # Splits: Concept B:N | Unit cost O:S | Units T:V | Subtotal W:Z
    across(11, "B", "N", "Concepto", fmt_th)
    across(11, "O", "S", "Costo unitario", fmt_th)
    across(11, "T", "V", "Unidades", fmt_th)
    across(11, "W", "Z", "Subtotal", fmt_th)

    # Items body starting row 12
    start = 12
    items = payload.items or [ODCItem(concept="", unit_cost=0.0, units=0.0)]
    row = start
    for i, it in enumerate(items):
        zebra = (row % 2 == 1)  # 13,15,... gray (like your mock)
        left_fmt = fmt_td_left_gray if zebra else fmt_td_left
        cen_fmt = fmt_td_center_gray if zebra else fmt_td_center
        mon_fmt = money_fmt_gray if zebra else money_fmt

        # paint row blocks (borders handled by formats)
        across(row, "B", "N", it.concept, left_fmt)
        across(row, "O", "S", (it.unit_cost if it.unit_cost else ""), mon_fmt)
        across(row, "T", "V", (it.units if it.units else ""), cen_fmt)
        subtotal = (it.unit_cost or 0.0) * (it.units or 0.0)
        across(row, "W", "Z", (subtotal if subtotal else ""), mon_fmt)

        row += 1

    last_row = row - 1

    # Print setup (Excel will use this when exporting to PDF)
    ws.set_paper(1)            # 1 = Letter
    ws.set_landscape(False)    # portrait
    ws.set_margins(0.3, 0.3, 0.35, 0.35)
    ws.fit_to_pages(1, 1)

    # Print area B2:Z<last_row>
    ws.print_area(r0(2), c0("B"), r0(last_row), c0("Z"))

    wb.close()
    return out.getvalue()

