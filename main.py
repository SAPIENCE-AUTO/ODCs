from __future__ import annotations

import io
from typing import List, Optional, Any, Dict

import requests
from fastapi import FastAPI, Body, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, Field, ConfigDict


# ----------------------------
# Helpers
# ----------------------------

def safe_float(v: Any) -> float:
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    try:
        return float(s)
    except Exception:
        return 0.0


def range_a1(row1: int, col1: int, row2: int, col2: int) -> str:
    """1-based rows/cols -> Excel A1 range. Example: A1:Z10"""
    def col_to_name(c: int) -> str:
        name = ""
        while c > 0:
            c, rem = divmod(c - 1, 26)
            name = chr(65 + rem) + name
        return name

    return f"{col_to_name(col1)}{row1}:{col_to_name(col2)}{row2}"


# ----------------------------
# Payload models
# ----------------------------

class ODCItem(BaseModel):
    model_config = ConfigDict(extra="ignore")
    concepto: str
    costo_unitario: float = 0
    unidades: float = 0
    subtotal: float = 0


class FacturarA(BaseModel):
    model_config = ConfigDict(extra="ignore")
    razon_social: str
    rfc: str
    direccion_linea1: str
    direccion_linea2: str


class ODCPayload(BaseModel):
    model_config = ConfigDict(extra="ignore")

    logo_url: Optional[str] = None
    odc_number: str = Field(...)

    # fecha: tolera variantes
    date_str: Optional[str] = None
    issue_date: Optional[str] = None
    fecha: Optional[str] = None

    # proveedor: tolera variantes
    provider: Optional[str] = None
    supplier: Optional[str] = None
    proveedor: Optional[str] = None

    service: Optional[str] = None
    servicio: Optional[str] = None

    project: Optional[str] = None
    proyecto: Optional[str] = None

    # bill-to plano
    bill_to_name: Optional[str] = None
    bill_to_rfc: Optional[str] = None
    bill_to_address_1: Optional[str] = None
    bill_to_address_2: Optional[str] = None

    # bill-to objeto
    facturar_a: Optional[FacturarA] = None

    items: List[ODCItem] = Field(default_factory=list)
    total: float = 0.0

    def normalized(self) -> Dict[str, Any]:
        final_date = self.date_str or self.issue_date or self.fecha or ""
        final_provider = self.provider or self.supplier or self.proveedor or ""
        final_service = self.service or self.servicio or ""
        final_project = self.project or self.proyecto or ""

        if self.facturar_a:
            bill_name = self.facturar_a.razon_social
            bill_rfc = self.facturar_a.rfc
            bill_a1 = self.facturar_a.direccion_linea1
            bill_a2 = self.facturar_a.direccion_linea2
        else:
            bill_name = self.bill_to_name or ""
            bill_rfc = self.bill_to_rfc or ""
            bill_a1 = self.bill_to_address_1 or ""
            bill_a2 = self.bill_to_address_2 or ""

        return {
            "logo_url": self.logo_url,
            "odc_number": self.odc_number,
            "date": final_date,
            "provider": final_provider,
            "service": final_service,
            "project": final_project,
            "bill_to_name": bill_name,
            "bill_to_rfc": bill_rfc,
            "bill_to_address_1": bill_a1,
            "bill_to_address_2": bill_a2,
            "items": self.items,
            "total": safe_float(self.total),
        }


# ----------------------------
# Excel generator
# ----------------------------

def build_odc_excel(payload: ODCPayload) -> bytes:
    import xlsxwriter

    data = payload.normalized()

    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = wb.add_worksheet("ODC")

    # Colors
    BLUE = "#0E4C5A"
    LIGHT_GRAY = "#EFEFEF"
    WHITE = "#FFFFFF"
    RED = "#E10600"
    GRID = "#5A5A5A"

    # Fonts (tus tamaños)
    BASE_FONT = "Montserrat"
    FONT_BADGE = 9
    FONT_LABEL = 8
    FONT_FACTURAR = 10
    FONT_BILLTO = 8
    FONT_TABLE_HDR = 8
    FONT_TABLE_BODY = 11
    FONT_TOTAL_LABEL = 10
    FONT_TOTAL_NUM = 11

    # Base backgrounds
    fmt_white_bg = wb.add_format({"bg_color": WHITE})

    # Badge formats
    fmt_badge_label = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_BADGE,
        "bold": True, "align": "center", "valign": "vcenter",
        "bg_color": WHITE, "font_color": BLUE,
        "border": 1, "border_color": GRID
    })
    fmt_badge_value = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_BADGE,
        "bold": True, "align": "center", "valign": "vcenter",
        "bg_color": WHITE, "font_color": RED,
        "border": 1, "border_color": GRID
    })

    # Left block formats
    fmt_left_label = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_LABEL,
        "bold": True, "align": "right", "valign": "vcenter",
        "font_color": BLUE
    })

    fmt_value_white = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_BILLTO,
        "align": "left", "valign": "vcenter",
        "bg_color": WHITE, "text_wrap": True
    })
    fmt_value_gray = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_BILLTO,
        "align": "left", "valign": "vcenter",
        "bg_color": LIGHT_GRAY, "text_wrap": True
    })

    # Bill-to formats
    fmt_facturar_title = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_FACTURAR,
        "bold": True, "align": "left", "valign": "vcenter",
        "font_color": BLUE
    })
    fmt_billto_bold = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_BILLTO,
        "bold": True, "align": "left", "valign": "vcenter",
        "font_color": "#000000"
    })
    fmt_billto = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_BILLTO,
        "align": "left", "valign": "vcenter",
        "font_color": "#000000",
        "text_wrap": True
    })

    # Table formats
    fmt_table_hdr = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_TABLE_HDR,
        "bold": True, "align": "center", "valign": "vcenter",
        "bg_color": BLUE, "font_color": WHITE,
        "border": 1, "border_color": GRID
    })

    # Texto concepto (dos versiones para zebra)
    fmt_concept_white = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_TABLE_BODY,
        "align": "left", "valign": "vcenter",
        "text_wrap": True,
        "border": 1, "border_color": GRID,
        "bg_color": WHITE
    })
    fmt_concept_gray = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_TABLE_BODY,
        "align": "left", "valign": "vcenter",
        "text_wrap": True,
        "border": 1, "border_color": GRID,
        "bg_color": LIGHT_GRAY
    })

    fmt_table_num = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_TABLE_BODY,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": GRID,
        "bg_color": WHITE,
        "num_format": "0"
    })
    fmt_table_money = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_TABLE_BODY,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": GRID,
        "bg_color": WHITE,
        "num_format": "$#,##0.00"
    })

    # Total formats
    fmt_total_label = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_TOTAL_LABEL,
        "bold": True, "align": "right", "valign": "vcenter",
        "font_color": BLUE
    })
    fmt_total_num = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_TOTAL_NUM,
        "bold": True, "align": "left", "valign": "vcenter",
        "font_color": RED,
        "num_format": "$#,##0.00"
    })

    # -------------------- Layout base --------------------
    ws.set_default_row(18)

    # Columns A..Z
    ws.set_column(0, 0, 2.5)   # A
    ws.set_column(1, 1, 1.5)   # B
    ws.set_column(2, 2, 12)    # C labels
    ws.set_column(3, 12, 3.0)  # D..M
    ws.set_column(13, 18, 3.0) # N..S
    ws.set_column(19, 25, 4.2) # T..Z

    # Fondo blanco real
    for r in range(0, 80):
        ws.write_row(r, 0, [""] * 26, fmt_white_bg)

    # -------------------- Banner --------------------
    fmt_blue_fill = wb.add_format({"bg_color": BLUE})
    ws.merge_range(1, 1, 3, 25, "", fmt_blue_fill)
    ws.set_row(1, 30)
    ws.set_row(2, 30)
    ws.set_row(3, 22)

    # Logo (si falla, no rompe)
    if data["logo_url"]:
        try:
            resp = requests.get(data["logo_url"], timeout=12)
            resp.raise_for_status()
            img_bytes = io.BytesIO(resp.content)
            ws.insert_image(1, 1, "logo.png", {
                "image_data": img_bytes,
                "x_offset": 8,
                "y_offset": 8,
                "x_scale": 0.55,
                "y_scale": 0.55,
            })
        except Exception:
            pass

    # Badge arriba derecha
    ws.merge_range(1, 19, 2, 22, "ODC #:", fmt_badge_label)
    ws.merge_range(1, 23, 2, 25, data["odc_number"], fmt_badge_value)

    # -------------------- Left info block --------------------
    start_row = 4
    label_col = 2
    val_c1, val_c2 = 3, 12  # D..M

    ws.set_row(start_row, 20)
    ws.write(start_row, label_col, "ODC #", fmt_left_label)
    ws.merge_range(start_row, val_c1, start_row, val_c2, data["odc_number"], fmt_value_gray)

    ws.set_row(start_row + 1, 20)
    ws.write(start_row + 1, label_col, "FECHA:", fmt_left_label)
    ws.merge_range(start_row + 1, val_c1, start_row + 1, val_c2, data["date"], fmt_value_white)

    ws.set_row(start_row + 2, 22)
    ws.write(start_row + 2, label_col, "PROVEEDOR:", fmt_left_label)
    ws.merge_range(start_row + 2, val_c1, start_row + 2, val_c2, data["provider"], fmt_value_gray)

    ws.set_row(start_row + 3, 32)  # aire
    ws.write(start_row + 3, label_col, "SERVICIO:", fmt_left_label)
    ws.merge_range(start_row + 3, val_c1, start_row + 3, val_c2, data["service"], fmt_value_white)

    ws.set_row(start_row + 4, 22)
    ws.write(start_row + 4, label_col, "PROYECTO:", fmt_left_label)
    ws.merge_range(start_row + 4, val_c1, start_row + 4, val_c2, data["project"], fmt_value_gray)

    # -------------------- Bill to block --------------------
    bill_col1, bill_col2 = 13, 25
    bill_row = start_row

    ws.set_row(bill_row, 22)
    ws.merge_range(bill_row, bill_col1, bill_row, bill_col2, "FACTURAR A:", fmt_facturar_title)

    ws.set_row(bill_row + 1, 20)
    ws.merge_range(bill_row + 1, bill_col1, bill_row + 1, bill_col2, data["bill_to_name"], fmt_billto_bold)

    ws.set_row(bill_row + 2, 20)
    ws.merge_range(bill_row + 2, bill_col1, bill_row + 2, bill_col2, f"RFC: {data['bill_to_rfc']}".strip(), fmt_billto_bold)

    ws.set_row(bill_row + 3, 20)
    ws.merge_range(bill_row + 3, bill_col1, bill_row + 3, bill_col2, data["bill_to_address_1"], fmt_billto)

    ws.set_row(bill_row + 4, 20)
    ws.merge_range(bill_row + 4, bill_col1, bill_row + 4, bill_col2, data["bill_to_address_2"], fmt_billto)

    # -------------------- Items table --------------------
    table_header_row = start_row + 6
    ws.set_row(table_header_row, 28)

    c_concept_1, c_concept_2 = 1, 13
    c_cost_1, c_cost_2 = 14, 18
    c_units_1, c_units_2 = 19, 20
    c_subt_1, c_subt_2 = 21, 25

    ws.merge_range(table_header_row, c_concept_1, table_header_row, c_concept_2, "Concepto", fmt_table_hdr)
    ws.merge_range(table_header_row, c_cost_1, table_header_row, c_cost_2, "Costo unitario", fmt_table_hdr)
    ws.merge_range(table_header_row, c_units_1, table_header_row, c_units_2, "Unidades", fmt_table_hdr)
    ws.merge_range(table_header_row, c_subt_1, table_header_row, c_subt_2, "Subtotal", fmt_table_hdr)

    body_start = table_header_row + 1
    row_h = 40

    for i, it in enumerate(data["items"]):
        r = body_start + i
        ws.set_row(r, row_h)
        concept_fmt = fmt_concept_gray if (i % 2 == 1) else fmt_concept_white

        ws.merge_range(r, c_concept_1, r, c_concept_2, it.concepto, concept_fmt)
        ws.merge_range(r, c_cost_1, r, c_cost_2, safe_float(it.costo_unitario), fmt_table_money)
        ws.merge_range(r, c_units_1, r, c_units_2, safe_float(it.unidades), fmt_table_num)
        ws.merge_range(r, c_subt_1, r, c_subt_2, safe_float(it.subtotal), fmt_table_money)

    last_item_row = body_start + max(len(data["items"]) - 1, 0)

    # -------------------- TOTAL --------------------
    total_row = last_item_row + 2
    ws.set_row(total_row, 26)

    ws.merge_range(total_row, 19, total_row, 21, "TOTAL:", fmt_total_label)
    ws.merge_range(total_row, 22, total_row, 25, data["total"], fmt_total_num)

    # -------------------- Print settings --------------------
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    ws.fit_to_pages(1, 0)
    ws.print_area(range_a1(1, 1, total_row + 3, 26))

    wb.close()
    out.seek(0)
    return out.getvalue()


# ----------------------------
# FastAPI app
# ----------------------------

app = FastAPI(title="ODC Excel Generator", version="1.0.1")


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/generate-odc-excel")
def generate_odc_excel(payload: ODCPayload = Body(...)):
    data = payload.normalized()

    missing = []
    if not data["odc_number"]:
        missing.append("odc_number")
    if not data["date"]:
        missing.append("date_str/issue_date/fecha")
    if not data["provider"]:
        missing.append("provider/supplier/proveedor")
    if not data["service"]:
        missing.append("service/servicio")
    if not data["project"]:
        missing.append("project/proyecto")
    if not data["bill_to_name"]:
        missing.append("bill_to_name o facturar_a.razon_social")
    if not data["bill_to_rfc"]:
        missing.append("bill_to_rfc o facturar_a.rfc")

    if missing:
        raise HTTPException(status_code=422, detail={"missing_fields": missing})

    xls_bytes = build_odc_excel(payload)
    filename = f"ODC_{data['odc_number']}.xlsx"

    return StreamingResponse(
        io.BytesIO(xls_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
