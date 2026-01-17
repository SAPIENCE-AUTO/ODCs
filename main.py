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
    """1-based rows/cols -> Excel A1 range. Example: range_a1(1,1,10,26) => A1:Z10"""
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
    """
    Este modelo acepta tanto las llaves "finales" como tus llaves en español.
    Así evitas 422 cuando el JSON cambie un poco.
    """
    model_config = ConfigDict(extra="ignore")

    logo_url: Optional[str] = None
    odc_number: str = Field(..., description="Ej: RI-02497")

    # Soportamos ambos: date_str (lo que te estaba pidiendo) y issue_date/fecha
    date_str: Optional[str] = None
    issue_date: Optional[str] = None
    fecha: Optional[str] = None  # alias “legacy” (por si llega así)

    # Soportamos ambos: provider (lo que te estaba pidiendo) y supplier/proveedor
    provider: Optional[str] = None
    supplier: Optional[str] = None
    proveedor: Optional[str] = None

    service: Optional[str] = None
    servicio: Optional[str] = None

    project: Optional[str] = None
    proyecto: Optional[str] = None

    # Facturar a: puede venir “flat” o en objeto
    bill_to_name: Optional[str] = None
    bill_to_rfc: Optional[str] = None
    bill_to_address_1: Optional[str] = None
    bill_to_address_2: Optional[str] = None

    facturar_a: Optional[FacturarA] = None

    items: List[ODCItem] = Field(default_factory=list)
    total: float = 0.0

    def normalized(self) -> Dict[str, Any]:
        # Date
        final_date = self.date_str or self.issue_date or self.fecha or ""

        # Provider
        final_provider = self.provider or self.supplier or self.proveedor or ""

        # Service / Project
        final_service = self.service or self.servicio or ""
        final_project = self.project or self.proyecto or ""

        # Bill-to (flat vs object)
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
# Excel generator (XlsxWriter)
# ----------------------------

def build_odc_excel(payload: ODCPayload) -> bytes:
    import xlsxwriter  # keep local import for Render cold starts

    data = payload.normalized()

    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    ws = wb.add_worksheet("ODC")

    # ----- Colors (ajusta si ya tienes hex exactos) -----
    BLUE = "#0E4C5A"
    LIGHT_GRAY = "#EFEFEF"
    WHITE = "#FFFFFF"
    RED = "#E10600"
    GRID = "#5A5A5A"

    # ----- Font system (tus tamaños) -----
    BASE_FONT = "Montserrat"

    FONT_BADGE = 9
    FONT_LABEL = 8
    FONT_FACTURAR = 10
    FONT_BILLTO = 8
    FONT_TABLE_HDR = 8
    FONT_TABLE_BODY = 11
    FONT_TOTAL_LABEL = 10
    FONT_TOTAL_NUM = 11

    # ----- Base formats -----
    fmt_white = wb.add_format({"bg_color": WHITE})  # “blanco real”
    fmt_gray = wb.add_format({"bg_color": LIGHT_GRAY})
    fmt_blue_fill = wb.add_format({"bg_color": BLUE})

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

    fmt_left_label = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_LABEL,
        "bold": True, "align": "right", "valign": "vcenter",
        "font_color": BLUE
    })
    fmt_left_value = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_BILLTO,
        "align": "left", "valign": "vcenter",
        "font_color": "#000000",
        "text_wrap": True
    })

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

    fmt_table_hdr = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_TABLE_HDR,
        "bold": True, "align": "center", "valign": "vcenter",
        "bg_color": BLUE, "font_color": WHITE,
        "border": 1, "border_color": GRID
    })

    fmt_table_body_text = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_TABLE_BODY,
        "align": "left", "valign": "vcenter",
        "text_wrap": True,
        "border": 1, "border_color": GRID,
        "bg_color": WHITE
    })
    fmt_table_body_num = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_TABLE_BODY,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": GRID,
        "bg_color": WHITE,
        "num_format": "0"
    })
    fmt_table_body_money = wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_TABLE_BODY,
        "align": "center", "valign": "vcenter",
        "border": 1, "border_color": GRID,
        "bg_color": WHITE,
        "num_format": "$#,##0.00"
    })

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

    # -------------------- Sheet layout --------------------
    # Columns A..Z (26 cols)
    # Ajusta anchos para que “Concepto” tenga aire y no se sienta apretado
    ws.set_default_row(18)

    # Base: deja A un poco de margen visual
    ws.set_column(0, 0, 2.5)   # A
    ws.set_column(1, 1, 1.5)   # B
    ws.set_column(2, 2, 12)    # C (labels)
    ws.set_column(3, 12, 3.0)  # D..M (grid para merges)
    ws.set_column(13, 18, 3.0) # N..S
    ws.set_column(19, 25, 4.2) # T..Z (zona total / badge)

    # Fondo blanco “real” en un área amplia
    for r in range(0, 60):
        ws.set_row(r, 18)
        ws.write_row(r, 0, [""] * 26, fmt_white)

    # -------------------- Banner --------------------
    # Banner: filas 1-3 (Excel 1-based). Aquí usamos 0-based.
    # Area azul: B2:Z4 -> rows 1..3, cols 1..25
    ws.merge_range(1, 1, 3, 25, "", fmt_blue_fill)
    ws.set_row(1, 30)  # más alto para el logo
    ws.set_row(2, 30)
    ws.set_row(3, 22)

    # Insert logo (más chico)
    # Lo colocamos cerca de B2
    if data["logo_url"]:
        try:
            resp = requests.get(data["logo_url"], timeout=12)
            resp.raise_for_status()
            img_bytes = io.BytesIO(resp.content)
            # Ajuste de escala (baja el tamaño)
            ws.insert_image(1, 1, "logo.png", {
                "image_data": img_bytes,
                "x_offset": 8,
                "y_offset": 8,
                "x_scale": 0.55,
                "y_scale": 0.55,
            })
        except Exception:
            # si falla el logo, no rompemos el excel
            pass

    # -------------------- Badge ODC (arriba derecha) --------------------
    # Bloque blanco con ODC#: cols V..Z aprox (22..25) + etiqueta (19..21)
    # Fila 2-3 para que se vea centrado
    badge_row1, badge_row2 = 1, 2
    ws.merge_range(badge_row1, 19, badge_row2, 22, "ODC #:", fmt_badge_label)
    ws.merge_range(badge_row1, 23, badge_row2, 25, data["odc_number"], fmt_badge_value)

    # -------------------- Left info block --------------------
    # Layout base:
    # Labels en C (col 2), values merge en D..M (3..12)
    start_row = 4  # fila 5 visual
    label_col = 2
    val_c1, val_c2 = 3, 12  # D..M

    # ODC# row (fondo gris para value)
    ws.set_row(start_row, 20)
    ws.write(start_row, label_col, "ODC #", fmt_left_label)
    ws.merge_range(start_row, val_c1, start_row, val_c2, data["odc_number"], wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_BILLTO,
        "align": "left", "valign": "vcenter", "bg_color": LIGHT_GRAY
    }))

    # FECHA
    ws.set_row(start_row + 1, 20)
    ws.write(start_row + 1, label_col, "FECHA:", fmt_left_label)
    ws.merge_range(start_row + 1, val_c1, start_row + 1, val_c2, data["date"], wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_BILLTO,
        "align": "left", "valign": "vcenter", "bg_color": WHITE
    }))

    # PROVEEDOR
    ws.set_row(start_row + 2, 22)
    ws.write(start_row + 2, label_col, "PROVEEDOR:", fmt_left_label)
    ws.merge_range(start_row + 2, val_c1, start_row + 2, val_c2, data["provider"], wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_BILLTO,
        "align": "left", "valign": "vcenter", "bg_color": LIGHT_GRAY,
        "text_wrap": True
    }))

    # SERVICIO (puede ser multi-line)
    ws.set_row(start_row + 3, 32)  # más alto para no “apretar”
    ws.write(start_row + 3, label_col, "SERVICIO:", fmt_left_label)
    ws.merge_range(start_row + 3, val_c1, start_row + 3, val_c2, data["service"], wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_BILLTO,
        "align": "left", "valign": "vcenter", "bg_color": WHITE,
        "text_wrap": True
    }))

    # PROYECTO
    ws.set_row(start_row + 4, 22)
    ws.write(start_row + 4, label_col, "PROYECTO:", fmt_left_label)
    ws.merge_range(start_row + 4, val_c1, start_row + 4, val_c2, data["project"], wb.add_format({
        "font_name": BASE_FONT, "font_size": FONT_BILLTO,
        "align": "left", "valign": "vcenter", "bg_color": LIGHT_GRAY,
        "text_wrap": True
    }))

    # -------------------- Bill to block --------------------
    # Col N..Z (13..25). “FACTURAR A:” left aligned.
    bill_col1, bill_col2 = 13, 25  # N..Z
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
    # Tabla empieza en fila 10 visual aprox
    table_header_row = start_row + 6  # deja un espacio
    ws.set_row(table_header_row, 28)

    # Column plan:
    # Concepto: B..N (1..13)  -> ancho grande
    # Costo unitario: O..S (14..18)
    # Unidades: T..U (19..20)
    # Subtotal: V..Z (21..25)
    c_concept_1, c_concept_2 = 1, 13
    c_cost_1, c_cost_2 = 14, 18
    c_units_1, c_units_2 = 19, 20
    c_subt_1, c_subt_2 = 21, 25

    ws.merge_range(table_header_row, c_concept_1, table_header_row, c_concept_2, "Concepto", fmt_table_hdr)
    ws.merge_range(table_header_row, c_cost_1, table_header_row, c_cost_2, "Costo unitario", fmt_table_hdr)
    ws.merge_range(table_header_row, c_units_1, table_header_row, c_units_2, "Unidades", fmt_table_hdr)
    ws.merge_range(table_header_row, c_subt_1, table_header_row, c_subt_2, "Subtotal", fmt_table_hdr)

    # body rows
    body_start = table_header_row + 1
    row_h = 40  # aire arriba/abajo (para “no apretado”)
    for i, it in enumerate(data["items"]):
        r = body_start + i
        ws.set_row(r, row_h)

        # Zebra: gris muy leve en filas alternas (solo concepto, como referencia)
        zebra = (i % 2 == 1)
        concept_fmt = wb.add_format(fmt_table_body_text.properties)
        if zebra:
            concept_fmt.set_bg_color(LIGHT_GRAY)

        ws.merge_range(r, c_concept_1, r, c_concept_2, it.concepto, concept_fmt)
        ws.merge_range(r, c_cost_1, r, c_cost_2, safe_float(it.costo_unitario), fmt_table_body_money)
        ws.merge_range(r, c_units_1, r, c_units_2, safe_float(it.unidades), fmt_table_body_num)
        ws.merge_range(r, c_subt_1, r, c_subt_2, safe_float(it.subtotal), fmt_table_body_money)

    last_item_row = body_start + max(len(data["items"]) - 1, 0)

    # -------------------- TOTAL --------------------
    # TOTAL en fila debajo de la tabla
    total_row = last_item_row + 2
    ws.set_row(total_row, 26)

    # “TOTAL:” en azul 10pt, número rojo 11pt, ancho suficiente para no ####
    # Label: T..V (19..21)
    # Amount: W..Z (22..25)
    ws.merge_range(total_row, 19, total_row, 21, "TOTAL:", fmt_total_label)
    ws.merge_range(total_row, 22, total_row, 25, data["total"], fmt_total_num)

    # -------------------- Print settings --------------------
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.set_margins(left=0.25, right=0.25, top=0.35, bottom=0.35)
    ws.fit_to_pages(1, 0)

    # Print area A1:Z(total_row+2)
    ws.print_area(range_a1(1, 1, total_row + 3, 26))

    wb.close()
    out.seek(0)
    return out.getvalue()


# ----------------------------
# FastAPI app
# ----------------------------

app = FastAPI(title="ODC Excel Generator", version="1.0.0")


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/generate-odc-excel")
def generate_odc_excel(payload: ODCPayload = Body(...)):
    # Validaciones mínimas (para no generar archivos basura)
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
