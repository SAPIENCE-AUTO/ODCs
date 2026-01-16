# main.py
import os
import re
import tempfile
from typing import List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import Response
from pydantic import BaseModel, Field

from openpyxl import load_workbook


# -----------------------------
# Config
# -----------------------------
TEMPLATE_PATH = os.getenv("ODC_TEMPLATE_PATH", "templates/odc_template.xlsx")
SHEET_NAME = os.getenv("ODC_SHEET_NAME", "ODC")  # cambia si tu hoja se llama distinto


# -----------------------------
# Helpers
# -----------------------------
def _safe_filename(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[^\w\-\.]+", "_", s, flags=re.UNICODE)
    return s[:120] or "odc"


def _as_money(value: float) -> float:
    # Mantén numérico; el formato ($, separadores, decimales) lo define el template
    return float(value)


def _as_int(value: int) -> int:
    return int(value)


# -----------------------------
# Request schema
# -----------------------------
class OdcItem(BaseModel):
    concepto: str
    costo_unitario: float
    unidades: int
    subtotal: Optional[float] = None  # si no lo mandas, se calcula


class OdcPayload(BaseModel):
    odc_num: str = Field(..., example="RI-02497")
    fecha: str = Field(..., example="08 ene 2026")

    proveedor: str
    servicio: str
    proyecto: str

    facturar_a: str
    rfc: str
    direccion: str

    items: List[OdcItem]


# -----------------------------
# App
# -----------------------------
app = FastAPI(title="ODC XLSX Generator", version="1.0")


@app.get("/")
def health():
    return {"ok": True, "service": "odc-xlsx-generator", "template": TEMPLATE_PATH}


@app.post("/generate-odc-xlsx")
def generate_odc_xlsx(payload: OdcPayload):
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(
            status_code=500,
            detail=f"No encuentro el template Excel en: {TEMPLATE_PATH}. Súbelo al repo (templates/odc_template.xlsx).",
        )

    # ---- calcula items + totales
    items = []
    subtotal_total = 0.0
    for it in payload.items:
        st = it.subtotal if it.subtotal is not None else float(it.costo_unitario) * int(it.unidades)
        subtotal_total += float(st)
        items.append(
            {
                "concepto": it.concepto,
                "costo_unitario": float(it.costo_unitario),
                "unidades": int(it.unidades),
                "subtotal": float(st),
            }
        )

    # ---- carga template
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    # =============================
    # MAPPING (AJUSTAR A TU TEMPLATE)
    # =============================
    # Pon aquí las celdas reales del template.
    # Si tu template ya está "idéntico" al diseño, esto es lo único que cambia.
    mapping = {
        "odc_num": "C8",
        "fecha": "C9",
        "proveedor": "C10",
        "servicio": "C11",
        "proyecto": "C12",
        "facturar_a": "H8",
        "rfc": "H10",
        "direccion": "H12",
    }

    # ---- escribe header
    ws[mapping["odc_num"]].value = payload.odc_num
    ws[mapping["fecha"]].value = payload.fecha
    ws[mapping["proveedor"]].value = payload.proveedor
    ws[mapping["servicio"]].value = payload.servicio
    ws[mapping["proyecto"]].value = payload.proyecto
    ws[mapping["facturar_a"]].value = payload.facturar_a
    ws[mapping["rfc"]].value = payload.rfc
    ws[mapping["direccion"]].value = payload.direccion

    # =============================
    # ITEMS TABLE (AJUSTAR A TU TEMPLATE)
    # =============================
    # Ejemplo: tabla empieza en fila 18
    start_row = 18

    # columnas (AJUSTA)
    col_concepto = "A"
    col_unit = "F"
    col_units = "H"
    col_sub = "J"

    # Limpia (por si el template tiene líneas dummy)
    # Ajusta el rango si tu tabla puede ser más grande/chica
    for r in range(start_row, start_row + 40):
        ws[f"{col_concepto}{r}"].value = None
        ws[f"{col_unit}{r}"].value = None
        ws[f"{col_units}{r}"].value = None
        ws[f"{col_sub}{r}"].value = None

    # Escribe filas
    for i, it in enumerate(items):
        r = start_row + i
        ws[f"{col_concepto}{r}"].value = it["concepto"]
        ws[f"{col_unit}{r}"].value = _as_money(it["costo_unitario"])
        ws[f"{col_units}{r}"].value = _as_int(it["unidades"])
        ws[f"{col_sub}{r}"].value = _as_money(it["subtotal"])

    # =============================
    # TOTALS (AJUSTAR A TU TEMPLATE)
    # =============================
    # Ejemplo: donde cae el subtotal/anticipo/total en tu layout
    totals_cells = {
        "subtotal": "J30",
        "anticipo": "J31",
        "total": "J32",
    }

    ws[totals_cells["subtotal"]].value = _as_money(subtotal_total)
    ws[totals_cells["anticipo"]].value = _as_money(0.0)
    ws[totals_cells["total"]].value = _as_money(subtotal_total)

    # ---- guarda y responde como XLSX
    with tempfile.TemporaryDirectory() as tmpdir:
        filename_base = _safe_filename(payload.odc_num)
        out_xlsx = os.path.join(tmpdir, f"ODC_{filename_base}.xlsx")
        wb.save(out_xlsx)
        xlsx_bytes = open(out_xlsx, "rb").read()

    filename = f"ODC_{_safe_filename(payload.odc_num)}.xlsx"
    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
