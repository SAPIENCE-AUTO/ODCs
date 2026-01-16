import os
import re
import json
import tempfile
import subprocess
from datetime import datetime
from typing import List, Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import Response, JSONResponse
from pydantic import BaseModel, Field

from openpyxl import load_workbook


# -----------------------------
# Config
# -----------------------------
TEMPLATE_PATH = os.getenv("ODC_TEMPLATE_PATH", "templates/odc_template.xlsx")
SHEET_NAME = os.getenv("ODC_SHEET_NAME", "ODC")  # cámbialo si tu sheet se llama distinto


# -----------------------------
# Helpers
# -----------------------------
def _safe_filename(s: str) -> str:
    s = s.strip()
    s = re.sub(r"[^\w\-\.]+", "_", s, flags=re.UNICODE)
    return s[:120] or "odc"


def _fmt_money(value: float) -> float:
    # En Excel el formato lo define el template; aquí solo dejamos numérico.
    return float(value)


def _excel_date(value: str) -> str:
    """
    Te acepto '08 enero 2026', '2026-01-08', etc.
    Si el template espera string, devolvemos string.
    Si el template espera date, cámbialo aquí para regresar datetime.date.
    """
    return value


def _convert_xlsx_to_pdf(input_xlsx: str, out_dir: str) -> str:
    """
    Convierte XLSX a PDF usando LibreOffice headless.
    Regresa la ruta al PDF.
    """
    # Render: lo normal es que soffice quede disponible si instalas libreoffice.
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        out_dir,
        input_xlsx,
    ]
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except FileNotFoundError:
        raise HTTPException(
            status_code=500,
            detail="LibreOffice (soffice) no está instalado en el runtime. Instálalo en Render.",
        )
    except subprocess.CalledProcessError as e:
        raise HTTPException(
            status_code=500,
            detail=f"Fallo al convertir a PDF con LibreOffice: {e.stderr.decode('utf-8', errors='ignore')}",
        )

    base = os.path.splitext(os.path.basename(input_xlsx))[0]
    pdf_path = os.path.join(out_dir, f"{base}.pdf")
    if not os.path.exists(pdf_path):
        raise HTTPException(status_code=500, detail="LibreOffice no generó el PDF esperado.")
    return pdf_path


# -----------------------------
# Request schema (JSON)
# -----------------------------
class OdcItem(BaseModel):
    concepto: str
    costo_unitario: float
    unidades: int
    subtotal: Optional[float] = None  # si no viene, lo calculamos


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

    # si luego quieres totales explícitos, se pueden aceptar también
    # subtotal: Optional[float] = None
    # anticipo: Optional[float] = 0
    # total: Optional[float] = None


# -----------------------------
# App
# -----------------------------
app = FastAPI(title="ODC Generator", version="2.0")


@app.get("/")
def health():
    return {"ok": True, "service": "odc-generator", "template": TEMPLATE_PATH}


@app.post("/generate-odc")
def generate_odc(payload: OdcPayload):
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(
            status_code=500,
            detail=f"No encuentro el template Excel en: {TEMPLATE_PATH}. Súbelo al repo (templates/odc_template.xlsx).",
        )

    # ---- preparar data
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

    # ---- abrir template y escribir celdas
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    # ⚠️ IMPORTANTE:
    # Estas celdas son EJEMPLO. Las tienes que ajustar a TU template real.
    # Lo correcto es que me digas (o tú fijes) el "mapping".
    mapping = {
        "odc_num": "C8",     # ejemplo
        "fecha": "C9",
        "proveedor": "C10",
        "servicio": "C11",
        "proyecto": "C12",

        "facturar_a": "H8",
        "rfc": "H10",
        "direccion": "H12",
    }

    ws[mapping["odc_num"]].value = payload.odc_num
    ws[mapping["fecha"]].value = _excel_date(payload.fecha)
    ws[mapping["proveedor"]].value = payload.proveedor
    ws[mapping["servicio"]].value = payload.servicio
    ws[mapping["proyecto"]].value = payload.proyecto

    ws[mapping["facturar_a"]].value = payload.facturar_a
    ws[mapping["rfc"]].value = payload.rfc
    ws[mapping["direccion"]].value = payload.direccion

    # ---- tabla de items
    # Ejemplo: empieza en fila 18
    start_row = 18
    col_concepto = "A"
    col_unit = "F"
    col_units = "H"
    col_sub = "J"

    # limpia filas (por si el template tiene datos dummy)
    for r in range(start_row, start_row + 30):
        ws[f"{col_concepto}{r}"].value = None
        ws[f"{col_unit}{r}"].value = None
        ws[f"{col_units}{r}"].value = None
        ws[f"{col_sub}{r}"].value = None

    for i, it in enumerate(items):
        r = start_row + i
        ws[f"{col_concepto}{r}"].value = it["concepto"]
        ws[f"{col_unit}{r}"].value = _fmt_money(it["costo_unitario"])
        ws[f"{col_units}{r}"].value = it["unidades"]
        ws[f"{col_sub}{r}"].value = _fmt_money(it["subtotal"])

    # ---- totales (ejemplo)
    ws["J30"].value = _fmt_money(subtotal_total)  # subtotal
    ws["J31"].value = _fmt_money(0.0)             # anticipo
    ws["J32"].value = _fmt_money(subtotal_total)  # total

    # ---- guardar a tmp y convertir a PDF
    with tempfile.TemporaryDirectory() as tmpdir:
        base = _safe_filename(payload.odc_num)
        out_xlsx = os.path.join(tmpdir, f"{base}.xlsx")
        wb.save(out_xlsx)

        pdf_path = _convert_xlsx_to_pdf(out_xlsx, tmpdir)
        pdf_bytes = open(pdf_path, "rb").read()

    filename = f"ODC_{_safe_filename(payload.odc_num)}.pdf"
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
