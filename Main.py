from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor, black, white
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from io import BytesIO


app = FastAPI()


# ---------- UTILIDADES ----------
PAGE_WIDTH, PAGE_HEIGHT = A4

SAP_BLUE = HexColor("#173344")
SAP_GRAY = HexColor("#F2F2F2")
SAP_RED = HexColor("#E53935")


def draw_text(c, text, x, y, size=10, bold=False, color=black):
    font = "Helvetica-Bold" if bold else "Helvetica"
    c.setFont(font, size)
    c.setFillColor(color)
    c.drawString(x, y, text)


# ---------- ENDPOINT ----------
@app.post("/generate-odc")
def generate_odc(payload: dict):

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    # =========================
    # PÁGINA 1
    # =========================

    # --- HEADER ---
    c.setFillColor(SAP_BLUE)
    c.rect(0, PAGE_HEIGHT - 3.5 * cm, PAGE_WIDTH, 3.5 * cm, fill=1)

    draw_text(c, "SAPIENCE", 2 * cm, PAGE_HEIGHT - 2 * cm, 24, bold=True, color=white)
    draw_text(c, "Human Insights Strategy", 2 * cm, PAGE_HEIGHT - 2.7 * cm, 10, color=white)

    # Caja ODC derecha
    box_x = PAGE_WIDTH - 7 * cm
    box_y = PAGE_HEIGHT - 2.8 * cm
    c.setFillColor(white)
    c.rect(box_x, box_y, 5.5 * cm, 1.2 * cm, fill=1)

    draw_text(c, "ODC #:", box_x + 0.4 * cm, box_y + 0.35 * cm, 12, bold=True)
    draw_text(c, "RI-XXXX", box_x + 2.2 * cm, box_y + 0.35 * cm, 12, bold=True, color=SAP_RED)

    # --- BLOQUE ADMINISTRATIVO ---
    start_y = PAGE_HEIGHT - 4.5 * cm
    row_h = 0.9 * cm
    left_x = 2 * cm
    mid_x = PAGE_WIDTH / 2 - 1 * cm

    labels = ["ODC #", "FECHA", "PROVEEDOR", "SERVICIO", "PROYECTO"]
    values = ["RI-XXXX", "DD MMM YYYY", "NOMBRE PROVEEDOR", "SERVICIO", "PROYECTO"]

    for i, (label, value) in enumerate(zip(labels, values)):
        y = start_y - i * row_h
        c.setFillColor(SAP_GRAY)
        c.rect(left_x, y, mid_x - left_x, row_h, fill=1)

        draw_text(c, f"{label}:", left_x + 0.2 * cm, y + 0.3 * cm, 10, bold=True, color=SAP_BLUE)
        draw_text(c, value, left_x + 3.2 * cm, y + 0.3 * cm, 10)

    # --- FACTURAR A ---
    draw_text(c, "FACTURAR A:", mid_x + 1 * cm, start_y + 0.3 * cm, 14, bold=True, color=SAP_BLUE)
    draw_text(c, "ASESORES GLOBALES CORPORATIVOS", mid_x + 1 * cm, start_y - 0.6 * cm, 11, bold=True)
    draw_text(c, "RFC: XXXXXXXX", mid_x + 1 * cm, start_y - 1.4 * cm, 10, bold=True)
    draw_text(
        c,
        "Dirección completa de facturación\nCiudad, CP, País",
        mid_x + 1 * cm,
        start_y - 2.3 * cm,
        10,
    )

    # --- TABLA CONCEPTOS ---
    table_x = 3 * cm
    table_y = start_y - 6 * cm
    table_w = PAGE_WIDTH - 6 * cm
    row_h = 0.9 * cm

    cols = [0.6, 0.15, 0.1, 0.15]
    col_x = [table_x]
    for w in cols[:-1]:
        col_x.append(col_x[-1] + table_w * w)

    headers = ["Concepto", "Precio unitario", "Unidades", "Precio Total"]

    c.setFillColor(SAP_BLUE)
    c.rect(table_x, table_y, table_w, row_h, fill=1)

    for i, h in enumerate(headers):
        draw_text(c, h, col_x[i] + 0.3 * cm, table_y + 0.3 * cm, 10, bold=True, color=white)

    # Filas dummy
    for r in range(4):
        y = table_y - (r + 1) * row_h
        c.setFillColor(SAP_GRAY if r % 2 == 0 else white)
        c.rect(table_x, y, table_w, row_h, fill=1)

        draw_text(c, "Concepto de ejemplo", col_x[0] + 0.3 * cm, y + 0.3 * cm)
        draw_text(c, "$0.00", col_x[1] + 0.3 * cm, y + 0.3 * cm)
        draw_text(c, "0", col_x[2] + 0.3 * cm, y + 0.3 * cm)
        draw_text(c, "$0.00", col_x[3] + 0.3 * cm, y + 0.3 * cm)

    # --- TOTALES ---
    totals_x = table_x + table_w - 7 * cm
    totals_y = table_y - 6 * row_h

    labels = ["Subtotal", "Anticipo", "Total"]
    values = ["$0.00", "-$0.00", "$0.00"]

    for i, (l, v) in enumerate(zip(labels, values)):
        y = totals_y - i * row_h
        c.setFillColor(SAP_BLUE)
        c.rect(totals_x, y, 3.5 * cm, row_h, fill=1)
        c.setFillColor(white)
        c.rect(totals_x + 3.5 * cm, y, 3.5 * cm, row_h, fill=1)

        draw_text(c, l, totals_x + 0.3 * cm, y + 0.3 * cm, 10, bold=True, color=white)
        draw_text(c, v, totals_x + 3.7 * cm, y + 0.3 * cm, 10, bold=True)

    c.showPage()

    # =========================
    # PÁGINA 2 – CONDICIONES
    # =========================

    draw_text(
        c,
        "NOTAS Y CONDICIONES DE LA ORDEN DE COMPRA",
        2 * cm,
        PAGE_HEIGHT - 2.5 * cm,
        14,
        bold=True,
        color=SAP_BLUE,
    )

    text = c.beginText(2 * cm, PAGE_HEIGHT - 4 * cm)
    text.setFont("Helvetica", 9)

    dummy_conditions = [
        "Emisión y entrega de factura",
        "Revisión, validación y aceptación",
        "Tiempos de pago",
        "Condiciones de servicio",
        "Cambio de alcance, cantidades o precios",
        "Cancelaciones y reprogramaciones",
        "Confidencialidad y manejo de información",
        "Protección de datos personales",
        "Comunicación y soporte",
        "Cierre administrativo del proyecto",
    ]

    for section in dummy_conditions:
        text.setFont("Helvetica-Bold", 9)
        text.textLine(section)
        text.setFont("Helvetica", 9)
        text.textLine("– Texto legal de ejemplo.")
        text.textLine(" ")
    c.drawText(text)

    c.save()
    buffer.seek(0)

    return StreamingResponse(buffer, media_type="application/pdf")
