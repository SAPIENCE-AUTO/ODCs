from fastapi import FastAPI
from fastapi.responses import Response
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from io import BytesIO
from datetime import datetime

app = FastAPI(title="ODC PDF Generator")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post(
    "/generate-odc",
    responses={
        200: {
            "content": {"application/pdf": {}},
            "description": "ODC PDF generated"
        }
    },
)
def generate_odc(payload: dict):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # -------------------------
    # Header
    # -------------------------
    c.setFont("Helvetica-Bold", 16)
    c.drawString(30 * mm, height - 25 * mm, "ORDEN DE COMPRA")

    c.setFont("Helvetica", 10)
    c.drawRightString(
        width - 30 * mm,
        height - 25 * mm,
        f"ODC: {payload.get('odc_number', '')}"
    )

    c.drawRightString(
        width - 30 * mm,
        height - 32 * mm,
        f"Fecha: {payload.get('date', '')}"
    )

    # -------------------------
    # Provider & Project Info
    # -------------------------
    y = height - 45 * mm

    c.setFont("Helvetica-Bold", 10)
    c.drawString(30 * mm, y, "Proveedor:")
    c.setFont("Helvetica", 10)
    c.drawString(30 * mm, y - 5 * mm, payload.get("provider", ""))

    c.setFont("Helvetica-Bold", 10)
    c.drawString(110 * mm, y, "Proyecto:")
    c.setFont("Helvetica", 10)
    c.drawString(110 * mm, y - 5 * mm, payload.get("project", ""))

    c.setFont("Helvetica-Bold", 10)
    c.drawString(110 * mm, y - 12 * mm, "Servicio:")
    c.setFont("Helvetica", 10)
    c.drawString(110 * mm, y - 17 * mm, payload.get("service", ""))

    # -------------------------
    # Bill To
    # -------------------------
    bill_to = payload.get("bill_to", {})

    y -= 30 * mm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(30 * mm, y, "Facturar a:")

    c.setFont("Helvetica", 10)
    c.drawString(30 * mm, y - 5 * mm, bill_to.get("company", ""))
    c.drawString(30 * mm, y - 10 * mm, f"RFC: {bill_to.get('rfc', '')}")
    c.drawString(30 * mm, y - 15 * mm, bill_to.get("address", ""))

    # -------------------------
    # Items Table
    # -------------------------
    y -= 30 * mm
    c.setFont("Helvetica-Bold", 10)
    c.drawString(30 * mm, y, "Concepto")
    c.drawRightString(140 * mm, y, "Cantidad")
    c.drawRightString(165 * mm, y, "Precio")
    c.drawRightString(190 * mm, y, "Importe")

    c.line(30 * mm, y - 2 * mm, 190 * mm, y - 2 * mm)

    y -= 8 * mm
    c.setFont("Helvetica", 10)

    subtotal = 0
    for item in payload.get("items", []):
        concept = item.get("concept", "")
        units = item.get("units", 0)
        unit_price = item.get("unit_price", 0)
        amount = units * unit_price
        subtotal += amount

        c.drawString(30 * mm, y, concept[:60])
        c.drawRightString(140 * mm, y, str(units))
        c.drawRightString(165 * mm, y, f"${unit_price:,.2f}")
        c.drawRightString(190 * mm, y, f"${amount:,.2f}")

        y -= 7 * mm
        if y < 40 * mm:
            c.showPage()
            y = height - 40 * mm
            c.setFont("Helvetica", 10)

    # -------------------------
    # Totals
    # -------------------------
    advance = payload.get("advance", 0)
    total = payload.get("total", subtotal - advance)

    y -= 10 * mm
    c.setFont("Helvetica-Bold", 10)
    c.drawRightString(165 * mm, y, "Subtotal:")
    c.drawRightString(190 * mm, y, f"${subtotal:,.2f}")

    y -= 6 * mm
    c.drawRightString(165 * mm, y, "Anticipo:")
    c.drawRightString(190 * mm, y, f"${advance:,.2f}")

    y -= 6 * mm
    c.drawRightString(165 * mm, y, "Total:")
    c.drawRightString(190 * mm, y, f"${total:,.2f}")

    # -------------------------
    # Footer
    # -------------------------
    c.setFont("Helvetica", 8)
    c.drawString(
        30 * mm,
        20 * mm,
        "Esta Orden de Compra constituye el acuerdo formal para la prestación del servicio descrito."
    )

    c.save()

    pdf_bytes = buffer.getvalue()
    buffer.close()

    filename = f"ODC_{payload.get('odc_number', 'ODC')}.pdf"

    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={
            "Content-Disposition": f'attachment; filename=\"{filename}\"',
            "Cache-Control": "no-store",
        },
    )
