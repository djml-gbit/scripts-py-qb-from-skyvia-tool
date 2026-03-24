"""
Transformación de pagos QuickBooks Online → Zoho Books (formato importación)
- Cruza TxnId del pago con InvoiceId del crudo de facturas QB para obtener DocNumber
- Solo incluye pagos cuya factura existe en los CSVs de facturas QB proporcionados
- Pagos sin factura encontrada se loguean en un archivo separado para revisión
"""

import csv
import json
from datetime import datetime
from openpyxl import Workbook

# ─── CONFIGURACIÓN ───────────────────────────────────────────────────────────
# Agregar o quitar archivos según los trimestres disponibles
FACTURAS_CSVS = [
     "Facturas_Q1_2025.csv",
"Facturas_Q2_2025.csv",
"Facturas_Q3_2025.csv",
"Facturas_Oct_Nov_2025.csv",
"Facturas_Q1_2024.csv",
"Facturas_Q2_2024.csv",
"Facturas_Q3_2024.csv",
"Facturas_Q4_2024.csv",
"Facturas_Q1_2023.csv",
"Facturas_Q2_2023.csv",
"Facturas_Q3_2023.csv",
"Facturas_Q4_2023.csv",
"Facturas_Q1_Q2_2022.csv",
"Facturas_Q3_Q4_2022.csv",
"Facturas_Q1_2021.csv",
"Facturas_Q2_2021.csv",
"Facturas_Q3_2021.csv",
"Facturas_Q4_2021.csv",
"Facturas_Q1_2020.csv",
"Facturas_Q2_2020.csv",
"Facturas_Q3_2020.csv",
"Facturas_Q4_2020.csv",
"Facturas_Q1_2019.csv",
"Facturas_Q2_2019.csv",
"Facturas_Q3_2019.csv",
"Facturas_Q4_2019.csv"
]

PAGOS_CSVS = [
    "Pagos_Q1_2025.csv",
"Pagos_Q2_Q3_2025.csv",
"Pagos_Oct_Nov_2025.csv",
"Pagos_Q1_Q2_2024.csv",
"Pagos_Q3_Q4_2024.csv",
"Pagos_Q1_Q2_2023.csv",
"Pagos_Q3_Q4_2023.csv",
"Pagos_Q1_Q2_2022.csv",
"Pagos_Q3_Q4_2022.csv",
"Pagos_Q1_Q2_2021.csv",
"Pagos_Q3_Q4_2021.csv",
"Pagos_Q1_Q2_2020.csv",
"Pagos_Q3_Q4_2020.csv",
"Pagos_Q1_Q2_2019.csv",
"Pagos_Q3_Q4_2019.csv",
"PaymentLineItem_Payment_Export_2026-03-10_13-54-28"
]

OUTPUT_XLSX    = "Pagos_Zoho.xlsx"
OUTPUT_SKIPPED = "Pagos_sin_factura.csv"   # pagos sin factura encontrada
CUENTA         = "Cuenta GRUPO ECO WASTE S.A."
NOTAS          = "Migrado desde QuickBooks"

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def format_date(date_str):
    """Convierte 'M/D/YYYY' → 'DD/MM/YYYY' para Zoho Books"""
    if not date_str:
        return ""
    for fmt in ("%m/%d/%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(date_str.strip(), fmt).strftime("%d/%m/%Y")
        except:
            continue
    return date_str.strip()

# ─── PASO 1: Construir mapa InvoiceId (QB) → DocNumber ───────────────────────

print("Cargando facturas QB...")
invoice_map = {}   # txn_id (str) → doc_number (str)

for fname in FACTURAS_CSVS:
    try:
        with open(fname, encoding="utf-8") as f:
            reader = csv.DictReader(f, delimiter=";")
            for row in reader:
                inv_id = row["InvoiceId"].strip()
                doc    = row["DocNumber"].strip()
                if inv_id and doc:
                    invoice_map[inv_id] = doc
        print(f"  ✅ {fname}: {len(invoice_map)} facturas acumuladas")
    except FileNotFoundError:
        print(f"  ⚠️  {fname}: archivo no encontrado, omitido")

print(f"Total facturas en mapa: {len(invoice_map)}")

# ─── PASO 2: Procesar pagos ───────────────────────────────────────────────────

print("\nProcesando pagos...")

wb = Workbook()
ws = wb.active
ws.title = "Payments"

HEADERS = [
    "Payment Date", "Payment Number", "Customer Name",
    "Invoice Number", "Amount", "Importe factura",
    "cuenta", "sufijo", "Notes"
]
ws.append(HEADERS)

skipped = []
sufijo  = 1
stats   = {"incluidos": 0, "omitidos": 0, "pagos_procesados": 0}

for fname in PAGOS_CSVS:
    try:
        with open(fname, encoding="utf-8") as f:
            rows = list(csv.DictReader(f, delimiter=";"))

        for row in rows:
            stats["pagos_procesados"] += 1
            payment_date = format_date(row["PaymentDate"])
            customer     = row["CustomerRefName"].strip()
            amount       = row["AmountApplied"].strip()

            # Saltar pagos de $0
            try:
                if float(amount) == 0:
                    stats["omitidos"] += 1
                    continue
            except:
                pass

            try:
                txns = json.loads(row["LinkedTxn"])
            except:
                txns = []

            found = False
            for txn in txns:
                txn_id = str(txn.get("TxnId", "")).strip()
                doc_number = invoice_map.get(txn_id)

                if doc_number:
                    ws.append([
                        payment_date,
                        "",           # Payment Number — se genera automático en Zoho
                        customer,
                        doc_number,
                        amount,
                        amount,       # Importe factura = mismo monto
                        CUENTA,
                        sufijo,
                        NOTAS,
                    ])
                    sufijo += 1
                    stats["incluidos"] += 1
                    found = True
                else:
                    skipped.append({
                        "PaymentId":    row["PaymentId"],
                        "PaymentDate":  row["PaymentDate"],
                        "Customer":     customer,
                        "Amount":       amount,
                        "TxnId_QB":     txn_id,
                        "Archivo":      fname,
                    })
                    stats["omitidos"] += 1

        print(f"  ✅ {fname}: {len(rows)} pagos procesados")

    except FileNotFoundError:
        print(f"  ⚠️  {fname}: archivo no encontrado, omitido")

wb.save(OUTPUT_XLSX)

# ─── PASO 3: Guardar pagos omitidos ──────────────────────────────────────────

if skipped:
    with open(OUTPUT_SKIPPED, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=skipped[0].keys())
        writer.writeheader()
        writer.writerows(skipped)

# ─── RESUMEN ─────────────────────────────────────────────────────────────────

print(f"\n{'='*50}")
print(f"✅ Archivo generado  : {OUTPUT_XLSX}")
print(f"   Pagos procesados  : {stats['pagos_procesados']}")
print(f"   Incluidos         : {stats['incluidos']}")
print(f"   Sin factura (skip): {stats['omitidos']}")
if skipped:
    print(f"   Log omitidos      : {OUTPUT_SKIPPED}")
print(f"{'='*50}")