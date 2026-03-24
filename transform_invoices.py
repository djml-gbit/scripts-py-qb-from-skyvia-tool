"""
Transformación de facturas QuickBooks Online → Zoho Books (formato importación)
- Expande bundles (GroupLineDetail) en sus componentes individuales
- Mantiene items normales (SalesItemLineDetail) tal cual
- Maneja descuentos a nivel de entidad (DiscountLineDetail)
- Ignora SubTotalLineDetail
"""

import csv
import json
from datetime import datetime
import openpyxl
from openpyxl import Workbook

# ─── CONFIGURACIÓN ───────────────────────────────────────────────────────────
INPUT_CSV   = "Facturas_Q4_2019.csv"
OUTPUT_XLSX = "Facturas_Q4_2019_zoho.xlsx"
LOTE        = "Q4_2019"   # Valor fijo para identificar lote de importación
TAX_MAP     = {"9": "ITBMS DGI", "": ""}   # ID QB → nombre en Zoho

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def clean_item_name(name):
    """Remueve prefijo de categoría QB: 'Categoria:Nombre' → 'Nombre'"""
    if ":" in name:
        return name.split(":", 1)[1].strip()
    return name.strip()

def format_date(date_str):
    """Convierte 'M/D/YYYY' → 'DD/MM/YYYY' para Zoho Books"""
    if not date_str:
        return ""
    try:
        return datetime.strptime(date_str.strip(), "%m/%d/%Y").strftime("%d/%m/%Y")
    except:
        return date_str

def invoice_status(linked_txn):
    """Si tiene pagos vinculados → Sent, si no → Draft (ajustar según necesidad)"""
    try:
        txns = json.loads(linked_txn) if linked_txn else []
        return "Sent" if txns else "Sent"  # todas van como Sent
    except:
        return "Sent"

# ─── PASO 1: Agrupar filas del CSV por número de factura ─────────────────────

invoices = {}   # doc_number → { header: {...}, items: [...], discount: {...} }

with open(INPUT_CSV, encoding="utf-8") as f:
    reader = csv.DictReader(f, delimiter=";")
    for row in reader:
        doc = row["DocNumber"].strip()
        if not doc:
            continue

        if doc not in invoices:
            invoices[doc] = {"header": None, "items": [], "discount": None}

        detail = row["DetailType"].strip()

        if detail == "SubTotalLineDetail":
            # Cabecera de la factura (fecha, cliente, due date, etc.)
            invoices[doc]["header"] = row

        elif detail == "SalesItemLineDetail":
            invoices[doc]["items"].append({
                "type": "normal",
                "name": clean_item_name(row["SalesItemLineDetail_ItemRefName"]),
                "description": row["Description"],
                "qty": row["SalesItemLineDetail_Qty"] or "1",
                "price": row["SalesItemLineDetail_UnitPrice"] or "0",
                "tax_code": TAX_MAP.get(row["SalesItemLineDetail_TaxCodeRefId"].strip(), ""),
                "tax_pct": "7" if row["SalesItemLineDetail_TaxCodeRefId"].strip() == "9" else "",
            })

        elif detail == "GroupLineDetail":
            # Expandir componentes del bundle
            try:
                components = json.loads(row["GroupLineDetail_Line"])
                for comp in components:
                    tax_id = str(comp.get("SalesItemLineDetail_TaxCodeRefId", "")).strip()
                    invoices[doc]["items"].append({
                        "type": "bundle_component",
                        "name": clean_item_name(comp.get("SalesItemLineDetail_ItemRefName", "")),
                        "description": comp.get("Description", ""),
                        "qty": str(comp.get("SalesItemLineDetail_Qty", 1)),
                        "price": str(comp.get("SalesItemLineDetail_UnitPrice", 0)),
                        "tax_code": TAX_MAP.get(tax_id, ""),
                        "tax_pct": "7" if tax_id == "9" else "",
                    })
            except Exception as e:
                print(f"  ⚠️  Error parseando bundle en factura {doc}: {e}")

        elif detail == "DiscountLineDetail":
            invoices[doc]["discount"] = {
                "percent_based": row["DiscountLineDetail_PercentBased"].strip(),
                "percent": row["DiscountLineDetail_DiscountPercent"].strip(),
                "amount": row["Amount"].strip(),
            }

# ─── PASO 2: Generar XLSX de Zoho Books ──────────────────────────────────────

wb = Workbook()
ws = wb.active
ws.title = "Invoices"

HEADERS = [
    "Invoice Date", "Invoice Number", "Invoice Status", "Customer Name",
    "Due Date", "Item Name", "Item Desc", "Quantity", "Item Price",
    "Discount", "Discount Amount", "Discount Type", "Is Discount Before Tax",
    "Item Tax", "Item Tax %", "Lote Import"
]
ws.append(HEADERS)

stats = {"total": 0, "bundles_expanded": 0, "normal": 0, "no_header": 0, "no_items": 0}

for doc, inv in sorted(invoices.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 0):
    header = inv["header"]
    items  = inv["items"]
    disc   = inv["discount"]

    if not header:
        print(f"  ⚠️  Factura {doc}: sin cabecera, omitida")
        stats["no_header"] += 1
        continue

    if not items:
        print(f"  ⚠️  Factura {doc}: sin ítems, omitida")
        stats["no_items"] += 1
        continue

    invoice_date = format_date(header["TxnDate"])
    due_date     = format_date(header["DueDate"])
    customer     = header["CustomerRefName"].strip()
    status       = invoice_status(header["LinkedTxn1"])

    # Descuento a nivel de entidad (aplica igual a todas las líneas)
    disc_value      = ""
    disc_amount     = ""
    disc_type       = ""
    disc_before_tax = ""

    if disc:
        disc_type       = "entity_level"
        disc_before_tax = "True"
        if disc["percent_based"].lower() == "true":
            disc_value = disc["percent"]   # % aplicado en columna Discount
        else:
            disc_amount = disc["amount"]   # monto fijo en columna Discount Amount

    for item in items:
        if item["type"] == "bundle_component":
            stats["bundles_expanded"] += 1
        else:
            stats["normal"] += 1

        ws.append([
            invoice_date,
            doc,
            status,
            customer,
            due_date,
            item["name"],
            item["description"],
            item["qty"],
            item["price"],
            disc_value,
            disc_amount,
            disc_type,
            disc_before_tax,
            item["tax_code"],
            item["tax_pct"],
            LOTE,
        ])

    stats["total"] += 1

wb.save(OUTPUT_XLSX)

print(f"\n✅ Archivo generado: {OUTPUT_XLSX}")
print(f"   Facturas procesadas : {stats['total']}")
print(f"   Ítems normales      : {stats['normal']}")
print(f"   Componentes bundle  : {stats['bundles_expanded']}")
print(f"   Sin cabecera (skip) : {stats['no_header']}")
print(f"   Sin ítems (skip)    : {stats['no_items']}")
