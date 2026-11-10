"""Reglas de facturas domiciliadas frente a documentos de pago."""

from collections import defaultdict
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

from ..common import (COL_VIA_PAGO, add_summary, build_rows, cents, date_distance,
                      detect_system, ensure_output_columns, header_map, write_matches)


def find_subset_for_target(rows, candidates, target_abs, tolerance, max_items):
    """Implementación heredada: conserva la primera combinación determinista."""
    states = {(0, 0): ()}; limit = target_abs + tolerance
    for row in candidates:
        value = abs(cents(rows[row]["amount"]))
        if value == 0 or value > limit: continue
        for (count, total), combo in list(states.items()):
            if count >= max_items or total + value > limit: continue
            states.setdefault((count + 1, total + value), combo + (row,))
        for count in range(2, max_items + 1):
            for delta in range(tolerance + 1):
                totals = {target_abs - delta, target_abs + delta} if delta else {target_abs}
                for total in totals:
                    if (count, total) in states: return list(states[(count, total)])
    return None


def process(input_path: Path, output_path: Path, payment_class: str | None = None, tolerance_cents: int = 3, max_invoices: int = 4) -> dict:
    workbook = load_workbook(input_path); ws = workbook[workbook.sheetnames[0]]; headers = header_map(ws)
    system = detect_system(headers)
    if COL_VIA_PAGO not in headers: raise ValueError("El modo domiciliados requiere la columna 'Vía de pago'.")
    headers = ensure_output_columns(ws, headers, "domiciliados"); rows = build_rows(ws, headers, system, "domiciliados")
    auto_class = "TU" if system == "GPC" else "TP"
    payment_class = (payment_class or auto_class).strip().upper()
    source = "parametro" if payment_class != auto_class else ("regla confirmada GPC" if system == "GPC" else "SUPUESTO DE PRUEBA VIP: confirmar TP con BPS")
    assigned = {}; match_type = {}; reason = {}; difference = {}; sequence = 1
    def assign(group, kind, why):
        nonlocal sequence
        code = f"CMP-DOM-{sequence:05d}"; sequence += 1; diff = round(sum(rows[row]["amount"] for row in group), 2)
        for row in group: assigned[row] = code; match_type[row] = kind; reason[row] = why; difference[row] = diff
    candidates = [row for row, item in rows.items() if not item["cleared"]]; buckets = defaultdict(list)
    for row in candidates: buckets[(rows[row]["soc"], rows[row]["acc"])].append(row)
    for group in buckets.values():
        payments = [row for row in group if rows[row]["class"] == payment_class]
        invoices = [row for row in group if rows[row]["class"] != payment_class and rows[row]["payment_method"]]
        for payment in sorted(payments, key=lambda row: (rows[row]["date"] or date.min, row)):
            if payment in assigned: continue
            options = []
            for invoice in invoices:
                if invoice in assigned: continue
                difference_cents = abs(cents(rows[payment]["amount"]) + cents(rows[invoice]["amount"]))
                if cents(rows[payment]["amount"]) * cents(rows[invoice]["amount"]) < 0 and difference_cents <= tolerance_cents:
                    options.append((difference_cents, date_distance(rows[payment]["date"], rows[invoice]["date"]), invoice))
            if options:
                _, _, invoice = min(options); assign([payment, invoice], "DOM01_1A1", f"{payment_class} contra factura con vía de pago; misma sociedad y acreedor; diferencia <= {tolerance_cents} céntimos")
    for group in buckets.values():
        payments = [row for row in group if rows[row]["class"] == payment_class and row not in assigned]
        invoices = [row for row in group if rows[row]["class"] != payment_class and rows[row]["payment_method"] and row not in assigned]
        for payment in payments:
            opposite = [row for row in invoices if row not in assigned and cents(rows[payment]["amount"]) * cents(rows[row]["amount"]) < 0]
            opposite.sort(key=lambda row: (date_distance(rows[payment]["date"], rows[row]["date"]), row))
            combo = find_subset_for_target(rows, opposite, abs(cents(rows[payment]["amount"])), tolerance_cents, max_invoices)
            if combo: assign([payment] + combo, "DOM02_1AN", f"Un {payment_class} contra {len(combo)} facturas con vía de pago; misma sociedad y acreedor; diferencia <= {tolerance_cents} céntimos")
    write_matches(ws, headers, rows, assigned, match_type, reason, difference)
    stats = add_summary(workbook, "domiciliados", system, rows, assigned, match_type, [("Clase de pago utilizada", payment_class), ("Origen clase de pago", source), ("Tolerancia", f"{tolerance_cents} céntimos"), ("Máximo facturas por pago en búsqueda 1:N", max_invoices)])
    stats.update(payment_class=payment_class, payment_class_source=source); workbook.save(output_path); workbook.close(); return stats
