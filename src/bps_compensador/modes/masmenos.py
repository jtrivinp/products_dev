"""Regla 1:1 de factura frente a anulación/abono autorizados."""

from collections import defaultdict
from pathlib import Path
from openpyxl import load_workbook
from ..common import add_summary, build_rows, cents, date_distance, detect_system, ensure_output_columns, header_map, load_class_table, write_matches


def process(input_path: Path, classes_path: Path, output_path: Path) -> dict:
    allowed = load_class_table(classes_path); workbook = load_workbook(input_path); ws = workbook[workbook.sheetnames[0]]; headers = ensure_output_columns(ws, header_map(ws), "masmenos")
    system = detect_system(headers); rows = build_rows(ws, headers, system, "masmenos"); assigned = {}; match_type = {}; reason = {}; difference = {}; sequence = 1
    def assign(group):
        nonlocal sequence
        code = f"CMP-MM-{sequence:05d}"; sequence += 1; diff = round(sum(rows[row]["amount"] for row in group), 2)
        for row in group: assigned[row] = code; match_type[row] = "MM01_1A1_EXACTO"; reason[row] = "Misma sociedad y acreedor/cuenta; clases autorizadas; importes exactamente opuestos"; difference[row] = diff
    buckets = defaultdict(lambda: defaultdict(list))
    for row, item in rows.items():
        if not item["cleared"] and item["class"] in allowed and cents(item["amount"]) != 0: buckets[(item["soc"], item["acc"])][cents(item["amount"])].append(row)
    for amounts in buckets.values():
        for amount in sorted(value for value in amounts if value > 0):
            if -amount not in amounts: continue
            positive = [row for row in amounts[amount] if row not in assigned]; negative = [row for row in amounts[-amount] if row not in assigned]
            while positive and negative:
                _, first, second = min((date_distance(rows[first]["date"], rows[second]["date"]), first, second) for first in positive for second in negative)
                assign([first, second]); positive.remove(first); negative.remove(second)
    write_matches(ws, headers, rows, assigned, match_type, reason, difference)
    stats = add_summary(workbook, "masmenos", system, rows, assigned, match_type, [("Clases autorizadas utilizadas", ", ".join(sorted(allowed))), ("Nota", "Lista de prueba/parametrizable; sustituir por listado oficial de BPS para producción")])
    stats["allowed_classes"] = sorted(allowed); workbook.save(output_path); return stats
