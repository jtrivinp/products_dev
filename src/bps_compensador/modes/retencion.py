"""Modo Retención (antes denominado ``abiertas`` en el prototipo)."""

from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

from ..common import (COL_ASIGNACION, COL_BASE_RET, COL_DOCUMENTO, COL_FECHA_CONT,
                      COL_IND_RET, OUT_IMP_AJUST, OUT_IMP_RET, OUT_MOTIVO, OUT_PCT_RET,
                      account_column, add_summary, as_float, build_rows, cents,
                      date_distance, detect_system, ensure_output_columns, header_map,
                      load_retention_table, norm_text, write_matches)


def process(input_path: Path, retention_path: Path, output_path: Path, day_window: int = 31, tolerance_cents: int = 12) -> dict:
    rates = load_retention_table(retention_path)
    workbook = load_workbook(input_path)
    ws = workbook[workbook.sheetnames[0]]
    headers = header_map(ws)
    system = detect_system(headers)
    required = ["Sociedad", account_column(system), COL_ASIGNACION, COL_DOCUMENTO, COL_FECHA_CONT, COL_BASE_RET, COL_IND_RET]
    missing = [name for name in required if name not in headers]
    if missing:
        raise ValueError(f"Faltan columnas obligatorias para abiertas: {missing}")
    headers = ensure_output_columns(ws, headers, "abiertas")
    rows = build_rows(ws, headers, system, "abiertas")
    unknown_codes = set()
    for row, item in rows.items():
        base = as_float(ws.cell(row, headers[COL_BASE_RET]).value)
        code = norm_text(ws.cell(row, headers[COL_IND_RET]).value).upper()
        rate = rates.get(code) if code else None
        if code and rate is None:
            unknown_codes.add(code)
        retention_amount = base * rate if rate is not None else 0.0
        adjusted = round(item["amount"] - retention_amount, 2) if rate is not None else item["amount"]
        item.update({"effective": adjusted, "ret_code": code, "has_retention": bool(code and rate is not None and abs(base) > 1e-9)})
        ws.cell(row, headers[OUT_PCT_RET]).value = rate
        ws.cell(row, headers[OUT_PCT_RET]).number_format = "0.00%"
        ws.cell(row, headers[OUT_IMP_RET]).value = round(retention_amount, 2) if rate is not None else None
        ws.cell(row, headers[OUT_IMP_AJUST]).value = adjusted

    assigned = {}; match_type = {}; reason = {}; difference = {}; sequence = 1
    def assign(group, kind, why, prefix="CMP"):
        nonlocal sequence
        code = f"{prefix}-{sequence:05d}"; sequence += 1
        diff = round(sum(rows[row]["effective"] for row in group), 2)
        for row in group:
            assigned[row] = code; match_type[row] = kind; reason[row] = why; difference[row] = diff

    candidates = [row for row, item in rows.items() if not item["cleared"]]
    groups = defaultdict(list)
    for row in candidates:
        if rows[row]["assignment"]:
            groups[(rows[row]["soc"], rows[row]["acc"], rows[row]["assignment"])].append(row)
    for group in groups.values():
        if len(group) >= 2 and sum(cents(rows[row]["effective"]) for row in group) == 0:
            assign(group, "R01_GRUPO_EXACTO", "Misma sociedad, acreedor y asignacion; suma efectiva = 0")

    buckets = defaultdict(list)
    for row in candidates:
        if row not in assigned: buckets[(rows[row]["soc"], rows[row]["acc"])].append(row)
    for group in buckets.values():
        by_amount = defaultdict(list)
        for row in group: by_amount[cents(rows[row]["effective"])].append(row)
        for amount in sorted(value for value in by_amount if value > 0):
            if -amount not in by_amount: continue
            positive = [row for row in by_amount[amount] if row not in assigned]
            negative = [row for row in by_amount[-amount] if row not in assigned]
            while positive and negative:
                _, first, second = min((date_distance(rows[first]["date"], rows[second]["date"]), first, second) for first in positive for second in negative)
                retained = rows[first]["has_retention"] or rows[second]["has_retention"]
                assign([first, second], "R02_1A1_RETENCION" if retained else "R02_1A1_EXACTO", "Importes efectivos opuestos" + ("; incluye ajuste por retencion parametrizada" if retained else ""), "CMP-RET" if retained else "CMP")
                positive.remove(first); negative.remove(second)

    # Casos 1:1 dentro de la tolerancia funcional confirmada.
    if tolerance_cents:
        buckets = defaultdict(list)
        for row in candidates:
            if row not in assigned: buckets[(rows[row]["soc"], rows[row]["acc"])].append(row)
        for group in buckets.values():
            positive = [row for row in group if cents(rows[row]["effective"]) > 0]
            negative = [row for row in group if cents(rows[row]["effective"]) < 0]
            while positive and negative:
                options = [(abs(cents(rows[p]["effective"]) + cents(rows[n]["effective"])), date_distance(rows[p]["date"], rows[n]["date"]), p, n) for p in positive for n in negative if abs(cents(rows[p]["effective"]) + cents(rows[n]["effective"])) <= tolerance_cents]
                if not options: break
                _, _, first, second = min(options)
                retained = rows[first]["has_retention"] or rows[second]["has_retention"]
                assign([first, second], "R02_1A1_RETENCION" if retained else "R02_1A1_EXACTO", "Importes efectivos opuestos dentro de tolerancia", "CMP-RET" if retained else "CMP")
                positive.remove(first); negative.remove(second)

    buckets = defaultdict(list)
    for row in candidates:
        if row not in assigned: buckets[(rows[row]["soc"], rows[row]["acc"])].append(row)
    for group in buckets.values():
        for target in list(group):
            if target in assigned or cents(rows[target]["effective"]) == 0: continue
            others = [row for row in group if row != target and row not in assigned and cents(rows[row]["effective"]) * cents(rows[target]["effective"]) < 0 and date_distance(rows[row]["date"], rows[target]["date"]) <= day_window]
            seen = {}; wanted = -cents(rows[target]["effective"]); found = None
            for row in others:
                value = cents(rows[row]["effective"]); complement = wanted - value
                if complement in seen and seen[complement] != row:
                    found = [target, seen[complement], row]; break
                seen[value] = row
            if found:
                assign(found, "R03_1A2", f"Una partida cuadra con dos contrapartidas; ventana <= {day_window} dias", "CMP-1N")

            if not found and tolerance_cents:
                pairs = [(abs(cents(rows[target]["effective"]) + cents(rows[a]["effective"]) + cents(rows[b]["effective"])), a, b) for i, a in enumerate(others) for b in others[i + 1:] if abs(cents(rows[target]["effective"]) + cents(rows[a]["effective"]) + cents(rows[b]["effective"])) <= tolerance_cents]
                if pairs:
                    _, first, second = min(pairs)
                    assign([target, first, second], "R03_1A2", f"Una partida cuadra con dos contrapartidas dentro de tolerancia; ventana <= {day_window} dias", "CMP-1N")

    write_matches(ws, headers, rows, assigned, match_type, reason, difference)
    for row, item in rows.items():
        if row not in assigned and item.get("ret_code") and item["ret_code"] not in rates:
            ws.cell(row, headers[OUT_MOTIVO]).value = f"Indicador de retencion no parametrizado: {item['ret_code']}"
    stats = add_summary(workbook, "abiertas", system, rows, assigned, match_type, [("Indicadores sin parametrizar", ", ".join(sorted(unknown_codes)) or "Ninguno")])
    stats["unknown_retention_codes"] = sorted(unknown_codes)
    workbook.save(output_path)
    return stats
