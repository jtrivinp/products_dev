"""Normalización SAP, libros Excel y primitivas compartidas del motor."""

from __future__ import annotations

from collections import defaultdict
from copy import copy
from datetime import date, datetime
from pathlib import Path
from typing import Dict, Optional, Sequence, Tuple

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

COL_SOCIEDAD = "Sociedad"
COL_ACREEDOR_VIP = "Acreedor"
COL_ACREEDOR_GPC = "Cuenta"
COL_ASIGNACION = "Asignación"
COL_DOCUMENTO = "Nº documento"
COL_REFERENCIA = "Referencia"
COL_CLASE = "Clase de documento"
COL_VIA_PAGO = "Vía de pago"
COL_FECHA_CONT = "Fe.contabilización"
COL_IMPORTE_DOC = "Importe en moneda doc."
COL_IMPORTE_LOCAL = "Importe en moneda local"
COL_BASE_RET = "Importe base de retención"
COL_IND_RET = "Indicador retención"
COL_DOC_COMP = "Doc.compensación"

OUT_CODIGO = "PROPUESTA_COMPENSACION"
OUT_TIPO = "TIPO_MATCH"
OUT_MOTIVO = "MOTIVO_PROPUESTA"
OUT_DIFERENCIA = "DIFERENCIA_MATCH"
OUT_PCT_RET = "% RETENCION"
OUT_IMP_RET = "IMPORTE_RETENCION"
OUT_IMP_AJUST = "IMPORTE_AJUSTADO_RETENCION"


def norm_text(value) -> str:
    return "" if value is None else str(value).strip()


def as_float(value) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    return float(text)


def as_rate(value) -> float:
    if value in (None, ""):
        raise ValueError("Porcentaje vacio")
    if isinstance(value, str):
        value = float(value.strip().replace("%", "").replace(",", "."))
    return float(value) / 100.0 if float(value) > 1 else float(value)


def to_date(value) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    return value if isinstance(value, date) else None


def cents(value: float) -> int:
    return int(round(float(value) * 100))


def date_distance(first: Optional[date], second: Optional[date]) -> int:
    return 999999 if first is None or second is None else abs((first - second).days)


def header_map(ws) -> Dict[str, int]:
    return {norm_text(cell.value): cell.column for cell in ws[1] if cell.value is not None}


def detect_system(headers: Dict[str, int]) -> str:
    if COL_ACREEDOR_VIP in headers:
        return "VIP"
    if COL_ACREEDOR_GPC in headers:
        return "GPC"
    raise ValueError("No se puede detectar el sistema: falta 'Acreedor' (VIP) o 'Cuenta' (GPC).")


def account_column(system: str) -> str:
    return COL_ACREEDOR_VIP if system == "VIP" else COL_ACREEDOR_GPC


def amount_column(headers: Dict[str, int], modo: str) -> str:
    if modo == "abiertas" and COL_IMPORTE_DOC in headers:
        return COL_IMPORTE_DOC
    if COL_IMPORTE_LOCAL in headers:
        return COL_IMPORTE_LOCAL
    if COL_IMPORTE_DOC in headers:
        return COL_IMPORTE_DOC
    raise ValueError("No se encuentra una columna de importe reconocible.")


def copy_style_from_left(ws, new_col: int) -> None:
    src = ws.cell(1, new_col - 1)
    dst = ws.cell(1, new_col)
    if src.has_style:
        dst._style = copy(src._style)
    dst.alignment = copy(src.alignment)
    ws.column_dimensions[get_column_letter(new_col)].width = max(18, ws.column_dimensions[get_column_letter(new_col - 1)].width or 12)


def ensure_output_columns(ws, headers: Dict[str, int], modo: str) -> Dict[str, int]:
    output = dict(headers)
    names = [OUT_CODIGO, OUT_TIPO, OUT_MOTIVO, OUT_DIFERENCIA]
    if modo == "abiertas":
        names = [OUT_PCT_RET, OUT_IMP_RET, OUT_IMP_AJUST] + names
    for name in names:
        if name not in output:
            column = ws.max_column + 1
            copy_style_from_left(ws, column)
            ws.cell(1, column).value = name
            output[name] = column
    for table in ws.tables.values():
        start, end = table.ref.split(":")
        table.ref = f"{start}:{get_column_letter(ws.max_column)}{''.join(filter(str.isdigit, end))}"
    return output


def load_retention_table(path: Path) -> Dict[str, float]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    ws = workbook[workbook.sheetnames[0]]
    headers = {norm_text(cell.value).lower(): cell.column for cell in ws[1]}
    code_col = next((col for name, col in headers.items() if name in {"indicativo", "indicador retención", "indicador retencion", "codigo retención", "codigo retencion", "código retención", "código retencion"}), None)
    pct_col = next((col for name, col in headers.items() if name in {"porcentaje", "%", "% retención", "% retencion", "porcentaje retención", "porcentaje retencion"}), None)
    if not code_col or not pct_col:
        raise ValueError("La tabla de retenciones debe incluir 'Indicador retención' y 'Porcentaje'.")
    result = {}
    for row in range(2, ws.max_row + 1):
        code = norm_text(ws.cell(row, code_col).value).upper()
        pct = ws.cell(row, pct_col).value
        if not code and pct in (None, ""):
            continue
        if not code:
            raise ValueError(f"Fila {row}: porcentaje sin codigo de retencion")
        rate = as_rate(pct)
        if not 0 <= rate <= 1:
            raise ValueError(f"Fila {row}: porcentaje invalido para {code}: {pct}")
        if code in result:
            if result[code] != rate:
                raise ValueError(f"Código de retención duplicado con porcentajes distintos: {code} ({result[code]:.2%} y {rate:.2%})")
            continue
        result[code] = rate
    return result


def load_class_table(path: Path) -> set[str]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    ws = workbook[workbook.sheetnames[0]]
    if ws.max_column < 1 or "clase" not in norm_text(ws.cell(1, 1).value).lower():
        raise ValueError("El Excel de clases debe tener en la primera columna una cabecera como 'Clase documento'.")
    classes = {norm_text(ws.cell(row, 1).value).upper() for row in range(2, ws.max_row + 1) if norm_text(ws.cell(row, 1).value)}
    if not classes:
        raise ValueError("El Excel de clases no contiene ninguna clase autorizada.")
    return classes


def build_rows(ws, headers: Dict[str, int], system: str, modo: str) -> Dict[int, dict]:
    account = account_column(system)
    amount = amount_column(headers, modo)
    required = [COL_SOCIEDAD, account, COL_DOCUMENTO, COL_CLASE, COL_FECHA_CONT, amount]
    missing = [name for name in required if name not in headers]
    if missing:
        raise ValueError(f"Faltan columnas obligatorias: {missing}")
    rows = {}
    for row in range(2, ws.max_row + 1):
        society = norm_text(ws.cell(row, headers[COL_SOCIEDAD]).value)
        creditor = norm_text(ws.cell(row, headers[account]).value)
        if not society or not creditor:
            continue
        rows[row] = {"row": row, "soc": society, "acc": creditor, "doc": norm_text(ws.cell(row, headers[COL_DOCUMENTO]).value), "class": norm_text(ws.cell(row, headers[COL_CLASE]).value).upper(), "payment_method": norm_text(ws.cell(row, headers[COL_VIA_PAGO]).value) if COL_VIA_PAGO in headers else "", "assignment": norm_text(ws.cell(row, headers[COL_ASIGNACION]).value) if COL_ASIGNACION in headers else "", "reference": norm_text(ws.cell(row, headers[COL_REFERENCIA]).value) if COL_REFERENCIA in headers else "", "date": to_date(ws.cell(row, headers[COL_FECHA_CONT]).value), "amount": round(as_float(ws.cell(row, headers[amount]).value), 2), "cleared": bool(norm_text(ws.cell(row, headers[COL_DOC_COMP]).value)) if COL_DOC_COMP in headers else False}
    return rows


def write_matches(ws, headers, rows, assigned, match_type, reason, difference) -> None:
    highlight = PatternFill(fill_type="solid", fgColor="FFF2CC")
    for row in rows:
        if row in assigned:
            ws.cell(row, headers[OUT_CODIGO]).value = assigned[row]
            ws.cell(row, headers[OUT_TIPO]).value = match_type[row]
            ws.cell(row, headers[OUT_MOTIVO]).value = reason[row]
            ws.cell(row, headers[OUT_DIFERENCIA]).value = difference.get(row, 0.0)
            ws.cell(row, headers[OUT_DIFERENCIA]).number_format = "0.00"
            for col in (headers[OUT_CODIGO], headers[OUT_TIPO], headers[OUT_MOTIVO], headers[OUT_DIFERENCIA]):
                ws.cell(row, col).fill = highlight


def add_summary(workbook, modo, system, rows, assigned, match_type, extra: Sequence[Tuple[str, object]] = ()) -> dict:
    if "RESUMEN_PROCESO" in workbook.sheetnames:
        del workbook["RESUMEN_PROCESO"]
    ws = workbook.create_sheet("RESUMEN_PROCESO")
    ws.append(["Concepto", "Valor"])
    for item in (("Modo", modo), ("Sistema detectado", system), ("Partidas de datos analizadas", len(rows)), ("Partidas ya compensadas en origen", sum(1 for item in rows.values() if item["cleared"])), ("Partidas propuestas", len(assigned)), ("Grupos de propuesta", len(set(assigned.values())))):
        ws.append(item)
    counts = defaultdict(set)
    for row, code in assigned.items():
        counts[match_type[row]].add(code)
    for kind in sorted(counts):
        ws.append([f"Grupos {kind}", len(counts[kind])])
    for item in extra:
        ws.append(item)
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 70
    return {"rows": len(rows), "proposed_rows": len(assigned), "groups": len(set(assigned.values())), "by_type": {kind: len(codes) for kind, codes in counts.items()}}
