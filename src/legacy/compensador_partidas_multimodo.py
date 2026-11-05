#!/usr/bin/env python3
"""Motor de apoyo a la compensacion de partidas SAP.

Modos soportados:
  abiertas       Matching de partidas abiertas con retenciones y 1:N.
  domiciliados   Facturas con via de pago contra pago TU (GPC) / TP (VIP, supuesto configurable).
  masmenos       Anulaciones/abonos 1:1 por importe opuesto y clases autorizadas.

El proceso NO compensa en SAP. Genera un Excel enriquecido con propuestas para revision humana.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from copy import copy
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# Columnas comunes / aliases SAP
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

# Columnas generadas comunes
OUT_CODIGO = "PROPUESTA_COMPENSACION"
OUT_TIPO = "TIPO_MATCH"
OUT_MOTIVO = "MOTIVO_PROPUESTA"
OUT_DIFERENCIA = "DIFERENCIA_MATCH"

# Columnas generadas solo para abiertas
OUT_PCT_RET = "% RETENCION"
OUT_IMP_RET = "IMPORTE_RETENCION"
OUT_IMP_AJUST = "IMPORTE_AJUSTADO_RETENCION"


def norm_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


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
    """Convierte 15, 15%, 0.15 a 0.15."""
    if value in (None, ""):
        raise ValueError("Porcentaje vacio")
    if isinstance(value, str):
        text = value.strip().replace("%", "").replace(",", ".")
        n = float(text)
        return n / 100.0 if n > 1 else n
    n = float(value)
    return n / 100.0 if n > 1 else n


def to_date(value) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def cents(value: float) -> int:
    return int(round(float(value) * 100))


def date_distance(a: Optional[date], b: Optional[date]) -> int:
    if a is None or b is None:
        return 999999
    return abs((a - b).days)


def header_map(ws) -> Dict[str, int]:
    return {norm_text(c.value): c.column for c in ws[1] if c.value is not None}


def detect_system(headers: Dict[str, int]) -> str:
    if COL_ACREEDOR_VIP in headers:
        return "VIP"
    if COL_ACREEDOR_GPC in headers:
        return "GPC"
    raise ValueError("No se puede detectar el sistema: falta 'Acreedor' (VIP) o 'Cuenta' (GPC).")


def account_column(system: str) -> str:
    return COL_ACREEDOR_VIP if system == "VIP" else COL_ACREEDOR_GPC


def amount_column(headers: Dict[str, int], modo: str) -> str:
    # Para abiertas preservamos el comportamiento del prototipo original.
    if modo == "abiertas" and COL_IMPORTE_DOC in headers:
        return COL_IMPORTE_DOC
    if COL_IMPORTE_LOCAL in headers:
        return COL_IMPORTE_LOCAL
    if COL_IMPORTE_DOC in headers:
        return COL_IMPORTE_DOC
    raise ValueError("No se encuentra una columna de importe reconocible.")


def copy_style_from_left(ws, new_col: int) -> None:
    """Copia solo el estilo de cabecera y el ancho.

    Evita copiar estilos celda a celda sobre miles de filas, lo que ralentizaba
    innecesariamente el procesamiento de los Excel grandes. Las tablas de Excel
    aplican su estilo visual al ampliar el rango y los valores de salida reciben
    sus formatos de numero explicitamente.
    """
    src_col = new_col - 1
    src = ws.cell(1, src_col)
    dst = ws.cell(1, new_col)
    if src.has_style:
        dst._style = copy(src._style)
    dst.alignment = copy(src.alignment)
    ws.column_dimensions[get_column_letter(new_col)].width = max(
        18, ws.column_dimensions[get_column_letter(src_col)].width or 12
    )


def ensure_output_columns(ws, headers: Dict[str, int], modo: str) -> Dict[str, int]:
    out = dict(headers)
    names = [OUT_CODIGO, OUT_TIPO, OUT_MOTIVO, OUT_DIFERENCIA]
    if modo == "abiertas":
        names = [OUT_PCT_RET, OUT_IMP_RET, OUT_IMP_AJUST] + names

    for name in names:
        if name not in out:
            c = ws.max_column + 1
            copy_style_from_left(ws, c)
            ws.cell(1, c).value = name
            out[name] = c

    # Si hay tabla Excel, incluir las columnas nuevas en su rango.
    for table in ws.tables.values():
        start, end = table.ref.split(":")
        import re
        m = re.match(r"([A-Z]+)(\d+)", start)
        n = re.match(r"([A-Z]+)(\d+)", end)
        if m and n:
            table.ref = f"{m.group(1)}{m.group(2)}:{get_column_letter(ws.max_column)}{n.group(2)}"
    return out


def load_retention_table(path: Path) -> Dict[str, float]:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    headers = {norm_text(c.value).lower(): c.column for c in ws[1]}
    code_col = pct_col = None
    for h, col in headers.items():
        if h in {"indicador retención", "indicador retencion", "codigo retención", "codigo retencion", "código retención", "código retencion"}:
            code_col = col
        if h in {"porcentaje", "%", "% retención", "% retencion", "porcentaje retención", "porcentaje retencion"}:
            pct_col = col
    if not code_col or not pct_col:
        raise ValueError("La tabla de retenciones debe incluir 'Indicador retención' y 'Porcentaje'.")

    result: Dict[str, float] = {}
    for r in range(2, ws.max_row + 1):
        code = norm_text(ws.cell(r, code_col).value).upper()
        pct = ws.cell(r, pct_col).value
        if not code and pct in (None, ""):
            continue
        if not code:
            raise ValueError(f"Fila {r}: porcentaje sin codigo de retencion")
        rate = as_rate(pct)
        if not 0 <= rate <= 1:
            raise ValueError(f"Fila {r}: porcentaje invalido para {code}: {pct}")
        if code in result:
            raise ValueError(f"Codigo de retencion duplicado: {code}")
        result[code] = rate
    return result


def load_class_table(path: Path) -> set[str]:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    if ws.max_column < 1:
        raise ValueError("El fichero de clases esta vacio.")
    header = norm_text(ws.cell(1, 1).value).lower()
    if "clase" not in header:
        raise ValueError("El Excel de clases debe tener en la primera columna una cabecera como 'Clase documento'.")
    classes = {norm_text(ws.cell(r, 1).value).upper() for r in range(2, ws.max_row + 1) if norm_text(ws.cell(r, 1).value)}
    if not classes:
        raise ValueError("El Excel de clases no contiene ninguna clase autorizada.")
    return classes


def build_rows(ws, headers: Dict[str, int], system: str, modo: str) -> Dict[int, dict]:
    acc_col = account_column(system)
    imp_col = amount_column(headers, modo)
    required = [COL_SOCIEDAD, acc_col, COL_DOCUMENTO, COL_CLASE, COL_FECHA_CONT, imp_col]
    missing = [c for c in required if c not in headers]
    if missing:
        raise ValueError(f"Faltan columnas obligatorias: {missing}")

    rows: Dict[int, dict] = {}
    for r in range(2, ws.max_row + 1):
        soc = norm_text(ws.cell(r, headers[COL_SOCIEDAD]).value)
        acc = norm_text(ws.cell(r, headers[acc_col]).value)
        # Las extracciones contienen subtotales/fila-resumen sin sociedad.
        if not soc or not acc:
            continue
        amount = round(as_float(ws.cell(r, headers[imp_col]).value), 2)
        rows[r] = {
            "row": r,
            "soc": soc,
            "acc": acc,
            "doc": norm_text(ws.cell(r, headers[COL_DOCUMENTO]).value),
            "class": norm_text(ws.cell(r, headers[COL_CLASE]).value).upper(),
            "payment_method": norm_text(ws.cell(r, headers[COL_VIA_PAGO]).value) if COL_VIA_PAGO in headers else "",
            "assignment": norm_text(ws.cell(r, headers[COL_ASIGNACION]).value) if COL_ASIGNACION in headers else "",
            "reference": norm_text(ws.cell(r, headers[COL_REFERENCIA]).value) if COL_REFERENCIA in headers else "",
            "date": to_date(ws.cell(r, headers[COL_FECHA_CONT]).value),
            "amount": amount,
            "cleared": bool(norm_text(ws.cell(r, headers[COL_DOC_COMP]).value)) if COL_DOC_COMP in headers else False,
        }
    return rows


def write_matches(ws, headers: Dict[str, int], rows: Dict[int, dict], assigned: Dict[int, str], match_type: Dict[int, str], reason: Dict[int, str], difference: Dict[int, float]) -> None:
    highlight = PatternFill(fill_type="solid", fgColor="FFF2CC")
    for r in rows:
        if r in assigned:
            ws.cell(r, headers[OUT_CODIGO]).value = assigned[r]
            ws.cell(r, headers[OUT_TIPO]).value = match_type[r]
            ws.cell(r, headers[OUT_MOTIVO]).value = reason[r]
            ws.cell(r, headers[OUT_DIFERENCIA]).value = difference.get(r, 0.0)
            ws.cell(r, headers[OUT_DIFERENCIA]).number_format = "0.00"
            for c in [headers[OUT_CODIGO], headers[OUT_TIPO], headers[OUT_MOTIVO], headers[OUT_DIFERENCIA]]:
                ws.cell(r, c).fill = highlight


def add_summary(wb, modo: str, system: str, rows: Dict[int, dict], assigned: Dict[int, str], match_type: Dict[int, str], extra: Sequence[Tuple[str, object]] = ()) -> dict:
    if "RESUMEN_PROCESO" in wb.sheetnames:
        del wb["RESUMEN_PROCESO"]
    ws = wb.create_sheet("RESUMEN_PROCESO")
    ws.append(["Concepto", "Valor"])
    ws.append(["Modo", modo])
    ws.append(["Sistema detectado", system])
    ws.append(["Partidas de datos analizadas", len(rows)])
    ws.append(["Partidas ya compensadas en origen", sum(1 for x in rows.values() if x["cleared"])])
    ws.append(["Partidas propuestas", len(assigned)])
    ws.append(["Grupos de propuesta", len(set(assigned.values()))])
    counts = defaultdict(set)
    for r, code in assigned.items():
        counts[match_type[r]].add(code)
    for kind in sorted(counts):
        ws.append([f"Grupos {kind}", len(counts[kind])])
    for k, v in extra:
        ws.append([k, v])
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 70
    return {
        "rows": len(rows),
        "proposed_rows": len(assigned),
        "groups": len(set(assigned.values())),
        "by_type": {k: len(v) for k, v in counts.items()},
    }


def process_abiertas(input_path: Path, retention_path: Path, output_path: Path, day_window: int = 31) -> dict:
    rates = load_retention_table(retention_path)
    wb = load_workbook(input_path)
    ws = wb[wb.sheetnames[0]]
    headers = header_map(ws)
    system = detect_system(headers)
    acc_col = account_column(system)

    required = [COL_SOCIEDAD, acc_col, COL_ASIGNACION, COL_DOCUMENTO, COL_FECHA_CONT, COL_BASE_RET, COL_IND_RET]
    missing = [c for c in required if c not in headers]
    if missing:
        raise ValueError(f"Faltan columnas obligatorias para abiertas: {missing}")

    headers = ensure_output_columns(ws, headers, "abiertas")
    rows = build_rows(ws, headers, system, "abiertas")
    unknown_codes = set()

    for r, x in rows.items():
        base = as_float(ws.cell(r, headers[COL_BASE_RET]).value)
        code = norm_text(ws.cell(r, headers[COL_IND_RET]).value).upper()
        rate = rates.get(code) if code else None
        if code and rate is None:
            unknown_codes.add(code)
        retention_amount = base * rate if rate is not None else 0.0
        adjusted = round(x["amount"] - retention_amount, 2) if rate is not None else x["amount"]
        x.update({"effective": adjusted, "ret_code": code, "has_retention": bool(code and rate is not None and abs(base) > 1e-9)})
        ws.cell(r, headers[OUT_PCT_RET]).value = rate
        ws.cell(r, headers[OUT_PCT_RET]).number_format = "0.00%"
        ws.cell(r, headers[OUT_IMP_RET]).value = round(retention_amount, 2) if rate is not None else None
        ws.cell(r, headers[OUT_IMP_AJUST]).value = adjusted

    assigned: Dict[int, str] = {}; match_type: Dict[int, str] = {}; reason: Dict[int, str] = {}; difference: Dict[int, float] = {}
    seq = 1
    def assign(group: List[int], kind: str, why: str, prefix: str = "CMP"):
        nonlocal seq
        code = f"{prefix}-{seq:05d}"; seq += 1
        diff = round(sum(rows[r]["effective"] for r in group), 2)
        for rr in group:
            assigned[rr] = code; match_type[rr] = kind; reason[rr] = why; difference[rr] = diff

    candidates = [r for r,x in rows.items() if not x["cleared"]]
    groups = defaultdict(list)
    for r in candidates:
        x=rows[r]
        if x["assignment"]:
            groups[(x["soc"],x["acc"],x["assignment"])].append(r)
    for grp in groups.values():
        if len(grp)>=2 and sum(cents(rows[r]["effective"]) for r in grp)==0:
            assign(grp,"R01_GRUPO_EXACTO","Misma sociedad, acreedor y asignacion; suma efectiva = 0")

    buckets=defaultdict(list)
    for r in candidates:
        if r not in assigned: buckets[(rows[r]["soc"],rows[r]["acc"])].append(r)
    for grp in buckets.values():
        by_amt=defaultdict(list)
        for r in grp: by_amt[cents(rows[r]["effective"])].append(r)
        for a in sorted(x for x in by_amt if x>0):
            if -a not in by_amt: continue
            pos=[r for r in by_amt[a] if r not in assigned]; neg=[r for r in by_amt[-a] if r not in assigned]
            while pos and neg:
                _,p,n=min((date_distance(rows[p]["date"],rows[n]["date"]),p,n) for p in pos for n in neg)
                has_ret=rows[p]["has_retention"] or rows[n]["has_retention"]
                assign([p,n],"R02_1A1_RETENCION" if has_ret else "R02_1A1_EXACTO", "Importes efectivos opuestos" + ("; incluye ajuste por retencion parametrizada" if has_ret else ""), "CMP-RET" if has_ret else "CMP")
                pos.remove(p); neg.remove(n)

    buckets=defaultdict(list)
    for r in candidates:
        if r not in assigned: buckets[(rows[r]["soc"],rows[r]["acc"])].append(r)
    for grp in buckets.values():
        for target in list(grp):
            if target in assigned or cents(rows[target]["effective"])==0: continue
            t=rows[target]
            others=[r for r in grp if r!=target and r not in assigned and cents(rows[r]["effective"])*cents(t["effective"])<0 and date_distance(rows[r]["date"],t["date"])<=day_window]
            seen={}; wanted=-cents(t["effective"]); found=None
            for r in others:
                v=cents(rows[r]["effective"]); comp=wanted-v
                if comp in seen and seen[comp]!=r:
                    found=[target,seen[comp],r]; break
                seen[v]=r
            if found:
                assign(found,"R03_1A2",f"Una partida cuadra con dos contrapartidas; ventana <= {day_window} dias","CMP-1N")

    write_matches(ws,headers,rows,assigned,match_type,reason,difference)
    for r,x in rows.items():
        if r not in assigned and x.get("ret_code") and x["ret_code"] not in rates:
            ws.cell(r, headers[OUT_MOTIVO]).value=f"Indicador de retencion no parametrizado: {x['ret_code']}"

    stats=add_summary(wb,"abiertas",system,rows,assigned,match_type,[("Indicadores sin parametrizar",", ".join(sorted(unknown_codes)) or "Ninguno")])
    stats["unknown_retention_codes"]=sorted(unknown_codes)
    wb.save(output_path)
    return stats


def find_subset_for_target(rows: Dict[int, dict], candidate_rows: List[int], target_abs: int, tolerance: int, max_items: int) -> Optional[List[int]]:
    """Busca de 2..max_items partidas cuya suma absoluta cuadre con target.

    Usa programacion dinamica y conserva una sola combinacion por (n, suma),
    suficiente para generar una propuesta determinista y evitar explosion combinatoria.
    """
    states: Dict[Tuple[int,int], Tuple[int,...]] = {(0,0): ()}
    limit = target_abs + tolerance
    for r in candidate_rows:
        v=abs(cents(rows[r]["amount"]))
        if v==0 or v>limit: continue
        current=list(states.items())
        for (n,s), combo in current:
            if n>=max_items: continue
            ns=s+v; nn=n+1
            if ns>limit: continue
            key=(nn,ns)
            if key not in states:
                states[key]=combo+(r,)
        for n in range(2,max_items+1):
            for delta in range(0,tolerance+1):
                for s in ({target_abs-delta,target_abs+delta} if delta else {target_abs}):
                    if (n,s) in states:
                        return list(states[(n,s)])
    return None


def process_domiciliados(input_path: Path, output_path: Path, payment_class: Optional[str] = None, tolerance_cents: int = 3, max_invoices: int = 4) -> dict:
    wb=load_workbook(input_path); ws=wb[wb.sheetnames[0]]; headers=header_map(ws)
    system=detect_system(headers)
    if COL_VIA_PAGO not in headers:
        raise ValueError("El modo domiciliados requiere la columna 'Vía de pago'.")
    headers=ensure_output_columns(ws,headers,"domiciliados")
    rows=build_rows(ws,headers,system,"domiciliados")
    auto_class = "TU" if system=="GPC" else "TP"
    payment_class=(payment_class or auto_class).strip().upper()
    payment_class_source="parametro" if payment_class != auto_class else ("regla confirmada GPC" if system=="GPC" else "SUPUESTO DE PRUEBA VIP: confirmar TP con BPS")

    assigned={}; match_type={}; reason={}; difference={}; seq=1
    def assign(group: List[int], kind: str, why: str):
        nonlocal seq
        code=f"CMP-DOM-{seq:05d}"; seq+=1
        diff=round(sum(rows[r]["amount"] for r in group),2)
        for rr in group:
            assigned[rr]=code; match_type[rr]=kind; reason[rr]=why; difference[rr]=diff

    candidates=[r for r,x in rows.items() if not x["cleared"]]
    buckets=defaultdict(list)
    for r in candidates: buckets[(rows[r]["soc"],rows[r]["acc"])].append(r)

    # DOM01 1:1: pago TU/TP contra factura con Via de pago, misma sociedad+acreedor.
    for grp in buckets.values():
        payments=[r for r in grp if rows[r]["class"]==payment_class]
        invoices=[r for r in grp if rows[r]["class"]!=payment_class and rows[r]["payment_method"]]
        for p in sorted(payments, key=lambda r:(rows[r]["date"] or date.min, r)):
            if p in assigned: continue
            pc=cents(rows[p]["amount"])
            options=[]
            for i in invoices:
                if i in assigned: continue
                ic=cents(rows[i]["amount"])
                if pc*ic>=0: continue
                diff=abs(pc+ic)
                if diff<=tolerance_cents:
                    options.append((diff,date_distance(rows[p]["date"],rows[i]["date"]),i))
            if options:
                diff,_,i=min(options)
                assign([p,i],"DOM01_1A1",f"{payment_class} contra factura con vía de pago; misma sociedad y acreedor; diferencia <= {tolerance_cents} céntimos")

    # DOM02 1:N: un TU/TP contra varias facturas con via de pago.
    for grp in buckets.values():
        payments=[r for r in grp if rows[r]["class"]==payment_class and r not in assigned]
        invoices=[r for r in grp if rows[r]["class"]!=payment_class and rows[r]["payment_method"] and r not in assigned]
        for p in payments:
            pc=cents(rows[p]["amount"])
            opposite=[i for i in invoices if i not in assigned and pc*cents(rows[i]["amount"])<0]
            # Prioriza facturas temporalmente cercanas al pago.
            opposite.sort(key=lambda i:(date_distance(rows[p]["date"],rows[i]["date"]),i))
            combo=find_subset_for_target(rows,opposite,abs(pc),tolerance_cents,max_invoices)
            if combo:
                assign([p]+combo,"DOM02_1AN",f"Un {payment_class} contra {len(combo)} facturas con vía de pago; misma sociedad y acreedor; diferencia <= {tolerance_cents} céntimos")

    write_matches(ws,headers,rows,assigned,match_type,reason,difference)
    stats=add_summary(wb,"domiciliados",system,rows,assigned,match_type,[
        ("Clase de pago utilizada",payment_class),
        ("Origen clase de pago",payment_class_source),
        ("Tolerancia",f"{tolerance_cents} céntimos"),
        ("Máximo facturas por pago en búsqueda 1:N",max_invoices),
    ])
    stats["payment_class"]=payment_class; stats["payment_class_source"]=payment_class_source
    wb.save(output_path); return stats


def process_masmenos(input_path: Path, classes_path: Path, output_path: Path) -> dict:
    allowed=load_class_table(classes_path)
    wb=load_workbook(input_path); ws=wb[wb.sheetnames[0]]; headers=header_map(ws)
    system=detect_system(headers); headers=ensure_output_columns(ws,headers,"masmenos")
    rows=build_rows(ws,headers,system,"masmenos")
    assigned={}; match_type={}; reason={}; difference={}; seq=1
    def assign(group: List[int]):
        nonlocal seq
        code=f"CMP-MM-{seq:05d}"; seq+=1
        diff=round(sum(rows[r]["amount"] for r in group),2)
        why="Misma sociedad y acreedor/cuenta; clases autorizadas; importes exactamente opuestos"
        for rr in group:
            assigned[rr]=code; match_type[rr]="MM01_1A1_EXACTO"; reason[rr]=why; difference[rr]=diff

    buckets=defaultdict(lambda:defaultdict(list))
    for r,x in rows.items():
        if x["cleared"] or x["class"] not in allowed or cents(x["amount"])==0: continue
        buckets[(x["soc"],x["acc"])][cents(x["amount"])].append(r)

    for by_amt in buckets.values():
        for a in sorted(x for x in by_amt if x>0):
            if -a not in by_amt: continue
            pos=[r for r in by_amt[a] if r not in assigned]
            neg=[r for r in by_amt[-a] if r not in assigned]
            while pos and neg:
                # No hace falta misma clase; priorizamos fechas cercanas solo para que la propuesta sea estable.
                _,p,n=min((date_distance(rows[p]["date"],rows[n]["date"]),p,n) for p in pos for n in neg)
                assign([p,n]); pos.remove(p); neg.remove(n)

    write_matches(ws,headers,rows,assigned,match_type,reason,difference)
    stats=add_summary(wb,"masmenos",system,rows,assigned,match_type,[
        ("Clases autorizadas utilizadas",", ".join(sorted(allowed))),
        ("Nota","Lista de prueba/parametrizable; sustituir por listado oficial de BPS para producción"),
    ])
    stats["allowed_classes"]=sorted(allowed)
    wb.save(output_path); return stats


def process(modo: str, partidas: Path, output: Path, retenciones: Optional[Path] = None, clases: Optional[Path] = None, ventana_dias: int = 31, clase_pago_domiciliado: Optional[str] = None, tolerancia_centimos: int = 3, max_facturas_domiciliado: int = 4) -> dict:
    if modo=="abiertas":
        if retenciones is None: raise ValueError("El modo 'abiertas' requiere --retenciones.")
        return process_abiertas(partidas,retenciones,output,ventana_dias)
    if modo=="domiciliados":
        return process_domiciliados(partidas,output,clase_pago_domiciliado,tolerancia_centimos,max_facturas_domiciliado)
    if modo=="masmenos":
        if clases is None: raise ValueError("El modo 'masmenos' requiere --clases.")
        return process_masmenos(partidas,clases,output)
    raise ValueError(f"Modo no soportado: {modo}")


def main():
    parser=argparse.ArgumentParser(description="Motor de propuestas de compensacion SAP")
    parser.add_argument("partidas",type=Path,help="Excel exportado desde SAP")
    parser.add_argument("retenciones_legacy",nargs="?",type=Path,help="Compatibilidad: Excel de retenciones para modo abiertas")
    parser.add_argument("--modo",choices=["abiertas","domiciliados","masmenos"],default="abiertas",help="Tipo de automatizacion (default: abiertas)")
    parser.add_argument("--retenciones",type=Path,help="Excel Indicador retencion + Porcentaje (solo abiertas)")
    parser.add_argument("--clases",type=Path,help="Excel con clases de documento autorizadas (solo masmenos)")
    parser.add_argument("-o","--output",type=Path,default=Path("PARTIDAS_PROCESADAS.xlsx"),help="Excel de salida")
    parser.add_argument("--ventana-dias",type=int,default=31,help="Ventana maxima para 1:2 en abiertas")
    parser.add_argument("--clase-pago-domiciliado",help="Sobrescribe clase del pago. Auto: TU en GPC, TP en VIP (TP pendiente de confirmar)")
    parser.add_argument("--tolerancia-centimos",type=int,default=3,help="Tolerancia para domiciliados (default: 3)")
    parser.add_argument("--max-facturas-domiciliado",type=int,default=4,help="Maximo N en 1:N de domiciliados (default: 4)")
    args=parser.parse_args()

    retention=args.retenciones or args.retenciones_legacy
    stats=process(args.modo,args.partidas,args.output,retention,args.clases,args.ventana_dias,args.clase_pago_domiciliado,args.tolerancia_centimos,args.max_facturas_domiciliado)
    print(f"Modo: {args.modo}")
    print(f"Salida: {args.output}")
    print(f"Partidas propuestas: {stats['proposed_rows']}")
    print(f"Grupos: {stats['groups']}")
    if stats.get("payment_class"):
        print(f"Clase pago domiciliado: {stats['payment_class']} ({stats['payment_class_source']})")
    print("Por tipo:")
    for k,v in sorted(stats["by_type"].items()): print(f"  {k}: {v}")
    if stats.get("unknown_retention_codes"):
        print("Indicadores sin parametrizar:",", ".join(stats["unknown_retention_codes"]))


if __name__=="__main__":
    main()
