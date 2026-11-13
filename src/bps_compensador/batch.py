"""Procesamiento de todos los Excel pendientes y consolidación de propuestas."""
from __future__ import annotations
from datetime import datetime
from pathlib import Path
import shutil
import tempfile
from openpyxl import Workbook, load_workbook
from .common import detect_system, header_map, OUT_CODIGO, OUT_TIPO, OUT_MOTIVO, OUT_DIFERENCIA, OUT_PCT_RET, OUT_IMP_RET, OUT_IMP_AJUST
from .engine import process
from .html_report import generate

HELPERS=[OUT_CODIGO,OUT_TIPO,OUT_MOTIVO,OUT_DIFERENCIA,OUT_PCT_RET,OUT_IMP_RET,OUT_IMP_AJUST]

def pending_files(input_dir: Path) -> list[Path]:
    processed=input_dir / "procesados"
    return sorted(p for p in input_dir.glob("*.xls*") if p.is_file() and not p.name.startswith("~$") and p.parent != processed)

def run_batch(settings, logger):
    files=pending_files(settings.input_dir)
    if not files: raise ValueError("No hay ficheros Excel pendientes en la carpeta de entrada.")
    run_id=datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir=settings.output_dir / run_id
    processed=settings.input_dir / "procesados"; processed.mkdir(exist_ok=True)
    rows=[]; columns={"FICHERO_ORIGEN","SISTEMA_ORIGEN"}; total_groups=0
    temp=Path(tempfile.mkdtemp(prefix="bps-lote-")); staged=temp / "entrega"; staged.mkdir()
    try:
        for file_index, source in enumerate(files,1):
            wb=load_workbook(source,read_only=True); headers=header_map(wb[wb.sheetnames[0]]); system=detect_system(headers); wb.close()
            logger.info("Procesando %s (%s)",source.name,system)
            outputs=[]
            # Cascada confirmada: las propuestas se fusionan por primera coincidencia.
            classes_file = settings.classes_gpc_file if system == "GPC" else settings.classes_vip_file
            retentions_file = settings.retentions_gpc_file if system == "GPC" else settings.retentions_vip_file
            outputs.append(process("masmenos",source,temp/f"{file_index}-mm.xlsx",clases=settings.auxiliary(classes_file)))
            outputs.append(process("domiciliados",source,temp/f"{file_index}-dom.xlsx",clase_pago_domiciliado=settings.gpc_class if system=="GPC" else settings.vip_class,tolerancia_centimos=settings.tolerance_cents,max_facturas_domiciliado=settings.max_invoices))
            outputs.append(process("abiertas",source,temp/f"{file_index}-ret.xlsx",retenciones=settings.auxiliary(retentions_file),ventana_dias=settings.day_window,tolerancia_retencion_centimos=settings.retention_tolerance_cents))
            total_groups += sum(x["groups"] for x in outputs)
            base=load_workbook(source,data_only=True,read_only=True); base_ws=base[base.sheetnames[0]]; base_headers=[str(c.value or "") for c in base_ws[1]]
            columns.update(base_headers)
            proposals={}
            for name in (f"{file_index}-mm.xlsx",f"{file_index}-dom.xlsx",f"{file_index}-ret.xlsx"):
                owb=load_workbook(temp/name,data_only=True,read_only=True); ws=owb[owb.sheetnames[0]]; h=header_map(ws)
                for number, values in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
                    if number not in proposals and h.get(OUT_CODIGO) and values[h[OUT_CODIGO]-1]:
                        proposals[number]={key: values[h[key]-1] if key in h else None for key in HELPERS}
                owb.close()
            for number, values in enumerate(base_ws.iter_rows(min_row=2,values_only=True),2):
                item=dict(zip(base_headers,values)); item.update({"FICHERO_ORIGEN":source.name,"SISTEMA_ORIGEN":system}); item.update(proposals.get(number,{})); rows.append(item)
            base.close()
        out=Workbook(); ws=out.active; ws.title="RESULTADOS_CONSOLIDADOS"; ordered=["FICHERO_ORIGEN","SISTEMA_ORIGEN"]+sorted(columns-{"FICHERO_ORIGEN","SISTEMA_ORIGEN"})+HELPERS
        ordered=list(dict.fromkeys(ordered)); ws.append(ordered)
        for item in rows: ws.append([item.get(col) for col in ordered])
        excel=staged/f"Compensacion_{run_id}.xlsx"; out.save(excel); out.close()
        proposed=sum(1 for item in rows if item.get(OUT_CODIGO)); stats={"rows":len(rows),"proposed_rows":proposed,"groups":total_groups}
        html=generate(excel,staged/f"Compensacion_{run_id}.html","Ejecución consolidada",f"{len(files)} fichero(s)",stats)
        for source in files:
            archived = processed / f"{source.stem}_{run_id}{source.suffix}"
            shutil.move(str(source), str(archived))
        shutil.move(str(staged), str(output_dir))
        excel=output_dir / excel.name; html=output_dir / html.name
        return stats,excel,html,len(files)
    finally:
        shutil.rmtree(temp, ignore_errors=True)
