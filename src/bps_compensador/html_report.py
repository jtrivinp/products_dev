"""Visor HTML local autocontenido para resultados ya calculados."""
from __future__ import annotations
from collections import defaultdict
from html import escape
from pathlib import Path
from openpyxl import load_workbook

def generate(excel: Path, destination: Path, mode: str, source_name: str, stats: dict) -> Path:
    wb=load_workbook(excel,data_only=True,read_only=True); ws=wb[wb.sheetnames[0]]
    headers=[str(c.value or "") for c in ws[1]]; index={name:i for i,name in enumerate(headers)}
    code=index.get("PROPUESTA_COMPENSACION"); doc=index.get("Nº documento"); soc=index.get("Sociedad"); acc=index.get("Acreedor",index.get("Cuenta")); typ=index.get("TIPO_MATCH"); reason=index.get("MOTIVO_PROPUESTA")
    groups=defaultdict(list); unmatched=[]
    for values in ws.iter_rows(min_row=2,values_only=True):
        row=["" if v is None else str(v) for v in values]
        if code is not None and row[code]: groups[row[code]].append(row)
        else: unmatched.append(row)
    def cell(row,i): return escape(row[i]) if i is not None and i<len(row) else ""
    rows=[]
    for key,items in groups.items():
        docs=', '.join(cell(x,doc) for x in items); rows.append(f"<tr data-docs='{escape(docs,quote=True)}'><td>{escape(key)}</td><td>{cell(items[0],typ)}</td><td>{cell(items[0],soc)}</td><td>{cell(items[0],acc)}</td><td>{docs}</td><td>{cell(items[0],reason)}</td><td><button onclick=\"copyDocs(this)\">Copiar documentos</button></td></tr>")
    no_match=''.join(f"<tr><td>{cell(row,soc)}</td><td>{cell(row,acc)}</td><td>{cell(row,doc)}</td><td>{cell(row,reason) or 'Sin coincidencia según las reglas ejecutadas'}</td></tr>" for row in unmatched)
    html=f'''<!doctype html><html lang="es"><meta charset="utf-8"><title>Resultado BPS</title><style>body{{font:14px Segoe UI,Arial;margin:28px;color:#17233b}}table{{border-collapse:collapse;width:100%;margin:12px 0}}th,td{{border:1px solid #d7dce5;padding:8px;text-align:left;vertical-align:top}}th{{background:#183b68;color:white;cursor:pointer}}th:hover{{background:#24558e}}input{{padding:8px;width:340px}}button{{padding:6px 10px}}.cards{{display:flex;gap:12px;flex-wrap:wrap}}.card{{background:#edf4fb;padding:12px;border-radius:6px;min-width:130px}}.hint{{color:#4b5563}}</style><h1>Propuestas de compensación BPS</h1><p><b>Modo:</b> {escape(mode)} · <b>Origen:</b> {escape(source_name)}</p><p>Las coincidencias son propuestas. La validación y ejecución final se realiza manualmente en SAP.</p><div class="cards"><div class="card">Partidas analizadas<br><b>{stats.get('rows',0)}</b></div><div class="card">Partidas propuestas<br><b>{stats.get('proposed_rows',0)}</b></div><div class="card">Grupos encontrados<br><b>{stats.get('groups',0)}</b></div><div class="card">Sin match<br><b>{len(unmatched)}</b></div></div><h2>Casos encontrados</h2><input oninput="filterTable('matches',this.value)" placeholder="Buscar por cualquier campo"><p class="hint">Pulsa una cabecera para ordenar.</p><table id="matches"><tr><th>Código</th><th>Tipo</th><th>Sociedad</th><th>Acreedor</th><th>Documentos</th><th>Motivo</th><th></th></tr>{''.join(rows)}</table><h2>Partidas sin propuesta</h2><input oninput="filterTable('unmatched',this.value)" placeholder="Buscar por cualquier campo"><p class="hint">Pulsa una cabecera para ordenar.</p><table id="unmatched"><tr><th>Sociedad</th><th>Acreedor</th><th>Documento</th><th>Explicación</th></tr>{no_match}</table><script>function filterTable(id,q){{q=q.toLowerCase();document.querySelectorAll('#'+id+' tr').forEach((r,i)=>{{if(i)r.style.display=r.innerText.toLowerCase().includes(q)?'':'none'}})}}document.querySelectorAll('table').forEach(t=>t.querySelectorAll('th').forEach((h,i)=>h.onclick=()=>{{let rows=[...t.rows].slice(1),asc=h.dataset.asc!=='1';rows.sort((a,b)=>a.cells[i].innerText.localeCompare(b.cells[i].innerText,undefined,{{numeric:true}})*(asc?1:-1));rows.forEach(r=>t.appendChild(r));t.querySelectorAll('th').forEach(x=>delete x.dataset.asc);h.dataset.asc=asc?'1':'0'}}));function copyDocs(b){{navigator.clipboard.writeText(b.parentElement.parentElement.dataset.docs);b.textContent='Copiado';setTimeout(()=>b.textContent='Copiar documentos',1200)}}</script></html>'''
    destination.write_text(html,encoding="utf-8"); return destination
