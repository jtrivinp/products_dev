"""Interfaz ttk local para usuarios de negocio."""
from __future__ import annotations
import os, threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from .application import MODE_LABELS, candidates, run
from .settings import load

class MainWindow(ttk.Frame):
    def __init__(self, root):
        super().__init__(root, padding=18); self.root=root; self.grid(sticky="nsew"); root.title("BPS — Propuestas de compensación"); root.columnconfigure(0, weight=1)
        self.settings = load(); self.mode=tk.StringVar(value="retencion"); self.file=tk.StringVar(); self.status=tk.StringVar(value="Listo para seleccionar un fichero."); self.last_output=None
        ttk.Label(self, text="Propuestas de compensación BPS", font=("Segoe UI", 16, "bold")).grid(row=0,column=0,sticky="w")
        ttk.Label(self, text="Las coincidencias son propuestas; la validación y ejecución final se realiza manualmente en SAP.", wraplength=700).grid(row=1,column=0,sticky="w",pady=(3,15))
        box=ttk.LabelFrame(self,text="Modo de procesamiento",padding=10); box.grid(row=2,column=0,sticky="ew")
        for i,(value,label) in enumerate(MODE_LABELS.items()): ttk.Radiobutton(box,text=label,value=value,variable=self.mode).grid(row=0,column=i,sticky="w",padx=(0,18))
        ttk.Label(self,text=f"Entrada configurada: {self.settings.input_dir}").grid(row=3,column=0,sticky="w",pady=(14,3))
        row=ttk.Frame(self); row.grid(row=4,column=0,sticky="ew"); row.columnconfigure(0,weight=1); ttk.Entry(row,textvariable=self.file,state="readonly").grid(row=0,column=0,sticky="ew"); ttk.Button(row,text="Seleccionar fichero…",command=self.select).grid(row=0,column=1,padx=(8,0))
        self.process_button=ttk.Button(self,text="Procesar",command=self.start); self.process_button.grid(row=5,column=0,sticky="w",pady=14)
        ttk.Label(self,textvariable=self.status,wraplength=700).grid(row=6,column=0,sticky="w")
        actions=ttk.Frame(self); actions.grid(row=7,column=0,sticky="w",pady=(12,0))
        self.open_button=ttk.Button(actions,text="Abrir último Excel",command=self.open_output,state="disabled"); self.open_button.grid(row=0,column=0)
        self.report_button=ttk.Button(actions,text="Abrir visor",command=self.open_report,state="disabled"); self.report_button.grid(row=0,column=1,padx=(8,0))
        ttk.Button(actions,text="Abrir carpeta de salida",command=self.open_folder).grid(row=0,column=2,padx=(8,0))
        self.refresh_candidates()
    def refresh_candidates(self):
        files=candidates(self.settings.input_dir) if self.settings.input_dir.is_dir() else []
        if len(files)==1: self.file.set(str(files[0]))
        elif files: self.status.set("Hay varios Excel en entrada: seleccione el que desea procesar.")
    def select(self):
        filename=filedialog.askopenfilename(initialdir=self.settings.input_dir,filetypes=[("Excel", "*.xlsx *.xlsm")])
        if filename: self.file.set(filename)
    def start(self):
        if not self.file.get(): messagebox.showwarning("Fichero necesario","Seleccione un Excel de entrada."); return
        self.process_button.configure(state="disabled"); self.status.set("Procesando… no cierre la aplicación.")
        threading.Thread(target=self.worker,args=(Path(self.file.get()),),daemon=True).start()
    def worker(self, source):
        try: result=run(self.settings,self.mode.get(),source); self.root.after(0,lambda:self.done(result))
        except Exception as exc: self.root.after(0, lambda exc=exc: self.failed(exc))
    def done(self,result):
        stats,output,report,log=result; self.last_output=output; self.last_report=report; self.status.set(f"Proceso terminado: {stats['groups']} grupos y {stats['proposed_rows']} partidas. Resultado: {output}"); self.process_button.configure(state="normal"); self.open_button.configure(state="normal"); self.report_button.configure(state="normal")
    def failed(self,exc): self.status.set("No se pudo completar el proceso. Consulte el log técnico."); self.process_button.configure(state="normal"); messagebox.showerror("Error de procesamiento",str(exc))
    def open_output(self):
        if self.last_output: os.startfile(self.last_output) if os.name=="nt" else None
    def open_report(self):
        if getattr(self,"last_report",None) and os.name=="nt": os.startfile(self.last_report)
    def open_folder(self):
        if os.name == "nt": os.startfile(self.settings.output_dir)
def main():
    root=tk.Tk(); MainWindow(root); root.mainloop()
