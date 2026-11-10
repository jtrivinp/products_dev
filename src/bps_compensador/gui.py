"""Interfaz ttk local para usuarios de negocio."""
from __future__ import annotations
import os, threading
from dataclasses import replace
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from .application import candidates, run
from .settings import load

class MainWindow(ttk.Frame):
    def __init__(self, root):
        super().__init__(root, padding=18); self.root=root; self.grid(sticky="nsew"); root.title("BPS — Propuestas de compensación"); root.columnconfigure(0, weight=1)
        self.settings = load(); self.input_path=tk.StringVar(value=str(self.settings.input_dir)); self.status=tk.StringVar(value="Listo para procesar los Excel pendientes."); self.last_output=None
        ttk.Label(self, text="Propuestas de compensación BPS", font=("Segoe UI", 16, "bold")).grid(row=0,column=0,sticky="w")
        ttk.Label(self, text="Las coincidencias son propuestas; la validación y ejecución final se realiza manualmente en SAP.", wraplength=700).grid(row=1,column=0,sticky="w",pady=(3,15))
        input_row=ttk.Frame(self); input_row.grid(row=3,column=0,sticky="ew",pady=(14,3)); input_row.columnconfigure(0,weight=1)
        ttk.Label(input_row,text="Carpeta de entrada:").grid(row=0,column=0,sticky="w")
        ttk.Entry(input_row,textvariable=self.input_path,state="readonly",width=75).grid(row=1,column=0,sticky="ew")
        ttk.Button(input_row,text="Cambiar ruta…",command=self.choose_input).grid(row=1,column=1,padx=(8,0))
        ttk.Label(self,text="Se procesarán todos los Excel pendientes (VIP y GPC). Los correctos se moverán a procesados/.").grid(row=4,column=0,sticky="w")
        self.process_button=ttk.Button(self,text="Procesar",command=self.start); self.process_button.grid(row=5,column=0,sticky="w",pady=14)
        ttk.Label(self,textvariable=self.status,wraplength=700).grid(row=6,column=0,sticky="w")
        actions=ttk.Frame(self); actions.grid(row=7,column=0,sticky="w",pady=(12,0))
        self.open_button=ttk.Button(actions,text="Abrir último Excel",command=self.open_output,state="disabled"); self.open_button.grid(row=0,column=0)
        self.report_button=ttk.Button(actions,text="Abrir visor",command=self.open_report,state="disabled"); self.report_button.grid(row=0,column=1,padx=(8,0))
        ttk.Button(actions,text="Abrir carpeta de salida",command=self.open_folder).grid(row=0,column=2,padx=(8,0))
        self.refresh_candidates()
    def refresh_candidates(self): pass
    def choose_input(self):
        folder=filedialog.askdirectory(initialdir=self.settings.input_dir)
        if folder:
            self.settings=replace(self.settings,input_dir=Path(folder)); self.input_path.set(folder)
    def start(self):
        self.process_button.configure(state="disabled"); self.status.set("Procesando… no cierre la aplicación.")
        threading.Thread(target=self.worker,daemon=True).start()
    def worker(self):
        try: result=run(self.settings); self.root.after(0,lambda:self.done(result))
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
