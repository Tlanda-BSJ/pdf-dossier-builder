"""
PDF Dossier Builder v2 - Compatible con Nitro PDF
- Fusiona todos los PDFs siguiendo vinculos /Launch y /GoToR
- Recalcula TODOS los hipervinculos para que apunten a la pagina
  correcta dentro del PDF fusionado final
- Ignora vinculos de retroceso automaticamente
- Sin instalacion requerida (empaquetado en .exe)
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import sys
import urllib.parse
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    DictionaryObject, ArrayObject, NumberObject,
    NameObject, NullObject
)


# ─── RESOLUCIÓN DE RUTAS ────────────────────────────────────────────────────

def normalizar_ruta(ruta):
    if not ruta:
        return ""
    ruta = str(ruta)
    try:
        ruta = urllib.parse.unquote(ruta)
    except Exception:
        pass
    return ruta.replace("/", os.sep).replace("\\", os.sep).strip()


def resolver_ruta(ruta_raw, dir_padre, dir_raiz):
    """
    Resuelve la ruta con prioridad:
    1. Absoluta directa
    2. Relativa al PDF padre
    3. Busqueda por nombre en el arbol de carpetas
    """
    ruta = normalizar_ruta(ruta_raw)
    if not ruta:
        return None

    if os.path.isabs(ruta) and os.path.isfile(ruta):
        return os.path.normpath(ruta)

    candidato = os.path.normpath(os.path.join(dir_padre, ruta))
    if os.path.isfile(candidato):
        return candidato

    nombre = os.path.basename(ruta)
    for root, dirs, files in os.walk(dir_raiz):
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        if nombre in files:
            return os.path.normpath(os.path.join(root, nombre))

    return None


# ─── EXTRACCIÓN DE VÍNCULOS ─────────────────────────────────────────────────

def extraer_vinculos_pdf(pdf_path):
    """
    Extrae vinculos /Launch y /GoToR que abren archivos .pdf externos.
    Ignora /GoTo (saltos internos) y vinculos a no-PDF.
    Devuelve lista de rutas raw sin resolver.
    """
    vinculos = []
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            if "/Annots" not in page:
                continue
            for annot_ref in page["/Annots"]:
                try:
                    annot = annot_ref.get_object()
                except Exception:
                    continue
                if annot.get("/Subtype") != "/Link":
                    continue
                action = annot.get("/A")
                if not action:
                    continue
                try:
                    action = action.get_object()
                except Exception:
                    pass

                subtype = action.get("/S")
                ruta = None

                if subtype == "/Launch":
                    win = action.get("/Win")
                    if win:
                        try:
                            win = win.get_object()
                        except Exception:
                            pass
                        f = win.get("/F") or win.get("/UF") or ""
                    else:
                        f = action.get("/F") or action.get("/UF") or ""
                    if hasattr(f, "get_object"):
                        try:
                            f = f.get_object()
                        except Exception:
                            pass
                    if isinstance(f, dict):
                        f = f.get("/F") or f.get("/UF") or ""
                    ruta = str(f).strip()

                elif subtype == "/GoToR":
                    f = action.get("/F") or ""
                    if hasattr(f, "get_object"):
                        try:
                            f = f.get_object()
                        except Exception:
                            pass
                    if isinstance(f, dict):
                        f = f.get("/F") or f.get("/UF") or ""
                    ruta = str(f).strip()

                # /GoTo → interno, ignorar
                # /URI → ignorar

                if ruta and ruta.lower().endswith(".pdf"):
                    vinculos.append(ruta)
    except Exception:
        pass
    return vinculos


# ─── PROCESADO RECURSIVO ────────────────────────────────────────────────────

def procesar_pdf(pdf_path, dir_raiz, visitados, lista, log_cb, nivel=0):
    """
    Sigue recursivamente los vinculos de un PDF.
    Los retrocesos (PDF ya visitado) se ignoran automaticamente.
    """
    pdf_path = os.path.normpath(pdf_path)

    if pdf_path in visitados:
        log_cb(f"{'  ' * nivel}retroceso ignorado: {os.path.basename(pdf_path)}")
        return
    if not os.path.isfile(pdf_path):
        log_cb(f"{'  ' * nivel}NO ENCONTRADO: {pdf_path}")
        return

    visitados.add(pdf_path)
    lista.append(pdf_path)
    log_cb(f"{'  ' * nivel}OK  {os.path.relpath(pdf_path, dir_raiz)}")

    dir_padre = os.path.dirname(pdf_path)
    for ruta_raw in extraer_vinculos_pdf(pdf_path):
        res = resolver_ruta(ruta_raw, dir_padre, dir_raiz)
        if res:
            procesar_pdf(res, dir_raiz, visitados, lista, log_cb, nivel + 1)
        else:
            log_cb(f"{'  ' * (nivel + 1)}no resuelto: {os.path.basename(ruta_raw)}")


# ─── CÁLCULO DE OFFSETS ─────────────────────────────────────────────────────

def calcular_offsets(lista_pdfs):
    """
    Devuelve dict {ruta_norm -> pagina_inicio_base0} y total de paginas.
    """
    pdf_a_inicio = {}
    offset = 0
    for p in lista_pdfs:
        pdf_a_inicio[os.path.normpath(p)] = offset
        try:
            offset += len(PdfReader(p).pages)
        except Exception:
            pass
    return pdf_a_inicio, offset


# ─── FUSIÓN CON RECÁLCULO DE VÍNCULOS ───────────────────────────────────────

def fusionar_con_vinculos(lista_pdfs, salida_path, dir_raiz, log_cb):
    """
    1. Fusiona todos los PDFs en uno.
    2. Reescribe todos los vinculos /Launch y /GoToR como /GoTo internos
       apuntando a la pagina correcta dentro del dossier.
    3. Vinculos a PDFs no incluidos en el dossier se eliminan.
    """
    pdf_a_inicio, total = calcular_offsets(lista_pdfs)
    log_cb(f"Total paginas en dossier: {total}")
    log_cb("Fusionando y recalculando vinculos...")

    writer = PdfWriter()

    # 1. Añadir todas las paginas
    for pdf_path in lista_pdfs:
        try:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            log_cb(f"  Error leyendo {os.path.basename(pdf_path)}: {e}")

    # Obtener el array de referencias a paginas del writer
    kids = writer._pages.get_object()["/Kids"]

    # 2. Recorrer y reescribir anotaciones
    pagina_global = 0
    vinculos_recalculados = 0
    vinculos_eliminados = 0

    for pdf_path in lista_pdfs:
        dir_padre = os.path.dirname(pdf_path)
        try:
            num_paginas = len(PdfReader(pdf_path).pages)
        except Exception:
            continue

        for i in range(num_paginas):
            page = writer.pages[pagina_global + i]

            if "/Annots" not in page:
                continue

            nuevas = ArrayObject()
            for annot_ref in page["/Annots"]:
                try:
                    annot = annot_ref.get_object()
                except Exception:
                    continue

                if annot.get("/Subtype") != "/Link":
                    nuevas.append(annot_ref)
                    continue

                action = annot.get("/A")
                if not action:
                    nuevas.append(annot_ref)
                    continue
                try:
                    action = action.get_object()
                except Exception:
                    pass

                subtype = action.get("/S")

                # Solo procesar vinculos externos
                if subtype not in ("/Launch", "/GoToR"):
                    nuevas.append(annot_ref)
                    continue

                # Extraer ruta destino
                ruta_raw = None
                if subtype == "/Launch":
                    win = action.get("/Win")
                    if win:
                        try:
                            win = win.get_object()
                        except Exception:
                            pass
                        f = win.get("/F") or win.get("/UF") or ""
                    else:
                        f = action.get("/F") or action.get("/UF") or ""
                    if hasattr(f, "get_object"):
                        try:
                            f = f.get_object()
                        except Exception:
                            pass
                    if isinstance(f, dict):
                        f = f.get("/F") or f.get("/UF") or ""
                    ruta_raw = str(f).strip()

                elif subtype == "/GoToR":
                    f = action.get("/F") or ""
                    if hasattr(f, "get_object"):
                        try:
                            f = f.get_object()
                        except Exception:
                            pass
                    if isinstance(f, dict):
                        f = f.get("/F") or f.get("/UF") or ""
                    ruta_raw = str(f).strip()

                if not ruta_raw or not ruta_raw.lower().endswith(".pdf"):
                    nuevas.append(annot_ref)
                    continue

                # Resolver ruta
                res = resolver_ruta(ruta_raw, dir_padre, dir_raiz)
                if res:
                    res_norm = os.path.normpath(res)
                else:
                    res_norm = None

                if res_norm and res_norm in pdf_a_inicio:
                    # Calcular pagina destino y construir GoTo interno
                    pag_dest = pdf_a_inicio[res_norm]
                    page_ref = kids[pag_dest]  # referencia indirecta correcta

                    nueva_accion = DictionaryObject({
                        NameObject("/S"): NameObject("/GoTo"),
                        NameObject("/D"): ArrayObject([
                            page_ref,
                            NameObject("/XYZ"),
                            NullObject(),
                            NullObject(),
                            NullObject(),
                        ])
                    })

                    # Reconstruir anotacion conservando apariencia
                    nueva_annot = DictionaryObject()
                    for k, v in annot.items():
                        if k != "/A":
                            nueva_annot[NameObject(k)] = v
                    nueva_annot[NameObject("/A")] = nueva_accion
                    nuevas.append(nueva_annot)
                    vinculos_recalculados += 1

                else:
                    # PDF no incluido en dossier → eliminar vinculo
                    vinculos_eliminados += 1

            page[NameObject("/Annots")] = nuevas

        pagina_global += num_paginas

    log_cb(f"Vinculos recalculados: {vinculos_recalculados}")
    if vinculos_eliminados:
        log_cb(f"Vinculos eliminados (PDF no incluido): {vinculos_eliminados}")

    with open(salida_path, "wb") as f:
        writer.write(f)

    log_cb(f"Guardado: {os.path.basename(salida_path)}")
    return vinculos_recalculados


# ─── INTERFAZ GRÁFICA ───────────────────────────────────────────────────────

class App(tk.Tk):

    BG      = "#1e1e2e"
    PANEL   = "#2a2a3d"
    ACCENT  = "#0078d4"
    SUCCESS = "#28a745"
    FG      = "#e0e0e0"
    FG_DIM  = "#888888"
    ENTRY   = "#313148"

    def __init__(self):
        super().__init__()
        self.title("PDF Dossier Builder v2")
        self.configure(bg=self.BG)
        self.resizable(False, False)
        self._build_ui()

    def _build_ui(self):
        # Cabecera
        hdr = tk.Frame(self, bg=self.ACCENT)
        hdr.grid(row=0, column=0, sticky="ew")
        tk.Label(hdr, text="  PDF Dossier Builder",
                 font=("Segoe UI", 14, "bold"),
                 bg=self.ACCENT, fg="white",
                 anchor="w").pack(fill="x", padx=12, pady=10)

        tk.Label(self,
                 text="Selecciona el indice general y genera el dossier completo\n"
                      "con todos los hipervinculos recalculados automaticamente.",
                 font=("Segoe UI", 9), bg=self.BG, fg=self.FG_DIM,
                 justify="left").grid(row=1, column=0, sticky="w", padx=20, pady=(12, 4))

        # Selector de archivo
        frame_sel = tk.Frame(self, bg=self.PANEL)
        frame_sel.grid(row=2, column=0, sticky="ew", padx=20, pady=4)
        frame_sel.columnconfigure(1, weight=1)

        tk.Label(frame_sel, text="Indice general:",
                 font=("Segoe UI", 10, "bold"),
                 bg=self.PANEL, fg=self.FG
                 ).grid(row=0, column=0, padx=(12, 8), pady=10)

        self.ruta_var = tk.StringVar()
        tk.Entry(frame_sel, textvariable=self.ruta_var,
                 width=50, bg=self.ENTRY, fg=self.FG,
                 insertbackground=self.FG, relief="flat",
                 font=("Segoe UI", 9)
                 ).grid(row=0, column=1, padx=4, pady=10, sticky="ew")

        tk.Button(frame_sel, text="Examinar...",
                  command=self._seleccionar,
                  bg=self.ACCENT, fg="white", relief="flat",
                  font=("Segoe UI", 9), padx=10, cursor="hand2"
                  ).grid(row=0, column=2, padx=(4, 12), pady=10)

        # Boton principal
        self.btn = tk.Button(self, text="Generar Dossier",
                             command=self._iniciar,
                             bg=self.SUCCESS, fg="white", relief="flat",
                             font=("Segoe UI", 12, "bold"),
                             padx=20, pady=10, cursor="hand2")
        self.btn.grid(row=3, column=0, pady=(10, 6))

        # Barra de progreso
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("accent.Horizontal.TProgressbar",
                        troughcolor=self.ENTRY, background=self.ACCENT,
                        bordercolor=self.BG, lightcolor=self.ACCENT,
                        darkcolor=self.ACCENT)
        self.progress = ttk.Progressbar(self, mode="indeterminate",
                                        length=500,
                                        style="accent.Horizontal.TProgressbar")
        self.progress.grid(row=4, column=0, padx=20, pady=(0, 4))

        # Log
        tk.Label(self, text="Registro de proceso:",
                 font=("Segoe UI", 9),
                 bg=self.BG, fg=self.FG_DIM
                 ).grid(row=5, column=0, sticky="w", padx=20)

        frame_log = tk.Frame(self, bg=self.BG)
        frame_log.grid(row=6, column=0, padx=20, pady=(2, 4), sticky="ew")

        self.log = tk.Text(frame_log, width=70, height=18,
                           bg=self.ENTRY, fg=self.FG,
                           font=("Consolas", 8), relief="flat",
                           state="disabled", wrap="none")
        sb = tk.Scrollbar(frame_log, command=self.log.yview, bg=self.ENTRY)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both")
        sb.pack(side="right", fill="y")

        # Estado
        self.estado = tk.StringVar(value="Listo")
        tk.Label(self, textvariable=self.estado,
                 font=("Segoe UI", 9, "italic"),
                 bg=self.BG, fg=self.FG_DIM
                 ).grid(row=7, column=0, pady=(0, 14))

    def _seleccionar(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el indice general PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if ruta:
            self.ruta_var.set(ruta)

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def _iniciar(self):
        ruta = self.ruta_var.get().strip()
        if not ruta or not os.path.isfile(ruta):
            messagebox.showerror("Error", "Selecciona un fichero PDF valido.")
            return
        self.btn.configure(state="disabled")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.progress.start(10)
        self.estado.set("Procesando...")
        threading.Thread(target=self._proceso, args=(ruta,), daemon=True).start()

    def _proceso(self, indice_path):
        try:
            dir_raiz = os.path.dirname(indice_path)
            visitados = set()
            lista = []

            self._log("=" * 62)
            self._log(f"Indice:  {os.path.basename(indice_path)}")
            self._log(f"Carpeta: {dir_raiz}")
            self._log("=" * 62)
            self._log("Resolviendo vinculos...\n")

            procesar_pdf(indice_path, dir_raiz, visitados, lista, self._log)

            self._log(f"\n{len(lista)} PDFs encontrados")
            self._log("=" * 62)

            nombre = os.path.splitext(os.path.basename(indice_path))[0]
            salida = os.path.join(dir_raiz, nombre + "_DOSSIER_COMPLETO.pdf")

            n = fusionar_con_vinculos(lista, salida, dir_raiz, self._log)

            self._log("=" * 62)
            self.estado.set(f"Completado - {len(lista)} docs, {n} vinculos recalculados")
            messagebox.showinfo(
                "Completado",
                f"Dossier generado con {len(lista)} documentos.\n"
                f"{n} hipervinculos recalculados correctamente.\n\n"
                f"Guardado en:\n{salida}"
            )
        except Exception as e:
            import traceback
            self._log(f"\nError: {e}")
            self._log(traceback.format_exc())
            self.estado.set("Error durante el proceso")
            messagebox.showerror("Error inesperado", str(e))
        finally:
            self.progress.stop()
            self.btn.configure(state="normal")


if __name__ == "__main__":
    app = App()
    app.mainloop()
