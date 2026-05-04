"""
PDF Dossier Builder - Compatible con Nitro PDF
- Extrae vínculos /Launch (Abrir archivo) e ignora /GoTo (saltos de página)
- Resuelve rutas absolutas Windows; si no existen, busca relativo a la carpeta del PDF padre
- Ignora vínculos de retroceso (PDF ya visitado)
- Sin instalación requerida (todo empaquetado en el .exe)
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import sys
import urllib.parse
from pypdf import PdfReader, PdfWriter


# ─── RESOLUCIÓN DE RUTAS ────────────────────────────────────────────────────

def normalizar_ruta(ruta):
    """Limpia la ruta: decodifica URL encoding y normaliza separadores."""
    if not ruta:
        return ""
    ruta = str(ruta)
    try:
        ruta = urllib.parse.unquote(ruta)
    except Exception:
        pass
    ruta = ruta.replace("/", os.sep).replace("\\", os.sep).strip()
    return ruta


def resolver_ruta(ruta_raw, directorio_padre, directorio_raiz):
    """
    Intenta resolver la ruta del vínculo con esta prioridad:
    1. Ruta absoluta tal cual (si existe)
    2. Ruta relativa desde el directorio del PDF padre
    3. Búsqueda del nombre de archivo en todo el árbol desde directorio_raiz
    Devuelve la ruta resuelta o None si no se encuentra.
    """
    ruta = normalizar_ruta(ruta_raw)
    if not ruta:
        return None

    # 1. Ruta absoluta directa
    if os.path.isabs(ruta) and os.path.isfile(ruta):
        return os.path.normpath(ruta)

    # 2. Relativa al directorio del PDF padre
    candidato = os.path.normpath(os.path.join(directorio_padre, ruta))
    if os.path.isfile(candidato):
        return candidato

    # 3. Solo el nombre de archivo buscado en el árbol de directorio_raiz
    nombre = os.path.basename(ruta)
    for root, dirs, files in os.walk(directorio_raiz):
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        if nombre in files:
            return os.path.normpath(os.path.join(root, nombre))

    return None


# ─── EXTRACCIÓN DE VÍNCULOS ─────────────────────────────────────────────────

def extraer_vinculos_archivos(pdf_path):
    """
    Extrae SOLO vínculos que abren archivos externos (/Launch o /GoToR).
    Ignora /GoTo (saltos de página internos).
    Devuelve lista de rutas raw (sin resolver).
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

                # /Launch → Nitro usa esto para "Abrir archivo"
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
                    if ruta.lower().endswith(".pdf"):
                        vinculos.append(ruta)

                # /GoToR → vínculo a otro PDF con destino específico
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
                    if ruta.lower().endswith(".pdf"):
                        vinculos.append(ruta)

                # /GoTo → salto de página INTERNO, ignorar siempre
                elif subtype == "/GoTo":
                    pass

                # URI → solo si apunta a un .pdf local (no http)
                elif subtype == "/URI":
                    uri = action.get("/URI", "")
                    if isinstance(uri, bytes):
                        uri = uri.decode("utf-8", errors="ignore")
                    uri = str(uri).strip()
                    if uri.lower().endswith(".pdf") and not uri.lower().startswith("http"):
                        vinculos.append(uri)

    except Exception:
        pass

    return vinculos


# ─── PROCESADO RECURSIVO ────────────────────────────────────────────────────

def procesar_pdf(pdf_path, directorio_raiz, visitados, lista_ordenada, log_cb, nivel=0):
    """
    Procesa un PDF: lo añade a la lista y sigue recursivamente sus vínculos.
    Los vínculos a PDFs ya visitados (retroceso) se ignoran automáticamente.
    """
    pdf_path = os.path.normpath(pdf_path)

    if pdf_path in visitados:
        log_cb(f"{'  ' * nivel}↩ Retroceso ignorado: {os.path.basename(pdf_path)}")
        return

    if not os.path.isfile(pdf_path):
        log_cb(f"{'  ' * nivel}✗ No encontrado: {pdf_path}")
        return

    visitados.add(pdf_path)
    lista_ordenada.append(pdf_path)
    log_cb(f"{'  ' * nivel}✔ {os.path.relpath(pdf_path, directorio_raiz)}")

    directorio_padre = os.path.dirname(pdf_path)
    vinculos_raw = extraer_vinculos_archivos(pdf_path)

    for ruta_raw in vinculos_raw:
        ruta_resuelta = resolver_ruta(ruta_raw, directorio_padre, directorio_raiz)
        if ruta_resuelta:
            procesar_pdf(ruta_resuelta, directorio_raiz, visitados,
                         lista_ordenada, log_cb, nivel + 1)
        else:
            log_cb(f"{'  ' * (nivel+1)}✗ No resuelto: {os.path.basename(ruta_raw)}")


# ─── FUSIÓN ─────────────────────────────────────────────────────────────────

def fusionar_pdfs(lista_pdfs, salida_path, log_cb):
    writer = PdfWriter()
    errores = 0
    for pdf_path in lista_pdfs:
        try:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            log_cb(f"  ⚠ Error leyendo {os.path.basename(pdf_path)}: {e}")
            errores += 1
    with open(salida_path, "wb") as f:
        writer.write(f)
    return errores


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
        self.title("PDF Dossier Builder")
        self.configure(bg=self.BG)
        self.resizable(False, False)
        self._build_ui()

    def _build_ui(self):
        # ── Cabecera ──
        hdr = tk.Frame(self, bg=self.ACCENT)
        hdr.grid(row=0, column=0, sticky="ew")
        tk.Label(hdr, text="  PDF Dossier Builder",
                 font=("Segoe UI", 14, "bold"),
                 bg=self.ACCENT, fg="white",
                 anchor="w").pack(fill="x", padx=12, pady=10)

        tk.Label(self,
                 text="Selecciona el indice general y genera el dossier completo\n"
                      "siguiendo todos los vinculos automaticamente.",
                 font=("Segoe UI", 9), bg=self.BG, fg=self.FG_DIM,
                 justify="left").grid(row=1, column=0, sticky="w", padx=20, pady=(12, 4))

        # ── Selector ──
        frame_sel = tk.Frame(self, bg=self.PANEL)
        frame_sel.grid(row=2, column=0, sticky="ew", padx=20, pady=4)
        frame_sel.columnconfigure(1, weight=1)

        tk.Label(frame_sel, text="Indice general:",
                 font=("Segoe UI", 10, "bold"),
                 bg=self.PANEL, fg=self.FG).grid(row=0, column=0, padx=(12, 8), pady=10)

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

        # ── Botón principal ──
        self.btn = tk.Button(self, text="Generar Dossier",
                             command=self._iniciar,
                             bg=self.SUCCESS, fg="white", relief="flat",
                             font=("Segoe UI", 12, "bold"),
                             padx=20, pady=10, cursor="hand2")
        self.btn.grid(row=3, column=0, pady=(10, 6))

        # ── Barra de progreso ──
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

        # ── Log ──
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

        # ── Estado ──
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

            procesar_pdf(indice_path, dir_raiz, visitados, lista, self._log)

            self._log("")
            self._log("=" * 62)
            self._log(f"PDFs encontrados: {len(lista)}")
            self._log("Fusionando documentos...")

            nombre = os.path.splitext(os.path.basename(indice_path))[0]
            salida = os.path.join(dir_raiz, nombre + "_DOSSIER_COMPLETO.pdf")
            errores = fusionar_pdfs(lista, salida, self._log)

            self._log(f"")
            self._log(f"Guardado: {os.path.basename(salida)}")
            if errores:
                self._log(f"AVISO: {errores} archivo(s) con errores al leer.")
            self._log("=" * 62)

            self.estado.set(f"Completado - {len(lista)} documentos fusionados")
            messagebox.showinfo(
                "Completado",
                f"Dossier generado con {len(lista)} documentos.\n\n"
                f"Guardado en:\n{salida}"
            )
        except Exception as e:
            self._log(f"\nError inesperado: {e}")
            self.estado.set("Error durante el proceso")
            messagebox.showerror("Error inesperado", str(e))
        finally:
            self.progress.stop()
            self.btn.configure(state="normal")


if __name__ == "__main__":
    app = App()
    app.mainloop()
