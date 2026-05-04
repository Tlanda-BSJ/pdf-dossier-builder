"""
PDF Dossier Builder - versión mejorada con PyMuPDF
- Fusiona PDFs siguiendo estructura recursiva existente
- Reescribe enlaces a PDFs como navegación interna correcta
- Compatible con Nitro PDF (/Launch, /GoToR, URI)
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import urllib.parse

import fitz  # PyMuPDF (CLAVE)

# ─────────────────────────────────────────────────────────────
# RUTAS
# ─────────────────────────────────────────────────────────────

def normalizar_ruta(ruta):
    if not ruta:
        return ""
    ruta = str(ruta)
    try:
        ruta = urllib.parse.unquote(ruta)
    except Exception:
        pass
    return ruta.replace("/", os.sep).replace("\\", os.sep).strip()


def resolver_ruta(ruta_raw, directorio_padre, directorio_raiz):
    ruta = normalizar_ruta(ruta_raw)
    if not ruta:
        return None

    # absoluta
    if os.path.isabs(ruta) and os.path.isfile(ruta):
        return os.path.normpath(ruta)

    # relativa
    candidato = os.path.normpath(os.path.join(directorio_padre, ruta))
    if os.path.isfile(candidato):
        return candidato

    # búsqueda global
    nombre = os.path.basename(ruta)
    for root, dirs, files in os.walk(directorio_raiz):
        if nombre in files:
            return os.path.normpath(os.path.join(root, nombre))

    return None


# ─────────────────────────────────────────────────────────────
# EXTRACCIÓN DE VÍNCULOS (TU LÓGICA CONSERVADA)
# ─────────────────────────────────────────────────────────────

def extraer_vinculos_archivos(pdf_path):
    from pypdf import PdfReader

    vinculos = []

    try:
        reader = PdfReader(pdf_path)

        for page in reader.pages:
            if "/Annots" not in page:
                continue

            for annot_ref in page["/Annots"]:
                try:
                    annot = annot_ref.get_object()
                except:
                    continue

                if annot.get("/Subtype") != "/Link":
                    continue

                action = annot.get("/A")
                if not action:
                    continue

                try:
                    action = action.get_object()
                except:
                    pass

                subtype = action.get("/S")

                # Launch
                if subtype == "/Launch":
                    f = action.get("/F") or action.get("/UF") or ""
                    vinculos.append(str(f))

                # GoToR
                elif subtype == "/GoToR":
                    f = action.get("/F") or ""
                    vinculos.append(str(f))

                # URI
                elif subtype == "/URI":
                    uri = action.get("/URI", "")
                    if isinstance(uri, bytes):
                        uri = uri.decode("utf-8", errors="ignore")
                    if str(uri).lower().endswith(".pdf"):
                        vinculos.append(str(uri))

    except:
        pass

    return vinculos


# ─────────────────────────────────────────────────────────────
# PROCESADO RECURSIVO (NO TOCADO)
# ─────────────────────────────────────────────────────────────

def procesar_pdf(pdf_path, directorio_raiz, visitados, lista_ordenada, log_cb, nivel=0):
    pdf_path = os.path.normpath(pdf_path)

    if pdf_path in visitados:
        log_cb(f"{'  '*nivel}↩ ignorado: {os.path.basename(pdf_path)}")
        return

    if not os.path.isfile(pdf_path):
        log_cb(f"{'  '*nivel}✗ no existe: {pdf_path}")
        return

    visitados.add(pdf_path)
    lista_ordenada.append(pdf_path)

    log_cb(f"{'  '*nivel}✔ {os.path.basename(pdf_path)}")

    directorio_padre = os.path.dirname(pdf_path)
    vinculos = extraer_vinculos_archivos(pdf_path)

    for v in vinculos:
        ruta = resolver_ruta(v, directorio_padre, directorio_raiz)
        if ruta:
            procesar_pdf(ruta, directorio_raiz, visitados, lista_ordenada, log_cb, nivel+1)


# ─────────────────────────────────────────────────────────────
# FUSIÓN CON OFFSETS (NUEVO CORE)
# ─────────────────────────────────────────────────────────────

def fusionar_con_offsets(lista_pdfs, log_cb):
    doc = fitz.open()
    offsets = {}

    pagina_actual = 0

    for pdf in lista_pdfs:
        try:
            d = fitz.open(pdf)

            offsets[os.path.normpath(pdf)] = pagina_actual

            doc.insert_pdf(d)

            pagina_actual += len(d)

        except Exception as e:
            log_cb(f"⚠ error {os.path.basename(pdf)}: {e}")

    return doc, offsets


# ─────────────────────────────────────────────────────────────
# REESCRIBIR ENLACES (CLAVE DEL PROYECTO)
# ─────────────────────────────────────────────────────────────

def reescribir_links(doc, offsets, log_cb):
    total = 0

    for i in range(len(doc)):
        page = doc[i]
        links = page.get_links()

        for link in links:
            uri = link.get("uri")

            if not uri:
                continue

            uri = str(uri)

            if ".pdf" not in uri.lower():
                continue

            nombre = os.path.normpath(os.path.basename(uri).split("#")[0])

            if nombre not in offsets:
                continue

            pagina_dest = 0

            if "#page=" in uri:
                try:
                    pagina_dest = int(uri.split("#page=")[1]) - 1
                except:
                    pass

            nueva_pagina = offsets[nombre] + pagina_dest

            page.insert_link({
                "kind": fitz.LINK_GOTO,
                "page": nueva_pagina,
                "from": link["from"]
            })

            page.delete_link(link)
            total += 1

    log_cb(f"✔ links corregidos: {total}")


# ─────────────────────────────────────────────────────────────
# INTERFAZ (SIN CAMBIOS IMPORTANTES)
# ─────────────────────────────────────────────────────────────

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("PDF Dossier Builder")
        self.geometry("600x400")

        self.ruta = tk.StringVar()

        tk.Entry(self, textvariable=self.ruta, width=60).pack(pady=10)
        tk.Button(self, text="Seleccionar", command=self.sel).pack()
        tk.Button(self, text="Generar", command=self.run).pack()

        self.log = tk.Text(self, height=15)
        self.log.pack(fill="both", expand=True)

    def sel(self):
        self.ruta.set(filedialog.askopenfilename())

    def logf(self, msg):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.update()

    def run(self):
        threading.Thread(target=self.process).start()

    def process(self):
        indice = self.ruta.get()
        raiz = os.path.dirname(indice)

        visitados = set()
        lista = []

        self.logf("Iniciando...")

        procesar_pdf(indice, raiz, visitados, lista, self.logf)

        self.logf("Fusionando...")

        doc, offsets = fusionar_con_offsets(lista, self.logf)

        self.logf("Reescribiendo links...")

        reescribir_links(doc, offsets, self.logf)

        salida = os.path.join(raiz, "DOSSIER_FINAL.pdf")
        doc.save(salida)

        self.logf("✔ listo: " + salida)
        messagebox.showinfo("OK", "PDF generado")


if __name__ == "__main__":
    App().mainloop()
