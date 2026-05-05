"""
PDF Dossier Builder v3 - Edicion Profesional
- Portada automatica bilingue con tabla de contenidos
- Marcadores/Bookmarks navegables en panel lateral
- Numeracion de paginas (pie de pagina, bilingue)
- Vista previa del arbol con reordenacion y exclusion
- Barra de progreso real documento a documento
- Log exportable a .txt
- Compatible con Nitro PDF (/Launch, /GoToR)
- Rutas absolutas y relativas, busqueda por nombre
- Retrocesos ignorados automaticamente
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import sys
import io
import datetime
import urllib.parse

from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
    DictionaryObject, ArrayObject, NumberObject,
    NameObject, NullObject, TextStringObject, BooleanObject
)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas as rl_canvas


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
                if ruta and ruta.lower().endswith(".pdf"):
                    vinculos.append(ruta)
    except Exception:
        pass
    return vinculos


# ─── PROCESADO RECURSIVO ────────────────────────────────────────────────────

def procesar_pdf(pdf_path, dir_raiz, visitados, lista, log_cb, nivel=0):
    pdf_path = os.path.normpath(pdf_path)
    if pdf_path in visitados:
        return
    if not os.path.isfile(pdf_path):
        log_cb(f"{'  ' * nivel}NO ENCONTRADO: {pdf_path}")
        return
    visitados.add(pdf_path)

    try:
        num_pags = len(PdfReader(pdf_path).pages)
    except Exception:
        num_pags = 0

    lista.append({
        "path": pdf_path,
        "nombre": os.path.splitext(os.path.basename(pdf_path))[0],
        "paginas": num_pags,
        "nivel": nivel,
        "incluir": True,
    })
    log_cb(f"{'  ' * nivel}OK  {os.path.relpath(pdf_path, dir_raiz)}")

    dir_padre = os.path.dirname(pdf_path)
    for ruta_raw in extraer_vinculos_pdf(pdf_path):
        res = resolver_ruta(ruta_raw, dir_padre, dir_raiz)
        if res:
            procesar_pdf(res, dir_raiz, visitados, lista, log_cb, nivel + 1)
        else:
            log_cb(f"{'  ' * (nivel + 1)}no resuelto: {os.path.basename(ruta_raw)}")


# ─── PORTADA ────────────────────────────────────────────────────────────────

def generar_portada(titulo_dossier, entradas_toc, dir_raiz):
    """
    Genera la portada + tabla de contenidos como PDF en memoria.
    Devuelve bytes del PDF.
    entradas_toc: lista de dicts {nombre, pagina_dossier, nivel}
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        rightMargin=20*mm,
        leftMargin=20*mm,
        topMargin=25*mm,
        bottomMargin=25*mm,
    )

    estilos = getSampleStyleSheet()
    W, H = A4

    # Estilos personalizados
    estilo_titulo = ParagraphStyle(
        "Titulo",
        parent=estilos["Normal"],
        fontSize=26,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#1a3a5c"),
        alignment=TA_CENTER,
        spaceAfter=6,
    )
    estilo_subtitulo = ParagraphStyle(
        "Subtitulo",
        parent=estilos["Normal"],
        fontSize=13,
        fontName="Helvetica",
        textColor=colors.HexColor("#4a4a4a"),
        alignment=TA_CENTER,
        spaceAfter=4,
    )
    estilo_fecha = ParagraphStyle(
        "Fecha",
        parent=estilos["Normal"],
        fontSize=10,
        fontName="Helvetica",
        textColor=colors.HexColor("#888888"),
        alignment=TA_CENTER,
        spaceAfter=2,
    )
    estilo_toc_header = ParagraphStyle(
        "TocHeader",
        parent=estilos["Normal"],
        fontSize=14,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#1a3a5c"),
        alignment=TA_CENTER,
        spaceBefore=10,
        spaceAfter=6,
    )
    estilo_toc_0 = ParagraphStyle(
        "Toc0",
        parent=estilos["Normal"],
        fontSize=10,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#1a3a5c"),
        leftIndent=0,
        spaceAfter=2,
    )
    estilo_toc_1 = ParagraphStyle(
        "Toc1",
        parent=estilos["Normal"],
        fontSize=9,
        fontName="Helvetica",
        textColor=colors.HexColor("#333333"),
        leftIndent=12,
        spaceAfter=1,
    )
    estilo_toc_2 = ParagraphStyle(
        "Toc2",
        parent=estilos["Normal"],
        fontSize=8,
        fontName="Helvetica",
        textColor=colors.HexColor("#666666"),
        leftIndent=24,
        spaceAfter=1,
    )

    fecha_hoy = datetime.datetime.now()
    fecha_str = fecha_hoy.strftime("%d / %m / %Y")

    elementos = []

    # Espacio superior
    elementos.append(Spacer(1, 18*mm))

    # Línea decorativa superior
    elementos.append(HRFlowable(
        width="100%", thickness=3,
        color=colors.HexColor("#0078d4"), spaceAfter=8*mm
    ))

    # Título
    elementos.append(Paragraph(titulo_dossier.upper(), estilo_titulo))
    elementos.append(Spacer(1, 3*mm))

    # Subtítulo bilingüe
    elementos.append(Paragraph("TECHNICAL DOSSIER / DOSSIER TÉCNICO", estilo_subtitulo))
    elementos.append(Spacer(1, 6*mm))

    # Fecha
    elementos.append(Paragraph(f"Date / Fecha: {fecha_str}", estilo_fecha))
    elementos.append(Spacer(1, 3*mm))

    # Línea decorativa inferior
    elementos.append(HRFlowable(
        width="100%", thickness=1,
        color=colors.HexColor("#cccccc"), spaceAfter=10*mm
    ))

    # Encabezado tabla de contenidos bilingüe
    elementos.append(Paragraph(
        "TABLE OF CONTENTS / ÍNDICE DE CONTENIDOS",
        estilo_toc_header
    ))
    elementos.append(HRFlowable(
        width="80%", thickness=0.5,
        color=colors.HexColor("#0078d4"), spaceAfter=5*mm
    ))

    # Entradas TOC
    estilos_nivel = [estilo_toc_0, estilo_toc_1, estilo_toc_2]
    for entrada in entradas_toc:
        nivel = min(entrada["nivel"], 2)
        nombre = entrada["nombre"]
        pagina = entrada["pagina_dossier"] + 2  # +1 portada +1 base1
        estilo_actual = estilos_nivel[nivel]

        # Fila: nombre ... página
        datos = [[
            Paragraph(nombre, estilo_actual),
            Paragraph(str(pagina), ParagraphStyle(
                "PagNum",
                parent=estilo_actual,
                alignment=TA_RIGHT,
            ))
        ]]
        t = Table(datos, colWidths=["85%", "15%"])
        t.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ]))
        elementos.append(t)

    elementos.append(Spacer(1, 10*mm))
    elementos.append(HRFlowable(
        width="100%", thickness=1,
        color=colors.HexColor("#0078d4"), spaceAfter=3*mm
    ))

    doc.build(elementos)
    buf.seek(0)
    return buf.read()


# ─── NUMERACIÓN DE PÁGINAS ───────────────────────────────────────────────────

def generar_pagina_numero(total_paginas, ancho, alto):
    """
    Genera un PDF de una sola página con el pie de numeración.
    Se superpone sobre cada página del dossier.
    Devuelve bytes.
    """
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(ancho, alto))
    c.setFont("Helvetica", 7)
    c.setFillColor(colors.HexColor("#888888"))
    return buf, c


def anadir_numeracion(writer, offset_paginas, total_paginas):
    """
    Añade numeracion de paginas como overlay en cada pagina del writer
    a partir de offset_paginas.
    """
    from reportlab.pdfgen import canvas as rl_canvas
    from pypdf import PdfReader as PR

    paginas_writer = len(writer.pages)

    for i in range(paginas_writer):
        page = writer.pages[i]
        # Obtener tamaño de página
        mb = page.mediabox
        ancho = float(mb.width)
        alto = float(mb.height)

        num_pag = offset_paginas + i + 1  # numero real en dossier
        texto_en = f"Page {num_pag} of {total_paginas}"
        texto_es = f"Página {num_pag} de {total_paginas}"
        texto = f"{texto_en}   |   {texto_es}"

        # Crear overlay con reportlab
        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(ancho, alto))
        c.setFont("Helvetica", 7)
        c.setFillColor(colors.HexColor("#888888"))
        # Centrar en la parte inferior
        c.drawCentredString(ancho / 2, 8*mm, texto)
        # Línea fina sobre el texto
        c.setStrokeColor(colors.HexColor("#cccccc"))
        c.setLineWidth(0.3)
        c.line(20*mm, 12*mm, ancho - 20*mm, 12*mm)
        c.save()
        buf.seek(0)

        overlay_reader = PR(buf)
        overlay_page = overlay_reader.pages[0]
        page.merge_page(overlay_page)


# ─── FUSIÓN COMPLETA ────────────────────────────────────────────────────────

def calcular_offsets(lista_items):
    pdf_a_inicio = {}
    offset = 0
    for item in lista_items:
        if not item["incluir"]:
            continue
        pdf_a_inicio[os.path.normpath(item["path"])] = offset
        offset += item["paginas"]
    return pdf_a_inicio, offset


def fusionar_dossier(lista_items, salida_path, dir_raiz, titulo, log_cb, progress_cb):
    """
    Pipeline completo:
    1. Portada + TOC
    2. Fusion de PDFs con recalculo de vinculos
    3. Numeracion de paginas
    4. Marcadores/Bookmarks
    """
    items_activos = [it for it in lista_items if it["incluir"]]
    pdf_a_inicio, total_content = calcular_offsets(lista_items)
    total_con_portada = total_content + 1  # +1 portada

    # ── 1. PORTADA ──
    log_cb("Generando portada...")
    entradas_toc = []
    for it in items_activos:
        entradas_toc.append({
            "nombre": it["nombre"],
            "nivel": it["nivel"],
            "pagina_dossier": pdf_a_inicio[os.path.normpath(it["path"])] + 1,  # +1 portada
        })

    portada_bytes = generar_portada(titulo, entradas_toc, dir_raiz)
    portada_reader = PdfReader(io.BytesIO(portada_bytes))
    num_pags_portada = len(portada_reader.pages)

    # Recalcular offsets reales (portada puede tener más de 1 página)
    pdf_a_inicio_real = {}
    offset = num_pags_portada
    for it in items_activos:
        pdf_a_inicio_real[os.path.normpath(it["path"])] = offset
        offset += it["paginas"]
    total_real = offset

    # ── 2. FUSIÓN ──
    log_cb("Fusionando documentos...")
    writer = PdfWriter()

    # Añadir portada
    for page in portada_reader.pages:
        writer.add_page(page)

    # Añadir documentos
    total_items = len(items_activos)
    for idx, item in enumerate(items_activos):
        progress_cb(idx + 1, total_items, item["nombre"])
        log_cb(f"  [{idx+1}/{total_items}] {item['nombre']}")
        try:
            reader = PdfReader(item["path"])
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            log_cb(f"    Error: {e}")

    # ── 3. RECÁLCULO DE VÍNCULOS ──
    log_cb("Recalculando hipervinculos...")
    kids = writer._pages.get_object()["/Kids"]
    vinculos_ok = 0
    vinculos_eliminados = 0

    pagina_global = num_pags_portada
    for item in items_activos:
        dir_padre = os.path.dirname(item["path"])
        num_p = item["paginas"]
        for i in range(num_p):
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
                if subtype not in ("/Launch", "/GoToR"):
                    nuevas.append(annot_ref)
                    continue

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

                res = resolver_ruta(ruta_raw, dir_padre, dir_raiz)
                res_norm = os.path.normpath(res) if res else None

                if res_norm and res_norm in pdf_a_inicio_real:
                    pag_dest = pdf_a_inicio_real[res_norm]
                    nueva_accion = DictionaryObject({
                        NameObject("/S"): NameObject("/GoTo"),
                        NameObject("/D"): ArrayObject([
                            kids[pag_dest],
                            NameObject("/XYZ"),
                            NullObject(),
                            NullObject(),
                            NullObject(),
                        ])
                    })
                    nueva_annot = DictionaryObject()
                    for k, v in annot.items():
                        if k != "/A":
                            nueva_annot[NameObject(k)] = v
                    nueva_annot[NameObject("/A")] = nueva_accion
                    nuevas.append(nueva_annot)
                    vinculos_ok += 1
                else:
                    vinculos_eliminados += 1

            page[NameObject("/Annots")] = nuevas
        pagina_global += num_p

    log_cb(f"  Vinculos recalculados: {vinculos_ok}")
    if vinculos_eliminados:
        log_cb(f"  Vinculos eliminados (no incluidos): {vinculos_eliminados}")

    # ── 4. NUMERACIÓN DE PÁGINAS ──
    log_cb("Anadiendo numeracion de paginas...")
    anadir_numeracion(writer, 0, total_real)

    # ── 5. MARCADORES / BOOKMARKS ──
    log_cb("Generando marcadores...")

    # Marcador de portada
    writer.add_outline_item(
        "Cover / Portada",
        0,
        parent=None
    )

    # Marcadores de documentos con jerarquía
    stack = []  # (nivel, bookmark_ref)
    for item in items_activos:
        pag = pdf_a_inicio_real[os.path.normpath(item["path"])]
        nivel = item["nivel"]

        # Limpiar stack hasta el nivel padre
        while stack and stack[-1][0] >= nivel:
            stack.pop()

        parent_ref = stack[-1][1] if stack else None

        bm = writer.add_outline_item(
            item["nombre"],
            pag,
            parent=parent_ref
        )
        stack.append((nivel, bm))

    # ── 6. GUARDAR ──
    log_cb("Guardando dossier...")
    with open(salida_path, "wb") as f:
        writer.write(f)

    return vinculos_ok, total_real


# ─── INTERFAZ GRÁFICA ───────────────────────────────────────────────────────

class App(tk.Tk):

    BG      = "#1e1e2e"
    PANEL   = "#2a2a3d"
    ACCENT  = "#0078d4"
    SUCCESS = "#28a745"
    WARN    = "#e67e22"
    FG      = "#e0e0e0"
    FG_DIM  = "#888888"
    ENTRY   = "#313148"

    def __init__(self):
        super().__init__()
        self.title("PDF Dossier Builder v3")
        self.configure(bg=self.BG)
        self.resizable(False, False)
        self._lista_items = []
        self._dir_raiz = ""
        self._log_lines = []
        self._build_ui()

    def _build_ui(self):
        # ── Cabecera ──
        hdr = tk.Frame(self, bg=self.ACCENT)
        hdr.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(hdr, text="  PDF Dossier Builder  v3",
                 font=("Segoe UI", 14, "bold"),
                 bg=self.ACCENT, fg="white", anchor="w"
                 ).pack(side="left", padx=12, pady=10)
        tk.Label(hdr, text="Professional Edition",
                 font=("Segoe UI", 9, "italic"),
                 bg=self.ACCENT, fg="#cce4ff", anchor="e"
                 ).pack(side="right", padx=12)

        # ── Selector de índice ──
        frame_sel = tk.Frame(self, bg=self.PANEL)
        frame_sel.grid(row=1, column=0, columnspan=2, sticky="ew", padx=16, pady=(12, 4))
        frame_sel.columnconfigure(1, weight=1)

        tk.Label(frame_sel, text="Indice general:",
                 font=("Segoe UI", 9, "bold"),
                 bg=self.PANEL, fg=self.FG
                 ).grid(row=0, column=0, padx=(10, 6), pady=8)

        self.ruta_var = tk.StringVar()
        tk.Entry(frame_sel, textvariable=self.ruta_var,
                 width=46, bg=self.ENTRY, fg=self.FG,
                 insertbackground=self.FG, relief="flat",
                 font=("Segoe UI", 9)
                 ).grid(row=0, column=1, padx=4, pady=8, sticky="ew")

        tk.Button(frame_sel, text="Examinar...",
                  command=self._seleccionar,
                  bg=self.ACCENT, fg="white", relief="flat",
                  font=("Segoe UI", 9), padx=8, cursor="hand2"
                  ).grid(row=0, column=2, padx=(4, 6), pady=8)

        tk.Button(frame_sel, text="Analizar",
                  command=self._analizar,
                  bg=self.WARN, fg="white", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=8, cursor="hand2"
                  ).grid(row=0, column=3, padx=(0, 10), pady=8)

        # ── Nombre del dossier ──
        frame_nom = tk.Frame(self, bg=self.PANEL)
        frame_nom.grid(row=2, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 4))
        frame_nom.columnconfigure(1, weight=1)

        tk.Label(frame_nom, text="Nombre dossier:",
                 font=("Segoe UI", 9, "bold"),
                 bg=self.PANEL, fg=self.FG
                 ).grid(row=0, column=0, padx=(10, 6), pady=6)

        self.nombre_var = tk.StringVar(value="DOSSIER_COMPLETO")
        tk.Entry(frame_nom, textvariable=self.nombre_var,
                 width=46, bg=self.ENTRY, fg=self.FG,
                 insertbackground=self.FG, relief="flat",
                 font=("Segoe UI", 9)
                 ).grid(row=0, column=1, padx=4, pady=6, sticky="ew", columnspan=3)

        # ── Vista previa del árbol ──
        tk.Label(self, text="Vista previa — documentos a incluir:",
                 font=("Segoe UI", 9), bg=self.BG, fg=self.FG_DIM
                 ).grid(row=3, column=0, sticky="w", padx=16, pady=(6, 2))

        tk.Label(self, text="(doble clic para excluir/incluir)",
                 font=("Segoe UI", 8, "italic"), bg=self.BG, fg=self.FG_DIM
                 ).grid(row=3, column=1, sticky="e", padx=16)

        frame_tree = tk.Frame(self, bg=self.BG)
        frame_tree.grid(row=4, column=0, columnspan=2, padx=16, pady=(0, 4), sticky="ew")

        self.tree = ttk.Treeview(frame_tree, columns=("paginas", "estado"),
                                 show="tree headings", height=8,
                                 selectmode="browse")
        self.tree.heading("#0", text="Documento")
        self.tree.heading("paginas", text="Pags")
        self.tree.heading("estado", text="Estado")
        self.tree.column("#0", width=380)
        self.tree.column("paginas", width=50, anchor="center")
        self.tree.column("estado", width=80, anchor="center")

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview",
                        background=self.ENTRY, foreground=self.FG,
                        fieldbackground=self.ENTRY, rowheight=22,
                        font=("Segoe UI", 9))
        style.configure("Treeview.Heading",
                        background=self.PANEL, foreground=self.FG,
                        font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", self.ACCENT)])

        sb_tree = tk.Scrollbar(frame_tree, command=self.tree.yview, bg=self.ENTRY)
        self.tree.configure(yscrollcommand=sb_tree.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb_tree.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._toggle_item)

        # ── Botón generar ──
        frame_btns = tk.Frame(self, bg=self.BG)
        frame_btns.grid(row=5, column=0, columnspan=2, pady=(6, 4))

        self.btn_generar = tk.Button(frame_btns, text="Generar Dossier",
                                     command=self._iniciar,
                                     bg=self.SUCCESS, fg="white", relief="flat",
                                     font=("Segoe UI", 12, "bold"),
                                     padx=20, pady=8, cursor="hand2",
                                     state="disabled")
        self.btn_generar.pack(side="left", padx=6)

        self.btn_log = tk.Button(frame_btns, text="Exportar Log",
                                 command=self._exportar_log,
                                 bg=self.PANEL, fg=self.FG, relief="flat",
                                 font=("Segoe UI", 9), padx=10, pady=8,
                                 cursor="hand2", state="disabled")
        self.btn_log.pack(side="left", padx=6)

        # ── Barra de progreso real ──
        frame_prog = tk.Frame(self, bg=self.BG)
        frame_prog.grid(row=6, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 2))
        frame_prog.columnconfigure(0, weight=1)

        style.configure("accent.Horizontal.TProgressbar",
                        troughcolor=self.ENTRY, background=self.ACCENT,
                        bordercolor=self.BG, lightcolor=self.ACCENT,
                        darkcolor=self.ACCENT)
        self.progress = ttk.Progressbar(frame_prog, mode="determinate",
                                        length=560,
                                        style="accent.Horizontal.TProgressbar")
        self.progress.grid(row=0, column=0, sticky="ew")

        self.lbl_prog = tk.Label(frame_prog, text="",
                                 font=("Segoe UI", 8), bg=self.BG, fg=self.FG_DIM)
        self.lbl_prog.grid(row=1, column=0, sticky="w", pady=(2, 0))

        # ── Log ──
        tk.Label(self, text="Registro:",
                 font=("Segoe UI", 9), bg=self.BG, fg=self.FG_DIM
                 ).grid(row=7, column=0, sticky="w", padx=16)

        frame_log = tk.Frame(self, bg=self.BG)
        frame_log.grid(row=8, column=0, columnspan=2, padx=16, pady=(2, 4), sticky="ew")

        self.log = tk.Text(frame_log, width=72, height=10,
                           bg=self.ENTRY, fg=self.FG,
                           font=("Consolas", 8), relief="flat",
                           state="disabled", wrap="none")
        sb_log = tk.Scrollbar(frame_log, command=self.log.yview, bg=self.ENTRY)
        self.log.configure(yscrollcommand=sb_log.set)
        self.log.pack(side="left", fill="both")
        sb_log.pack(side="right", fill="y")

        self.estado = tk.StringVar(value="Selecciona el indice general y pulsa Analizar")
        tk.Label(self, textvariable=self.estado,
                 font=("Segoe UI", 9, "italic"),
                 bg=self.BG, fg=self.FG_DIM
                 ).grid(row=9, column=0, columnspan=2, pady=(0, 12))

    # ── Acciones ──

    def _seleccionar(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el indice general PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if ruta:
            self.ruta_var.set(ruta)
            nombre_base = os.path.splitext(os.path.basename(ruta))[0]
            self.nombre_var.set(nombre_base + "_DOSSIER")

    def _analizar(self):
        ruta = self.ruta_var.get().strip()
        if not ruta or not os.path.isfile(ruta):
            messagebox.showerror("Error", "Selecciona un fichero PDF valido.")
            return
        self._lista_items = []
        self.tree.delete(*self.tree.get_children())
        self._limpiar_log()
        self.btn_generar.configure(state="disabled")
        self.estado.set("Analizando vinculos...")
        threading.Thread(target=self._tarea_analizar, args=(ruta,), daemon=True).start()

    def _tarea_analizar(self, ruta):
        self._dir_raiz = os.path.dirname(ruta)
        visitados = set()
        procesar_pdf(ruta, self._dir_raiz, visitados, self._lista_items, self._log)
        self.after(0, self._poblar_arbol)

    def _poblar_arbol(self):
        self.tree.delete(*self.tree.get_children())
        nodos = {}
        for idx, item in enumerate(self._lista_items):
            nivel = item["nivel"]
            nombre = item["nombre"]
            pags = item["paginas"]
            padre_id = ""
            if nivel > 0:
                # Buscar el ultimo nodo de nivel anterior
                for j in range(idx - 1, -1, -1):
                    if self._lista_items[j]["nivel"] < nivel:
                        padre_id = nodos.get(j, "")
                        break
            nid = self.tree.insert(
                padre_id, "end",
                text=("  " * nivel) + nombre,
                values=(pags, "Incluido"),
                tags=("incluido",)
            )
            nodos[idx] = nid

        self.tree.tag_configure("incluido", foreground=self.FG)
        self.tree.tag_configure("excluido", foreground="#555577")

        total_docs = len(self._lista_items)
        total_pags = sum(it["paginas"] for it in self._lista_items)
        self.estado.set(f"{total_docs} documentos encontrados  |  {total_pags} paginas totales  |  Listo para generar")
        self.btn_generar.configure(state="normal")

    def _toggle_item(self, event):
        """Doble clic: incluir/excluir documento del dossier."""
        sel = self.tree.selection()
        if not sel:
            return
        nid = sel[0]
        # Encontrar índice
        todos = self._obtener_todos_nids()
        if nid not in todos:
            return
        idx = todos.index(nid)
        item = self._lista_items[idx]
        item["incluir"] = not item["incluir"]
        estado_str = "Incluido" if item["incluir"] else "EXCLUIDO"
        tag = "incluido" if item["incluir"] else "excluido"
        self.tree.item(nid, values=(item["paginas"], estado_str), tags=(tag,))

    def _obtener_todos_nids(self):
        resultado = []
        def recorrer(padre=""):
            for nid in self.tree.get_children(padre):
                resultado.append(nid)
                recorrer(nid)
        recorrer()
        return resultado

    def _log(self, msg):
        self._log_lines.append(msg)
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    def _limpiar_log(self):
        self._log_lines = []
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _exportar_log(self):
        ruta = filedialog.asksaveasfilename(
            title="Guardar registro",
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt")]
        )
        if ruta:
            with open(ruta, "w", encoding="utf-8") as f:
                f.write("\n".join(self._log_lines))
            messagebox.showinfo("Log exportado", f"Guardado en:\n{ruta}")

    def _iniciar(self):
        if not self._lista_items:
            messagebox.showerror("Error", "Primero analiza un indice.")
            return
        items_activos = [it for it in self._lista_items if it["incluir"]]
        if not items_activos:
            messagebox.showerror("Error", "No hay documentos incluidos.")
            return

        self.btn_generar.configure(state="disabled")
        self.btn_log.configure(state="disabled")
        self._limpiar_log()
        self.progress["value"] = 0
        self.progress["maximum"] = len(items_activos)
        self.estado.set("Generando dossier...")
        threading.Thread(target=self._proceso, daemon=True).start()

    def _proceso(self):
        try:
            ruta_indice = self.ruta_var.get().strip()
            dir_raiz = os.path.dirname(ruta_indice)
            nombre = self.nombre_var.get().strip() or "DOSSIER_COMPLETO"
            if not nombre.lower().endswith(".pdf"):
                nombre += ".pdf"
            salida = os.path.join(dir_raiz, nombre)
            titulo = os.path.splitext(os.path.basename(nombre))[0].replace("_", " ")

            self._log("=" * 62)
            self._log(f"Dossier: {titulo}")
            self._log(f"Carpeta: {dir_raiz}")
            self._log("=" * 62)

            def progress_cb(actual, total, nombre_doc):
                self.progress["value"] = actual
                self.lbl_prog.configure(
                    text=f"Procesando {actual}/{total}: {nombre_doc}"
                )
                self.update_idletasks()

            vinculos_ok, total_pags = fusionar_dossier(
                self._lista_items, salida, dir_raiz,
                titulo, self._log, progress_cb
            )

            self.progress["value"] = self.progress["maximum"]
            self.lbl_prog.configure(text="Completado")
            self._log("=" * 62)
            self.estado.set(f"Completado — {total_pags} paginas, {vinculos_ok} vinculos recalculados")
            self.btn_log.configure(state="normal")

            abre = messagebox.askyesno(
                "Completado",
                f"Dossier generado correctamente.\n"
                f"  Paginas totales: {total_pags}\n"
                f"  Vinculos recalculados: {vinculos_ok}\n\n"
                f"Guardado en:\n{salida}\n\n"
                f"¿Abrir el dossier ahora?"
            )
            if abre:
                import subprocess
                subprocess.Popen(["start", "", salida], shell=True)

        except Exception as e:
            import traceback
            self._log(f"\nError: {e}")
            self._log(traceback.format_exc())
            self.estado.set("Error durante el proceso")
            messagebox.showerror("Error inesperado", str(e))
        finally:
            self.btn_generar.configure(state="normal")


if __name__ == "__main__":
    app = App()
    app.mainloop()
