import tkinter as tk
from tkinter import filedialog, messagebox
import os
from pypdf import PdfReader, PdfWriter


def extraer_pdfs_desde_indice(pdf_path):
    """Busca enlaces a otros PDFs dentro del documento"""
    reader = PdfReader(pdf_path)
    encontrados = []

    for page in reader.pages:
        if "/Annots" not in page:
            continue

        for annot_ref in page["/Annots"]:
            try:
                annot = annot_ref.get_object()

                if "/A" not in annot:
                    continue

                action = annot["/A"]
                ruta = action.get("/F")

                if ruta and isinstance(ruta, str) and ruta.lower().endswith(".pdf"):
                    ruta_abs = os.path.join(os.path.dirname(pdf_path), ruta)

                    if os.path.exists(ruta_abs):
                        encontrados.append(os.path.normpath(ruta_abs))

            except Exception:
                continue

    # eliminar duplicados manteniendo orden
    vistos = set()
    resultado = []
    for p in encontrados:
        if p not in vistos:
            vistos.add(p)
            resultado.append(p)

    return resultado


def fusionar_pdfs(lista_pdfs, salida_path):
    writer = PdfWriter()

    for pdf in lista_pdfs:
        try:
            reader = PdfReader(pdf)
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            print(f"[WARN] No se pudo añadir {pdf}: {e}")

    with open(salida_path, "wb") as f:
        writer.write(f)


def ejecutar():
    indice = filedialog.askopenfilename(
        title="Selecciona el PDF índice",
        filetypes=[("PDF files", "*.pdf")]
    )

    if not indice:
        return

    try:
        pdfs = extraer_pdfs_desde_indice(indice)

        if not pdfs:
            messagebox.showwarning("Aviso", "No se encontraron PDFs enlazados")
            return

        lista_final = [indice] + pdfs

        salida = os.path.join(os.path.dirname(indice), "dossier_final.pdf")

        fusionar_pdfs(lista_final, salida)

        messagebox.showinfo("Listo", f"Dossier generado:\n{salida}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# GUI
app = tk.Tk()
app.title("PDF Dossier Builder (Base)")
app.geometry("350x150")

btn = tk.Button(app, text="Seleccionar índice y generar dossier", command=ejecutar)
btn.pack(expand=True)

app.mainloop()
