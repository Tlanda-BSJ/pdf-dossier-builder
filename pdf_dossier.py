import tkinter as tk
from tkinter import filedialog, messagebox
import os
from pypdf import PdfReader, PdfWriter


def normalizar_ruta(ruta):
    if not ruta:
        return ""
    return ruta.replace("/", os.sep).replace("\\\\", os.sep).strip()


def fusionar_pdfs(lista_pdfs, salida_path):
    writer = PdfWriter()

    for pdf in lista_pdfs:
        reader = PdfReader(pdf)
        for page in reader.pages:
            writer.add_page(page)

    with open(salida_path, "wb") as f:
        writer.write(f)


def ejecutar():
    archivos = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
    if not archivos:
        return

    salida = os.path.join(os.path.dirname(archivos[0]), "dossier_final.pdf")
    fusionar_pdfs(archivos, salida)

    messagebox.showinfo("OK", "PDF creado correctamente")


app = tk.Tk()
app.title("PDF Builder")
app.geometry("300x120")

tk.Button(app, text="Seleccionar PDFs", command=ejecutar).pack(expand=True)

app.mainloop()
