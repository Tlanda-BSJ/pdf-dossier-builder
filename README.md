# PDF Dossier Builder

Herramienta para generar un único PDF a partir de un índice, siguiendo automáticamente los enlaces a otros documentos y convirtiéndolos en enlaces internos.

## 🚀 Características

* Fusiona múltiples PDFs automáticamente
* Detecta enlaces externos (`/GoToR`, `/Launch`)
* Convierte enlaces externos en internos dentro del dossier
* Genera un ejecutable `.exe` sin instalación
* Compatible con rutas relativas y absolutas

## 🧠 Cómo funciona

1. Seleccionas un PDF índice
2. El programa detecta enlaces a otros PDFs
3. Fusiona todos los documentos
4. Reescribe los enlaces para que funcionen dentro del PDF final

## 📦 Instalación (modo desarrollo)

pip install pypdf
python pdf_dossier.py

## 🪟 Descargar ejecutable (.exe)

1. Ve a la pestaña **Actions**
2. Ejecuta el workflow
3. Descarga el archivo generado

## ⚠️ Limitaciones

* Los PDFs se identifican por nombre de archivo
* Puede fallar si hay nombres duplicados
* No todos los tipos de enlaces PDF están soportados

## 🛠️ Tecnologías

* Python 3.11
* pypdf
* Tkinter
* PyInstaller
