# PDF Dossier Builder

Genera un dossier PDF completo siguiendo automáticamente los hipervínculos desde un índice general.

## Cómo obtener el .exe (sin instalar nada)

1. Crea una cuenta gratuita en https://github.com
2. Crea un repositorio nuevo (New repository) → nombre: `pdf-dossier-builder` → Public → Create
3. Sube estos tres archivos manteniendo la estructura de carpetas:
   - `pdf_dossier.py`
   - `.github/workflows/build.yml`
4. GitHub compilará el .exe automáticamente en ~3 minutos
5. Ve a la pestaña **Actions** → haz clic en el workflow → descarga **PDFDossierBuilder-Windows**

## Uso del programa

1. Ejecuta `PDFDossierBuilder.exe` (no requiere instalación)
2. Pulsa **Examinar** y selecciona tu índice general PDF
3. Pulsa **Generar Dossier**
4. El dossier se guarda en la misma carpeta que el índice

## Compatibilidad

- Creado con Nitro PDF (vínculos /Launch y /GoToR)
- Rutas absolutas Windows y relativas
- Vínculos de retroceso ignorados automáticamente
- Windows 10 / 11
