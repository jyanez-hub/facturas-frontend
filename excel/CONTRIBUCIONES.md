# Guía para compartir módulos y formularios VBA

Sigue estas pautas para adjuntar módulos (`.bas`) y formularios (`.frm/.frx`) al repositorio y que podamos incorporarlos fácilmente a la plantilla.

## 1. Exportar módulos estándar
1. Abre el Editor de VBA (`Alt + F11`).
2. En el Explorador de proyectos, selecciona el módulo que quieras compartir.
3. Ve a **Archivo > Exportar archivo…** y guarda el archivo con extensión `.bas`.
4. Copia el archivo exportado dentro de la carpeta `excel/` del repositorio (por ejemplo `excel/modRenombrar.bas`).
5. Sube o agrega el archivo al control de versiones para que pueda revisarlo.

## 2. Exportar formularios de usuario
1. En el Editor de VBA, selecciona el UserForm.
2. Usa **Archivo > Exportar archivo…**. Esto generará dos archivos con el mismo nombre: `.frm` (definición) y `.frx` (recursos y controles).
3. Crea en el repositorio la carpeta `excel/forms/` si aún no existe y coloca allí ambos archivos (por ejemplo `excel/forms/frmImportador.frm` y `excel/forms/frmImportador.frx`).
4. Añade los dos archivos al repositorio para que puedan importarse desde Excel con **Archivo > Importar archivo…**.

## 3. Compartir varios archivos a la vez
Si prefieres comprimir todo en un solo paquete:
1. Agrupa los `.bas`, `.frm` y `.frx` en una carpeta de tu equipo.
2. Crea un ZIP (por ejemplo `REPOSITORIO.zip` en `C:\Users\Mayra\OneDriveOutlook\OneDrive\Desktop\EXCELBOT\`).
3. Copia ese ZIP dentro del repositorio (puedes crear una carpeta `attachments/` para mantenerlo organizado) y súbelo. Yo me encargaré de extraerlo y mover cada archivo a su ubicación final.

## 4. Confirmar la importación
Una vez compartidos los archivos:
- Yo los importaré en el libro desde la misma ruta, revisando `Option Explicit` y el nombre del módulo.
- Cualquier dependencia adicional (por ejemplo, tablas o parámetros nuevos) la documentaré para que quede alineada con `Setup_Estructura`.

Con estos pasos tendremos un flujo ordenado para integrar tus módulos y formularios en la plantilla automatizada.
