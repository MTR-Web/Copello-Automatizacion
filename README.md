# Copello — Centro de correos

Automatización inicial para registrar correos de Gmail en una planilla de Google Sheets que puede descargarse como Excel.

## Qué incluye

- `Code.gs`: procesa correos etiquetados `Pagos` (también admite la etiqueta anterior `Copello_Para_Excel`), guarda adjuntos en Drive, evita volver a procesar IDs ya registrados, intenta extraer datos, registra comprobantes y crea borradores de respuesta sin enviarlos.
- `Index.html`: panel para preparar, importar, revisar registros, abrir la planilla y exportarla a Excel.
- `index.html`: página informativa del proyecto.

## Instalación

1. En Apps Script, reemplazá el contenido de `Code.gs` por el de este repositorio.
2. Creá o reemplazá el archivo HTML llamado exactamente `Index` con el contenido de `Index.html`.
3. En el editor de Apps Script, abrí **Servicios +**, agregá **Drive API** (v2), necesaria para intentar OCR de PDF e imágenes.
4. Guardá. Ejecutá `prepararCopello` desde el editor y autorizá Gmail, Drive y Sheets.
5. En Gmail, etiquetá un correo de prueba con `Pagos`. No uses esa etiqueta para quejas.
6. Implementá como aplicación web, ejecutá como tu cuenta y limitá el acceso a las cuentas autorizadas.
7. Abrí la URL `/exec`, tocá **Preparar sistema** y luego **Importar correos**. Revisá tanto la planilla como la carpeta `Copello_Comprobantes`.
8. Activá la revisión automática solo después de comprobar una prueba manual.

## Qué registra y qué no

- La hoja `Comprobantes` guarda un registro por adjunto, el enlace al archivo de Drive, texto extraído, campos detectados y estado de revisión.
- La hoja `Correos` conserva el registro del mensaje.
- Copello intenta detectar importe, fecha, número de operación y pagador mediante patrones de texto. La extracción puede fallar o interpretar mal formatos; los datos deben verificarse.
- Los adjuntos se guardan en una carpeta de Drive llamada `Copello_Comprobantes`.
- Si el correo tiene adjuntos, el sistema intenta crear un borrador de respuesta. **No envía mensajes.** El texto del borrador se configura en `CONFIG.textoBorrador`.
- El OCR necesita el servicio avanzado Drive API v2. Si no está habilitado, los adjuntos se guardan y quedan marcados para revisión manual.
- Todavía no se compara contra el Excel maestro de consorcios/unidades: hace falta integrar ese archivo y conocer sus columnas.
- La automatización cada 15 minutos es opcional y debe probarse con la cuenta real antes de depender de ella.

## Archivos

- [Backend Apps Script](Code.gs)
- [Panel web para Apps Script](Index.html)
- [Página informativa](index.html)


## Página pública y conexión al panel

La página informativa está en `index.html`. Se agregó un workflow en `.github/workflows/pages.yml` para publicar el sitio mediante GitHub Pages. Si Pages todavía no está habilitado para el repositorio, entrá a **Settings → Pages** y seleccioná **GitHub Actions** como origen de publicación. Después revisá **Actions** para ver el resultado del despliegue.

URL esperada del sitio: https://mtr-web.github.io/Copello-Automatizacion/

La página pública tiene un campo para guardar la URL de la aplicación web de Apps Script y abrir el panel. Esa URL solo se obtiene después de desplegar el proyecto desde la cuenta de Google; no se puede inventar ni crear con la integración de GitHub.

### Desplegar el panel de Apps Script

1. Abrí https://script.google.com/ y creá un proyecto.
2. Copiá el archivo `Code.gs` de este repositorio al editor de Apps Script.
3. Creá un archivo HTML llamado exactamente `Index` y copiá el contenido de `Index.html`.
4. Guardá y ejecutá `prepararCopello` una vez desde el editor. Aceptá los permisos.
5. Seleccioná **Implementar → Nueva implementación → Aplicación web**.
6. Elegí **Ejecutar como: usuario que accede a la aplicación web** y restringí el acceso a la cuenta o las cuentas autorizadas. No habilites acceso público a una app que puede leer Gmail.
7. Implementá y copiá la URL terminada en `/exec`.
8. Abrí la página de GitHub Pages, pegá esa URL en **Acceso operativo** y guardala. Desde ahí se abrirá el panel.

La página de GitHub Pages no puede acceder directamente a Gmail. El panel operativo debe ejecutarse desde la URL de Apps Script, que solicita los permisos de Google. No compartas la URL de una implementación con permisos amplios.
