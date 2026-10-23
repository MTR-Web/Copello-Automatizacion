# Copello — Centro de correos

Automatización inicial para registrar correos de Gmail en una planilla de Google Sheets que puede descargarse como Excel.

## Qué incluye

- `Code.gs`: backend de Gmail/Sheets, detección de duplicados, estado del sistema y disparador opcional cada 15 minutos.
- `Index.html`: panel web adaptable a computadora y celular, con estado de configuración, total de correos, lista de mensajes recientes y botones para preparar, importar y activar/desactivar la revisión automática.
- `index.html`: página informativa del repositorio.

## Instalación y puesta en marcha

1. Abrí [Google Apps Script](https://script.google.com/) e iniciá un proyecto.
2. En el archivo `Code.gs` del proyecto, pegá el contenido de este repositorio.
3. Creá un archivo HTML llamado exactamente `Index` y pegá el contenido de `Index.html`.
4. Guardá los archivos. Ejecutá `prepararCopello` una vez desde el editor y aceptá los permisos solicitados.
5. En Gmail, aplicá la etiqueta `Copello_Para_Excel` a un correo de prueba.
6. En Apps Script, elegí **Implementar → Nueva implementación → Aplicación web**. Para que procese el Gmail de quien la usa, seleccioná **Ejecutar como: usuario que accede a la aplicación web** y restringí **Quién tiene acceso** a tu cuenta o a las cuentas autorizadas. Cada usuario deberá conceder sus propios permisos. No la publiques con acceso abierto si puede leer correos personales.
7. Abrí la URL de la aplicación web, usá **Importar correos** y verificá la planilla creada.
8. Solo después de comprobar el resultado, activá **Activar automático** en el panel.

## Qué registra

Cada mensaje importado genera una fila con fecha, remitente, asunto, cuerpo, enlace al correo original, ID del mensaje y fecha de importación. El ID se usa para evitar duplicados.

La hoja creada es Google Sheets. Para obtener un archivo Excel, abrí la hoja y elegí **Archivo → Descargar → Microsoft Excel (.xlsx)**.

## Límites actuales

- La importación busca mensajes que tengan la etiqueta indicada. El sistema no lee todos los correos de Gmail indiscriminadamente.
- La versión actual guarda el correo completo; todavía no extrae automáticamente campos de negocio específicos. Para eso hacen falta ejemplos del correo y los encabezados de la planilla de destino.
- No envía respuestas, no elimina mensajes ni confirma pagos.
- La revisión automática es opcional. Se ejecuta cada 15 minutos cuando se activa desde el panel.
- La importación actual revisa hasta 100 conversaciones etiquetadas por ejecución. Si se necesita procesar volúmenes mayores, hay que implementar paginación y probarla con la cuenta real.
- No está conectada a un archivo Excel local de una computadora. Usa Google Sheets, que se puede exportar a Excel.

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
