# Copello — correo a planilla

Objetivo inicial: registrar en una planilla los correos que se marquen con una etiqueta específica de Gmail.

## Qué hace esta versión

- Busca correos dentro de la etiqueta `Copello_Para_Excel`.
- Agrega una fila por mensaje con fecha, remitente, asunto, cuerpo, enlace e ID.
- Evita volver a importar mensajes ya registrados.
- Crea una Google Sheet que se puede descargar como Excel (.xlsx).
- No envía respuestas ni modifica o elimina correos.
- La revisión automática cada 15 minutos es opcional y se activa manualmente.

## Instalación

1. Abrí [Google Apps Script](https://script.google.com/).
2. Creá un proyecto y pegá el contenido de `Code.gs`.
3. Ejecutá `prepararCopello` y autorizá los permisos solicitados.
4. En Gmail, aplicá la etiqueta `Copello_Para_Excel` a un correo de prueba.
5. Ejecutá `importarCorreosCopello`.
6. En los registros de ejecución, abrí el enlace de la planilla.
7. Cuando funcione correctamente, podés ejecutar `instalarRevisionAutomaticaCopello` para automatizar la revisión. Para desactivarla, ejecutá `quitarRevisionAutomaticaCopello`.

## Importante: falta definir las columnas definitivas

Esta es una base mínima: guarda el contenido del correo, pero todavía no extrae datos específicos del texto para colocarlos en columnas particulares. Para completar esa parte sin inventar requisitos, hace falta un ejemplo real del correo (podés ocultar nombres, direcciones y otros datos privados) y los encabezados de la planilla Excel que debe completar.

La hoja creada es Google Sheets; podés descargarla desde Archivo → Descargar → Microsoft Excel (.xlsx). Esta versión no escribe directamente en un archivo Excel almacenado en una computadora.
