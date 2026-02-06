# Excel/VBA – Creador de PDFs (Demo)

Macro en Excel + VBA que genera un PDF por correo usando una plantilla Word. Reemplaza campos de texto (campaña y fecha) e inserta una tabla con VINs asociados a cada correo.

Confidencialidad:
El desarrollo original se utilizó con información y plantillas de una empresa (correos reales, identificadores/series y bases internas). Por lo mismo, no se publica el archivo original ni datos reales. Este repositorio contiene una versión demo con datos ficticios y lógica equivalente para fines de portafolio.

## Para qué sirve
Cuando se necesita generar comunicaciones individuales por cliente (correo) con una lista de identificadores (VINs), esta macro permite:
- Consolidar VINs por correo y eliminar duplicados
- Generar automáticamente un documento por correo a partir de una plantilla Word
- Insertar una tabla de VINs en 3 columnas para mejorar legibilidad
- Exportar un PDF por correo con nombre estandarizado

## Cómo funciona
1. Seleccionas la celda donde inicia la base (columna Correo) en Excel.
   - Se asume que el VIN está en la columna inmediata a la derecha.
2. Seleccionas la celda donde inicia la lista de correos (una columna).
3. Capturas dos textos: campaña y fecha.
4. Seleccionas una plantilla Word (.docx/.dotx) y una carpeta de salida.
5. La macro agrupa VINs por correo y genera un PDF por cada correo de la lista:
   - Reemplaza los marcadores <<CAMPANIA>> y <<FECHA>>
   - Inserta la tabla de VINs en el marcador <<VIN_TABLA>>
   - Exporta a PDF

## Estructura del repositorio
- src/ contiene el módulo VBA exportado (.bas) y contiene la plantilla Word demo
- data/ contiene archivos de entrada ficticios
- output/ incluye capturas del input/output y contiene ejemplos de PDFs generados

## Uso
1. Abrir los archivos de entrada en data/
2. Importar src/GenerarPDFsPorCorreo.bas en el Editor VBA (Alt + F11)
3. Ejecutar la macro GenerarPDFsPorCorreo
4. Seguir los cuadros de diálogo (base, lista, campaña, fecha, plantilla Word, carpeta destino)

## Salida
Se genera un PDF por correo en la carpeta seleccionada, con nombre:
cantidadVINs_correo.pdf

Ejemplo:
130_cliente01.pdf

## Supuestos y notas
- Se filtra por VINs únicos: si un VIN se repite para el mismo correo, cuenta una sola vez.
- Los marcadores de la plantilla deben existir: <<CAMPANIA>>, <<FECHA>>, <<VIN_TABLA>>.
- La tabla de VINs en Word se inserta en 3 columnas y se ajusta al ancho de la página para evitar texto encimado.
