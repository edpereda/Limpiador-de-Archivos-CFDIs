# Limpiador de Archivos CFDI's
Programa realizado en lenguaje Java, junto con la libreria Apache Poi, para la lectura y creación de archivos xls.

### Realizado en el año 2018. Lenguaje JAVA con IDE Netbeans. Librería POI Apache. Cliente: Salas Perez Contadores.

## Problemática
El cliente trabaja con muchos CFDI (Comprobante Fiscal Digital por Internet) en formato xls, que éstos mismos son descargados desde el portal en línea del SAT (Servicio de Administración Tributaria).
Lo que el cliente realiza con cada uno de estos CFDI's descargados es eliminar columnas que no necesita, para posteriormente realizar sumas totales de aspectos en específico y finalmente utilizar estas tablas para poder envíar información al cliente y al portal en línea del SAT. Pero el proceso para realizar esto excede más de 30 minutos por cada CFDI, por lo que relantiza las actividades del cliente, se muestra de manera monotona para el cliente realizar estos pasos con todos los archivos que descarga y ocupa la mayor parte de su tiempo en el despacho.

## Objetivos
- Presentar una interfaz sencilla y fácil de utilizar para el cliente.
- El cliente solo debe seleccionar las columnas que necesita en el nuevo archivo xls
- El programa identificará columnas en específico, las cuales necesita tratar los datos para sumar o realizar totales en cada columna.
- Botón configurable que le permita seleccionar columnas en automático, sin la necesidad de volver a elegir las columnas del archivo.

## Justificación
- Software hecho a la medida para el cliente, además de una interfaz sencilla y fácil de utilizar.
- No se podían utilizar macros de Excel, ya que el cliente lo consideraba un poco más complicado.
- Existen herramientas que cuentan con esta funcionalidad, pero tiene bastantes funciones que el cliente no utiliza y el costo mensual por éste es elevado.
