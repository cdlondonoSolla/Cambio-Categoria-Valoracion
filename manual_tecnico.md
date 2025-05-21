📄 **MANUAL TÉCNICO – Documentación Técnica Detallada**

**1. Objetivo del Script**
Automatizar el cambio de la categoría de valoración en SAP para materiales específicos listados en una hoja de cálculo Excel. El script gestiona automáticamente excepciones y errores comunes.

**2. Estructura del Código**
a. Inicialización de SAP GUI Scripting
vbscript

Set SapGuiAuto = GetObject("SAPGUI")
Set application = SapGuiAuto.GetScriptingEngine
...
Establece conexión con la sesión activa de SAP.

b. Conexión con Excel
vb

Set objExcel = GetObject(,"Excel.Application")
Set objSheet = objExcel.ActiveWorkbook.ActiveSheet
Conecta con el libro de Excel actualmente abierto y toma la hoja activa.

c. Bucle de procesamiento
vbscript

For i=2 To objSheet.UsedRange.Rows.Count
Itera sobre cada fila del Excel desde la segunda, procesando solo aquellas con estado "Pendiente".

d. Ingreso a SAP y cambio de categoría
Transacción MM02

Acceso a vista de contabilidad 1

Modificación del campo MBEW-BKLAS (Categoría de valoración)

e. Manejo de errores
vbscript

If mensaje = "No se puede modificar categoría valoración..." Then
Cuando SAP no permite el cambio por razones como pedidos abiertos o stock valorado, se genera automáticamente un PDF de error, que se guarda con el nombre Material-Centro.pdf.

f. Automatización del guardado
vbscript

WshShell.SendKeys "D:\...\PDF\Material-Centro.pdf"
El script automatiza la interacción con la ventana de “Guardar como...” mediante SendKeys.

g. Actualización del Excel
Si se ejecuta con éxito, actualiza la columna D a "OK"

En caso de error, guarda el mensaje en esa misma columna.

**3. Parámetros configurables**
Ruta de guardado de PDF (SendKeys)

Estructura esperada del Excel

**4. Manejo de errores**
Se usan bloques On Error Resume Next para evitar que errores detengan la ejecución.

Se documentan los errores encontrados en la misma hoja de Excel.

**5. Limitaciones**
Dependencia del foco de ventana (por uso de SendKeys)

No apto para ejecución desatendida (sin interfaz gráfica)

No incluye logs persistentes fuera de Excel

**6. Mejoras sugeridas**
Reemplazar SendKeys por una herramienta más robusta como AutoIt

Incorporar logs de texto independientes

Parametrizar rutas y hojas de Excel vía archivo de configuración externo
