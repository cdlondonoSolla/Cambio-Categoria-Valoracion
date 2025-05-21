📘 **README – Resumen Operativo para Usuarios Técnicos**

📄 **Nombre del Script:**
CambiarCategoriaValoracion.vbs

🧾 **Propósito:**
Este script automatiza el proceso de modificación de la categoría de valoración de materiales en SAP a través de SAP GUI Scripting, leyendo los datos desde una hoja de Excel predefinida.

⚙️ **Requisitos del Sistema:**
Sistema operativo: Windows 10/11

Aplicaciones instaladas:

SAP GUI con scripting habilitado

Microsoft Excel con un libro abierto y hoja activa

Permisos:

Acceso autorizado a la transacción MM02 en SAP

Permisos de lectura sobre la hoja de Excel y escritura en el directorio de destino de los PDF

Carpeta destino para guardar PDF:
D:\tu\ruta\especifica\Cambio Categoria Valoracion\PDF\

El sistema detecta automaticamente la ruta para guardar los documentos.

▶️ **Instrucciones de Ejecución:**
Abrir SAP GUI y realizar login con el usuario correspondiente.

Abrir Microsoft Excel con el archivo que contiene:

Columna A: Código de material

Columna B: Centro

Columna C: Nueva categoría de valoración

Columna D: Estado ("Pendiente" para procesar)

Ejecutar el script con doble clic o desde consola con cscript.

El script procesará fila por fila todos los materiales con estado "Pendiente".

📝 **Log de salida**
Durante la ejecución del script, se genera un log de salida que documenta el estado de cada operación realizada, el cual se registra en la Columna D (Estado). Este log puede contener tres tipos principales de mensajes, que permiten identificar rápidamente si el proceso fue exitoso o si se presentaron errores:

**Éxito**

Mensaje: Se modifica Categoria de Valoracion
Descripción: Indica que el script se ejecutó correctamente y sin errores. Este mensaje se registra cuando todas las operaciones finalizan de forma satisfactoria.

**Error controlado**

Mensaje: No se puede modificar categoría valoración, seleccione "Visualizar error"
Descripción: Señala que se detectó un error específico durante la ejecución, pero el script logró manejarlo sin detener el proceso completo.
Acción adicional: Además del mensaje en el log, el sistema genera y descarga automáticamente un archivo PDF que contiene el detalle completo de los errores detectados. Este documento puede ser utilizado para análisis, seguimiento o soporte técnico.

**Error del sistema / no controlado**

Mensaje: Mensaje de error capturado del sistema ERP.
Descripción: Ocurrió un error inesperado. El script captura y muestra el mensaje de error tal como lo proporciona el sistema operativo o el motor de VBScript. Estos errores requieren análisis detallado ya que podrían detener la ejecución del script.

📥 **Archivos relacionados:**
Plantilla.xlsx – archivo de Excel que contiene los materiales a modificar.

Carpeta de destino: contiene los PDF generados cuando hay errores relacionados con stock/pedidos.

⚠️ **Notas importantes:**
El script interactúa directamente con ventanas emergentes. Asegúrese de no usar el equipo durante la ejecución.

Las ventanas deben estar en primer plano para que SendKeys funcione correctamente.

El script usa funciones del sistema como clip.exe y cmd, lo cual puede estar restringido por políticas corporativas.
