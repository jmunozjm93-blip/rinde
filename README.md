# Rinde

App móvil en GitHub Pages con cuentas personales y Google Sheets/Drive como almacenamiento.

- Inicio con carrusel Converse, Umbro y Fila, reutilizado de la portada Sell Out del propietario. Incluye pausa y respeta movimiento reducido.
- Login individual; contraseñas verificadas en Apps Script. El repositorio no contiene contraseñas ni hashes de usuarios.
- Sesiones de hasta seis horas: token en sessionStorage, nunca la contraseña. Logout revoca el token y limpia la interfaz. La caché de Google puede terminar la sesión antes.
- Cada cuenta accede exclusivamente a sus fondos, boletas y fotos. Jonathan administra todas y puede consultar el resumen global. Los permisos se validan en el servidor en cada operación.
- La propiedad privada RINDE_USERS contiene cuentas, sales y hashes. La antigua ACCESS_DIGEST ya no habilita ningún endpoint.
- Cinco intentos incorrectos bloquean temporalmente el usuario durante un minuto.
- La elección de claves predecibles fue solicitada por el propietario. Este esquema no comprueba la identidad real de quien conoce una clave; para mayor protección deben sustituirse por claves únicas o Google Identity.
- Fondos: hoja Fondos, columnas Persona y Fondo CLP. Boletas: hoja de ID 0, columnas A:J. Solo Estado=Vigente se incluye en saldos. Anulada y Prueba anterior conservan historial sin afectar saldos.
- Configuración/B2 conserva el antiguo fondo común como referencia; no participa en los fondos personales.
- Fotos privadas en Drive, entregadas por la API después de comprobar titular o administrador.
- OCR en el navegador y sugerencia por palabras del detalle; las categorías ambiguas quedan en Otros y la elección manual se respeta.

El despliegue de Apps Script ejecuta como propietario con acceso web público, pero cada acción exige una sesión válida salvo login. config.js contiene solo la URL pública. Nunca subir el archivo privado de accesos ni los scripts privados de configuración.

Se admiten boletas con fondo cero; el saldo negativo indica el monto pendiente de cubrir. Cada persona tiene una pestaña automática en Sheets con su fondo, total, saldo y boletas vigentes, vinculada a los registros centrales mediante fórmulas.

Descarga Excel: el botón abre una revisión del encabezado y la clasificación de cada gasto. Genera .xlsx en el navegador con el formato de Comercial Depor: columnas contables, totales con fórmulas y resultados guardados, fondo, saldo y firmas. Los campos adicionales se usan solo para la descarga; no modifican Sheets ni las boletas. El administrador debe seleccionar una persona. Incluye todas sus boletas vigentes, sin filtro de periodo; el periodo del encabezado es una etiqueta editable. No se incluyen archivos ni datos históricos de la referencia. ExcelJS 4.4.0 se sirve desde vendor/ con su licencia MIT.

Clasificación en origen: las siete columnas contables están ahora en el formulario de la boleta. La categoría elegida se persiste en la columna G de Sheets y determina la columna del Excel. Solo registros anteriores ambiguos requieren revisión al exportar. Los catálogos de centro de costo (28 activos), marca (9), canal (6) y sucursal (64 activas) se extrajeron de las hojas auxiliares del formato Oficina 11-02, sin gastos históricos. Los códigos se conservan como texto; nombres distintos con código compartido siguen separados en los desplegables. El encabezado está en el formulario y se recuerda por persona en localStorage de este dispositivo; no se sincroniza con Sheets.

Pagos al descargar: preparar el Excel no modifica la rendición. El botón final registra Pagado para la selección exacta bajo bloqueo, con lote idempotente y estado en el Excel; después descarga el archivo. Al cerrar un pago el fondo queda en cero. Las columnas K:N conservan lote, fecha, clasificación y encabezado para regenerar copias privadas desde el historial. Los pagos salen de la lista pendiente y no se pueden anular desde ella. No se permite volver a registrar el mismo número/comercio ya vigente o pagado. Una copia del historial no paga otra vez ni reinicia un fondo nuevo. El navegador no permite confirmar que el archivo terminó de guardarse: el registro se confirma antes de iniciar la descarga y se puede recuperar desde el historial si esta falla.
