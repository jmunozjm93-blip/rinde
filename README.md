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
