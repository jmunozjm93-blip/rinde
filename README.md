# Rinde

Interfaz HTML para GitHub Pages. Registros y fotos almacenados en Google Sheets y Drive, nunca en el repositorio.

## Arquitectura

- index.html: interfaz móvil, cámara, lector, personas y saldo.
- config.js: URL pública de la API (no contiene claves).
- apps-script/Code.gs: API de Google Apps Script con clave verificada en servidor.
- manifest.webmanifest e iconos: acceso desde pantalla de inicio.

La API requiere la propiedad de script ACCESS_DIGEST con el SHA-256 de una clave aleatoria privada. No guardar claves en el código ni en GitHub. La clave se introduce en el navegador y permanece solo en memoria.

Para activar: configurar ACCESS_DIGEST, desplegar como aplicación web y actualizar apiUrl. La URL pública por sí sola no permite leer o escribir: todas las operaciones validan la clave en doPost. doGet solo muestra el estado genérico del servicio.

El selector de persona asigna el gasto; no es un sistema de identidad ni limita lo que ve cada persona. El fondo es común. Las fotos se guardan en Drive y las boletas anuladas conservan historial.

No hay modo offline de escritura ni sincronización automática de registros del HTML anterior.

