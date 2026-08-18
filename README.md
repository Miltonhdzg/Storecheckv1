# Storecheckv1

Aplicación de captura de precios con funcionamiento offline.

## Flujo offline

- En la primera apertura con conexión se descarga el catálogo completo.
- El catálogo queda disponible en el dispositivo durante 7 días.
- Las capturas se guardan localmente con **Guardar captura**.
- **Enviar pendientes** se habilita cuando existe conexión y conserva cualquier captura que el servidor no confirme.
- Si el catálogo vence, se requiere conexión para descargarlo nuevamente.

La aplicación debe publicarse mediante HTTPS para que el Service Worker pueda habilitar las recargas sin conexión. Después de actualizar `code.gs`, se debe publicar una nueva versión del Web App de Google Apps Script antes de publicar el frontend.
