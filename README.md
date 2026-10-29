# Administración Predial - H&C

App Android simple para registrar visitas a terrenos: fotos, reporte, observaciones y PDF firmado por quien lo genera.
Sin cuentas ni servidores: cada reporte queda guardado en el teléfono, y su PDF se guarda en **Documentos › Predial HyC** y se comparte por WhatsApp, correo, etc.

## Descargar la app
Entra a **[Releases](../../releases/latest)** y descarga **Predial-HyC.apk** en el teléfono. Ábrelo para instalar (o actualizar).
Cada cambio que se sube a la rama `main` genera automáticamente un APK nuevo.

## Cambiar terrenos o preguntas en todos los teléfonos
En `web/app.js`, edita `TERRENOS_INICIALES` o `PREGUNTAS_INICIALES` (y sube en 1 `VERSION_PREGUNTAS`). Al instalar el APK nuevo, las preguntas se actualizan en cada teléfono.

## Estructura
- `web/` — código de la app.
- `assets/` — ícono y pantalla de inicio de H&C.
- `.github/workflows/android.yml` — compila y firma el APK en GitHub.
- `android-signing/` — llave con la que se firma el APK. Mantén el repositorio **privado**; si se pierde la llave, el teléfono no aceptará actualizaciones y habrá que reinstalar.
