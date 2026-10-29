# Administración Predial - H&C

App Android para registrar visitas a terrenos: fotos, reporte, observaciones y PDF firmado por quien lo genera.

## Descargar la app
Entra a **[Releases](../../releases/latest)** y descarga **Predial-HyC.apk** en el teléfono. Ábrelo para instalar (o actualizar).
Cada cambio que se sube a la rama `main` genera automáticamente un APK nuevo.

## Estructura
- `web/` — código de la app (pantallas, reporte, PDF). `web/firebase-config.js` lleva los datos del proyecto Firebase.
- `firebase/firestore.rules` — reglas de seguridad; se pegan en Firebase › Firestore › Reglas.
- `assets/` — ícono y pantalla de inicio de H&C.
- `.github/workflows/android.yml` — compila y firma el APK en GitHub.
- `android-signing/` — llave con la que se firma el APK. Mantén el repositorio **privado**; si se pierde la llave, el teléfono no aceptará actualizaciones y habrá que reinstalar.
