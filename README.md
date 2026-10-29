# Administración Predial - H&C

App Android para registrar visitas a terrenos: fotos, reporte, observaciones y PDF firmado por quien lo genera.
Cada persona entra con su cuenta de Microsoft de H&C; los datos quedan en SharePoint.

## Descargar la app
Entra a **[Releases](../../releases/latest)** y descarga **Predial-HyC.apk** en el teléfono. Ábrelo para instalar (o actualizar).
Cada cambio que se sube a la rama `main` genera automáticamente un APK nuevo.

## Dónde quedan los datos (sitio de SharePoint configurado)
- Listas: **Predial - Terrenos**, **Predial - Preguntas**, **Predial - Reportes**, **Predial - Administradores** (se crean solas la primera vez que entra un administrador).
- Documentos › **Administración Predial › Reportes › <Terreno> › <AAAA MM DD TERRENO>**: fotos y PDF de cada visita.

## Configuración
`web/m365-config.js`: dominio de la empresa, Id. de aplicación de Microsoft Entra y dirección del sitio de SharePoint.

## Estructura
- `web/` — código de la app.
- `assets/` — ícono y pantalla de inicio de H&C.
- `.github/workflows/android.yml` — compila y firma el APK en GitHub.
- `android-signing/` — llave con la que se firma el APK. Mantén el repositorio **privado**; si se pierde la llave, el teléfono no aceptará actualizaciones y habrá que reinstalar.
