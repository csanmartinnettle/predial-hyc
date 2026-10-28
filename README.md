# Administración Predial - H&C

App instalable (Android y iPhone) para registrar visitas a terrenos: fotos, reporte, observaciones y PDF firmado por quien lo genera.

## Puesta en marcha
1. **Firebase** (datos): crea el proyecto, activa Authentication › Anónimo y Firestore, y crea el documento `privado/acceso` con el campo `codigo`.
2. **Reglas:** en Firestore › Reglas, pega el contenido de [`firebase/firestore.rules`](firebase/firestore.rules) y publica.
3. **Configuración:** edita [`firebase-config.js`](firebase-config.js) con los datos de tu app web de Firebase.
4. **Dominio autorizado:** en Authentication › Configuración › Dominios autorizados, agrega `csanmartinnettle.github.io`.
5. **Publicación:** GitHub Pages publica la rama `main`. Cada cambio que se sube aquí llega solo a los teléfonos.

La guía completa está en el documento "Guía de instalación · Administración Predial H&C".
