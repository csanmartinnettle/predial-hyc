// Conexión con Microsoft 365 de H&C (ver la guía de instalación).
export const m365Config = {
  // Dominio de la empresa: identifica el Microsoft 365 de H&C.
  tenant: "hycinversiones.cl",
  // "Id. de aplicación (cliente)" que entrega Microsoft Entra al registrar la app.
  clientId: "PEGAR_ID_DE_APLICACION",
  // Dirección del sitio de SharePoint donde se guardan los reportes.
  siteUrl: "PEGAR_URL_DEL_SITIO_SHAREPOINT",
  // Carpeta dentro de "Documentos" del sitio donde quedan las fotos y los PDF.
  carpeta: "Administración Predial"
};
