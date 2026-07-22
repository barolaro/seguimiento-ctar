const CARPETA_ID = 'REEMPLAZAR_CON_ID_CARPETA';
const TOKEN_SECRETO = 'REEMPLAZAR_CON_TOKEN_SEGURO';

function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents || '{}');
    if (datos.token !== TOKEN_SECRETO) {
      return responder({ok: false, error: 'Acceso no autorizado'});
    }

    const carpeta = DriveApp.getFolderById(CARPETA_ID);

    if (datos.accion === 'subir') {
      const bytes = Utilities.base64Decode(datos.contenido);
      const blob = Utilities.newBlob(bytes, datos.mime_type || 'application/octet-stream', datos.nombre);
      const archivo = carpeta.createFile(blob);
      return responder({
        ok: true,
        id: archivo.getId(),
        nombre: archivo.getName(),
        mime_type: archivo.getMimeType(),
        url: archivo.getUrl()
      });
    }

    if (datos.accion === 'descargar') {
      const archivo = DriveApp.getFileById(datos.id);
      if (!archivo.getParents().hasNext()) {
        return responder({ok: false, error: 'Archivo no disponible'});
      }
      return responder({
        ok: true,
        contenido: Utilities.base64Encode(archivo.getBlob().getBytes())
      });
    }

    return responder({ok: false, error: 'Acción no reconocida'});
  } catch (error) {
    return responder({ok: false, error: String(error)});
  }
}

function responder(datos) {
  return ContentService
    .createTextOutput(JSON.stringify(datos))
    .setMimeType(ContentService.MimeType.JSON);
}
