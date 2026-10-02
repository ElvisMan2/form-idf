function subirArchivo(fileName, archivoBase64, folderId) {
  try {
    const base64Limpio = archivoBase64.includes(',')
      ? archivoBase64.split(',')[1]
      : archivoBase64;

    const bytes = Utilities.base64Decode(base64Limpio);
    const blob = Utilities.newBlob(
      bytes,
      'application/octet-stream',
      fileName
    );

    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);

    Logger.log('Archivo subido exitosamente: ' + file.getUrl());
    return file.getUrl();

  } catch (error) {
    Logger.log('Error al subir archivo: ' + error.message);
    throw error;
  }
}


/**
 * Script para guardar datos en Google Sheets
 * ID de la hoja: 
 * Hoja: Registros
 */
const SPREADSHEET_ID = '1LO_m4h9QpqT0D-nL8_0G9tP0__sMEiNfwXJgF90klno';
const SHEET_NAME = 'Postulaciones';
const SHEET_PERSONAS = 'Alumnos';
const FOLDER_ID = '1kq_oMryCC04arWJzIQpwYJEUMJUP9IPk';//

/**
 * Guarda el estado de un envío en el cache
 */
function guardarEstadoEnvio(idEnvio, estado, mensaje) {
  const cache = CacheService.getScriptCache();
  const datos = {
    status: estado,
    message: mensaje,
    timestamp: new Date().getTime()
  };
  // Guardar por 10 minutos (600 segundos)
  cache.put(idEnvio, JSON.stringify(datos), 600);
}

/**
 * Obtiene el estado de un envío desde el cache
 */
function obtenerEstadoEnvio(idEnvio) {
  const cache = CacheService.getScriptCache();
  const datos = cache.get(idEnvio);
  if (datos) {
    return JSON.parse(datos);
  }
  return null;
}

/**
 * Guarda los inventores en la pestaña Personas
 * @param {Object} datos - Objeto con los datos incluyendo inventores
 * @param {string} fechaEnvio - Fecha de envío ya calculada en guardarDatos
 * @return {void}
 */
function guardarInventores(datos, fechaEnvio) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_PERSONAS);
    
    // Si la pestaña no existe, crearla
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_PERSONAS);
      sheet.appendRow(['Fecha de Envío', 'Título de la Invención', 'Apellidos', 'Nombres', 'Género', 'Edad', 'Grado']);
    } else {
      const encabezados = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (encabezados.indexOf('Género') === -1) {
        sheet.getRange(1, encabezados.length + 1).setValue('Género');
      }
    }
    
    const fecha = fechaEnvio || Utilities.formatDate(new Date(), 'America/Lima', 'dd/MM/yyyy HH:mm:ss');
    const titulo = datos.titulo || '';
    
    // Guardar cada inventor
    if (datos.inventores && Array.isArray(datos.inventores)) {
      for (const inventor of datos.inventores) {
        const fila = [
          fecha,
          titulo,
          inventor.apellidos || '',
          inventor.nombres || '',
          inventor.genero || '',
          inventor.edad || '',
          inventor.grado || ''
        ];
        sheet.appendRow(fila);
      }
    }
    
    Logger.log('Inventores guardados exitosamente en la pestaña Personas');
  } catch (error) {
    Logger.log('Error al guardar inventores: ' + error.message);
    throw error;
  }
}

/**
 * Guarda los datos en la hoja de cálculo
 * @param {Object} datos - Objeto con los datos a guardar
 * @return {Object} Resultado de la operación
 */
function guardarDatos(datos) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error('No se encontró la hoja "' + SHEET_NAME + '"');
    }
    
    const fecha = Utilities.formatDate(new Date(), 'America/Lima', 'dd/MM/yyyy HH:mm:ss');

    // Subir archivo fichaPostulante solo si existe
    let fichaPostulanteUrl = '';
    if (datos.fichaPostulanteBase64 && datos.fichaPostulanteNombreArchivo) {
      fichaPostulanteUrl = subirArchivo(
        datos.fichaPostulanteNombreArchivo,
        datos.fichaPostulanteBase64,
        FOLDER_ID
      );
    }

    // Subir archivo fichaInvencion solo si existe
    let fichaInvencionUrl = '';
    if (datos.fichaInvencionBase64 && datos.fichaInvencionNombreArchivo) {
      fichaInvencionUrl = subirArchivo(
        datos.fichaInvencionNombreArchivo,
        datos.fichaInvencionBase64,
        FOLDER_ID
      );
    }

    // Subir archivo declaracionParteAsesor solo si existe
    let declaracionParteAsesorUrl = '';
    if (datos.declaracionParteAsesorBase64 && datos.declaracionParteAsesorNombreArchivo) {
      declaracionParteAsesorUrl = subirArchivo(
        datos.declaracionParteAsesorNombreArchivo,
        datos.declaracionParteAsesorBase64,
        FOLDER_ID
      );
    }

    // Crear array con los datos
    const fila = [
      fecha,
      datos.idEnvio || '',
      datos.titulo,
      datos.categoriaParticipacion || '',
      datos.institucionEducativa || '',
      datos.direccionInstitucion || '',
      datos.regionInstitucion || '',
      datos.apellidosRepresentante,
      datos.nombreRepresentante,
      datos.dni,
      datos.telefono,
      datos.correo,
      fichaPostulanteUrl,
      fichaInvencionUrl,
      declaracionParteAsesorUrl,
      datos.descripcion || '',
      datos.autorizacionContacto || '',
      datos.numeroAlumnos || ''
    ];
    
    // Guardar inventores en la pestaña Personas
    guardarInventores(datos, fecha);
    
    // Agregar nueva fila al final
    sheet.appendRow(fila);
    
    Logger.log('Datos guardados exitosamente');
    return {
      success: true,
      message: 'Datos guardados correctamente',
      fecha: fecha,
      fila: sheet.getLastRow()
    };
    
  } catch (error) {
    Logger.log('Error al guardar datos: ' + error.message);
    throw error;
  }
}

/**
 * Maneja peticiones POST con seguridad mejorada
 */
function doPost(e) {
  let idEnvio = null;
  
  try {
    const params = JSON.parse(e.postData.contents);
    idEnvio = params.idEnvio;
    
    // Marcar como pendiente inmediatamente
    if (idEnvio) {
      guardarEstadoEnvio(idEnvio, 'pending', 'Procesando...');
    }
    
    const resultado = guardarDatos(params);
    
    // Guardar resultado exitoso en cache
    if (idEnvio) {
      guardarEstadoEnvio(idEnvio, 'success', 'Formulario enviado exitosamente');
    }
    
    return createResponseWithCORS(resultado);
    
  } catch (error) {
    Logger.log('Error en doPost: ' + error.message);
    if (idEnvio) {
      guardarEstadoEnvio(idEnvio, 'error', 'Error interno del servidor: ' + error.message);
    }
    return createResponseWithCORS({
      success: false,
      message: 'Error interno del servidor'
    });
  }
}

/**
 * Maneja peticiones GET para verificar estados
 */
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'checkStatus') {
    const idEnvio = e.parameter.id;
    if (!idEnvio) {
      return createResponseWithCORS({
        status: 'error',
        message: 'ID de envío no proporcionado'
      });
    }
    
    const estado = obtenerEstadoEnvio(idEnvio);
    if (estado) {
      return createResponseWithCORS(estado);
    } else {
      return createResponseWithCORS({
        status: 'notfound',
        message: 'Estado no encontrado'
      });
    }
  }
  
  return createResponseWithCORS({
    success: true,
    message: 'API funcionando correctamente'
  });
}

/**
 * Crea una respuesta JSON con headers CORS y seguridad apropiados
 */
function createResponseWithCORS(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  
  // Nota: Apps Script maneja CORS automáticamente cuando el proyecto está publicado como Web App
  // con acceso "Cualquier persona" o "Cualquier persona, incluso anónimos"
  
  return output;
}

/**
 * Crea una respuesta JSON con headers apropiados (mantiene compatibilidad)
 * @param {Object} data - Datos a retornar
 * @return {TextOutput} Respuesta formateada
 */
function createResponse(data) {
  return createResponseWithCORS(data);
}

