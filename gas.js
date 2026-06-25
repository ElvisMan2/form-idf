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
      datos.autorizacionContacto || ''
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
 * Configuración de seguridad
 */
const SECURITY_CONFIG = {
  MAX_REQUESTS_PER_IP: 100, // máximo 100 envíos por IP por hora
  MAX_REQUESTS_PER_EMAIL: 100, // máximo 100 envíos por email por día
  MIN_FORM_TIME: 15, // mínimo 15 segundos para completar el formulario
  MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB máximo por archivo
  ALLOWED_ORIGINS: ['https://www.patenta.pe'],
  TOKEN_VALIDITY: 1800000 // 30 minutos en millisegundos
};

/**
 * Genera un token único para cada sesión
 */
function generarToken() {
  const timestamp = new Date().getTime();
  const random = Math.random().toString(36).substring(2);
  return `${timestamp}_${random}`;
}

/**
 * Valida el token y el tiempo transcurrido
 */
function validarToken(token, tiempoFormulario) {
  // Si no hay token o es del cliente, solo validar tiempo
  if (!token || typeof token !== 'string' || token.startsWith('client_')) {
    // Verificar tiempo mínimo del formulario
    if (tiempoFormulario && tiempoFormulario < SECURITY_CONFIG.MIN_FORM_TIME * 1000) {
      return { valido: false, error: 'Formulario completado muy rápido' };
    }
    return { valido: true }; // Permitir tokens de cliente
  }
  
  const partes = token.split('_');
  if (partes.length !== 2) {
    return { valido: false, error: 'Token malformado' };
  }
  
  const timestamp = parseInt(partes[0]);
  if (isNaN(timestamp)) {
    return { valido: false, error: 'Timestamp inválido' };
  }
  
  const ahora = new Date().getTime();
  const tiempoTranscurrido = ahora - timestamp;
  
  // Verificar que el token no sea muy antiguo (30 minutos)
  if (tiempoTranscurrido > SECURITY_CONFIG.TOKEN_VALIDITY) {
    return { valido: false, error: 'Token expirado' };
  }
  
  // Verificar tiempo mínimo del formulario
  if (tiempoFormulario && tiempoFormulario < SECURITY_CONFIG.MIN_FORM_TIME * 1000) {
    return { valido: false, error: 'Formulario completado muy rápido' };
  }
  
  return { valido: true };
}

/**
 * Rate limiting usando PropertiesService
 */
function verificarRateLimit(ip, email) {
  const propiedades = PropertiesService.getScriptProperties();
  const ahora = new Date().getTime();
  const unaHora = 60 * 60 * 1000;
  const unDia = 24 * unaHora;
  
  // Verificar límite por IP
  const claveIP = `ip_${ip}`;
  const datosIP = propiedades.getProperty(claveIP);
  if (datosIP) {
    const { count, timestamp } = JSON.parse(datosIP);
    if (ahora - timestamp < unaHora) {
      if (count >= SECURITY_CONFIG.MAX_REQUESTS_PER_IP) {
        return { permitido: false, error: 'Demasiadas solicitudes desde esta IP' };
      }
      propiedades.setProperty(claveIP, JSON.stringify({ count: count + 1, timestamp }));
    } else {
      propiedades.setProperty(claveIP, JSON.stringify({ count: 1, timestamp: ahora }));
    }
  } else {
    propiedades.setProperty(claveIP, JSON.stringify({ count: 1, timestamp: ahora }));
  }
  
  // Verificar límite por email
  const claveEmail = `email_${email}`;
  const datosEmail = propiedades.getProperty(claveEmail);
  if (datosEmail) {
    const { count, timestamp } = JSON.parse(datosEmail);
    if (ahora - timestamp < unDia) {
      if (count >= SECURITY_CONFIG.MAX_REQUESTS_PER_EMAIL) {
        return { permitido: false, error: 'Demasiadas solicitudes para este email' };
      }
      propiedades.setProperty(claveEmail, JSON.stringify({ count: count + 1, timestamp }));
    } else {
      propiedades.setProperty(claveEmail, JSON.stringify({ count: 1, timestamp: ahora }));
    }
  } else {
    propiedades.setProperty(claveEmail, JSON.stringify({ count: 1, timestamp: ahora }));
  }
  
  return { permitido: true };
}

/**
 * Valida el origen de la petición
 */
function validarOrigen(headers) {
  const origen = headers['origin'] || headers['referer'] || '';
  const origenPermitido = SECURITY_CONFIG.ALLOWED_ORIGINS.some(allowed => 
    origen.includes(allowed) || allowed === '*'
  );
  
  if (!origenPermitido && origen !== '') {
    return { valido: false, error: 'Origen no autorizado' };
  }
  
  return { valido: true };
}

/**
 * Valida los datos del formulario
 */
function validarDatos(datos) {
  const camposRequeridos = [
    'titulo',
    'categoriaParticipacion',
    'autorizacionContacto',
    'institucionEducativa',
    'direccionInstitucion',
    'regionInstitucion',
    'apellidosRepresentante', 'nombreRepresentante', 'dni', 
    'telefono', 'correo',
    'descripcion'
  ];
  
  // Verificar campos requeridos
  for (const campo of camposRequeridos) {
    if (!datos[campo] || datos[campo].trim() === '') {
      return { valido: false, error: `Campo requerido faltante: ${campo}` };
    }
  }
  
  // Validar campos de texto adicionales

  // (Las declaraciones obligatorias fueron removidas del formulario)

  // Validar categoría de participación (A o B)
  if (!datos.categoriaParticipacion || !['A', 'B'].includes(datos.categoriaParticipacion)) {
    return { valido: false, error: 'Categoría de participación inválida' };
  }

  // Validar autorización de contacto (Si o No)
  if (!datos.autorizacionContacto || !['Si', 'No'].includes(datos.autorizacionContacto)) {
    return { valido: false, error: 'Autorización de contacto inválida' };
  }
  
  // Verificar honeypot
  if (datos.honeypot && datos.honeypot !== '') {
    return { valido: false, error: 'Envío de bot detectado' };
  }
  
  // Validar tamaño de archivos
  if (datos.fichaPostulanteBase64 && datos.fichaPostulanteBase64.length > SECURITY_CONFIG.MAX_FILE_SIZE * 1.4) {
    return { valido: false, error: 'Archivo de ficha postulante muy grande' };
  }
  
  if (datos.fichaInvencionBase64 && datos.fichaInvencionBase64.length > SECURITY_CONFIG.MAX_FILE_SIZE * 1.4) {
    return { valido: false, error: 'Archivo de ficha invención muy grande' };
  }

  if (datos.declaracionParteAsesorBase64 && datos.declaracionParteAsesorBase64.length > SECURITY_CONFIG.MAX_FILE_SIZE * 1.4) {
    return { valido: false, error: 'Archivo de declaración parte asesor muy grande' };
  }
  
  // Validar inventores: al menos 1 y campos requeridos por inventor
  if (!datos.inventores || !Array.isArray(datos.inventores) || datos.inventores.length === 0) {
    return { valido: false, error: 'Debe agregar al menos 1 inventor' };
  }

  for (const inventor of datos.inventores) {
    if (!inventor.apellidos || !inventor.nombres || !inventor.genero || !inventor.edad || !inventor.grado) {
      return { valido: false, error: 'Cada inventor debe tener apellidos, nombres, género, edad y grado' };
    }

    if (!['Masculino', 'Femenino'].includes(inventor.genero)) {
      return { valido: false, error: 'Género de inventor inválido' };
    }
    // Validar edad como entero positivo
    const edadNum = parseInt(inventor.edad, 10);
    if (isNaN(edadNum) || edadNum <= 0) {
      return { valido: false, error: `Edad de inventor inválida: ${inventor.edad}` };
    }
  }
  
  return { valido: true };
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
    
    // Validar origen
    const validacionOrigen = validarOrigen(e.parameter);
    if (!validacionOrigen.valido) {
      if (idEnvio) {
        guardarEstadoEnvio(idEnvio, 'error', 'Acceso denegado: ' + validacionOrigen.error);
      }
      return createResponseWithCORS({
        success: false,
        message: 'Acceso denegado: ' + validacionOrigen.error
      });
    }
    
    // Validar datos
    const validacionDatos = validarDatos(params);
    if (!validacionDatos.valido) {
      if (idEnvio) {
        guardarEstadoEnvio(idEnvio, 'error', validacionDatos.error);
      }
      return createResponseWithCORS({
        success: false,
        message: validacionDatos.error
      });
    }
    
    // Validar token
    const validacionToken = validarToken(params.token, params.tiempoFormulario);
    if (!validacionToken.valido) {
      if (idEnvio) {
        guardarEstadoEnvio(idEnvio, 'error', 'Token inválido: ' + validacionToken.error);
      }
      return createResponseWithCORS({
        success: false,
        message: 'Token inválido: ' + validacionToken.error
      });
    }
    
    // Obtener IP del usuario (aproximada)
    const ip = e.parameter['user_ip'] || 'unknown';
    
    // Verificar rate limiting
    const verificacionRate = verificarRateLimit(ip, params.correo);
    if (!verificacionRate.permitido) {
      if (idEnvio) {
        guardarEstadoEnvio(idEnvio, 'error', verificacionRate.error);
      }
      return createResponseWithCORS({
        success: false,
        message: verificacionRate.error
      });
    }
    
    // Procesar datos si todas las validaciones pasan
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
 * Maneja peticiones GET para generar tokens y verificar estados
 */
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'getToken') {
    return createResponseWithCORS({
      success: true,
      token: generarToken(),
      timestamp: new Date().getTime()
    });
  }
  
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

