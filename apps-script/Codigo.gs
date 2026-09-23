/**
 * API de la plataforma de Cotizaciones de Valorización.
 *
 * Este script vive dentro del Google Sheet y lo expone como servicio web,
 * para que la aplicación (publicada en GitHub Pages) pueda leer y escribir
 * en él. Sin esta pieza, una página estática no tiene forma de tocar Drive.
 *
 * Valida usuario + PIN contra la hoja USUARIOS en CADA petición. Los PIN
 * nunca salen del servidor: no se devuelven en ninguna respuesta.
 *
 * Instalación: ver README-apps-script.md
 */

/* ====================== CONFIGURACIÓN ====================== */

// ID del Google Sheet (el trozo de la URL entre /d/ y /edit).
var SHEET_ID = '1GWTZIHtO5E2SsiUoNhGyeDhqhkeKFHEIPMO1LNWu00w';

// Nombres de las hojas.
var HOJA_COTIZACIONES = 'COTIZACIONES';
var HOJA_DESTINATARIOS = 'DESTINATARIOS';
var HOJA_USUARIOS = 'USUARIOS';

// Último correlativo que vive en el Excel histórico (COT-01666).
// El Sheet parte vacío, así que el primer ID nuevo será COT-01667.
var ULTIMO_ID_HISTORICO = 1666;

// Zona horaria para la fecha de registro.
var ZONA_HORARIA = 'America/Santiago';

/* ====================== UTILIDADES ====================== */

function _libro() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function _hoja(nombre) {
  var h = _libro().getSheetByName(nombre);
  if (!h) throw new Error('No existe la hoja "' + nombre + '" en el Sheet.');
  return h;
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _ok(data) {
  var r = { ok: true };
  for (var k in data) r[k] = data[k];
  return _json(r);
}

function _error(mensaje) {
  return _json({ ok: false, error: mensaje });
}

/** Devuelve los encabezados de la fila 1 de una hoja. */
function _encabezados(hoja) {
  return hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
}

/** Convierte una hoja completa en un arreglo de objetos. */
function _filasComoObjetos(hoja) {
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];
  var ultimaCol = hoja.getLastColumn();
  var valores = hoja.getRange(1, 1, ultimaFila, ultimaCol).getValues();
  var enc = valores[0];
  var out = [];
  for (var i = 1; i < valores.length; i++) {
    var fila = valores[i];
    var vacia = fila.every(function (v) { return v === '' || v === null; });
    if (vacia) continue;
    var obj = {};
    for (var j = 0; j < enc.length; j++) {
      if (enc[j] === '' || enc[j] === null) continue;
      obj[String(enc[j]).trim()] = fila[j];
    }
    out.push(obj);
  }
  return out;
}

/**
 * Escribe una fila mapeando el objeto a las columnas POR NOMBRE de
 * encabezado. Así el orden de las columnas del Sheet puede cambiar sin
 * romper nada, y una columna que el objeto no traiga queda simplemente
 * en blanco.
 */
function _agregarFila(hoja, objeto) {
  var enc = _encabezados(hoja);
  var fila = enc.map(function (h) {
    var clave = String(h).trim();
    var v = objeto[clave];
    return (v === undefined || v === null) ? '' : v;
  });
  hoja.appendRow(fila);
}

/* ====================== AUTENTICACIÓN ====================== */

/**
 * Valida usuario + PIN contra la hoja USUARIOS.
 * Lanza una excepción si las credenciales no son válidas.
 *
 * Nota: la comparación del PIN es como texto, para que un PIN que parta
 * en cero (0472) funcione igual que uno que no.
 */
function _validarCredenciales(usuario, pin) {
  if (!usuario || !pin) throw new Error('Faltan credenciales.');

  var filas = _filasComoObjetos(_hoja(HOJA_USUARIOS));
  for (var i = 0; i < filas.length; i++) {
    var f = filas[i];
    var u = String(f['USUARIO'] || '').trim();
    if (u !== String(usuario).trim()) continue;

    var activo = String(f['ACTIVO'] || '').trim().toUpperCase();
    if (activo !== 'SI') throw new Error('El usuario está desactivado.');

    var pinGuardado = String(f['PIN'] || '').trim();
    if (pinGuardado === '') throw new Error('El usuario no tiene PIN asignado.');
    if (pinGuardado !== String(pin).trim()) throw new Error('Usuario o PIN incorrecto.');

    return { usuario: u, nombre: f['NOMBRE'] || '' };
  }
  throw new Error('Usuario o PIN incorrecto.');
}

/* ====================== CORRELATIVO DE ID ====================== */

/**
 * Calcula el próximo ID. Toma el mayor entre los IDs ya presentes en el
 * Sheet y el último del Excel histórico, y le suma 1.
 */
function _siguienteId(hojaCot) {
  var maximo = ULTIMO_ID_HISTORICO;
  var ultimaFila = hojaCot.getLastRow();

  if (ultimaFila >= 2) {
    var enc = _encabezados(hojaCot);
    var colId = enc.indexOf('ID') + 1;
    if (colId > 0) {
      var ids = hojaCot.getRange(2, colId, ultimaFila - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        var m = String(ids[i][0] || '').match(/COT-(\d+)/);
        if (m) {
          var n = parseInt(m[1], 10);
          if (n > maximo) maximo = n;
        }
      }
    }
  }

  var siguiente = maximo + 1;
  return 'COT-' + ('00000' + siguiente).slice(-5);
}

/* ====================== ACCIONES ====================== */

function accionLogin(p) {
  var u = _validarCredenciales(p.usuario, p.pin);
  return _ok({ usuario: u.usuario, nombre: u.nombre });
}

function accionGetDestinatarios(p) {
  _validarCredenciales(p.usuario, p.pin);
  var filas = _filasComoObjetos(_hoja(HOJA_DESTINATARIOS));
  // Se normalizan los nombres de columna que en el Excel original vienen
  // con espacios (y una con un typo heredado: ESTABLECIMIETNO).
  var out = filas.map(function (f) {
    return {
      NOMBRE_FANTASIA: f['NOMBRE_FANTASIA'],
      SISTEMA_DECLARACION: f['SISTEMA_DECLARACION'],
      CODIGO_ESTABLECIMIENTO: f['CODIGO_ESTABLECIMIENTO'],
      RAZON_SOCIAL: f['RAZON_SOCIAL'],
      RUT: f['RUT'],
      NOMBRE_ESTABLECIMIENTO: f['NOMBRE DE ESTABLECIMIENTO'],
      COMUNA_ESTABLECIMIENTO: f['COMUNA DE ESTABLECIMIENTO'],
      REGION_ESTABLECIMIENTO: f['REGION DE ESTABLECIMIETNO']
    };
  });
  return _ok({ destinatarios: out });
}

function accionGetCotizaciones(p) {
  _validarCredenciales(p.usuario, p.pin);
  var filas = _filasComoObjetos(_hoja(HOJA_COTIZACIONES));
  // Las fechas viajan como texto ISO para que no las altere la zona horaria.
  filas.forEach(function (f) {
    if (f['FECHA'] instanceof Date) {
      f['FECHA'] = Utilities.formatDate(f['FECHA'], ZONA_HORARIA, 'yyyy-MM-dd');
    }
  });
  return _ok({ cotizaciones: filas });
}

function accionGuardarCotizacion(p) {
  var u = _validarCredenciales(p.usuario, p.pin);
  if (!p.cotizacion) throw new Error('No se recibió la cotización.');

  // El bloqueo evita que dos personas guardando al mismo tiempo obtengan
  // el mismo correlativo.
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var hoja = _hoja(HOJA_COTIZACIONES);
    var id = _siguienteId(hoja);

    var registro = p.cotizacion;
    registro['ID'] = id;
    registro['FECHA'] = Utilities.formatDate(new Date(), ZONA_HORARIA, 'yyyy-MM-dd');
    // El responsable se toma de la sesión validada, nunca de lo que
    // mande el cliente.
    registro['RESPONSABLE'] = u.usuario;

    _agregarFila(hoja, registro);
    return _ok({ id: id });
  } finally {
    lock.releaseLock();
  }
}

function accionAgregarDestinatario(p) {
  _validarCredenciales(p.usuario, p.pin);
  if (!p.destinatario) throw new Error('No se recibió el destinatario.');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var hoja = _hoja(HOJA_DESTINATARIOS);
    var d = p.destinatario;
    var nombre = String(d.NOMBRE_FANTASIA || '').trim();
    if (!nombre) throw new Error('El destinatario no tiene nombre de fantasía.');

    // Evita duplicar un destinatario que ya existe con la misma comuna.
    var existentes = _filasComoObjetos(hoja);
    for (var i = 0; i < existentes.length; i++) {
      var e = existentes[i];
      if (String(e['NOMBRE_FANTASIA'] || '').trim().toUpperCase() === nombre.toUpperCase() &&
          String(e['COMUNA DE ESTABLECIMIENTO'] || '').trim().toUpperCase() ===
            String(d.COMUNA_ESTABLECIMIENTO || '').trim().toUpperCase()) {
        return _ok({ duplicado: true, mensaje: 'El destinatario ya existía; no se agregó de nuevo.' });
      }
    }

    _agregarFila(hoja, {
      'NOMBRE_FANTASIA': d.NOMBRE_FANTASIA,
      'SISTEMA_DECLARACION': d.SISTEMA_DECLARACION,
      'CODIGO_ESTABLECIMIENTO': d.CODIGO_ESTABLECIMIENTO,
      'RAZON_SOCIAL': d.RAZON_SOCIAL,
      'RUT': d.RUT,
      'NOMBRE DE ESTABLECIMIENTO': d.NOMBRE_ESTABLECIMIENTO,
      'COMUNA DE ESTABLECIMIENTO': d.COMUNA_ESTABLECIMIENTO,
      'REGION DE ESTABLECIMIETNO': d.REGION_ESTABLECIMIENTO
    });

    return _ok({ duplicado: false });
  } finally {
    lock.releaseLock();
  }
}

/* ====================== PUNTOS DE ENTRADA ====================== */

var ACCIONES = {
  login: accionLogin,
  getDestinatarios: accionGetDestinatarios,
  getCotizaciones: accionGetCotizaciones,
  guardarCotizacion: accionGuardarCotizacion,
  agregarDestinatario: accionAgregarDestinatario
};

/**
 * Todas las peticiones de la aplicación entran por acá.
 * El cuerpo llega como texto plano (no como JSON) a propósito: así el
 * navegador no dispara una petición OPTIONS previa, que Apps Script no
 * sabe responder.
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return _error('Petición vacía.');
    }
    var p = JSON.parse(e.postData.contents);
    var fn = ACCIONES[p.accion];
    if (!fn) return _error('Acción desconocida: ' + p.accion);
    return fn(p);
  } catch (err) {
    return _error(err.message || String(err));
  }
}

/** Solo para comprobar en el navegador que la implementación quedó viva. */
function doGet() {
  return _json({
    ok: true,
    servicio: 'API Cotizaciones de Valorización',
    mensaje: 'El servicio está activo. Las operaciones se hacen por POST.'
  });
}

/* ====================== PRUEBA MANUAL ====================== */

/**
 * Ejecuta esta función desde el editor de Apps Script (botón Ejecutar)
 * para comprobar que el script ve bien el Sheet. No modifica nada.
 * El resultado aparece en el Registro de ejecución.
 */
function probarConexion() {
  var libro = _libro();
  Logger.log('Sheet: %s', libro.getName());

  var usuarios = _filasComoObjetos(_hoja(HOJA_USUARIOS));
  var conPin = usuarios.filter(function (u) {
    return String(u['PIN'] || '').trim() !== '';
  });
  Logger.log('Usuarios: %s en total, %s con PIN asignado.', usuarios.length, conPin.length);

  Logger.log('Destinatarios: %s', _filasComoObjetos(_hoja(HOJA_DESTINATARIOS)).length);
  Logger.log('Cotizaciones en el Sheet: %s', _filasComoObjetos(_hoja(HOJA_COTIZACIONES)).length);
  Logger.log('Próximo ID que se asignaría: %s', _siguienteId(_hoja(HOJA_COTIZACIONES)));

  if (conPin.length === 0) {
    Logger.log('ATENCION: ningun usuario tiene PIN. Nadie podra entrar.');
  }
}
