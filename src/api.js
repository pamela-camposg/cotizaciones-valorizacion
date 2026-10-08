/**
 * Cliente de la API (Google Apps Script) que expone el Google Sheet.
 *
 * Reemplaza al antiguo storage.js basado en localStorage: ahora los datos
 * viven en el Sheet y son los mismos para todo el equipo.
 *
 * Nota técnica: el cuerpo se envía como 'text/plain' a propósito. Si se
 * enviara como 'application/json', el navegador dispararía una petición
 * OPTIONS previa (preflight de CORS) que Apps Script no sabe responder, y
 * todas las llamadas fallarían. El servidor igual lo interpreta como JSON.
 */

// URL de la implementación del Apps Script (termina en /exec).
// Si vuelves a implementar con "Nueva implementación", esta URL cambia y
// hay que actualizarla acá.
export const API_URL =
  'https://script.google.com/macros/s/AKfycbxBJlzINKoiB41pSn44KIU6cmxvBp60THPIrXNpV1dtsZ9BkgeeQoebU6ZBwdPVZx0_-g/exec';

/** Llamada genérica a la API. Devuelve los datos o lanza un error con el mensaje del servidor. */
const esperar = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * Un intento. Devuelve { datos } si el servidor contestó JSON, o
 * { reintentable:true } si contestó otra cosa.
 *
 * Google a veces responde con una página HTML en vez de datos cuando el
 * script está ocupado. Eso se pasa solo, así que conviene reintentar en vez
 * de darle un error al usuario.
 */
async function intentar(cuerpo) {
  let respuesta;
  try {
    respuesta = await fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(cuerpo),
      redirect: 'follow',
    });
  } catch (e) {
    return { reintentable: true, motivo: 'sin-conexion' };
  }

  const texto = await respuesta.text();
  try {
    return { datos: JSON.parse(texto) };
  } catch (e) {
    return { reintentable: true, motivo: 'no-json', texto };
  }
}

async function llamar(accion, credenciales, extra) {
  const cuerpo = {
    accion,
    usuario: credenciales ? credenciales.usuario : undefined,
    pin: credenciales ? credenciales.pin : undefined,
    ...(extra || {}),
  };

  // Tres intentos, esperando cada vez un poco más.
  let ultimo = null;
  for (let i = 0; i < 3; i++) {
    if (i > 0) await esperar(i * 1200);
    ultimo = await intentar(cuerpo);
    if (ultimo.datos) break;
  }

  if (!ultimo.datos) {
    if (ultimo.motivo === 'sin-conexion') {
      throw new Error(
        'No se pudo contactar el servidor. Revise su conexión a internet, ' +
        'o que la implementación del Apps Script tenga acceso "Cualquier usuario".'
      );
    }
    // Deja rastro en la consola del navegador para poder diagnosticar.
    console.error('Respuesta no reconocida del servidor:', (ultimo.texto || '').slice(0, 400));
    throw new Error(
      'El servidor respondió en un formato inesperado después de 3 intentos. ' +
      'Suele ser pasajero: vuelva a cargar la página en unos segundos.'
    );
  }

  const datos = ultimo.datos;
  if (!datos.ok) throw new Error(datos.error || 'Error desconocido del servidor.');
  return datos;
}

/** Valida usuario y PIN. Devuelve { usuario, nombre }. */
export async function login(usuario, pin) {
  const r = await llamar('login', { usuario, pin });
  return { usuario: r.usuario, nombre: r.nombre };
}

/** Trae la lista completa de destinatarios desde el Sheet. */
export async function getDestinatarios(credenciales) {
  const r = await llamar('getDestinatarios', credenciales);
  return r.destinatarios || [];
}

/** Trae las cotizaciones registradas en el Sheet. */
export async function getCotizaciones(credenciales) {
  const r = await llamar('getCotizaciones', credenciales);
  return r.cotizaciones || [];
}

/** Guarda una cotización. El ID, la fecha y el responsable los asigna el servidor. */
export async function guardarCotizacion(credenciales, cotizacion) {
  const r = await llamar('guardarCotizacion', credenciales, { cotizacion });
  return r.id;
}

/**
 * Guarda varias líneas de residuo de una misma cotización. Cada línea recibe
 * su propio código, y los códigos quedan consecutivos.
 * Devuelve el arreglo de códigos asignados, en el mismo orden.
 */
export async function guardarCotizaciones(credenciales, cotizaciones) {
  const r = await llamar('guardarCotizaciones', credenciales, { cotizaciones });
  return r.ids || [];
}

/** Agrega un destinatario nuevo. Devuelve true si ya existía (no se duplicó). */
export async function agregarDestinatario(credenciales, destinatario) {
  const r = await llamar('agregarDestinatario', credenciales, { destinatario });
  return r.duplicado === true;
}
