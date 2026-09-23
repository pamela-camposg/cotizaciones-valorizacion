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
  'https://script.google.com/macros/s/AKfycbwHJDpcFu2aMIIOm31WFdNyyCMNmy6IVs3ytFBtgy6l9IYgpiCX5dH7Wz45012YLsadFw/exec';

/** Llamada genérica a la API. Devuelve los datos o lanza un error con el mensaje del servidor. */
async function llamar(accion, credenciales, extra) {
  const cuerpo = {
    accion,
    usuario: credenciales ? credenciales.usuario : undefined,
    pin: credenciales ? credenciales.pin : undefined,
    ...(extra || {}),
  };

  let respuesta;
  try {
    respuesta = await fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(cuerpo),
      redirect: 'follow',
    });
  } catch (e) {
    throw new Error(
      'No se pudo contactar el servidor. Revise su conexión a internet, ' +
      'o que la implementación del Apps Script tenga acceso "Cualquier usuario".'
    );
  }

  let datos;
  try {
    datos = await respuesta.json();
  } catch (e) {
    throw new Error('El servidor respondió en un formato inesperado.');
  }

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

/** Agrega un destinatario nuevo. Devuelve true si ya existía (no se duplicó). */
export async function agregarDestinatario(credenciales, destinatario) {
  const r = await llamar('agregarDestinatario', credenciales, { destinatario });
  return r.duplicado === true;
}
