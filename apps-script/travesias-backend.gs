/**
 * TRAVESÍAS — backend de la convocatoria
 * ===============================================================
 * Recibe las inscripciones de la página y deja en Drive una carpeta
 * por participante con los archivos ya renombrados, más una fila en
 * el Sheet de inscripciones.
 *
 * POR QUÉ ESTÁ ESCRITO ASÍ (no tocar sin leer):
 *
 *  1. El participante NO necesita cuenta de Google. Este script se
 *     implementa con "Ejecutar como: yo" y "Acceso: cualquier
 *     persona", así que escribe en Drive con los permisos de
 *     equipo@, nunca con los del visitante.
 *
 *  2. Los bytes NO pasan por aquí. El script le pide a Drive una
 *     sesión de subida reanudable y devuelve esa URL; el navegador
 *     sube directo a Google. Por eso no hay techo de tamaño pese al
 *     límite de ejecución de Apps Script.
 *
 *  3. Las peticiones llegan como text/plain a propósito. Con
 *     application/json el navegador exige un preflight OPTIONS que
 *     Apps Script no sabe responder.
 *
 *  4. Google amarra el permiso CORS de la sesión al encabezado
 *     Origin declarado AL CREARLA. Si el dominio de la página no
 *     está en ORIGENES, el navegador rechaza la subida después,
 *     con todo lo demás correcto.
 *
 * DESPLIEGUE
 *   1. script.google.com → Nuevo proyecto → pegar esto
 *   2. Ejecutar `probarPermisos` una vez y autorizar
 *      (aparece "app no verificada": Configuración avanzada → Ir a…)
 *   3. Implementar → Nueva implementación → Aplicación web
 *        Ejecutar como:      Yo (equipo@estudiolaselva.com)
 *        Quién tiene acceso: CUALQUIER PERSONA
 *   4. Copiar la URL /exec y ponerla en la constante BACKEND de la página
 *
 *   OJO: cada cambio de código exige una implementación NUEVA (o editar
 *   la existente eligiendo "Nueva versión"). Si no, la URL sigue
 *   sirviendo el código viejo — causa número uno de "pero si ya lo arreglé".
 */

/* ===============================================================
   CONFIGURACIÓN
   =============================================================== */

/* Dominios autorizados a subir. Exactos: con esquema, con puerto si
   lo hay, SIN barra final. Los despliegues de vista previa de Vercel
   tienen otro dominio y no pasan por aquí. */
var ORIGENES = [
  'https://travesias-neon.vercel.app',
  'http://localhost:8000'
];

var CARPETA_ENTREGAS = '1FCaKaLh2o3-9L5Ji9-3-Zmn2aqiSQ4W6';   // Travesías/Entregas
var SHEET_ID         = '12A8JmULj-MvN6blZPVLB1woB0I8NKJmLqOd-aTTR-HI';
var HOJA             = 'Inscripciones';

/* Cierre: 29 de enero de 2027, 23:59 hora de Colombia (UTC-5) */
var CIERRE = new Date('2027-01-29T23:59:00-05:00');

/* Límites. PENDIENTES DE DEFINIR con Uniandes — estos son provisionales
   y el formulario de la página no los conoce todavía. */
var MAX_ARCHIVOS      = 20;
var MAX_BYTES_ARCHIVO = 5 * 1024 * 1024 * 1024;    //  5 GB por archivo
var MAX_BYTES_TOTAL   = 20 * 1024 * 1024 * 1024;   // 20 GB por propuesta

var EXT_FOTO  = ['.jpg', '.jpeg', '.png'];
var EXT_VIDEO = ['.mp4', '.mov'];

/* Correo de confirmación al participante, enviado desde la cuenta que
   corre el script. Al estar encendido, Apps Script pide un permiso más
   (enviar correo en tu nombre): la próxima vez que ejecutes
   `probarPermisos` habrá que autorizarlo de nuevo.
   Cuota diaria de Workspace: 1500 destinatarios, de sobra aquí. */
var ENVIAR_CONFIRMACION = true;
var RESPONDER_A = 'equipo@estudiolaselva.com';

/* ===============================================================
   ENTRADA
   =============================================================== */

/** Ejecutar a mano una vez: dispara el diálogo de permisos. */
function probarPermisos() {
  var c = DriveApp.getFolderById(CARPETA_ENTREGAS);
  var h = SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA);
  Logger.log('Carpeta: ' + c.getName());
  Logger.log('Hoja: ' + (h ? h.getName() + ' — ' + h.getLastRow() + ' fila(s)' : 'NO ENCONTRADA'));
  Logger.log('Token: ' + (ScriptApp.getOAuthToken().length > 20 ? 'ok' : 'falta'));
}

/** Abrir la URL /exec en el navegador debe mostrar esto. */
function doGet() {
  return json({
    ok: true,
    servicio: 'Travesías — inscripciones',
    abierta: Date.now() <= CIERRE.getTime(),
    cierre: CIERRE.toISOString()
  });
}

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    switch (p.accion) {
      case 'inscribir': return json(inscribir(p));
      case 'sesion':    return json(sesion(p));
      case 'cerrar':    return json(cerrar(p));
      default:          return json({ ok: false, error: 'acción desconocida: ' + p.accion });
    }
  } catch (err) {
    /* el detalle va al registro del script, no al participante */
    console.error(err);
    return json({ ok: false, error: 'error interno; vuelve a intentarlo' });
  }
}

/* ===============================================================
   1 · INSCRIBIR — crea carpeta, radicado y fila
   =============================================================== */

function inscribir(p) {
  if (ORIGENES.indexOf(p.origen) === -1)
    return { ok: false, error: 'origen no autorizado' };

  if (Date.now() > CIERRE.getTime())
    return { ok: false, error: 'la convocatoria cerró el 29 de enero de 2027' };

  var d = p.datos || {};
  var falta = [];
  if (!txt(d.nombre))   falta.push('nombre');
  if (!txt(d.apellido)) falta.push('apellido');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(txt(d.correo))) falta.push('correo');
  if (!txt(d.titulo))   falta.push('título');
  if (falta.length) return { ok: false, error: 'faltan datos: ' + falta.join(', ') };

  /* el cliente ya validó, pero el cliente no es de fiar */
  var archivos = p.archivos || [];
  if (!archivos.length)              return { ok: false, error: 'no hay archivos' };
  if (archivos.length > MAX_ARCHIVOS) return { ok: false, error: 'máximo ' + MAX_ARCHIVOS + ' archivos' };

  var total = 0;
  for (var i = 0; i < archivos.length; i++) {
    var a = archivos[i];
    var ext = extension(a.nombre);
    if (EXT_FOTO.indexOf(ext) === -1 && EXT_VIDEO.indexOf(ext) === -1)
      return { ok: false, error: 'formato no aceptado: ' + a.nombre };
    if (a.tamano > MAX_BYTES_ARCHIVO)
      return { ok: false, error: 'archivo demasiado grande: ' + a.nombre };
    total += a.tamano;
  }
  if (total > MAX_BYTES_TOTAL) return { ok: false, error: 'la propuesta supera el tamaño máximo' };

  /* radicado secuencial — bajo candado, o dos envíos simultáneos
     se llevan el mismo número */
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var hoja = SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA);
    var radicado = 'TRV-' + pad(hoja.getLastRow(), 4);   // fila 1 = encabezado
    var slug = slugNombre(d.apellido, d.nombre);

    var carpeta = DriveApp.getFolderById(CARPETA_ENTREGAS).createFolder(radicado + '_' + slug);
    var fotos = carpeta.createFolder('fotos');
    var videos = carpeta.createFolder('videos');

    var plan = planearArchivos(archivos, radicado, slug);
    var tecnico = {
      carpeta: carpeta.getId(),
      fotos: fotos.getId(),
      videos: videos.getId(),
      plan: plan
    };

    /* la propuesta también queda como texto dentro de la carpeta: si
       algún día el Sheet se pierde o se reorganiza, la carpeta sigue
       diciendo de quién es y qué contiene */
    carpeta.createFile('propuesta.txt', fichaTexto(radicado, d, plan), MimeType.PLAIN_TEXT);

    hoja.appendRow([
      radicado,
      new Date(),
      txt(d.nombre), txt(d.apellido), txt(d.correo), txt(d.telefono),
      txt(d.titulo), txt(d.descripcion),
      plan.filter(esFoto).map(nuevoNombre).join('\n'),
      plan.filter(esVideo).map(nuevoNombre).join('\n'),
      plan.map(function (x) { return x.original + '  →  ' + x.nuevo; }).join('\n'),
      carpeta.getUrl(),
      'sí',
      'subiendo',
      JSON.stringify(tecnico)
    ]);

    return { ok: true, radicado: radicado };
  } finally {
    lock.releaseLock();
  }
}

/* ===============================================================
   2 · SESIÓN — una URL de subida por archivo
   ===============================================================
   Sin radicado funciona como prueba suelta: sube a Entregas y sirve
   para verificar CORS sin inscribir a nadie (lo usa prueba-subida.html).
   =============================================================== */

function sesion(p) {
  if (ORIGENES.indexOf(p.origen) === -1)
    return { ok: false, error: 'origen no autorizado: ' + p.origen };

  var destino, nombre;

  if (p.radicado) {
    var fila = buscarFila(p.radicado);
    if (!fila) return { ok: false, error: 'radicado no encontrado' };
    var t = JSON.parse(fila.hoja.getRange(fila.n, 15).getValue());   // columna O: tecnico
    var item = t.plan[p.indice];
    if (!item) return { ok: false, error: 'índice de archivo fuera de rango' };
    if (item.original !== p.nombre)
      return { ok: false, error: 'el archivo no corresponde al índice' };
    destino = (item.carpeta === 'videos') ? t.videos : t.fotos;
    nombre = item.nuevo;
  } else {
    destino = CARPETA_ENTREGAS;
    nombre = 'PRUEBA_' + new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-') +
             '_' + p.nombre;
  }

  var r = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true',
    {
      method: 'post',
      contentType: 'application/json; charset=UTF-8',
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
        'X-Upload-Content-Type': p.tipo || 'application/octet-stream',
        'X-Upload-Content-Length': String(p.tamano),
        Origin: p.origen            /* ← la línea de la que depende todo */
      },
      payload: JSON.stringify({ name: nombre, parents: [destino] }),
      muteHttpExceptions: true,
      followRedirects: false
    }
  );

  var cod = r.getResponseCode();
  if (cod !== 200 && cod !== 201)
    return { ok: false, error: 'Drive respondió ' + cod,
             detalle: r.getContentText().slice(0, 300) };

  var h = r.getAllHeaders();
  var url = h['Location'] || h['location'];
  if (!url) return { ok: false, error: 'Drive no devolvió la URL de subida' };

  return { ok: true, url: url, nombre: nombre };
}

/* ===============================================================
   3 · CERRAR — marca la propuesta como completa
   =============================================================== */

function cerrar(p) {
  var fila = buscarFila(p.radicado);
  if (!fila) return { ok: false, error: 'radicado no encontrado' };

  fila.hoja.getRange(fila.n, 14).setValue('completa');   // columna N: estado

  if (ENVIAR_CONFIRMACION) {
    try {
      var v = fila.hoja.getRange(fila.n, 1, 1, 13).getValues()[0];
      var lista = [v[8], v[9]].filter(String).join('\n').split('\n').filter(String);
      MailApp.sendEmail({
        to: v[4],
        replyTo: RESPONDER_A,
        name: 'Travesías',
        subject: 'Travesías — recibimos tu propuesta (' + v[0] + ')',
        body: 'Hola ' + v[2] + ',\n\n' +
              'Recibimos tu propuesta "' + v[6] + '". Queda registrada con el\n' +
              'número de radicado ' + v[0] + ' — guárdalo, es tu comprobante de\n' +
              'entrega y el que usaremos para cualquier consulta.\n\n' +
              'Archivos recibidos (' + lista.length + '):\n' +
              lista.map(function (x) { return '  · ' + x; }).join('\n') + '\n\n' +
              'Si algo de esto no coincide con lo que enviaste, respóndenos a\n' +
              'este correo citando tu radicado.\n\n' +
              'Te escribiremos cuando haya resultados.\n\n' +
              '—\n' +
              'Travesías\n' +
              'Universidad de los Andes · Antropología × Estudio La Selva'
      });
    } catch (err) {
      console.error('no se pudo enviar la confirmación: ' + err);
      /* el envío del correo no invalida la inscripción */
    }
  }
  return { ok: true };
}

/* ===============================================================
   AUXILIARES
   =============================================================== */

function buscarFila(radicado) {
  var hoja = SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA);
  var col = hoja.getRange(1, 1, hoja.getLastRow(), 1).getValues();
  for (var i = col.length - 1; i >= 1; i--) {     // de abajo hacia arriba: lo reciente primero
    if (col[i][0] === radicado) return { hoja: hoja, n: i + 1 };
  }
  return null;
}

function planearArchivos(archivos, radicado, slug) {
  var nf = 0, nv = 0;
  return archivos.map(function (a) {
    var ext = extension(a.nombre);
    var video = EXT_VIDEO.indexOf(ext) !== -1;
    var n = video ? ++nv : ++nf;
    return {
      original: a.nombre,
      nuevo: radicado + '_' + slug + '_' + pad(n, 2) + ext,
      carpeta: video ? 'videos' : 'fotos',
      tamano: a.tamano
    };
  });
}

function fichaTexto(radicado, d, plan) {
  return [
    'TRAVESÍAS — propuesta ' + radicado,
    'Recibida: ' + new Date().toLocaleString('es-CO'),
    '',
    'Participante: ' + txt(d.nombre) + ' ' + txt(d.apellido),
    'Correo: ' + txt(d.correo),
    'Teléfono: ' + txt(d.telefono),
    '',
    'Título: ' + txt(d.titulo),
    '',
    'Descripción:',
    txt(d.descripcion),
    '',
    'Archivos (original → archivado):',
    plan.map(function (x) { return '  ' + x.original + '  →  ' + x.nuevo; }).join('\n'),
    '',
    'Autorizaciones de autoría, consentimiento y uso: aceptadas en el formulario.'
  ].join('\n');
}

/* el nombre del archivo tiene que sobrevivir a cualquier sistema:
   sin tildes, sin eñes, sin espacios */
function slugNombre(apellido, nombre) {
  var s = txt(apellido) + '-' + txt(nombre);
  return s.normalize('NFD').replace(/[\u0300-\u036f]/g, '')
          .replace(/ñ/gi, 'n')
          .replace(/[^A-Za-z0-9]+/g, '-')
          .replace(/^-+|-+$/g, '')
          .slice(0, 48) || 'sin-nombre';
}

function extension(n) {
  var m = String(n).match(/\.[A-Za-z0-9]+$/);
  return m ? m[0].toLowerCase() : '';
}
function esFoto(x)  { return x.carpeta === 'fotos'; }
function esVideo(x) { return x.carpeta === 'videos'; }
function nuevoNombre(x) { return x.nuevo; }
function pad(n, a) { var s = String(n); while (s.length < a) s = '0' + s; return s; }
function txt(v) { return (v == null ? '' : String(v)).trim(); }

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
