'use strict';
/**
 * Sirve las fichas tecnicas de cocina y bar del Club Lagos de Caujaral.
 *
 * La pagina es un archivo autonomo que se genera en el PC de Sergio con
 * "Generar Fichas Tecnicas.ps1" a partir del libro de recetas, y se sube al
 * repositorio. Aqui solo se entrega: no hay base de datos ni consultas vivas.
 *
 * Para actualizar lo que ve la cocina: correr el generador, hacer commit y push.
 * Railway redespliega solo y el link no cambia.
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
const crypto = require('crypto');

const PUERTO = Number(process.env.PORT || 3000);

/* Clave de acceso. CERRADO POR DEFECTO, por decision de Sergio del 09/08/2026.
   Si no existe la variable FICHAS_PASSWORD, la pagina NO se entrega: muestra un
   aviso de que falta configurarla. Es a proposito — el archivo lleva adentro las
   891 recetas del club y el costo de cada insumo, y un despliegue sin clave las
   deja al alcance de cualquiera que tenga la direccion.
   Para abrirla hay que crear FICHAS_PASSWORD en Railway; no hay forma de dejarla
   publica por descuido. */
const USUARIO = process.env.FICHAS_USER || 'caujaral';
const CLAVE = process.env.FICHAS_PASSWORD || '';
/* Segundo perfil: los operarios de cocina y bar. Entran con otra clave y
   reciben OTRO ARCHIVO, generado sin un solo costo adentro. No es que se les
   escondan los botones: los numeros no viajan en la pagina que descargan, asi
   que no hay nada que destapar con las herramientas del navegador.
   Si no se configura FICHAS_PASSWORD_OP, simplemente no existe ese perfil. */
/* Los operarios, uno por persona. Cada uno es una variable de entorno que
   empieza por FICHAS_OP_ : el nombre de usuario es lo que sigue, en minusculas,
   y el valor es su clave.

     FICHAS_OP_DGUERRA = <clave de D. Guerra>   ->  entra como "dguerra"
     FICHAS_OP_AMEZA   = <clave de A. Meza>     ->  entra como "ameza"

   Agregar a alguien es crear una variable; quitarle el acceso es borrarla, sin
   tocar a los demas. Se admite tambien el par FICHAS_USER_OP/FICHAS_PASSWORD_OP
   de antes, para no romper lo que ya estuviera configurado.

   Las claves quedan en texto plano en las variables de Railway: las ve quien
   tenga acceso al panel. Es aceptable mientras sean claves de consulta y cada
   persona tenga la suya, pero el destino es la tabla de usuarios con la clave
   cifrada (Argon2id) de la fase 1 del diagnostico. */
const OPERARIOS = (function () {
  const mapa = new Map();
  for (const nombre of Object.keys(process.env)) {
    if (!nombre.startsWith('FICHAS_OP_')) continue;
    const usuario = nombre.slice('FICHAS_OP_'.length).toLowerCase();
    const clave = process.env[nombre] || '';
    if (usuario && clave) mapa.set(usuario, clave);
  }
  // el par antiguo, si existe
  const viejoUsuario = (process.env.FICHAS_USER_OP || 'operario').toLowerCase();
  const viejaClave = process.env.FICHAS_PASSWORD_OP || '';
  if (viejaClave) mapa.set(viejoUsuario, viejaClave);
  return mapa;
})();
const ABIERTA_A_PROPOSITO = String(process.env.FICHAS_PUBLICA || '').toLowerCase() === 'si';

const ARCHIVO = path.join(__dirname, 'Fichas Tecnicas Caujaral.html');
const ARCHIVO_OP = path.join(__dirname, 'Fichas Tecnicas Caujaral - operarios.html');

/* La pagina se lee una vez al arrancar: cada despliegue trae su propia version. */
let PAGINA = '';
let PAGINA_GZ = null;
let PAGINA_OP = '';
let PAGINA_OP_GZ = null;
let ARRANQUE = new Date();
let RECETAS = 0;
let CORTE = '';
let LOGO = null;   // el logo en binario, para servirlo como favicon

function cargar() {
  PAGINA = fs.readFileSync(ARCHIVO, 'utf8');
  PAGINA_GZ = zlib.gzipSync(Buffer.from(PAGINA, 'utf8'), { level: 9 });

  // Se leen las cifras del propio HTML para poder reportarlas en /salud
  const corte = PAGINA.match(/"generado":"([^"]+)"/);
  CORTE = corte ? corte[1] : '';
  const recetas = PAGINA.match(/"c":"/g);
  RECETAS = recetas ? recetas.length : 0;

  // Se saca el logo de la propia pagina, del atributo del icono
  const icono = PAGINA.match(/rel="icon"[^>]*href="data:image\/png;base64,([^"]+)"/);
  LOGO = icono ? Buffer.from(icono[1], 'base64') : null;
  // La pagina del operario es opcional: si el generador todavia no la produjo,
  // el perfil simplemente no se ofrece, en vez de caerse al arrancar.
  if (fs.existsSync(ARCHIVO_OP)) {
    PAGINA_OP = fs.readFileSync(ARCHIVO_OP, 'utf8');
    PAGINA_OP_GZ = zlib.gzipSync(Buffer.from(PAGINA_OP, 'utf8'), { level: 9 });
  } else {
    PAGINA_OP = '';
    PAGINA_OP_GZ = null;
  }
}

/* Devuelve el PERFIL de quien entra, no un simple si/no:
     'admin'    -> ve todo, con costos
     'operario' -> recibe la pagina sin costos
     null       -> no entra
   La comparacion se hace con tiempo constante para no filtrar la clave por lo
   que se demora en responder. */
function comparaSeguro(a, b) {
  const x = Buffer.from(String(a));
  const y = Buffer.from(String(b));
  if (x.length !== y.length) {
    // se compara igual contra si mismo para no delatar el largo con el tiempo
    crypto.timingSafeEqual(x, x);
    return false;
  }
  return crypto.timingSafeEqual(x, y);
}

function perfilDe(req) {
  if (!CLAVE) return ABIERTA_A_PROPOSITO ? 'admin' : null;
  const cabecera = req.headers.authorization || '';
  if (!cabecera.startsWith('Basic ')) return null;
  let plano;
  try {
    plano = Buffer.from(cabecera.slice(6), 'base64').toString('utf8');
  } catch (e) {
    return null;
  }
  const corte = plano.indexOf(':');
  if (corte < 0) return null;
  const usuario = plano.slice(0, corte);
  const clave = plano.slice(corte + 1);

  if (comparaSeguro(usuario, USUARIO) && comparaSeguro(clave, CLAVE)) return 'admin';
  /* El perfil de operario solo existe si hay pagina sin costos que entregar.
     Se compara SIEMPRE, exista o no el usuario, para no delatar por el tiempo
     de respuesta cuales nombres estan dados de alta. */
  if (PAGINA_OP) {
    const guardada = OPERARIOS.get(usuario.toLowerCase());
    const contra = guardada || 'clave-que-no-existe';
    const coincide = comparaSeguro(clave, contra);
    if (guardada && coincide) return 'operario';
  }
  return null;
}

function autorizado(req) {
  return perfilDe(req) !== null;
}

function pedirClave(res) {
  if (!CLAVE) {
    // Ni siquiera se pide clave: no hay ninguna configurada, asi que no hay
    // manera de entrar. Se responde sin pistas sobre lo que hay detras.
    res.writeHead(503, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('Esta pagina no esta disponible.');
    return;
  }
  res.writeHead(401, {
    'WWW-Authenticate': 'Basic realm="Fichas tecnicas Caujaral", charset="UTF-8"',
    'Content-Type': 'text/plain; charset=utf-8'
  });
  res.end('Necesitas usuario y clave para ver las fichas tecnicas.');
}

const servidor = http.createServer((req, res) => {
  const ruta = (req.url || '/').split('?')[0];

  if (ruta === '/salud') {
    // No revela cuantas recetas hay ni desde cuando, si la pagina esta cerrada
    var salud = { estado: 'ok', protegidoConClave: Boolean(CLAVE), sirviendo: Boolean(CLAVE) || ABIERTA_A_PROPOSITO };
    if (salud.sirviendo) {
      salud.recetas = RECETAS;
      salud.corteDeDatos = CORTE;
      salud.arranque = ARRANQUE.toISOString();
    }
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify(salud, null, 2));
    return;
  }

  const perfil = perfilDe(req);
  if (perfil === null) { pedirClave(res); return; }

  if (ruta === '/' || ruta === '/index.html' || ruta === '/fichas') {
    const aceptaGzip = /\bgzip\b/.test(req.headers['accept-encoding'] || '');
    // cada perfil recibe SU archivo: el del operario no trae los costos
    const cuerpo = (perfil === 'operario') ? PAGINA_OP : PAGINA;
    const cuerpoGz = (perfil === 'operario') ? PAGINA_OP_GZ : PAGINA_GZ;
    const cabeceras = {
      'Content-Type': 'text/html; charset=utf-8',
      'Cache-Control': 'no-cache',
      'X-Content-Type-Options': 'nosniff',
      // que no lo indexen los buscadores
      'X-Robots-Tag': 'noindex, nofollow',
      // si se comparte el enlace, que la respuesta no se guarde en un cache
      // intermedio y termine entregandole a un operario la pagina con costos
      'Vary': 'Authorization, Accept-Encoding'
    };
    if (aceptaGzip && cuerpoGz) {
      cabeceras['Content-Encoding'] = 'gzip';
      cabeceras['Content-Length'] = cuerpoGz.length;
      res.writeHead(200, cabeceras);
      res.end(cuerpoGz);
    } else {
      res.writeHead(200, cabeceras);
      res.end(cuerpo);
    }
    return;
  }

  /* El icono de la pestaña. Va dentro de la propia pagina como data-uri, pero
     algunos navegadores lo piden aparte en /favicon.ico, y si no existe muestran
     el globo terraqueo por defecto. Aqui se extrae el mismo logo y se entrega. */
  if (ruta === '/favicon.ico' || ruta === '/favicon.png') {
    if (LOGO) {
      res.writeHead(200, {
        'Content-Type': 'image/png',
        'Cache-Control': 'public, max-age=86400',
        'Content-Length': LOGO.length
      });
      res.end(LOGO);
    } else {
      res.writeHead(404); res.end();
    }
    return;
  }

  if (ruta === '/robots.txt') {
    res.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('User-agent: *\nDisallow: /\n');
    return;
  }

  res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end('Aqui no hay nada. Las fichas estan en /');
});

try {
  cargar();
} catch (e) {
  console.error('No se pudo leer "Fichas Tecnicas Caujaral.html":', e.message);
  process.exit(1);
}

servidor.listen(PUERTO, () => {
  console.log('Fichas tecnicas escuchando en el puerto ' + PUERTO);
  console.log('  recetas: ' + RECETAS + '   corte de datos: ' + CORTE);
  if (CLAVE) {
    console.log('  acceso con clave, usuario "' + USUARIO + '"');
  } else if (ABIERTA_A_PROPOSITO) {
    console.log('  ABIERTA A PROPOSITO (FICHAS_PUBLICA=si). Cualquiera con el link entra.');
  } else {
    console.log('  CERRADA: no hay FICHAS_PASSWORD, no se entrega la pagina.');
    console.log('  Para abrirla, crear FICHAS_PASSWORD en Railway.');
  }
});
