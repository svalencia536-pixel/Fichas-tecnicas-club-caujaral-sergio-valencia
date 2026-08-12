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
const ABIERTA_A_PROPOSITO = String(process.env.FICHAS_PUBLICA || '').toLowerCase() === 'si';

const ARCHIVO = path.join(__dirname, 'Fichas Tecnicas Caujaral.html');

/* La pagina se lee una vez al arrancar: cada despliegue trae su propia version. */
let PAGINA = '';
let PAGINA_GZ = null;
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
}

function autorizado(req) {
  // sin clave configurada no se entrega nada, salvo que se pida abrirla a proposito
  if (!CLAVE) return ABIERTA_A_PROPOSITO;
  const cabecera = req.headers.authorization || '';
  if (!cabecera.startsWith('Basic ')) return false;
  let plano;
  try {
    plano = Buffer.from(cabecera.slice(6), 'base64').toString('utf8');
  } catch (e) {
    return false;
  }
  const corte = plano.indexOf(':');
  if (corte < 0) return false;
  return plano.slice(0, corte) === USUARIO && plano.slice(corte + 1) === CLAVE;
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

  if (!autorizado(req)) { pedirClave(res); return; }

  if (ruta === '/' || ruta === '/index.html' || ruta === '/fichas') {
    const aceptaGzip = /\bgzip\b/.test(req.headers['accept-encoding'] || '');
    const cabeceras = {
      'Content-Type': 'text/html; charset=utf-8',
      'Cache-Control': 'no-cache',
      'X-Content-Type-Options': 'nosniff',
      // que no lo indexen los buscadores
      'X-Robots-Tag': 'noindex, nofollow'
    };
    if (aceptaGzip && PAGINA_GZ) {
      cabeceras['Content-Encoding'] = 'gzip';
      cabeceras['Content-Length'] = PAGINA_GZ.length;
      res.writeHead(200, cabeceras);
      res.end(PAGINA_GZ);
    } else {
      res.writeHead(200, cabeceras);
      res.end(PAGINA);
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
