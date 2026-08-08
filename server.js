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

/* Clave de acceso.
   Si se define FICHAS_PASSWORD en Railway, la pagina pide usuario y clave.
   Si no se define, queda abierta a cualquiera con el link — y conviene tenerlo
   presente, porque el archivo lleva adentro las recetas y los costos. */
const USUARIO = process.env.FICHAS_USER || 'caujaral';
const CLAVE = process.env.FICHAS_PASSWORD || '';

const ARCHIVO = path.join(__dirname, 'Fichas Tecnicas Caujaral.html');

/* La pagina se lee una vez al arrancar: cada despliegue trae su propia version. */
let PAGINA = '';
let PAGINA_GZ = null;
let ARRANQUE = new Date();
let RECETAS = 0;
let CORTE = '';

function cargar() {
  PAGINA = fs.readFileSync(ARCHIVO, 'utf8');
  PAGINA_GZ = zlib.gzipSync(Buffer.from(PAGINA, 'utf8'), { level: 9 });

  // Se leen las cifras del propio HTML para poder reportarlas en /salud
  const corte = PAGINA.match(/"generado":"([^"]+)"/);
  CORTE = corte ? corte[1] : '';
  const recetas = PAGINA.match(/"c":"/g);
  RECETAS = recetas ? recetas.length : 0;
}

function autorizado(req) {
  if (!CLAVE) return true;
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
  res.writeHead(401, {
    'WWW-Authenticate': 'Basic realm="Fichas tecnicas Caujaral", charset="UTF-8"',
    'Content-Type': 'text/plain; charset=utf-8'
  });
  res.end('Necesitas usuario y clave para ver las fichas tecnicas.');
}

const servidor = http.createServer((req, res) => {
  const ruta = (req.url || '/').split('?')[0];

  if (ruta === '/salud') {
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({
      estado: 'ok',
      recetas: RECETAS,
      corteDeDatos: CORTE,
      arranque: ARRANQUE.toISOString(),
      protegidoConClave: Boolean(CLAVE)
    }, null, 2));
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
  } else {
    console.log('  ATENCION: sin clave. Cualquiera con el link ve las recetas y los costos.');
    console.log('  Para cerrarlo, crear la variable FICHAS_PASSWORD en Railway.');
  }
});
