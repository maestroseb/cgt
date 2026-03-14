// ============================================================
//  CONCURSO DE TRASLADOS — Automatización Google Sheets
//  v3.0 — Menú, importación CSV, Global, Web de búsqueda
// ============================================================

// ===================== CONFIGURACIÓN =====================

// Pestañas del sistema (no son especialidades, se ignoran al fusionar)
const TABS_SISTEMA = ['Global', 'Búsqueda', 'Centros', 'Config', 'Instrucciones', 'Especialidades'];

// Columnas de la pestaña Centros (ajustar si el fichero oficial cambia)
const CENTROS = {
  COL_CODIGO: 1,     // A
  COL_NOMBRE: 5,     // E
  COL_LOCALIDAD: 7,  // G
  COL_PROVINCIA: 11  // K
};

// ===================== MENÚ =====================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔄 Concurso de Traslados')
    .addItem('📥 Importar CSV de especialidad', 'mostrarDialogoImportar')
    .addSeparator()
    .addItem('🔄 Regenerar pestaña Global', 'generarGlobal')
    .addSeparator()
    .addItem('🌐 Obtener URL de búsqueda web', 'mostrarURLWeb')
    .addSeparator()
    .addItem('ℹ️ Ayuda', 'mostrarAyuda')
    .addToUi();
}

// ===================== WEB APP =====================

function doGet(e) {
  // Si piden datos por JSON (para fetch desde el propio HTML)
  if (e && e.parameter && e.parameter.action === 'getData') {
    const datos = obtenerDatosGlobal();
    return ContentService.createTextOutput(JSON.stringify(datos))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Si no, servir la página HTML
  return HtmlService.createHtmlOutput(getWebHTML())
    .setTitle('Búsqueda CdT')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function mostrarURLWeb() {
  const url = ScriptApp.getService().getUrl();
  if (url) {
    SpreadsheetApp.getUi().alert(
      '🌐 URL de la web de búsqueda:\n\n' + url +
      '\n\nComparte esta URL con quien necesite consultar los datos.\n' +
      'Recuerda desplegar como webapp: Implementar → Nueva implementación → App web.'
    );
  } else {
    SpreadsheetApp.getUi().alert(
      '⚠️ Primero debes desplegar como webapp:\n\n' +
      '1. Implementar → Nueva implementación\n' +
      '2. Tipo: Aplicación web\n' +
      '3. Ejecutar como: Yo\n' +
      '4. Acceso: Cualquier persona\n' +
      '5. Implementar y copia la URL'
    );
  }
}

// ===================== DATOS PARA LA WEB =====================

function obtenerDatosGlobal() {
  // Intentar cache primero (6 horas)
  const cache = CacheService.getScriptCache();
  const cached = cache.get('cdt_datos');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) { /* cache corrupta, recalcular */ }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const global = ss.getSheetByName('Global');
  if (!global || global.getLastRow() < 2) return { datos: [], especialidades: {}, centros: [] };

  const datos = global.getDataRange().getValues();
  const filas = [];

  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    if (!f[0] && !f[1]) continue;
    const espVal = f.length > 12 ? String(f[12] || '').trim() : '';
    filas.push({
      nif: String(f[0] || '').trim(),
      nombre: String(f[1] || '').trim(),
      centroO: String(f[2] || '').trim(),
      puestoO: String(f[3] || '').trim(),
      centroAd: String(f[4] || '').trim(),
      puestoAd: String(f[5] || '').trim(),
      peticion: String(f[6] || '').trim(),
      puntos: String(f[7] || '').trim(),
      provO: String(f[8] || '').trim(),
      nombreO: String(f[9] || '').trim(),
      provD: String(f[10] || '').trim(),
      nombreD: String(f[11] || '').trim(),
      esp: espVal,
      _idx: i  // posición original para respetar escalafón
    });
  }

  // Especialidades (col A=código, col C=resumen)
  const espMap = {};
  const espSheet = ss.getSheetByName('Especialidades');
  if (espSheet && espSheet.getLastRow() > 1) {
    const espData = espSheet.getDataRange().getValues();
    for (let i = 1; i < espData.length; i++) {
      const codigo = String(espData[i][0] || '').trim();
      const resumen = String(espData[i][2] || '').trim();
      if (codigo) espMap[codigo] = resumen;
    }
  }

  // Centros para autocompletado (solo código, nombre, localidad, provincia)
  const centrosList = [];
  const centrosSheet = ss.getSheetByName('Centros');
  if (centrosSheet && centrosSheet.getLastRow() > 1) {
    const cd = centrosSheet.getDataRange().getValues();
    for (let i = 1; i < cd.length; i++) {
      const codigo = String(cd[i][CENTROS.COL_CODIGO - 1] || '').trim();
      const nombre = String(cd[i][CENTROS.COL_NOMBRE - 1] || '').trim();
      const localidad = String(cd[i][CENTROS.COL_LOCALIDAD - 1] || '').trim();
      const provincia = String(cd[i][CENTROS.COL_PROVINCIA - 1] || '').trim();
      if (codigo) centrosList.push({ codigo, nombre, localidad, provincia });
    }
  }

  const result = { datos: filas, especialidades: espMap, centros: centrosList };

  // Guardar en cache (máx 100KB por chunk, 6h = 21600s)
  try {
    const json = JSON.stringify(result);
    if (json.length < 100000) {
      cache.put('cdt_datos', json, 21600);
    }
  } catch(e) { /* si no cabe en cache, no pasa nada */ }

  return result;
}

// Llamar tras regenerar Global para refrescar cache
function invalidarCache() {
  try { CacheService.getScriptCache().remove('cdt_datos'); } catch(e) {}
}

// ===================== HTML WEB DE BÚSQUEDA =====================

function getWebHTML() {
  // Incluir la URL del script para fallback en Safari
  const scriptUrl = ScriptApp.getService().getUrl() || '';

  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="mobile-web-app-capable" content="yes">
<title>Búsqueda CdT 2025/2026</title>
<style>
  :root { --verde: #1a5c2e; --verde-claro: #d9e2d0; --fondo: #f5f7fa; --borde: #ddd; }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  html { width: 100%; overflow-x: hidden; -webkit-text-size-adjust: 100%; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: var(--fondo); color: #333; width: 100%; overflow-x: hidden; }

  header { background: var(--verde); color: #fff; padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; }
  header h1 { font-size: 1.1rem; }
  header p { font-size: 0.8rem; opacity: 0.7; }

  .search-area {
    background: #fff; border-bottom: 1px solid var(--borde); padding: 12px 24px;
    position: sticky; top: 0; z-index: 100; box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    display: flex; gap: 10px; align-items: center; flex-wrap: wrap;
  }
  .search-input-wrap { position: relative; flex: 1; min-width: 200px; }
  .search-input-wrap input {
    width: 100%; padding: 9px 12px 9px 32px; border: 2px solid var(--borde);
    border-radius: 8px; font-size: 13px; transition: border-color 0.2s;
  }
  .search-input-wrap input:focus { border-color: var(--verde); outline: none; }
  .search-input-wrap .icon { position: absolute; left: 9px; top: 50%; transform: translateY(-50%); font-size: 14px; opacity: 0.35; }
  .prov-badge-inline { padding: 4px 10px; border-radius: 6px; font-weight: 700; font-size: 12px; color: #fff; white-space: nowrap; }

  .autocomplete-list {
    position: absolute; top: 100%; left: 0; right: 0; background: #fff;
    border: 1px solid var(--borde); border-top: none; border-radius: 0 0 8px 8px;
    max-height: 250px; overflow-y: auto; z-index: 200; display: none;
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
  }
  .autocomplete-list .item { padding: 7px 12px; cursor: pointer; font-size: 12px; border-bottom: 1px solid #f5f5f5; }
  .autocomplete-list .item:hover, .autocomplete-list .item.active { background: var(--verde-claro); }
  .autocomplete-list .item .code { color: #999; font-size: 11px; }

  select { padding: 8px 10px; border: 2px solid var(--borde); border-radius: 8px; font-size: 12px; background: #fff; }
  .btn { padding: 8px 16px; border: none; border-radius: 8px; font-size: 12px; font-weight: 600; cursor: pointer; }
  .btn-buscar { background: var(--verde); color: #fff; }
  .btn-buscar:hover { background: #155a24; }
  .btn-limpiar { background: #eee; color: #555; }

  .results-info { padding: 8px 24px; font-size: 12px; color: #666; background: #fff; border-bottom: 1px solid var(--borde); }

  .table-wrap { overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 11px; }
  th { padding: 7px 8px; text-align: left; font-weight: 600; white-space: nowrap; position: sticky; top: 0; color: #fff; }
  th.th-base { background: var(--verde); }
  th.th-origen { background: #2E86AB; }
  th.th-destino { background: #C73E1D; }
  td { padding: 5px 8px; border-bottom: 1px solid #f0f0f0; white-space: nowrap; }
  tr:hover td { background: #f8faf6; }

  /* Separadores verticales entre bloques */
  .sep-origen { border-left: 3px solid #2E86AB; }
  .sep-destino { border-left: 3px solid #C73E1D; }
  .sep-extra { border-left: 3px solid var(--verde); }

  .esp-badge { display: inline-block; padding: 2px 7px; border-radius: 4px; font-size: 9px; font-weight: 700; color: #fff; letter-spacing: 0.3px; }
  .prov { display: inline-block; padding: 1px 6px; border-radius: 3px; font-size: 10px; font-weight: 700; color: #fff; }
  .dir { display: inline-block; padding: 1px 5px; border-radius: 3px; font-size: 9px; font-weight: 700; margin-left: 3px; }
  .dir-sale { background: #f8d7da; color: #842029; }
  .dir-llega { background: #d1e7dd; color: #0f5132; }
  .centro-match { background: #fff3cd; }
  .puntos { font-weight: 700; font-variant-numeric: tabular-nums; }

  .loading { text-align: center; padding: 60px; color: #aaa; }
  .spin { display: inline-block; width: 20px; height: 20px; border: 3px solid #eee; border-top-color: var(--verde); border-radius: 50%; animation: spin 0.7s linear infinite; }
  @keyframes spin { to { transform: rotate(360deg); } }
  .placeholder { text-align: center; padding: 50px; color: #bbb; font-size: 14px; }

  .error-msg { text-align: center; padding: 40px 20px; color: #C73E1D; }
  .error-msg button { margin-top: 12px; padding: 8px 20px; border: 2px solid #C73E1D; background: #fff; color: #C73E1D; border-radius: 8px; font-weight: 600; cursor: pointer; }

  /* Mobile cards */
  .cards { display: none; }
  .card {
    background: #fff; border-radius: 10px; margin: 8px 12px; padding: 12px 14px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08); border-left: 4px solid var(--verde);
  }
  .card.card-sale { border-left-color: #dc3545; }
  .card.card-llega { border-left-color: #28a745; }
  .card-top { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }
  .card-nombre { font-weight: 700; font-size: 13px; }
  .card-nif { font-size: 11px; color: #888; }
  .card-puntos { font-weight: 700; font-size: 15px; color: var(--verde); }
  .card-row { display: flex; gap: 12px; font-size: 11px; margin-top: 4px; }
  .card-bloque { flex: 1; min-width: 0; }
  .card-label { font-size: 9px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 2px; }
  .card-label-origen { color: #2E86AB; }
  .card-label-destino { color: #C73E1D; }
  .card-centro { font-size: 11px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .card-meta { display: flex; gap: 4px; flex-wrap: wrap; align-items: center; margin-top: 2px; }
  .card-peticion { font-size: 10px; color: #888; margin-top: 4px; }

  @media (max-width: 900px) {
    header h1 { font-size: 0.95rem; }
    .search-area { flex-direction: column; align-items: stretch; padding: 10px 12px; gap: 8px; }
    .search-input-wrap { min-width: 100%; }
    .search-input-wrap input { font-size: 14px; padding: 10px 12px 10px 32px; }
    select { font-size: 13px; padding: 9px 10px; }
    .btn { padding: 10px 16px; font-size: 13px; }
    .results-info { padding: 8px 12px; }

    /* Filtros en grid 2x2 */
    .filtros-mobile { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; width: 100%; }
    .filtros-mobile select { width: 100%; }
    .filtros-mobile .btn-row { grid-column: 1 / -1; display: flex; gap: 6px; }
    .filtros-mobile .btn-row .btn { flex: 1; }

    /* Ocultar tabla, mostrar tarjetas */
    .table-wrap table { display: none; }
    .cards { display: block; }
  }
  @media (min-width: 901px) {
    .filtros-desktop { display: contents; }
    .filtros-mobile { display: contents; }
    .filtros-mobile .btn-row { display: contents; }
  }

  /* Forzar layout móvil via JS cuando el user-agent es móvil pero el iframe reporta ancho grande */
  body.force-mobile header h1 { font-size: 0.95rem; }
  body.force-mobile .search-area { flex-direction: column; align-items: stretch; padding: 10px 12px; gap: 8px; }
  body.force-mobile .search-input-wrap { min-width: 100%; }
  body.force-mobile .search-input-wrap input { font-size: 14px; padding: 10px 12px 10px 32px; }
  body.force-mobile select { font-size: 13px; padding: 9px 10px; }
  body.force-mobile .btn { padding: 10px 16px; font-size: 13px; }
  body.force-mobile .results-info { padding: 8px 12px; }
  body.force-mobile .filtros-mobile { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; width: 100%; }
  body.force-mobile .filtros-mobile select { width: 100%; }
  body.force-mobile .filtros-mobile .btn-row { grid-column: 1 / -1; display: flex; gap: 6px; }
  body.force-mobile .filtros-mobile .btn-row .btn { flex: 1; }
  body.force-mobile .table-wrap table { display: none; }
  body.force-mobile .cards { display: block; }
</style>
</head>
<body>

<header>
  <div>
    <h1>Concurso de Traslados 2025/2026 — Cuerpo de Maestros</h1>
    <p>Resolución provisional</p>
  </div>
</header>

<div class="search-area">
  <div class="search-input-wrap">
    <span class="icon">🔍</span>
    <input type="text" id="txtBuscar" placeholder="Código, nombre del centro o localidad" autocomplete="off">
    <div class="autocomplete-list" id="autocomplete"></div>
  </div>
  <span id="provBadge"></span>
  <div class="filtros-mobile">
    <select id="filtroDireccion">
      <option value="ambos">↔ Entradas y salidas</option>
      <option value="sale">↗ Solo SALEN</option>
      <option value="llega">↙ Solo LLEGAN</option>
    </select>
    <select id="filtroEsp">
      <option value="">Todas las especialidades</option>
    </select>
    <select id="filtroProv">
      <option value="">Todas las provincias</option>
      <option value="AL">Almería</option>
      <option value="CA">Cádiz</option>
      <option value="CO">Córdoba</option>
      <option value="GR">Granada</option>
      <option value="HU">Huelva</option>
      <option value="JA">Jaén</option>
      <option value="MA">Málaga</option>
      <option value="SE">Sevilla</option>
    </select>
    <div class="btn-row">
      <button class="btn btn-buscar" onclick="filtrar()">Buscar</button>
      <button class="btn btn-limpiar" onclick="limpiar()">Limpiar</button>
    </div>
  </div>
</div>

<div class="results-info" id="info" style="display:none;"></div>
<div class="table-wrap" id="tablaWrap">
  <div class="loading" id="loading"><div class="spin"></div><br>Cargando datos...</div>
</div>

<script>
// ============ VIEWPORT FIX (iframe de Google Apps Script) ============
(function() {
  // Forzar viewport correcto — el iframe de Google a veces no lo propaga
  try {
    if (window.self !== window.top) {
      // Estamos en un iframe, intentar ajustar el parent frame
      var fr = window.frameElement;
      if (fr) {
        fr.style.width = '100%';
        fr.style.maxWidth = '100vw';
      }
    }
  } catch(e) { /* cross-origin, ignorar */ }

  // Forzar el viewport meta tag dinámicamente
  var vp = document.querySelector('meta[name="viewport"]');
  if (!vp) {
    vp = document.createElement('meta');
    vp.name = 'viewport';
    document.head.appendChild(vp);
  }
  vp.content = 'width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes';

  // Si el user-agent indica móvil pero el media query no detecta <900px (problema del iframe de GAS),
  // forzar el layout móvil con una clase CSS
  var uaMobile = /Android|iPhone|iPad|iPod|webOS|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
  if (uaMobile) {
    document.body.classList.add('force-mobile');
  }
})();

// ============ STATE ============
let DATOS = [];
let ESPECIALIDADES = {};
let CENTROS = [];
let centroSeleccionado = null;

const COLORES_ESP = [
  '#2E86AB','#A23B72','#F18F01','#C73E1D','#3B1F2B',
  '#44BBA4','#E94F37','#393E41','#8E7DBE','#5FAD56',
  '#3D5A80','#EE6C4D','#293241','#D4A373','#6D6875',
  '#B5838D','#6930C3','#5390D9','#FF6B6B','#4ECDC4',
  '#2D3436','#636E72','#D63031','#E17055','#00B894'
];
const mapaColorEsp = {};
let colorIdx = 0;
function colorEsp(esp) {
  if (!esp) return '#999';
  if (!mapaColorEsp[esp]) { mapaColorEsp[esp] = COLORES_ESP[colorIdx % COLORES_ESP.length]; colorIdx++; }
  return mapaColorEsp[esp];
}

const COLORES_PROV = {
  'AL':'#7B2D8E','CA':'#C49000','CO':'#00796B','GR':'#C62828',
  'HU':'#1565C0','JA':'#E65100','MA':'#AD1457','SE':'#2E7D32'
};
function colorProv(p) { return COLORES_PROV[p] || '#666'; }
function esc(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function norm(s) { return String(s||'').toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,''); }

// Buscar especialidad a partir del código de puesto
// ESPECIALIDADES = { "597035": "Mus", "597036": "PT", ... }
function lookupEsp(puestoCodigo) {
  if (!puestoCodigo) return '';
  var cod = String(puestoCodigo).trim();
  // Buscar coincidencia exacta
  if (ESPECIALIDADES[cod]) return ESPECIALIDADES[cod];
  // Buscar si el código contiene el puesto (ej: "00597035" contiene "597035")
  for (var k in ESPECIALIDADES) {
    if (cod.includes(k) || k.includes(cod)) return ESPECIALIDADES[k];
  }
  return '';
}

// ============ AUTOCOMPLETE ============
const input = document.getElementById('txtBuscar');
const acList = document.getElementById('autocomplete');
let acIdx = -1;
let acResultados = [];

input.addEventListener('input', function() {
  const q = norm(this.value);
  centroSeleccionado = null;
  document.getElementById('provBadge').innerHTML = '';
  if (q.length < 2) { acList.style.display = 'none'; return; }

  acResultados = CENTROS.filter(c =>
    norm(c.codigo).includes(q) || norm(c.nombre).includes(q) || norm(c.localidad).includes(q)
  ).slice(0, 20);

  if (acResultados.length === 0) { acList.style.display = 'none'; return; }

  acList.innerHTML = acResultados.map((c, i) =>
    '<div class="item" data-idx="' + i + '">' +
    '<b>' + esc(c.nombre) + '</b> ' +
    '<span class="code">' + esc(c.codigo) + '</span> — ' +
    esc(c.localidad) + ' (' + esc(c.provincia) + ')' +
    '</div>'
  ).join('');
  acList.style.display = 'block';
  acIdx = -1;

  // Bind click events AFTER rendering
  acList.querySelectorAll('.item').forEach(function(item, i) {
    item.addEventListener('mousedown', function(e) {
      e.preventDefault(); // Prevent blur from hiding the list
      seleccionarCentro(i);
    });
  });
});

input.addEventListener('keydown', function(e) {
  const items = acList.querySelectorAll('.item');
  if (!items.length || acList.style.display === 'none') {
    if (e.key === 'Enter') e.preventDefault();
    return;
  }
  if (e.key === 'ArrowDown') { e.preventDefault(); acIdx = Math.min(acIdx + 1, items.length - 1); updateAcActive(items); }
  else if (e.key === 'ArrowUp') { e.preventDefault(); acIdx = Math.max(acIdx - 1, 0); updateAcActive(items); }
  else if (e.key === 'Enter') { e.preventDefault(); seleccionarCentro(acIdx >= 0 ? acIdx : 0); }
  else if (e.key === 'Escape') { acList.style.display = 'none'; }
});

function updateAcActive(items) {
  items.forEach(function(it, i) { it.classList.toggle('active', i === acIdx); });
  if (items[acIdx]) items[acIdx].scrollIntoView({ block: 'nearest' });
}

function seleccionarCentro(idx) {
  const c = acResultados[idx];
  if (!c) return;
  centroSeleccionado = c;
  input.value = c.nombre + ' ' + c.codigo + ' (' + c.localidad + ')';
  acList.style.display = 'none';
  document.getElementById('provBadge').innerHTML =
    '<span class="prov-badge-inline" style="background:' + colorProv(c.provincia) + '">' + esc(c.provincia) + '</span>';
  filtrar();
}

document.addEventListener('click', function(e) {
  if (!e.target.closest('.search-input-wrap')) acList.style.display = 'none';
});

// ============ FILTRAR ============
function filtrar() {
  const q = norm(input.value);
  const dir = document.getElementById('filtroDireccion').value;
  const espFiltro = document.getElementById('filtroEsp').value;
  const provFiltro = document.getElementById('filtroProv').value;
  if (!q && !espFiltro && !provFiltro) { renderTabla([]); return; }

  const codigoCentro = centroSeleccionado ? centroSeleccionado.codigo : '';
  const resultados = [];

  for (const r of DATOS) {
    // Filtro de especialidad
    if (espFiltro) {
      var espO = lookupEsp(r.puestoO);
      var espD = lookupEsp(r.puestoAd);
      if (espO !== espFiltro && espD !== espFiltro) continue;
    }

    // Filtro de provincia (origen o destino)
    if (provFiltro) {
      if (r.provO !== provFiltro && r.provD !== provFiltro) continue;
    }

    let coincide = false;
    let direccion = '';

    if (codigoCentro) {
      const esOrigen = r.centroO === codigoCentro;
      const esDestino = r.centroAd === codigoCentro;
      if (esOrigen) { coincide = true; direccion = 'sale'; }
      if (esDestino) { coincide = true; direccion = (direccion === 'sale') ? 'ambos' : 'llega'; }
    } else if (q) {
      if (norm(r.centroO).includes(q) || norm(r.nombreO).includes(q) || norm(r.provO).includes(q)) {
        coincide = true; direccion = 'sale';
      }
      if (norm(r.centroAd).includes(q) || norm(r.nombreD).includes(q) || norm(r.provD).includes(q)) {
        coincide = true; direccion = (direccion === 'sale') ? 'ambos' : 'llega';
      }
      if (norm(r.nombre).includes(q) || norm(r.nif).includes(q)) coincide = true;
    } else if (espFiltro || provFiltro) {
      coincide = true;
    }

    if (!coincide) continue;
    if (dir === 'sale' && direccion === 'llega') continue;
    if (dir === 'llega' && direccion === 'sale') continue;

    resultados.push({ ...r, direccion: direccion });
  }

  resultados.sort(function(a, b) {
    var pa = parseFloat(String(a.puntos).replace(',','.')) || 0;
    var pb = parseFloat(String(b.puntos).replace(',','.')) || 0;
    // Si ambos tienen puntos, ordenar por puntos DESC
    // Si empatan o ambos son 0, respetar orden original del escalafón
    if (pb !== pa) return pb - pa;
    return (a._idx || 0) - (b._idx || 0);
  });

  renderTabla(resultados);
}

// ============ RENDER ============
// Detectar móvil por user-agent como respaldo (el iframe de GAS puede reportar ancho incorrecto)
var uaEsMobile = /Android|iPhone|iPad|iPod|webOS|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
var esMobile = window.matchMedia('(max-width: 900px)').matches || (uaEsMobile && window.innerWidth < 1024);
// Actualizar en caso de cambio de orientación
window.addEventListener('resize', function() {
  esMobile = window.matchMedia('(max-width: 900px)').matches || (uaEsMobile && window.innerWidth < 1024);
});
var MAX_RESULTADOS_MOBILE = 200;

function renderTabla(filas) {
  const info = document.getElementById('info');
  const wrap = document.getElementById('tablaWrap');

  if (filas.length === 0) {
    info.style.display = 'none';
    wrap.innerHTML = '<div class="placeholder">Busca un centro, localidad o nombre para ver resultados.</div>';
    return;
  }

  var totalFilas = filas.length;
  var limitado = false;
  if (esMobile && filas.length > MAX_RESULTADOS_MOBILE) {
    filas = filas.slice(0, MAX_RESULTADOS_MOBILE);
    limitado = true;
  }

  info.style.display = 'block';
  info.innerHTML = '<b>' + filas.length + '</b> resultado(s)' + (limitado ? ' <span style="color:#C73E1D">(mostrando ' + MAX_RESULTADOS_MOBILE + ' de ' + totalFilas + ', filtra más para ver todos)</span>' : '');

  const codigoCentro = centroSeleccionado ? centroSeleccionado.codigo : '';
  var html = '';

  if (!esMobile) {
    // Desktop: render table
    html = '<table><thead><tr>';
    html += '<th class="th-base"></th>';
    html += '<th class="th-base">N.I.F.</th>';
    html += '<th class="th-base">Apellidos y nombre</th>';
    html += '<th class="th-origen sep-origen">Prov.O</th>';
    html += '<th class="th-origen">Puesto.O</th>';
    html += '<th class="th-origen"></th>';
    html += '<th class="th-origen">Centro Origen</th>';
    html += '<th class="th-destino sep-destino">Prov.D</th>';
    html += '<th class="th-destino">Puesto.D</th>';
    html += '<th class="th-destino"></th>';
    html += '<th class="th-destino">Centro Destino</th>';
    html += '<th class="th-base sep-extra">Petición</th>';
    html += '<th class="th-base">Puntos</th>';
    html += '</tr></thead><tbody>';

    for (var i = 0; i < filas.length; i++) {
      var r = filas[i];
      var oMatch = codigoCentro && r.centroO === codigoCentro;
      var dMatch = codigoCentro && r.centroAd === codigoCentro;
      var espO = lookupEsp(r.puestoO);
      var espD = lookupEsp(r.puestoAd);

      html += '<tr>';
      html += '<td style="text-align:center">';
      if (r.direccion === 'sale') html += '<span class="dir dir-sale">↗ SALE</span>';
      else if (r.direccion === 'llega') html += '<span class="dir dir-llega">↙ LLEGA</span>';
      html += '</td>';
      html += '<td>' + esc(r.nif) + '</td>';
      html += '<td><b>' + esc(r.nombre) + '</b></td>';
      html += '<td class="sep-origen">' + (r.provO ? '<span class="prov" style="background:' + colorProv(r.provO) + '">' + esc(r.provO) + '</span>' : '') + '</td>';
      html += '<td>' + esc(r.puestoO) + '</td>';
      html += '<td>' + (espO ? '<span class="esp-badge" style="background:' + colorEsp(espO) + '">' + esc(espO) + '</span>' : '') + '</td>';
      html += '<td' + (oMatch ? ' class="centro-match"' : '') + '>' + esc(r.nombreO || r.centroO) + '</td>';
      html += '<td class="sep-destino">' + (r.provD ? '<span class="prov" style="background:' + colorProv(r.provD) + '">' + esc(r.provD) + '</span>' : '') + '</td>';
      html += '<td>' + esc(r.puestoAd) + '</td>';
      html += '<td>' + (espD ? '<span class="esp-badge" style="background:' + colorEsp(espD) + '">' + esc(espD) + '</span>' : '') + '</td>';
      html += '<td' + (dMatch ? ' class="centro-match"' : '') + '>' + esc(r.nombreD || r.centroAd) + '</td>';
      html += '<td class="sep-extra" style="text-align:center">' + esc(r.peticion) + '</td>';
      html += '<td class="puntos">' + esc(r.puntos) + '</td>';
      html += '</tr>';
    }
    html += '</tbody></table>';
  } else {
    // Mobile: render cards only
    html = '<div class="cards" style="display:block">';
    for (var j = 0; j < filas.length; j++) {
      var m = filas[j];
      var mEspO = lookupEsp(m.puestoO);
      var mEspD = lookupEsp(m.puestoAd);
      var cardClass = 'card';
      if (m.direccion === 'sale') cardClass += ' card-sale';
      else if (m.direccion === 'llega') cardClass += ' card-llega';

      html += '<div class="' + cardClass + '">';
      html += '<div class="card-top">';
      html += '<div><span class="card-nombre">' + esc(m.nombre) + '</span>';
      if (m.direccion === 'sale') html += ' <span class="dir dir-sale">↗ SALE</span>';
      else if (m.direccion === 'llega') html += ' <span class="dir dir-llega">↙ LLEGA</span>';
      html += '<br><span class="card-nif">' + esc(m.nif) + '</span></div>';
      html += '<span class="card-puntos">' + esc(m.puntos) + '</span>';
      html += '</div>';

      html += '<div class="card-row">';
      html += '<div class="card-bloque">';
      html += '<div class="card-label card-label-origen">Origen</div>';
      html += '<div class="card-meta">';
      if (m.provO) html += '<span class="prov" style="background:' + colorProv(m.provO) + '">' + esc(m.provO) + '</span>';
      if (mEspO) html += '<span class="esp-badge" style="background:' + colorEsp(mEspO) + '">' + esc(mEspO) + '</span>';
      html += '</div>';
      html += '<div class="card-centro">' + esc(m.nombreO || m.centroO) + '</div>';
      html += '</div>';

      if (m.centroAd || m.nombreD) {
        html += '<div class="card-bloque">';
        html += '<div class="card-label card-label-destino">Destino</div>';
        html += '<div class="card-meta">';
        if (m.provD) html += '<span class="prov" style="background:' + colorProv(m.provD) + '">' + esc(m.provD) + '</span>';
        if (mEspD) html += '<span class="esp-badge" style="background:' + colorEsp(mEspD) + '">' + esc(mEspD) + '</span>';
        html += '</div>';
        html += '<div class="card-centro">' + esc(m.nombreD || m.centroAd) + '</div>';
        if (m.peticion) html += '<div class="card-peticion">Pet. ' + esc(m.peticion) + '</div>';
        html += '</div>';
      }

      html += '</div>'; // card-row
      html += '</div>'; // card
    }
    html += '</div>'; // cards
  }

  wrap.innerHTML = html;
}

function limpiar() {
  input.value = '';
  centroSeleccionado = null;
  document.getElementById('provBadge').innerHTML = '';
  document.getElementById('filtroEsp').value = '';
  document.getElementById('filtroProv').value = '';
  document.getElementById('filtroDireccion').value = 'ambos';
  document.getElementById('info').style.display = 'none';
  document.getElementById('tablaWrap').innerHTML = '<div class="placeholder">Busca un centro, localidad o nombre para ver resultados.</div>';
}

// ============ INIT (carga asíncrona con google.script.run) ============
function inicializarDatos(result) {
  DATOS = result.datos || [];
  for (var idx = 0; idx < DATOS.length; idx++) { DATOS[idx]._idx = idx; }
  ESPECIALIDADES = result.especialidades || {};
  CENTROS = result.centros || [];

  var espSelect = document.getElementById('filtroEsp');
  var espNombres = [];
  var espDone = {};
  for (var k in ESPECIALIDADES) {
    var v = ESPECIALIDADES[k];
    if (v && !espDone[v]) { espDone[v] = true; espNombres.push(v); }
  }
  espNombres.sort();
  for (var j = 0; j < espNombres.length; j++) {
    var opt = document.createElement('option');
    opt.value = espNombres[j];
    opt.textContent = espNombres[j];
    espSelect.appendChild(opt);
  }

  document.getElementById('loading').style.display = 'none';
  document.getElementById('tablaWrap').innerHTML = '<div class="placeholder">Busca un centro, localidad o nombre para ver resultados.</div>';
}

function errorCarga(err) {
  document.getElementById('loading').innerHTML =
    '<div class="error-msg">' +
    '<p>No se pudieron cargar los datos.</p>' +
    '<p style="font-size:12px;margin-top:8px">' + (err && err.message ? err.message : 'Error de conexión') + '</p>' +
    '<button onclick="cargarDatos()">Reintentar</button>' +
    '</div>';
}

var SCRIPT_URL = '${scriptUrl}';

function cargarDatosFetch() {
  if (!SCRIPT_URL) { errorCarga({ message: 'URL del script no disponible' }); return; }
  fetch(SCRIPT_URL + '?action=getData')
    .then(function(r) { return r.json(); })
    .then(inicializarDatos)
    .catch(function(e) { errorCarga({ message: 'Error de red: ' + e.message }); });
}

function cargarDatos() {
  document.getElementById('loading').innerHTML = '<div class="spin"></div><br>Cargando datos...';
  document.getElementById('loading').style.display = 'block';

  // Intentar google.script.run primero, si no existe (Safari bloquea cookies) usar fetch
  if (typeof google !== 'undefined' && google.script && google.script.run) {
    google.script.run
      .withSuccessHandler(inicializarDatos)
      .withFailureHandler(function(err) {
        // Si google.script.run falla, intentar fetch como respaldo
        console.warn('google.script.run falló, intentando fetch...', err);
        cargarDatosFetch();
      })
      .obtenerDatosGlobal();
  } else {
    // Safari u otro navegador donde google.script.run no está disponible
    cargarDatosFetch();
  }
}

cargarDatos();
</script>
</body>
</html>`;
}

// ===================== 1. IMPORTAR CSV =====================

function mostrarDialogoImportar() {
  const html = HtmlService.createHtmlOutput(getImportarHTML())
    .setWidth(480).setHeight(420).setTitle('Importar CSV de especialidad');
  SpreadsheetApp.getUi().showModalDialog(html, 'Importar CSV de especialidad');
}

function getImportarHTML() {
  const existentes = obtenerTabsEspecialidad().map(s => s.getName());
  return `<!DOCTYPE html><html><head><style>
  body{font-family:-apple-system,sans-serif;padding:16px;color:#333}
  h3{margin:0 0 16px;color:#1a5c2e}
  label{display:block;font-weight:600;margin:12px 0 4px}
  input[type="file"]{margin:8px 0}
  input[type="text"],select{width:100%;padding:8px;border:1px solid #ccc;border-radius:6px;font-size:14px;box-sizing:border-box}
  .btn{padding:10px 20px;border:none;border-radius:6px;font-size:14px;font-weight:bold;cursor:pointer;margin-top:16px}
  .btn-ok{background:#28a745;color:#fff}.btn-cancel{background:#eee;color:#333;margin-left:8px}
  .hint{font-size:12px;color:#888;margin-top:4px}
  .or{text-align:center;color:#aaa;margin:4px 0;font-size:12px}
  #status{margin-top:12px;font-size:13px}
</style></head><body>
  <h3>📥 Importar CSV</h3>
  <label>Archivo CSV:</label>
  <input type="file" id="csvFile" accept=".csv,.txt">
  <p class="hint">CSV del bookmarklet/Tampermonkey (separador: punto y coma)</p>
  <label>Nombre de la pestaña:</label>
  <select id="tabExistente">
    <option value="">— Seleccionar existente —</option>
    ${existentes.map(n => '<option value="' + n + '">' + n + '</option>').join('')}
  </select>
  <p class="or">— o crear nueva —</p>
  <input type="text" id="tabNueva" placeholder="Ej: Pri, Mus, Ing, PT...">
  <div>
    <button class="btn btn-ok" onclick="importar()">Importar</button>
    <button class="btn btn-cancel" onclick="google.script.host.close()">Cancelar</button>
  </div>
  <div id="status"></div>
  <script>
  function importar(){
    var tab=document.getElementById('tabNueva').value.trim()||document.getElementById('tabExistente').value;
    var st=document.getElementById('status');
    if(!tab){st.textContent='⚠️ Indica la pestaña.';return}
    var fi=document.getElementById('csvFile');
    if(!fi.files.length){st.textContent='⚠️ Selecciona un CSV.';return}
    st.textContent='⏳ Leyendo...';
    var r=new FileReader();
    r.onload=function(e){
      st.textContent='⏳ Importando...';
      google.script.run
        .withSuccessHandler(function(m){st.innerHTML='✅ '+m;setTimeout(function(){google.script.host.close()},2000)})
        .withFailureHandler(function(err){st.textContent='❌ '+err.message})
        .procesarCSVImportado(e.target.result,tab);
    };
    r.readAsText(fi.files[0],'UTF-8');
  }
  </script>
</body></html>`;
}

function procesarCSVImportado(contenidoCSV, nombreTab) {
  if (contenidoCSV.charCodeAt(0) === 0xFEFF) contenidoCSV = contenidoCSV.substring(1);
  const filas = parsearCSV(contenidoCSV);
  if (filas.length < 2) throw new Error('El CSV no contiene datos suficientes.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(nombreTab);
  if (hoja) hoja.clearContents();
  else hoja = ss.insertSheet(nombreTab);

  const maxCols = Math.max(...filas.map(f => f.length));
  const datos = filas.map(f => { while (f.length < maxCols) f.push(''); return f; });
  hoja.getRange(1, 1, datos.length, maxCols).setValues(datos);

  const cab = hoja.getRange(1, 1, 1, maxCols);
  cab.setFontWeight('bold').setBackground('#d9e2d0');
  for (let i = 1; i <= maxCols; i++) hoja.autoResizeColumn(i);

  return `${filas.length - 1} registros importados en "${nombreTab}"`;
}

function parsearCSV(texto) {
  const lineas = texto.split(/\r?\n/);
  const resultado = [];
  for (const linea of lineas) {
    if (!linea.trim()) continue;
    const campos = [];
    let actual = '', enComillas = false;
    for (let i = 0; i < linea.length; i++) {
      const c = linea[i];
      if (enComillas) {
        if (c === '"') { if (i + 1 < linea.length && linea[i + 1] === '"') { actual += '"'; i++; } else enComillas = false; }
        else actual += c;
      } else {
        if (c === '"') enComillas = true;
        else if (c === ';') { campos.push(actual.trim()); actual = ''; }
        else actual += c;
      }
    }
    campos.push(actual.trim());
    resultado.push(campos);
  }
  return resultado;
}

// ===================== 2. GENERAR GLOBAL =====================

function generarGlobal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const tabsEsp = obtenerTabsEspecialidad();

  if (tabsEsp.length === 0) {
    ui.alert('No se encontraron pestañas de especialidad.\nImporta al menos un CSV primero.');
    return;
  }

  const nombres = tabsEsp.map(s => s.getName());
  const resp = ui.alert('Regenerar Global',
    `Se fusionarán ${nombres.length} pestañas:\n\n• ${nombres.join('\n• ')}\n\n¿Continuar?`,
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  let global = ss.getSheetByName('Global');
  if (global) { global.clearContents(); global.clearFormats(); }
  else global = ss.insertSheet('Global', 0);

  // Recoger datos de todas las pestañas
  // CSV tiene: N., N.I.F., Apellidos, CentroO, PuestoO, CentroAd, PuestoAd, PetAd, Puntos
  // Columnas:  A    B        C         D        E        F          G        H       I
  // Queremos en Global: NIF, Nombre, CentroO, PuestoO, CentroAd, PuestoAd, PetAd, Puntos, [fórmulas I-L], Especialidad(M)
  // Tomamos desde col B (índice 2) del CSV = NIF hasta col I (índice 9) = Puntos
  const todosDatos = [];

  for (const tab of tabsEsp) {
    const lastRow = tab.getLastRow();
    if (lastRow < 2) continue;
    const lastCol = tab.getLastColumn();
    const colInicio = 2; // B = NIF
    const colFin = Math.min(lastCol, 9); // I = Puntos
    const numCols = colFin - colInicio + 1;
    if (numCols < 1) continue;

    const datos = tab.getRange(2, colInicio, lastRow - 1, numCols).getValues();
    const espNombre = tab.getName();

    for (const fila of datos) {
      if (!fila[0] || String(fila[0]).trim() === '') continue;
      // fila = [NIF, Nombre, CentroO, PuestoO, CentroAd, PuestoAd, PetAd, Puntos]
      const filaCompleta = [...fila];
      while (filaCompleta.length < 8) filaCompleta.push('');
      todosDatos.push(filaCompleta);
    }
  }

  if (todosDatos.length === 0) {
    ui.alert('No se encontraron datos en las pestañas de especialidad.');
    return;
  }

  // Deduplicar por NIF parcial + nombre
  const porClave = new Map();
  let idx = 0;
  for (const fila of todosDatos) {
    const nif = String(fila[0]).trim().toLowerCase();
    const nombre = String(fila[1]).trim().toLowerCase();
    if (!nif && !nombre) continue;
    const clave = nif + '|' + nombre;
    const puntos = parseFloat(String(fila[7]).replace(',', '.')) || 0;
    if (!porClave.has(clave) || puntos > porClave.get(clave).puntos) {
      porClave.set(clave, { fila, puntos, idx });
    }
    idx++;
  }

  // Ordenar: por puntos DESC, si empatan respetar orden original (escalafón)
  const datosUnicos = Array.from(porClave.values())
    .sort((a, b) => b.puntos !== a.puntos ? b.puntos - a.puntos : a.idx - b.idx)
    .map(item => item.fila);

  // Escribir datos en A:H (8 columnas de datos)
  const filasParaEscribir = datosUnicos.map(f => f.slice(0, 8));

  const cabeceraAH = ['N.I.F.', 'Apellidos y nombre', 'Centro O.', 'Puesto. Or.', 'Centro Ad.', 'Puesto. Ad', 'Pet. Ad', 'Puntos'];
  global.getRange(1, 1, 1, 8).setValues([cabeceraAH]);

  if (filasParaEscribir.length > 0) {
    global.getRange(2, 1, filasParaEscribir.length, 8).setValues(filasParaEscribir);
  }

  // Columna M (13): Especialidad — fórmula que busca PuestoO (col D) en Especialidades
  const hayEspecialidades = !!ss.getSheetByName('Especialidades');
  if (hayEspecialidades) {
    global.getRange(1, 13).setFormula(
      `={"Especialidad";ARRAYFORMULA(SI.ERROR(BUSCARV(D2:D;Especialidades!$A:$C;3;FALSO)))}`
    );
  } else {
    global.getRange(1, 13).setValue('Especialidad');
  }

  // Fórmulas I-L con cabecera incluida (estilo del usuario)
  // I = ProvO (busca por CentroO que está en col C)
  // J = Nombre centro ORIGEN
  // K = ProvD (busca por CentroAd que está en col E)
  // L = Nombre centro DESTINO
  const hayCentros = !!ss.getSheetByName('Centros');
  if (hayCentros) {
    global.getRange(1, 9).setFormula(
      `={"Prov.O";ARRAYFORMULA(SI.ERROR(BUSCARV(C2:C;Centros!A:L;${CENTROS.COL_PROVINCIA};FALSO)))}`
    );
    global.getRange(1, 10).setFormula(
      `={"Nombre centro ORIGEN";ARRAYFORMULA(SI.ERROR(BUSCARV(C2:C;Centros!A:L;${CENTROS.COL_NOMBRE};FALSO)&" ("&BUSCARV(C2:C;Centros!A:L;${CENTROS.COL_LOCALIDAD};FALSO)&")"))}`
    );
    global.getRange(1, 11).setFormula(
      `={"Prov.D";ARRAYFORMULA(SI.ERROR(BUSCARV(E2:E;Centros!A:L;${CENTROS.COL_PROVINCIA};FALSO)))}`
    );
    global.getRange(1, 12).setFormula(
      `={"Nombre centro DESTINO";ARRAYFORMULA(SI.ERROR(BUSCARV(E2:E;Centros!A:L;${CENTROS.COL_NOMBRE};FALSO)&" ("&BUSCARV(E2:E;Centros!A:L;${CENTROS.COL_LOCALIDAD};FALSO)&")"))}`
    );
  } else {
    // Sin centros: solo poner cabeceras vacías
    global.getRange(1, 9, 1, 4).setValues([['Prov.O', 'Nombre centro ORIGEN', 'Prov.D', 'Nombre centro DESTINO']]);
    ui.alert('⚠️ No se encontró la pestaña "Centros". Las fórmulas de búsqueda no se han creado.\nImporta el fichero de centros y regenera.');
  }

  // Formato cabecera
  const cabRange = global.getRange(1, 1, 1, 13);
  cabRange.setFontWeight('bold');
  cabRange.setBackground('#1a5c2e');
  cabRange.setFontColor('#ffffff');
  cabRange.setHorizontalAlignment('center');

  global.setFrozenRows(1);
  for (let i = 1; i <= 13; i++) global.autoResizeColumn(i);

  invalidarCache();
  ui.alert(`✅ Global generada: ${datosUnicos.length} registros únicos de ${nombres.length} especialidades.`);
}

function obtenerTabsEspecialidad() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sistemaLower = TABS_SISTEMA.map(n => n.toLowerCase());
  return ss.getSheets().filter(s => {
    const nombre = s.getName().toLowerCase();
    return !sistemaLower.includes(nombre) && !nombre.startsWith('_');
  });
}

// ===================== AYUDA =====================

function mostrarAyuda() {
  const html = HtmlService.createHtmlOutput(`
    <style>body{font-family:-apple-system,sans-serif;padding:20px;color:#333;line-height:1.6}h3{color:#1a5c2e}code{background:#e9ecef;padding:2px 6px;border-radius:3px}li{margin-bottom:6px}</style>
    <h3>Concurso de Traslados — Ayuda</h3>
    <h4>Pestañas</h4>
    <ul>
      <li><b>Pestañas de especialidad</b> (Pri, Mus, Ing...): datos CSV importados</li>
      <li><b>Global</b>: fusión automática con fórmulas de centros (I-L) y especialidad (M)</li>
      <li><b>Centros</b>: fichero oficial de la Junta (código en col A)</li>
      <li><b>Especialidades</b>: códigos (col A) y abreviaturas (col C)</li>
    </ul>
    <h4>Columnas de Global</h4>
    <p>A: NIF · B: Nombre · C: CentroO · D: PuestoO · E: CentroAd · F: PuestoAd · G: PetAd · H: Puntos · I-L: fórmulas · M: Especialidad</p>
    <h4>Flujo</h4>
    <ol>
      <li>Descarga CSV desde la web del CdT</li>
      <li><b>📥 Importar CSV</b> → asignar a pestaña</li>
      <li><b>🔄 Regenerar Global</b> → fusiona todo</li>
      <li><b>🌐 Web</b> → comparte la URL de búsqueda</li>
    </ol>
  `).setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Ayuda');
}