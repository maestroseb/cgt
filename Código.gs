// ============================================================
//  CONCURSO DE TRASLADOS — Automatización Google Sheets
//  v4.0 — Menú, importación CSV, Global, Web de búsqueda
//  HTML de la web en archivo separado: Página.html
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
  // Endpoint JSON para peticiones de datos
  if (e && e.parameter && e.parameter.action === 'getData') {
    return ContentService.createTextOutput(JSON.stringify(obtenerDatosGlobal()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Servir la página HTML desde archivo separado
  // IMPORTANTE: no ejecutar ninguna lógica aquí para evitar timeouts/errores
  return HtmlService.createHtmlOutputFromFile('Página')
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
