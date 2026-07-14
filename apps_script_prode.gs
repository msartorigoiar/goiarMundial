// ============================================================
//  PRODE MUNDIAL 2026 — GOIAR FINTECH
//  Google Apps Script v6 — partido_id normalizado a 1-indexed
// ============================================================

var SHEET_ID      = "1OXyMZ0YDF01yfqIRh9YPqEZCnSQ98eNj0XL3csP3JIA";
var SHEET_PRED    = "Predicciones";
var SHEET_FIXTURE = "Fixture";
var FD_TOKEN      = "e628108cdbfa41c9ada7fb44e9f0fcf4";
var FD_BASE       = "https://api.football-data.org/v4";
var WC_2026_CODE  = "WC";

// ============================================================
//  doGet — router
// ============================================================
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "ranking";
  try {
    if (action === "fixture")      return json(getFixture());
    if (action === "save")         return json(savePrediccion(e.parameter.data));
    if (action === "updateScores") return json(updateScoresFromAPI());
    if (action === "resultados")   return json(getResultados());
    return json(getRanking());
  } catch(err) {
    return json({error: err.message});
  }
}

// ============================================================
//  savePrediccion — guarda predicciones con partido_id 1-indexed
// ============================================================
function savePrediccion(dataStr) {
  if (!dataStr) return {error: "Sin datos"};

  var body, usuario, preds;
  try {
    body    = JSON.parse(decodeURIComponent(dataStr));
    usuario = body.usuario;
    preds   = body.predicciones;
  } catch(e) {
    return {error: "JSON inválido: " + e.message};
  }

  if (!usuario) return {error: "Falta el campo usuario"};
  if (!preds)   return {error: "Falta el campo predicciones"};

  // Lock para evitar race conditions entre usuarios guardando simultáneamente
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // esperar hasta 10 segundos
  } catch(e) {
    return {error: "El servidor está ocupado, intentá de nuevo en unos segundos"};
  }

  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(SHEET_PRED);
    if (!sheet) { lock.releaseLock(); return {error: "No se encontró la hoja: " + SHEET_PRED}; }

  var ts   = new Date().toISOString();
  var data = sheet.getDataRange().getValues();

  // Construir mapa de predicciones existentes del usuario: pid → fila (1-based)
  var existingRows = {}; // pid → row index (1-based en la sheet)
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(usuario).trim()) {
      var pid = parseInt(data[i][1]);
      if (!isNaN(pid)) existingRows[pid] = i + 1; // row 1-based
    }
  }

  // Leer resultados de la sheet Fixture para saber qué partidos ya finalizaron
  var sheetFix  = ss.getSheetByName(SHEET_FIXTURE);
  var resultados = {};
  if (sheetFix) {
    var fixData = sheetFix.getDataRange().getValues();
    for (var f = 1; f < fixData.length; f++) {
      var fpid = parseInt(fixData[f][0]);
      var gl   = fixData[f][10];
      var gv   = fixData[f][11];
      if (!isNaN(fpid) && gl !== "" && gl !== null && gv !== "" && gv !== null) {
        resultados[fpid] = true; // partido ya finalizado
      }
    }
  }

  var actualizadas = 0;
  var insertadas   = 0;
  var rowsToInsert = [];

  for (var key in preds) {
    var val = preds[key];
    if (val.h === "" || val.h === undefined || val.a === "" || val.a === undefined) continue;
    var pidRaw = parseInt(key.replace("m", ""));
    if (isNaN(pidRaw)) continue;
    var pid = pidRaw + 1; // convertir a 1-indexed

    // No guardar predicciones de partidos ya finalizados (tienen resultado real)
    if (resultados[pid]) continue;

    var h = parseInt(val.h);
    var a = parseInt(val.a);
    if (isNaN(h) || isNaN(a)) continue;

    if (existingRows[pid]) {
      // Actualizar fila existente (solo goles, no borrar)
      sheet.getRange(existingRows[pid], 3).setValue(h);
      sheet.getRange(existingRows[pid], 4).setValue(a);
      sheet.getRange(existingRows[pid], 5).setValue(ts);
      actualizadas++;
    } else {
      // Insertar nueva fila
      rowsToInsert.push([String(usuario), pid, h, a, ts]);
      insertadas++;
    }
  }

  if (rowsToInsert.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, 5).setValues(rowsToInsert);
  }

    SpreadsheetApp.flush();
    return {ok: true, actualizadas: actualizadas, insertadas: insertadas, usuario: usuario};
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  getRanking — calcula puntos con partido_id 1-indexed
// ============================================================
function getRanking() {
  var ss        = SpreadsheetApp.openById(SHEET_ID);
  var sheetPred = ss.getSheetByName(SHEET_PRED);
  var sheetFix  = ss.getSheetByName(SHEET_FIXTURE);

  if (!sheetPred) return {error: "No se encontró: " + SHEET_PRED};

  // Mapa partido_id (1-indexed) → {gl, gv}
  var resultados = {};
  if (sheetFix) {
    var fixData = sheetFix.getDataRange().getValues();
    for (var i = 1; i < fixData.length; i++) {
      var pid = parseInt(fixData[i][0]);
      var gl  = parseInt(fixData[i][10]);
      var gv  = parseInt(fixData[i][11]);
      if (!isNaN(pid) && !isNaN(gl) && !isNaN(gv)) {
        resultados[pid] = {gl: gl, gv: gv};
      }
    }
  }

  // Leer predicciones
  var predData = sheetPred.getDataRange().getValues();
  var usuarios = {};

  for (var j = 1; j < predData.length; j++) {
    var usuario = String(predData[j][0]).trim();
    if (!usuario) continue;

    var pid       = parseInt(predData[j][1]); // siempre 1-indexed post-migración
    var predLocal = parseInt(predData[j][2]);
    var predVisit = parseInt(predData[j][3]);
    if (isNaN(pid) || isNaN(predLocal) || isNaN(predVisit)) continue;

    if (!usuarios[usuario]) {
      usuarios[usuario] = {predicciones: 0, puntos: 0, exactos: 0, correctos: 0};
    }
    usuarios[usuario].predicciones++;

    var res = resultados[pid];
    if (res) {
      var pL = predLocal, pV = predVisit;
      var rL = res.gl,    rV = res.gv;
      if (pL === rL && pV === rV) {
        usuarios[usuario].puntos += 3;
        usuarios[usuario].exactos++;
      } else {
        var predGan = pL > pV ? 1 : pL < pV ? -1 : 0;
        var resGan  = rL > rV ? 1 : rL < rV ? -1 : 0;
        if (predGan === resGan) {
          usuarios[usuario].puntos += 1;
          usuarios[usuario].correctos++;
        }
      }
    }
  }

  var ranking = [];
  for (var nombre in usuarios) {
    ranking.push({
      usuario:      nombre,
      predicciones: usuarios[nombre].predicciones,
      puntos:       usuarios[nombre].puntos,
      exactos:      usuarios[nombre].exactos,
      correctos:    usuarios[nombre].correctos
    });
  }

  ranking.sort(function(a, b) {
    return b.puntos - a.puntos || b.predicciones - a.predicciones;
  });

  return {ranking: ranking, total: 103, resultados_cargados: Object.keys(resultados).length};
}

// ============================================================
//  getResultados — devuelve mapa {partido_id: {gl, gv}}
// ============================================================
function getResultados() {
  var ss       = SpreadsheetApp.openById(SHEET_ID);
  var sheetFix = ss.getSheetByName(SHEET_FIXTURE);
  if (!sheetFix) return {resultados: {}};

  var data = sheetFix.getDataRange().getValues();
  var resultados = {};

  for (var i = 1; i < data.length; i++) {
    var pid = data[i][0];
    var gl  = data[i][10];
    var gv  = data[i][11];
    if (pid !== "" && gl !== "" && gl !== null && gl !== undefined &&
        gv !== "" && gv !== null && gv !== undefined) {
      var pidInt = parseInt(pid);
      var glInt  = parseInt(gl);
      var gvInt  = parseInt(gv);
      if (!isNaN(pidInt) && !isNaN(glInt) && !isNaN(gvInt)) {
        resultados[pidInt] = {gl: glInt, gv: gvInt};
        resultados[String(pidInt)] = {gl: glInt, gv: gvInt};
      }
    }
  }

  return {resultados: resultados};
}

// ============================================================
//  updateScoresFromAPI — trae resultados de football-data.org
// ============================================================
function updateScoresFromAPI() {
  // Lock para evitar que dos ejecuciones del trigger se solapen
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return {error: "Otra actualización en curso, se reintentará en el próximo ciclo"};
  }

  try {

  var ss       = SpreadsheetApp.openById(SHEET_ID);
  var sheetFix = ss.getSheetByName(SHEET_FIXTURE);
  if (!sheetFix) { lock.releaseLock(); return {error: "No se encontró hoja Fixture"}; }

  sheetFix.getRange("N1").setValue("Última ejecución: " + new Date().toISOString());

  // Traer partidos finalizados Y en curso
  var url = FD_BASE + "/competitions/" + WC_2026_CODE + "/matches?status=FINISHED,IN_PLAY,PAUSED&season=2026";
  var options = {
    method: "get",
    headers: {"X-Auth-Token": FD_TOKEN},
    muteHttpExceptions: true
  };

  var response, code;
  try {
    response = UrlFetchApp.fetch(url, options);
    code     = response.getResponseCode();
  } catch(fetchErr) {
    sheetFix.getRange("N2").setValue("ERROR fetch: " + fetchErr.message);
    return {error: "Fetch error: " + fetchErr.message};
  }

  sheetFix.getRange("N2").setValue("API status: " + code);

  if (code !== 200) {
    sheetFix.getRange("N3").setValue("API error: " + response.getContentText().substring(0, 300));
    return {error: "API error " + code};
  }

  var apiData  = JSON.parse(response.getContentText());
  var matches  = apiData.matches || [];

  if (matches.length === 0) {
    sheetFix.getRange("N3").setValue("Sin partidos finalizados");
    return {ok: true, actualizados: 0};
  }

  var fixData      = sheetFix.getDataRange().getValues();
  var actualizados = 0;

  // Función para obtener el resultado real (sin penales)
  // La API guarda en fullTime el total acumulado incluyendo penales
  // El resultado real = fullTime - penalties
  function getRealScore(match) {
    var s = match.score;
    if (!s || !s.fullTime) return null;
    var h = s.fullTime.home;
    var a = s.fullTime.away;
    if (h === null || a === null) return null;
    // Si hubo penales, restar los goles de penales para obtener resultado real
    if (s.penalties && s.penalties.home !== null && s.penalties.away !== null) {
      h = h - s.penalties.home;
      a = a - s.penalties.away;
    }
    return {h: h, a: a};
  }

  // Diccionario de traducción inglés → español
  var TEAM_MAP = {
    "mexico": "méxico", "south africa": "sudáfrica", "south korea": "corea del sur",
    "czechia": "rep. checa", "czech republic": "rep. checa", "canada": "canadá",
    "bosnia and herzegovina": "bosnia y herz.", "bosnia-herzegovina": "bosnia y herz.", "bosnia & herzegovina": "bosnia y herz.", "united states": "estados unidos",
    "usa": "estados unidos", "paraguay": "paraguay", "qatar": "qatar",
    "switzerland": "suiza", "brazil": "brasil", "morocco": "marruecos",
    "haiti": "haití", "scotland": "escocia", "australia": "australia",
    "turkey": "turquía", "germany": "alemania", "curacao": "curazao", "curaçao": "curazao",
    "netherlands": "países bajos", "japan": "japón", "ivory coast": "costa de marfil", "côte d'ivoire": "costa de marfil",
    "côte d'ivoire": "costa de marfil", "ecuador": "ecuador", "sweden": "suecia",
    "tunisia": "túnez", "spain": "españa", "cape verde": "cabo verde", "cape verde islands": "cabo verde",
    "belgium": "bélgica", "egypt": "egipto", "saudi arabia": "arabia saudita",
    "uruguay": "uruguay", "iran": "irán", "new zealand": "nueva zelanda",
    "france": "francia", "senegal": "senegal", "iraq": "irak", "norway": "noruega",
    "argentina": "argentina", "algeria": "argelia", "austria": "austria",
    "jordan": "jordania", "portugal": "portugal", "dr congo": "rd congo",
    "congo dr": "rd congo", "england": "inglaterra", "croatia": "croacia",
    "ghana": "ghana", "panama": "panamá", "uzbekistan": "uzbekistán", "colombia": "colombia",
    "denmark": "dinamarca", "serbia": "serbia", "poland": "polonia",
    "nigeria": "nigeria", "cameroon": "camerún", "mali": "mali",
    "venezuela": "venezuela", "chile": "chile", "peru": "perú",
    "costa rica": "costa rica", "honduras": "honduras", "jamaica": "jamaica"
  };

  function translateTeam(name) {
    var lower = name.toLowerCase().trim();
    return TEAM_MAP[lower] || lower;
  }

  // Traer TODOS los partidos para actualizar nombres eliminatorios
  var urlAll = FD_BASE + "/competitions/" + WC_2026_CODE + "/matches?season=2026";
  var respAll;
  try { respAll = UrlFetchApp.fetch(urlAll, options); } catch(e) { respAll = null; }

  var allMatches = [];
  if (respAll && respAll.getResponseCode() === 200) {
    allMatches = JSON.parse(respAll.getContentText()).matches || [];
  }

  // Construir mapa por fecha+hora UTC → {home, away}
  // Para matchear con la sheet usamos fecha UTC del partido
  var nombresActualizados = 0;

  allMatches.forEach(function(match) {
    var homeTeam = match.homeTeam && match.homeTeam.name ? match.homeTeam.name : "";
    var awayTeam = match.awayTeam && match.awayTeam.name ? match.awayTeam.name : "";
    if (!homeTeam || !awayTeam || homeTeam === "TBD" || awayTeam === "TBD") return;

    var apiHomeEs = translateTeam(homeTeam);
    var apiAwayEs = translateTeam(awayTeam);

    // Fecha del partido en UTC: "2026-06-29T18:00:00Z" → "2026-06-29"
    var apiDateStr = match.utcDate ? match.utcDate.substring(0, 10) : "";
    if (!apiDateStr) return;

    // Buscar en filas eliminatorias (id >= 73) con nombre genérico
    for (var i = 1; i < fixData.length; i++) {
      var pid = parseInt(fixData[i][0]);
      if (isNaN(pid) || pid < 73) continue;

      var sheetHome = String(fixData[i][5]).trim();
      var sheetAway = String(fixData[i][6]).trim();

      // Solo actualizar si el nombre actual es genérico
      var isGeneric = function(n) {
        return n.indexOf("°") !== -1 || n.indexOf("Gan") !== -1 || n.indexOf("Per") !== -1 || n.indexOf("mejor") !== -1;
      };
      if (!isGeneric(sheetHome) && !isGeneric(sheetAway)) continue;

      // Matchear por fecha UTC de la sheet
      // La fecha en la sheet (col B, índice 1) es un objeto Date de Sheets
      var sheetDateRaw = fixData[i][1];
      var sheetDateStr = "";
      if (sheetDateRaw instanceof Date) {
        sheetDateStr = Utilities.formatDate(sheetDateRaw, "UTC", "yyyy-MM-dd");
      } else if (typeof sheetDateRaw === "string" && sheetDateRaw.indexOf("T") !== -1) {
        sheetDateStr = sheetDateRaw.substring(0, 10);
      }

      if (sheetDateStr && apiDateStr === sheetDateStr) {
        // Capitalizar primera letra
        var cap = function(s) { return s.charAt(0).toUpperCase() + s.slice(1); };
        sheetFix.getRange(i + 1, 6).setValue(cap(apiHomeEs));
        sheetFix.getRange(i + 1, 7).setValue(cap(apiAwayEs));
        nombresActualizados++;
        fixData[i][5] = cap(apiHomeEs); // actualizar en memoria para no re-matchear
        fixData[i][6] = cap(apiAwayEs);
        break;
      }
    }
  });

  sheetFix.getRange("N5").setValue("Nombres eliminatorios actualizados: " + nombresActualizados);

  // Actualizar resultados de partidos finalizados
  matches.forEach(function(match) {
    var realScore = getRealScore(match);
    if (!realScore) return;
    var homeGoals = realScore.h;
    var awayGoals = realScore.a;

    var homeTeam  = match.homeTeam.name;
    var awayTeam  = match.awayTeam.name;
    var apiHomeEs = translateTeam(homeTeam);
    var apiAwayEs = translateTeam(awayTeam);

    for (var i = 1; i < fixData.length; i++) {
      var sheetHome = String(fixData[i][5]).trim().toLowerCase();
      var sheetAway = String(fixData[i][6]).trim().toLowerCase();
      var fuente    = fixData[i][12] ? String(fixData[i][12]).toLowerCase() : "";

      if (fuente === "manual") continue;

      // Match por nombre (fase de grupos) o por fecha+posición (eliminatoria)
      var nameMatch = (apiHomeEs === sheetHome || apiHomeEs.indexOf(sheetHome) !== -1 || sheetHome.indexOf(apiHomeEs) !== -1) &&
                      (apiAwayEs === sheetAway || apiAwayEs.indexOf(sheetAway) !== -1 || sheetAway.indexOf(apiAwayEs) !== -1);

      if (nameMatch) {
        // NUNCA pisar resultado manual
        var glActual = fixData[i][10];
        var gvActual = fixData[i][11];
        if (fuente === "manual") continue;

        sheetFix.getRange(i + 1, 11).setValue(homeGoals);
        sheetFix.getRange(i + 1, 12).setValue(awayGoals);
        sheetFix.getRange(i + 1, 13).setValue("api");
        actualizados++;
        break;
      }
    }
  });

  SpreadsheetApp.flush();
  sheetFix.getRange("N3").setValue("Partidos API: " + matches.length + " | Actualizados: " + actualizados);
  var apiNames = matches.map(function(m) { return m.homeTeam.name + " vs " + m.awayTeam.name; }).join(" | ");
  sheetFix.getRange("N4").setValue("Nombres API: " + apiNames.substring(0, 500));

  return {ok: true, actualizados: actualizados, partidos_api: matches.length};

  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  Trigger automático — cada hora
// ============================================================
function autoUpdateScores() {
  updateScoresFromAPI();
  // autoUpdateNombresEliminatoria(); // DESHABILITADO — usar cargarNombresEliminatoria manualmente
}

// ============================================================
//  getFixture
// ============================================================
function getFixture() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_FIXTURE);
  if (!sheet) return {fixture: []};

  var data = sheet.getDataRange().getValues();
  var keys = data[0].map(function(h) { return h.toString().toLowerCase(); });
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    keys.forEach(function(k, j) { row[k] = data[i][j]; });
    rows.push(row);
  }
  return {fixture: rows};
}


// ============================================================
//  cargarNombresEliminatoria — carga nombres oficiales del bracket
//  Ejecutar manualmente cuando cambien los cruces
// ============================================================
function cargarNombresEliminatoria() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_FIXTURE);
  var data  = sheet.getDataRange().getValues();

  // Mapa partido_id → {local, visitante}
  var nombres = {
    // 16avos (ids 73-88)
    73:  {l:"Sudáfrica",      v:"Canadá"},
    74:  {l:"Brasil",         v:"Japón"},
    75:  {l:"Alemania",       v:"Paraguay"},
    76:  {l:"Países Bajos",   v:"Marruecos"},
    77:  {l:"Costa de Marfil",v:"Noruega"},
    78:  {l:"Francia",        v:"Suecia"},
    79:  {l:"México",         v:"Ecuador"},
    80:  {l:"Inglaterra",     v:"RD Congo"},
    81:  {l:"Bélgica",        v:"Senegal"},
    82:  {l:"Estados Unidos", v:"Bosnia y Herz."},
    83:  {l:"España",         v:"Austria"},
    84:  {l:"Portugal",       v:"Croacia"},
    85:  {l:"Suiza",          v:"Argelia"},
    86:  {l:"Australia",      v:"Egipto"},
    87:  {l:"Argentina",      v:"Cabo Verde"},
    88:  {l:"Colombia",       v:"Ghana"},
    // Octavos (ids 89-96) — orden según bracket FX
    89:  {l:"Canadá",         v:"Marruecos"},
    90:  {l:"Paraguay",       v:"Francia"},
    91:  {l:"Brasil",         v:"Noruega"},
    92:  {l:"Ecuador",        v:"Inglaterra"},
    93:  {l:"Estados Unidos", v:"Bélgica"},
    94:  {l:"Portugal",       v:"España"},
    95:  {l:"Suiza",          v:"Colombia"},
    96:  {l:"Argentina",      v:"Egipto"},
  };

  var actualizados = 0;
  for (var i = 1; i < data.length; i++) {
    var pid = parseInt(data[i][0]);
    if (!nombres[pid]) continue;
    var actual_l = data[i][5];
    var actual_v = data[i][6];
    var nuevo_l  = nombres[pid].l;
    var nuevo_v  = nombres[pid].v;
    // Solo actualizar si el nombre es genérico o vacío
    if (!actual_l || actual_l.indexOf('°') !== -1 || actual_l.indexOf('Gan') !== -1) {
      sheet.getRange(i+1, 6).setValue(nuevo_l);
      actualizados++;
    }
    if (!actual_v || actual_v.indexOf('°') !== -1 || actual_v.indexOf('Gan') !== -1) {
      sheet.getRange(i+1, 7).setValue(nuevo_v);
      actualizados++;
    }
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_FIXTURE)
    .getRange("N6").setValue("Nombres cargados: " + actualizados);
}

function guardarAPIIds() {
  var ss       = SpreadsheetApp.openById(SHEET_ID);
  var sheetFix = ss.getSheetByName(SHEET_FIXTURE);
  var data     = sheetFix.getDataRange().getValues();

  // Agregar header en columna N si no existe
  if (!data[0][13] || data[0][13] === "") {
    sheetFix.getRange(1, 14).setValue("api_id");
  }

  // Traer todos los partidos eliminatorios de la API
  var stages = ["LAST_32", "LAST_32", "LAST_16", "QUARTER_FINALS", "SEMI_FINALS", "FINAL", "THIRD_PLACE"];
  var allMatches = [];

  stages.forEach(function(stage) {
    var url = FD_BASE + "/competitions/" + WC_2026_CODE + "/matches?season=2026&stage=" + stage;
    try {
      var resp = UrlFetchApp.fetch(url, {
        headers: {"X-Auth-Token": FD_TOKEN},
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        var matches = JSON.parse(resp.getContentText()).matches || [];
        allMatches = allMatches.concat(matches);
      }
    } catch(e) {}
    Utilities.sleep(300);
  });

  // Ordenar partidos de la API por fecha UTC
  allMatches.sort(function(a, b) {
    return new Date(a.utcDate) - new Date(b.utcDate);
  });

  // Obtener filas eliminatorias de la sheet ordenadas por id
  var sheetRows = [];
  for (var i = 1; i < data.length; i++) {
    var pid = parseInt(data[i][0]);
    if (!isNaN(pid) && pid >= 73) {
      sheetRows.push({row: i, pid: pid, apiId: data[i][13]});
    }
  }
  sheetRows.sort(function(a, b) { return a.pid - b.pid; });

  // Mapear en orden: el partido N de la API corresponde al partido N de la sheet
  // dentro del mismo día calendario ARG
  var mapeados = 0;

  // Agrupar API matches por fecha ARG
  var apiByDate = {};
  allMatches.forEach(function(m) {
    var dt    = new Date(m.utcDate);
    var dtArg = new Date(dt.getTime() - 3 * 3600000);
    var dd    = Utilities.formatDate(dtArg, "GMT-3", "yyyy-MM-dd");
    if (!apiByDate[dd]) apiByDate[dd] = [];
    apiByDate[dd].push(m);
  });

  // Agrupar sheet rows por fecha ARG
  var sheetByDate = {};
  sheetRows.forEach(function(sr) {
    var raw = data[sr.row][1];
    if (!raw) return;
    var dt  = raw instanceof Date ? raw : new Date(raw);
    var dd  = Utilities.formatDate(dt, "GMT-3", "yyyy-MM-dd");
    if (!sheetByDate[dd]) sheetByDate[dd] = [];
    sheetByDate[dd].push(sr);
  });

  // Para cada día, mapear en orden de hora
  for (var dd in apiByDate) {
    var apiDayMatches   = apiByDate[dd];
    var sheetDayMatches = sheetByDate[dd] || [];
    if (sheetDayMatches.length === 0) continue;

    // Ordenar API por hora UTC
    apiDayMatches.sort(function(a, b) { return new Date(a.utcDate) - new Date(b.utcDate); });

    // Mapear en orden
    var count = Math.min(apiDayMatches.length, sheetDayMatches.length);
    for (var j = 0; j < count; j++) {
      if (sheetDayMatches[j].apiId) continue; // ya tiene api_id
      sheetFix.getRange(sheetDayMatches[j].row + 1, 14).setValue(apiDayMatches[j].id);
      data[sheetDayMatches[j].row][13] = apiDayMatches[j].id;
      mapeados++;
    }
  }

  SpreadsheetApp.flush();
  sheetFix.getRange("U1").setValue("API IDs mapeados: " + mapeados);
}

// ============================================================
//  autoUpdateNombresEliminatoria — actualiza nombres usando api_id
//  Se llama desde autoUpdateScores automáticamente
// ============================================================
function autoUpdateNombresEliminatoria() {
  var ss       = SpreadsheetApp.openById(SHEET_ID);
  var sheetFix = ss.getSheetByName(SHEET_FIXTURE);
  var data     = sheetFix.getDataRange().getValues();

  // Traer todos los partidos eliminatorios
  var stages = ["LAST_32", "LAST_32", "LAST_16", "QUARTER_FINALS", "SEMI_FINALS", "FINAL", "THIRD_PLACE"];
  var apiMatchMap = {}; // api_id → {home, away}

  var TEAM_MAP = {
    "mexico": "México", "south africa": "Sudáfrica", "south korea": "Corea del Sur",
    "czechia": "Rep. Checa", "canada": "Canadá", "bosnia and herzegovina": "Bosnia y Herz.",
    "bosnia-herzegovina": "Bosnia y Herz.", "united states": "Estados Unidos",
    "switzerland": "Suiza", "brazil": "Brasil", "morocco": "Marruecos",
    "haiti": "Haití", "scotland": "Escocia", "australia": "Australia",
    "turkey": "Turquía", "germany": "Alemania", "curacao": "Curazao", "curaçao": "Curazao",
    "netherlands": "Países Bajos", "japan": "Japón", "ivory coast": "Costa de Marfil",
    "côte d'ivoire": "Costa de Marfil", "ecuador": "Ecuador", "sweden": "Suecia",
    "tunisia": "Túnez", "spain": "España", "cape verde": "Cabo Verde",
    "cape verde islands": "Cabo Verde", "belgium": "Bélgica", "egypt": "Egipto",
    "saudi arabia": "Arabia Saudita", "uruguay": "Uruguay", "iran": "Irán",
    "new zealand": "Nueva Zelanda", "france": "Francia", "senegal": "Senegal",
    "iraq": "Irak", "norway": "Noruega", "argentina": "Argentina", "algeria": "Argelia",
    "austria": "Austria", "jordan": "Jordania", "portugal": "Portugal",
    "dr congo": "RD Congo", "congo dr": "RD Congo", "england": "Inglaterra",
    "croatia": "Croacia", "ghana": "Ghana", "panama": "Panamá",
    "uzbekistan": "Uzbekistán", "colombia": "Colombia", "paraguay": "Paraguay",
    "qatar": "Qatar", "costa rica": "Costa Rica", "mali": "Mali",
    "nigeria": "Nigeria", "cameroon": "Camerún", "cape verde": "Cabo Verde",
    "senegal": "Senegal", "norway": "Noruega", "sweden": "Suecia"
  };

  function translateTeam(name) {
    if (!name) return null;
    var key = name.toLowerCase().trim();
    return TEAM_MAP[key] || (name.charAt(0).toUpperCase() + name.slice(1));
  }

  stages.forEach(function(stage) {
    var url = FD_BASE + "/competitions/" + WC_2026_CODE + "/matches?season=2026&stage=" + stage;
    try {
      var resp = UrlFetchApp.fetch(url, {
        headers: {"X-Auth-Token": FD_TOKEN},
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        var matches = JSON.parse(resp.getContentText()).matches || [];
        matches.forEach(function(m) {
          var home = m.homeTeam && m.homeTeam.name && m.homeTeam.name !== "null" ? translateTeam(m.homeTeam.name) : null;
          var away = m.awayTeam && m.awayTeam.name && m.awayTeam.name !== "null" ? translateTeam(m.awayTeam.name) : null;
          if (home && away) {
            apiMatchMap[m.id] = {home: home, away: away};
          }
        });
      }
    } catch(e) {}
    Utilities.sleep(300);
  });

  // Actualizar nombres en sheet usando api_id (columna N, índice 13)
  var actualizados = 0;
  for (var i = 1; i < data.length; i++) {
    var pid    = parseInt(data[i][0]);
    var apiId  = data[i][13];
    var fuente = String(data[i][12] || "").toLowerCase();

    if (isNaN(pid) || pid < 73 || !apiId || fuente === "manual") continue;

    var match = apiMatchMap[apiId];
    if (!match) continue;

    var sheetHome = String(data[i][5]).trim();
    var sheetAway = String(data[i][6]).trim();

    // Solo actualizar si el nombre cambió
    if (sheetHome !== match.home || sheetAway !== match.away) {
      sheetFix.getRange(i + 1, 6).setValue(match.home);
      sheetFix.getRange(i + 1, 7).setValue(match.away);
      actualizados++;
    }
  }

  SpreadsheetApp.flush();
  return {nombres_actualizados: actualizados};
}


// ============================================================
//  corregirResultadosPenales — corrige partidos guardados con penales
// ============================================================
function corregirResultadosPenales() {
  var ss       = SpreadsheetApp.openById(SHEET_ID);
  var sheetFix = ss.getSheetByName(SHEET_FIXTURE);
  var data     = sheetFix.getDataRange().getValues();

  // Mapa de correcciones: partido_id → {gl, gv} real (sin penales)
  var correcciones = {
    75: {gl: 1, gv: 1}, // Alemania vs Paraguay: 1-1 (penales 3-4)
    76: {gl: 1, gv: 1}, // Países Bajos vs Marruecos: 1-1 (penales 2-3)
  };

  var corregidos = 0;
  for (var i = 1; i < data.length; i++) {
    var pid    = parseInt(data[i][0]);
    var fuente = String(data[i][12] || "").toLowerCase();
    if (fuente === "manual") continue;
    if (correcciones[pid]) {
      sheetFix.getRange(i + 1, 11).setValue(correcciones[pid].gl);
      sheetFix.getRange(i + 1, 12).setValue(correcciones[pid].gv);
      corregidos++;
    }
  }

  SpreadsheetApp.flush();
  sheetFix.getRange("N13").setValue("Corregidos: " + corregidos);
}


// ============================================================
//  instalarTrigger15min — ejecutar UNA VEZ manualmente
//  Instala un trigger que actualiza resultados cada 15 minutos
// ============================================================
function instalarTrigger15min() {
  // Borrar triggers existentes de autoUpdateScores
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'autoUpdateScores') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Crear nuevo trigger cada 15 minutos
  ScriptApp.newTrigger('autoUpdateScores')
    .timeBased()
    .everyMinutes(15)
    .create();
  SpreadsheetApp.openById(SHEET_ID)
    .getSheetByName(SHEET_FIXTURE)
    .getRange("N5").setValue("Trigger 15min instalado: " + new Date().toISOString());
}


// ============================================================
//  corregirResultadoOctavos — corrige el resultado mal mapeado
//  id 88 de la sheet tiene el resultado de Canadá-Marruecos
//  pero en el FX id 88 = Colombia-Ghana (16avos, no jugado aún)
//  Mover ese resultado al id correcto de Octavos
// ============================================================
function corregirResultadoOctavos() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_FIXTURE);
  var data  = sheet.getDataRange().getValues();
  var log   = "";

  // Buscar id 88 y id 89
  for (var i = 1; i < data.length; i++) {
    var pid = parseInt(data[i][0]);
    if (pid === 88) {
      var gl = data[i][10];
      var gv = data[i][11];
      log += "id=88 antes: gl="+gl+" gv="+gv+"\n";
      // Borrar resultado de id 88 (Colombia-Ghana, no se jugó)
      sheet.getRange(i+1, 11).setValue("");
      sheet.getRange(i+1, 12).setValue("");
      sheet.getRange(i+1, 13).setValue("");
      log += "id=88 limpiado\n";
    }
    if (pid === 89) {
      var gl89 = data[i][10];
      log += "id=89 antes: gl="+gl89+"\n";
      // Guardar resultado de Canadá-Marruecos en id 89
      sheet.getRange(i+1, 11).setValue(1);
      sheet.getRange(i+1, 12).setValue(0);
      sheet.getRange(i+1, 13).setValue("manual");
      log += "id=89 corregido: Canada 1-0 Marruecos\n";
    }
  }

  SpreadsheetApp.flush();
  sheet.getRange("P1").setValue(log);
}

// ============================================================
//  Helper JSON
// ============================================================
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
