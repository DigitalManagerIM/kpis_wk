/**
 * Actualizador mensual de KPIs — Workprotec / Sermaco
 * ----------------------------------------------------
 * Despliega esto como Web App (Ejecutar como: Yo / Acceso: Cualquiera con el enlace)
 * y Claude le hará un POST cada mes con los datos ya extraídos del PDF.
 * No requiere que compartas el Sheet ni conectes tu Drive con nadie.
 */

// ===================== CONFIG =====================

const SHEET_IDS = {
  workprotec: '1y4PZUu1bS-WMxYHLifefX6tuvhVOcBqol_bFqoyveO8',
  sermaco: '1Hw_zbJ4sTT--MUO9i-RYji-FmJzwr4QZrDyYk-Oltec'
};

// Nombre de la pestaña principal de KPIs dentro de cada spreadsheet.
// Si tu pestaña se llama distinto, cámbialo aquí.
const KPI_TAB_NAME = {
  workprotec: 'WORKPROTEC',
  sermaco: 'SERMACO'
};

// Textos EXACTOS (columna A) que marcan el inicio de cada bloque/sección.
// Estas filas llevan también los encabezados de mes y se actualizan igual
// que las demás filas de datos.
const SECTION_HEADERS = ['Rendimiento Página web', 'Procedencia', 'Objetivos cumplidos'];

// Fila donde están los encabezados de mes (para calcular la próxima columna libre).
const HEADER_ROW = 2;

// ===================== ENTRY POINTS =====================

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: 'KPI updater activo' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
    if (!secret || body.secret !== secret) {
      return jsonOut({ ok: false, error: 'Token inválido' });
    }

    const brand = body.brand; // 'workprotec' | 'sermaco'
    if (!SHEET_IDS[brand]) {
      return jsonOut({ ok: false, error: 'brand desconocido: ' + brand });
    }

    const result = { ok: true, brand: brand, kpiResult: null, altaResult: null };

    if (body.kpis && body.month) {
      result.kpiResult = updateKpiSheet(brand, body.month, body.kpis);
    }

    if (body.altaClientes) {
      result.altaResult = updateAltaClientes(brand, body.altaClientes);
    }

    return jsonOut(result);
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ===================== BLOQUE 1: tabla de KPIs (Rendimiento/Procedencia/Objetivos) =====================

/**
 * kpis: objeto { "Sesiones totales": 3570, "Visitantes únicos": 2827, ... }
 * Las claves deben coincidir (recortando espacios, sin distinguir mayúsc/minúsc)
 * con el texto de columna A de cada fila.
 */
function updateKpiSheet(brand, month, kpis) {
  const ss = SpreadsheetApp.openById(SHEET_IDS[brand]);
  const sheet = ss.getSheetByName(KPI_TAB_NAME[brand]);
  if (!sheet) throw new Error('No se encontró la pestaña ' + KPI_TAB_NAME[brand]);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headerRowValues = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

  // última columna con algo escrito en la fila de cabecera -> la siguiente es la nueva
  let lastUsedCol = 1;
  for (let c = lastCol; c >= 1; c--) {
    if (headerRowValues[c - 1] !== '' && headerRowValues[c - 1] !== null) { lastUsedCol = c; break; }
  }
  const newCol = lastUsedCol + 1;

  const colAValues = sheet.getRange(1, 1, lastRow, 1).getValues();
  const unmatched = [];
  let sectionsFound = 0;

  // normaliza para comparar
  const norm = s => String(s || '').trim().toLowerCase();
  const kpiMap = {};
  Object.keys(kpis).forEach(k => kpiMap[norm(k)] = kpis[k]);

  for (let r = 1; r <= lastRow; r++) {
    const label = String(colAValues[r - 1][0] || '').trim();
    if (!label) continue;

    if (SECTION_HEADERS.indexOf(label) !== -1) {
      // fila de cabecera de sección -> escribir el mes
      sheet.getRange(r, newCol).setValue(month);
      sectionsFound++;
      continue;
    }

    const key = norm(label);
    if (Object.prototype.hasOwnProperty.call(kpiMap, key)) {
      const val = kpiMap[key];
      if (val !== null && val !== undefined && val !== '') {
        sheet.getRange(r, newCol).setValue(val);
      }
      delete kpiMap[key]; // marcar como usado
    }
  }

  const notWritten = Object.keys(kpiMap); // valores del PDF que no encontraron fila
  return {
    newColumn: newCol,
    sectionsFound: sectionsFound,
    valuesNotMatched: notWritten
  };
}

// ===================== BLOQUE 2: pestaña "Alta clientes" =====================

/**
 * altaClientes: {
 *   prevLabel: "jun-25", prevValue: 48,
 *   curLabel:  "jun-26", curValue:  53
 * }
 * Busca automáticamente la pestaña que contenga "alta" en el nombre,
 * localiza las dos celdas de cabecera con formato mes-año (ej. "jun-25")
 * y escribe debajo el valor correspondiente.
 */
function updateAltaClientes(brand, alta) {
  const ss = SpreadsheetApp.openById(SHEET_IDS[brand]);
  const sheets = ss.getSheets();
  const altaSheet = sheets.find(s => /alta/i.test(s.getName()));
  if (!altaSheet) throw new Error('No se encontró ninguna pestaña con "alta" en el nombre');

  const range = altaSheet.getRange(1, 1, Math.min(15, altaSheet.getLastRow() || 15), Math.min(15, altaSheet.getLastColumn() || 15));
  const values = range.getValues();
  const monthRegex = /^[a-zA-Zñ]{3}-\d{2}$/;

  const headerCells = [];
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const v = String(values[r][c] || '').trim();
      if (monthRegex.test(v)) headerCells.push({ row: r + 1, col: c + 1, value: v });
    }
  }

  if (headerCells.length < 2) {
    return { warning: 'No se localizaron 2 celdas de cabecera con formato mes-aa (ej. jun-25). Revisa manualmente la pestaña ' + altaSheet.getName(), found: headerCells };
  }

  // Asumimos que las dos primeras encontradas (de izq a der / arriba a abajo) son prev/actual
  headerCells.sort((a, b) => a.row - b.row || a.col - b.col);
  const prevCell = headerCells[0];
  const curCell = headerCells[1];

  altaSheet.getRange(prevCell.row, prevCell.col).setValue(alta.prevLabel);
  altaSheet.getRange(curCell.row, curCell.col).setValue(alta.curLabel);
  altaSheet.getRange(prevCell.row + 1, prevCell.col).setValue(alta.prevValue);
  altaSheet.getRange(curCell.row + 1, curCell.col).setValue(alta.curValue);

  return {
    sheet: altaSheet.getName(),
    prevCell: prevCell, curCell: curCell
  };
}

// ===================== UTILIDAD: fijar el secreto una vez =====================
// Ejecuta esta función UNA VEZ manualmente desde el editor (▶) para guardar tu token.
function setSecretOnce() {
  PropertiesService.getScriptProperties().setProperty('SECRET', 'CAMBIA_ESTO_POR_TU_TOKEN');
}
