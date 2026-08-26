/**
 * Actualizador mensual de KPIs — Workprotec / Sermaco
 * ----------------------------------------------------
 * Despliega esto como Web App (Ejecutar como: Yo / Acceso: Cualquiera con el enlace)
 * y Claude le hará un POST cada mes con los datos ya extraídos del PDF.
 * No requiere que compartas el Sheet ni conectes tu Drive con nadie.
 *
 * Estructura real de cada spreadsheet:
 * - La mayoría de las pestañas son series mensuales: la fila 1 tiene los meses
 *   como cabecera (se va añadiendo una columna nueva cada mes) y la columna A
 *   tiene el nombre de cada métrica (ej. "Totales", "Orgánico", "Web"...).
 * - Cada spreadsheet tiene además UNA pestaña de comparación con solo 2
 *   columnas (mes actual vs. mismo mes del año anterior) — ver
 *   COMPARISON_TAB_NAME más abajo.
 */

// ===================== CONFIG =====================

const SHEET_IDS = {
  workprotec: '1y4PZUu1bS-WMxYHLifefX6tuvhVOcBqol_bFqoyveO8',
  sermaco: '1Hw_zbJ4sTT--MUO9i-RYji-FmJzwr4QZrDyYk-Oltec'
};

// Pestaña de comparación "mes actual vs. mismo mes del año anterior" (2 columnas).
// Todas las demás pestañas del spreadsheet se tratan como series mensuales genéricas.
const COMPARISON_TAB_NAME = {
  workprotec: 'T5. Comparativa altas',
  sermaco: 'T4.Alta Clientes'
};

// Fila donde están los encabezados de mes en las pestañas de serie mensual.
const HEADER_ROW = 1;

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
      result.kpiResult = updateAllKpiSheets(brand, body.month, body.kpis);
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

// ===================== BLOQUE 1: todas las pestañas de serie mensual =====================

/**
 * kpis: objeto { "Totales": 3570, "Orgánico": 2827, "Web": 44, ... } con los
 * nombres de fila (columna A) tal cual aparecen en cualquiera de las pestañas
 * de serie mensual. El script recorre TODAS las pestañas del spreadsheet
 * (excepto la de comparación) y reparte cada valor en la pestaña donde
 * encuentre una fila con ese nombre.
 */
function updateAllKpiSheets(brand, month, kpis) {
  const ss = SpreadsheetApp.openById(SHEET_IDS[brand]);
  const comparisonName = COMPARISON_TAB_NAME[brand];
  const norm = s => String(s || '').trim().toLowerCase();

  const kpiMap = {};
  Object.keys(kpis).forEach(k => kpiMap[norm(k)] = kpis[k]);

  const sheetsUpdated = [];

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name === comparisonName) return; // esa se actualiza aparte, ver updateAltaClientes

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return; // pestaña vacía

    const headerRowValues = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
    let lastUsedCol = 1;
    for (let c = lastCol; c >= 1; c--) {
      if (headerRowValues[c - 1] !== '' && headerRowValues[c - 1] !== null) { lastUsedCol = c; break; }
    }
    const newCol = lastUsedCol + 1;

    const colAValues = sheet.getRange(1, 1, lastRow, 1).getValues();
    let matchedInSheet = 0;

    for (let r = HEADER_ROW + 1; r <= lastRow; r++) {
      const label = String(colAValues[r - 1][0] || '').trim();
      if (!label) continue;

      const key = norm(label);
      if (Object.prototype.hasOwnProperty.call(kpiMap, key)) {
        const val = kpiMap[key];
        if (val !== null && val !== undefined && val !== '') {
          sheet.getRange(r, newCol).setValue(val);
        }
        delete kpiMap[key]; // marcar como usado
        matchedInSheet++;
      }
    }

    if (matchedInSheet > 0) {
      sheet.getRange(HEADER_ROW, newCol).setValue(month);
      sheetsUpdated.push({ sheet: name, newColumn: newCol, matched: matchedInSheet });
    }
  });

  const notWritten = Object.keys(kpiMap); // valores del PDF que no encontraron fila en ninguna pestaña
  return {
    sheetsUpdated: sheetsUpdated,
    valuesNotMatched: notWritten
  };
}

// ===================== BLOQUE 2: pestaña de comparación (mes actual vs. año anterior) =====================

/**
 * altaClientes: {
 *   prevLabel: "jun-25", prevValue: 48,
 *   curLabel:  "jun-26", curValue:  53
 * }
 * Localiza las dos celdas de cabecera con formato mes-año (ej. "jun-25") en
 * la pestaña de comparación configurada en COMPARISON_TAB_NAME y escribe
 * debajo el valor correspondiente.
 */
function updateAltaClientes(brand, alta) {
  const ss = SpreadsheetApp.openById(SHEET_IDS[brand]);
  const name = COMPARISON_TAB_NAME[brand];
  const altaSheet = ss.getSheetByName(name);
  if (!altaSheet) throw new Error('No se encontró la pestaña ' + name);

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
    return { warning: 'No se localizaron 2 celdas de cabecera con formato mes-aa (ej. jun-25). Revisa manualmente la pestaña ' + name, found: headerCells };
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
    sheet: name,
    prevCell: prevCell, curCell: curCell
  };
}

// ===================== UTILIDAD: fijar el secreto una vez =====================
// Ejecuta esta función UNA VEZ manualmente desde el editor (▶) para guardar tu token.
function setSecretOnce() {
  PropertiesService.getScriptProperties().setProperty('SECRET', 'CAMBIA_ESTO_POR_TU_TOKEN');
}
