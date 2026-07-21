/**
 * =================================================================================================
 * Formatos de Cotización | Sistema de cotizaciones Ventel
 * =================================================================================================
 * Gestiona los distintos formatos con los que se puede imprimir/enviar una cotización:
 *
 *  1. 'actual'        -> El PDF que se arma desde HTML (generateQuoteHtml en Correos.gs).
 *  2. 'ccl_liverpool' -> El formato oficial CCL. Se genera copiando la Google Sheet plantilla,
 *                        llenándola y exportándola a PDF, que es el mismo camino que se hacía
 *                        a mano y por eso conserva la fidelidad del formato.
 *
 * Cada formato se puede habilitar/deshabilitar desde el panel de administración.
 */

// =================================================================================================
// CONFIGURACIÓN — REVISAR ANTES DE USAR
// =================================================================================================

/**
 * ID de la Google Sheet que contiene la plantilla del formato CCL.
 * Se saca de la URL de la hoja:
 *   https://docs.google.com/spreadsheets/d/[ESTE_ES_EL_ID]/edit
 * Sin este valor el formato CCL no se puede generar.
 */
const CCL_TEMPLATE_SHEET_ID = "1zD_0TiN7EBKfYIjNWzAH77pYJ021GGqilI1jUJ000sI";

/** Nombre de la pestaña dentro de la plantilla que se usa como formato. */
const CCL_SHEET_NAME = "Liverpool";

/** Carpeta de Drive donde viven las hojas CCL generadas. Se crea sola la primera vez. */
const CCL_OUTPUT_FOLDER_NAME = "Cotizaciones CCL generadas";

/** Columna de la hoja 'Cotizaciones' donde se guarda el enlace al documento generado. */
const CCL_LINK_COLUMN = "LinkSheetCCL";

/**
 * Parámetros de exportación a PDF. Replican la configuración de impresión de la plantilla
 * (horizontal, ajustada al ancho, centrada). Si el PDF no sale idéntico al de referencia,
 * este es el objeto a ajustar.
 */
const CCL_EXPORT_OPTIONS = {
  size: "A4",              // Tamaño de papel de la plantilla. Cambiar a "letter" si aplica.
  portrait: "false",       // La plantilla es horizontal.
  fitw: "true",            // Ajustar al ancho de la página.
  gridlines: "false",
  printtitle: "false",
  sheetnames: "false",
  pagenum: "UNDEFINED",
  horizontal_alignment: "CENTER",
  top_margin: "0.75",
  bottom_margin: "0.75",
  left_margin: "0.7",
  right_margin: "0.7"
};

/**
 * Ubicación de los datos dentro de la plantilla CCL. Coincide con el importador de la
 * Bolsa de Liverpool. Si se reacomoda la plantilla, se ajusta aquí.
 */
const CCL_CONFIG = {
  CELDA_ASESOR: "A15",
  CELDA_CLIENTE: "J15",
  CELDA_CORREO: "K16",
  CELDA_TELEFONO: "K17",
  CELDA_OBSERVACIONES: "K18",

  TEXTO_CABECERA: "sku",          // Cabecera de la tabla, en la columna A.
  TEXTO_FIN_TABLA: "informacion", // Marca el final de la tabla, en la columna A.
  TEXTO_SUBTOTAL: "subtotal",     // Etiqueta de la fila de subtotal.
  COLUMNA_TOTALES_VALOR: 15,      // Columna O: ahí van las fórmulas de suma.
  ANCHO_TABLA: 16                 // Columnas A..P.
};

/** Catálogo de formatos disponibles. El orden es el que ve el asesor. */
const QUOTE_FORMATS = [
  {
    id: "actual",
    name: "Actual",
    description: "El formato que ya usa el sistema. Incluye fotos de los productos y el detalle de descuentos."
  },
  {
    id: "ccl_liverpool",
    name: "CCL Liverpool",
    description: "Formato oficial del Centro de Contacto Liverpool, generado desde la plantilla de Google Sheets."
  }
];

const DEFAULT_FORMAT_ID = "actual";
const FORMATS_PROP_KEY = "formatos_habilitados";

// =================================================================================================
// HABILITAR / DESHABILITAR FORMATOS
// =================================================================================================

/**
 * Lee del almacén de propiedades qué formatos están habilitados.
 * Un formato sin registro se considera habilitado, para que al agregar formatos nuevos
 * no queden invisibles por omisión.
 * @return {object} Mapa {formatId: boolean}.
 */
function readFormatFlags_() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(FORMATS_PROP_KEY);
    const stored = raw ? JSON.parse(raw) : {};
    const flags = {};
    QUOTE_FORMATS.forEach(f => {
      flags[f.id] = stored[f.id] === undefined ? true : stored[f.id] === true;
    });
    return flags;
  } catch (error) {
    Logger.log("Error en readFormatFlags_: " + error.message);
    const flags = {};
    QUOTE_FORMATS.forEach(f => flags[f.id] = true);
    return flags;
  }
}

/**
 * Indica si el correo pertenece a un usuario con permisos avanzados,
 * consultando la columna 'Avanzado' de la hoja de registros.
 * @param {string} email - Correo del usuario a verificar.
 * @return {boolean}
 */
function isAdvancedUser(email) {
  try {
    if (!email) return false;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REGISTROS_SHEET_NAME);
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();
    const headers = data.shift() || [];
    const emailIdx = headers.indexOf("Email");
    const avanzadoIdx = headers.indexOf("Avanzado");
    if (emailIdx === -1 || avanzadoIdx === -1) return false;

    const target = String(email).trim().toLowerCase();
    const row = data.find(r => String(r[emailIdx] || "").trim().toLowerCase() === target);
    return !!row && String(row[avanzadoIdx]).trim() === "Si";
  } catch (error) {
    Logger.log("Error en isAdvancedUser: " + error.message);
    return false;
  }
}

/**
 * Devuelve el catálogo completo de formatos con su estado. Para el panel de administración.
 * @param {string} email - Correo del usuario que consulta.
 * @return {object} { success, formats: [{id, name, description, enabled, available, unavailableReason}] }
 */
function getFormatSettings(email) {
  try {
    if (!isAdvancedUser(email)) {
      return { success: false, message: "No tienes permisos para administrar los formatos." };
    }

    const flags = readFormatFlags_();
    const formats = QUOTE_FORMATS.map(f => {
      const availability = checkFormatAvailability_(f.id);
      return {
        id: f.id,
        name: f.name,
        description: f.description,
        enabled: flags[f.id],
        available: availability.available,
        unavailableReason: availability.reason
      };
    });
    return { success: true, formats: formats };
  } catch (error) {
    Logger.log("Error en getFormatSettings: " + error.message);
    return { success: false, message: "Error al leer los formatos: " + error.message };
  }
}

/**
 * Habilita o deshabilita un formato. Solo para usuarios avanzados.
 * No permite dejar cero formatos habilitados, porque dejaría al sistema sin forma de imprimir.
 * @param {string} email - Correo del usuario que hace el cambio.
 * @param {string} formatId - Formato a modificar.
 * @param {boolean} enabled - Nuevo estado.
 * @return {object} { success, formats } o { success: false, message }
 */
function setQuoteFormatEnabled(email, formatId, enabled) {
  try {
    if (!isAdvancedUser(email)) {
      return { success: false, message: "No tienes permisos para administrar los formatos." };
    }
    if (!QUOTE_FORMATS.some(f => f.id === formatId)) {
      return { success: false, message: `El formato '${formatId}' no existe.` };
    }

    const flags = readFormatFlags_();
    flags[formatId] = enabled === true;

    if (!Object.keys(flags).some(id => flags[id])) {
      return { success: false, message: "Debe quedar al menos un formato habilitado." };
    }

    PropertiesService.getScriptProperties().setProperty(FORMATS_PROP_KEY, JSON.stringify(flags));
    Logger.log(`Formato ${formatId} ${enabled ? "habilitado" : "deshabilitado"} por ${email}.`);

    return getFormatSettings(email);
  } catch (error) {
    Logger.log("Error en setQuoteFormatEnabled: " + error.message);
    return { success: false, message: "Error al guardar el formato: " + error.message };
  }
}

/**
 * Verifica si un formato se puede usar realmente (no solo si está habilitado).
 * @param {string} formatId
 * @return {object} { available: boolean, reason: string }
 */
function checkFormatAvailability_(formatId) {
  if (formatId !== "ccl_liverpool") return { available: true, reason: "" };

  if (!CCL_TEMPLATE_SHEET_ID) {
    return { available: false, reason: "Falta configurar CCL_TEMPLATE_SHEET_ID en Formatos.gs." };
  }
  try {
    const ss = SpreadsheetApp.openById(CCL_TEMPLATE_SHEET_ID);
    if (!ss.getSheetByName(CCL_SHEET_NAME)) {
      return { available: false, reason: `La plantilla no tiene una pestaña llamada '${CCL_SHEET_NAME}'.` };
    }
    return { available: true, reason: "" };
  } catch (error) {
    return { available: false, reason: "No se pudo abrir la plantilla CCL: " + error.message };
  }
}

/**
 * Devuelve los formatos que el asesor puede elegir al cotizar: habilitados y utilizables.
 * @return {object} { success, formats: [{id, name, description}], defaultId }
 */
function getEnabledQuoteFormats() {
  try {
    const flags = readFormatFlags_();
    const formats = QUOTE_FORMATS
      .filter(f => {
        if (!flags[f.id]) {
          Logger.log(`getEnabledQuoteFormats: '${f.id}' está deshabilitado en el panel de administración.`);
          return false;
        }
        const availability = checkFormatAvailability_(f.id);
        if (!availability.available) {
          Logger.log(`getEnabledQuoteFormats: '${f.id}' no está disponible. Motivo: ${availability.reason}`);
          return false;
        }
        return true;
      })
      .map(f => ({ id: f.id, name: f.name, description: f.description }));

    const defaultId = formats.some(f => f.id === DEFAULT_FORMAT_ID)
      ? DEFAULT_FORMAT_ID
      : (formats.length > 0 ? formats[0].id : null);

    return { success: true, formats: formats, defaultId: defaultId };
  } catch (error) {
    Logger.log("Error en getEnabledQuoteFormats: " + error.message);
    return { success: false, message: "Error al obtener los formatos: " + error.message, formats: [] };
  }
}

// =================================================================================================
// GENERACIÓN DEL PDF
// =================================================================================================

/**
 * Genera el PDF de una cotización en el formato indicado.
 * @param {string} folio - Folio de la cotización.
 * @param {string} formatId - Formato a usar. Si se omite, se usa el guardado en la cotización.
 * @return {Blob} El PDF listo para adjuntar o descargar.
 */
function generateQuotePdfBlob(folio, formatId) {
  if (!folio) throw new Error("El folio es requerido para generar el PDF.");

  let format = formatId;
  if (!format) {
    const stored = getQuoteDetails(folio);
    format = (stored.success && stored.quote.format) ? stored.quote.format : DEFAULT_FORMAT_ID;
    Logger.log(`generateQuotePdfBlob: no llegó formato, se usa el guardado '${format}' (folio ${folio}).`);
  }

  // Se valida contra el catálogo en vez de caer al formato por omisión: un valor
  // inesperado significa que el cliente mandó basura, y hay que verlo, no taparlo.
  if (!QUOTE_FORMATS.some(f => f.id === format)) {
    throw new Error(`Formato desconocido: '${format}'. Los válidos son: ${QUOTE_FORMATS.map(f => f.id).join(", ")}.`);
  }

  Logger.log(`generateQuotePdfBlob: generando folio ${folio} con formato '${format}'.`);

  if (format === "ccl_liverpool") return generateCclPdfBlob_(folio);

  // Formato 'actual': el PDF que ya se armaba desde HTML.
  const htmlResponse = generateQuoteHtml(folio);
  if (!htmlResponse.success) throw new Error(htmlResponse.message);

  return Utilities.newBlob(htmlResponse.html, "text/html", `Cotizacion_${folio}.html`)
                  .getAs("application/pdf")
                  .setName(`Cotizacion_${folio}.pdf`);
}

/**
 * Genera el PDF del formato CCL: copia la plantilla, la llena y exporta la copia.
 * La copia se borra siempre, incluso si la exportación falla.
 * @param {string} folio - Folio de la cotización.
 * @return {Blob} El PDF del formato CCL.
 */
function generateCclPdfBlob_(folio) {
  const availability = checkFormatAvailability_("ccl_liverpool");
  if (!availability.available) throw new Error(availability.reason);

  const quoteResponse = getQuoteDetails(folio);
  if (!quoteResponse.success) throw new Error(quoteResponse.message);
  const quote = quoteResponse.quote;

  let copyFile = null;
  try {
    copyFile = DriveApp.getFileById(CCL_TEMPLATE_SHEET_ID).makeCopy(`TEMP_Cotizacion_${folio}`);
    const copySs = SpreadsheetApp.openById(copyFile.getId());
    const sheet = copySs.getSheetByName(CCL_SHEET_NAME);
    if (!sheet) throw new Error(`La copia no tiene la pestaña '${CCL_SHEET_NAME}'.`);

    fillCclSheet_(sheet, quote);
    SpreadsheetApp.flush();

    const pdfBlob = exportSheetAsPdf_(copyFile.getId(), sheet.getSheetId());
    return pdfBlob.setName(`Cotizacion_${folio}.pdf`);
  } finally {
    if (copyFile) {
      try {
        copyFile.setTrashed(true);
      } catch (cleanupError) {
        Logger.log(`No se pudo borrar la copia temporal del folio ${folio}: ${cleanupError.message}`);
      }
    }
  }
}

/**
 * Escribe los datos de una cotización dentro de una hoja con el formato CCL,
 * ajustando dinámicamente el número de filas de producto.
 * @param {Sheet} sheet - La pestaña a llenar (debe ser una copia, nunca la plantilla).
 * @param {object} quote - La cotización tal como la devuelve getQuoteDetails.
 */
function fillCclSheet_(sheet, quote) {
  const products = quote.products || [];
  if (products.length === 0) throw new Error("La cotización no tiene productos.");

  // --- 1. Datos generales ---
  sheet.getRange(CCL_CONFIG.CELDA_ASESOR).setValue(quote.advisorName || "");
  sheet.getRange(CCL_CONFIG.CELDA_CLIENTE).setValue("Dirigida a: " + (quote.clientName || "Cliente"));
  sheet.getRange(CCL_CONFIG.CELDA_CORREO).setValue("Correo: " + (quote.clientEmail || ""));
  sheet.getRange(CCL_CONFIG.CELDA_TELEFONO).setValue("Teléfono: " + (quote.clientPhone || ""));
  sheet.getRange(CCL_CONFIG.CELDA_OBSERVACIONES).setValue("Observación: " + (quote.observations || ""));

  // --- 2. Localizar la tabla ---
  const lastRow = sheet.getLastRow();
  const columnA = sheet.getRange(1, 1, Math.min(100, lastRow || 100), 1).getValues();

  let headerRow = -1;
  for (let r = 0; r < columnA.length; r++) {
    if (String(columnA[r][0]).trim().toLowerCase() === CCL_CONFIG.TEXTO_CABECERA) {
      headerRow = r + 1;
      break;
    }
  }
  if (headerRow === -1) {
    throw new Error(`No se encontró la cabecera '${CCL_CONFIG.TEXTO_CABECERA}' en la columna A de la plantilla.`);
  }

  let infoRow = -1;
  for (let r = headerRow; r < columnA.length; r++) {
    if (String(columnA[r][0]).replace(/\s+/g, "").toLowerCase().indexOf(CCL_CONFIG.TEXTO_FIN_TABLA) !== -1) {
      infoRow = r + 1;
      break;
    }
  }
  if (infoRow === -1) {
    throw new Error(`No se encontró el fin de tabla '${CCL_CONFIG.TEXTO_FIN_TABLA}' en la columna A de la plantilla.`);
  }

  const firstProductRow = headerRow + 1;
  const lastProductRow = infoRow - 2; // Hay una fila en blanco separadora antes del bloque final.
  const currentNumRows = lastProductRow - firstProductRow + 1;
  const targetNumRows = products.length;

  // --- 3. Ajustar el número de filas al de productos ---
  const templateRange = sheet.getRange(firstProductRow, 1, 1, CCL_CONFIG.ANCHO_TABLA);

  if (targetNumRows > currentNumRows) {
    sheet.insertRowsAfter(lastProductRow, targetNumRows - currentNumRows);
    for (let i = 1; i < targetNumRows; i++) {
      templateRange.copyTo(sheet.getRange(firstProductRow + i, 1, 1, CCL_CONFIG.ANCHO_TABLA),
                           { format: true, formulas: true });
    }
    infoRow += targetNumRows - currentNumRows;
  } else if (targetNumRows < currentNumRows) {
    sheet.deleteRows(firstProductRow + targetNumRows, currentNumRows - targetNumRows);
    infoRow -= currentNumRows - targetNumRows;
  }

  // --- 4. Escribir los productos ---
  const matrix = products.map((p, i) => buildCclProductRow_(p, firstProductRow + i));
  sheet.getRange(firstProductRow, 1, targetNumRows, CCL_CONFIG.ANCHO_TABLA).setValues(matrix);

  // --- 5. Recalcular los totales ---
  const endProductRow = firstProductRow + targetNumRows - 1;
  const subtotalRow = findCclSubtotalRow_(sheet, infoRow);

  sheet.getRange(subtotalRow, CCL_CONFIG.COLUMNA_TOTALES_VALOR)
       .setFormula(`=SUM(L${firstProductRow}:L${endProductRow})`);
  sheet.getRange(subtotalRow + 1, CCL_CONFIG.COLUMNA_TOTALES_VALOR)
       .setFormula(`=SUM(M${firstProductRow}:M${endProductRow})`);
  sheet.getRange(subtotalRow + 2, CCL_CONFIG.COLUMNA_TOTALES_VALOR)
       .setFormula(`=SUM(P${firstProductRow}:P${endProductRow})`);
}

/**
 * Traduce un producto del sistema a la fila de 16 columnas (A..P) del formato CCL.
 *
 * El sistema guarda los porcentajes de 0 a 100 y el CCL los espera como fracción de 0 a 1.
 * Cuando el producto tiene 'costPaymentUnique' (precio de pago único, que en el formato
 * actual pisa cualquier descuento) el CCL no tiene una celda equivalente, así que se
 * convierte a la fracción de descuento que produce ese mismo importe final.
 *
 * @param {object} p - Producto tal como lo guarda el sistema.
 * @param {number} R - Número de fila en la hoja, para construir las fórmulas.
 * @return {Array} Fila de 16 posiciones lista para setValues.
 */
function buildCclProductRow_(p, R) {
  const quantity = parseInt(p.quantity) || 0;
  const unitPrice = parseFloat(p.unitPrice) || 0;
  const costPaymentUnique = parseFloat(p.costPaymentUnique) || 0;
  const priceVolume = unitPrice * quantity;

  let discountFraction = (parseFloat(p.discountPublicPercent) || 0) / 100;
  let additionalApplied = p.additionalDiscountApplied === "Si" ? "Si" : "No";
  let additionalFraction = (parseFloat(p.additionalDiscountPercent) || 0) / 100;

  if (costPaymentUnique > 0 && priceVolume > 0) {
    if (additionalApplied === "Si") {
      // El asesor marcó descuento ADICIONAL (p.ej. producto de marketplace): el pago
      // único se traduce a la columna adicional (O), no a la pública (J), para que el
      // documento conserve "Aplica descuento adicional = Si" con su porcentaje.
      const priceWithPublic = priceVolume * (1 - discountFraction);
      additionalFraction = priceWithPublic > 0
        ? Math.min(1, Math.max(0, 1 - (costPaymentUnique / priceWithPublic)))
        : 0;
    } else {
      discountFraction = Math.max(0, 1 - (costPaymentUnique / priceVolume));
      additionalFraction = 0;
    }
  }

  const row = new Array(CCL_CONFIG.ANCHO_TABLA).fill("");
  row[0]  = p.sku || "";                                          // A: Sku
  row[1]  = quantity;                                             // B: Cantidad
  row[2]  = p.description || "";                                  // C: Descripción (combinada C:G)
  row[7]  = unitPrice;                                            // H: Precio Unitario
  row[8]  = `=H${R}*B${R}`;                                       // I: Precio por volumen
  row[9]  = discountFraction;                                     // J: Descuento
  row[10] = `=I${R}-(I${R}*J${R})`;                               // K: Precio con descuento
  row[11] = `=IF(UPPER(N${R})="NO",K${R}/1.16,P${R}/1.16)`;       // L: Subtotal sin IVA
  row[12] = `=IF(UPPER(N${R})="NO",K${R}-L${R},P${R}-L${R})`;     // M: IVA
  row[13] = additionalApplied;                                    // N: Aplica descuento adicional
  row[14] = additionalFraction;                                   // O: % Descuento adicional
  row[15] = `=K${R}-(O${R}*K${R})`;                               // P: Total
  return row;
}

/**
 * Localiza la fila donde va el subtotal, buscando su etiqueta debajo del fin de la tabla.
 * @param {Sheet} sheet - La hoja con el formato.
 * @param {number} infoRow - Fila donde arranca el bloque de información adicional.
 * @return {number} Número de fila del subtotal.
 */
function findCclSubtotalRow_(sheet, infoRow) {
  const lastSearchRow = Math.min(sheet.getMaxRows(), infoRow + 14);
  const block = sheet.getRange(infoRow, 1, lastSearchRow - infoRow + 1, 12).getValues();

  for (let r = 0; r < block.length; r++) {
    for (let c = 0; c < block[r].length; c++) {
      if (String(block[r][c]).replace(/\s+/g, "").toLowerCase() === CCL_CONFIG.TEXTO_SUBTOTAL) {
        return infoRow + r;
      }
    }
  }
  return infoRow + 2; // Posición por omisión si la plantilla no trae la etiqueta.
}

/**
 * Exporta una pestaña concreta de una hoja de cálculo como PDF.
 * @param {string} spreadsheetId - ID del archivo a exportar.
 * @param {number} gid - ID de la pestaña dentro del archivo.
 * @return {Blob} El PDF resultante.
 */
function exportSheetAsPdf_(spreadsheetId, gid) {
  const params = Object.keys(CCL_EXPORT_OPTIONS)
    .map(key => `${key}=${encodeURIComponent(CCL_EXPORT_OPTIONS[key])}`)
    .join("&");
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf&gid=${gid}&${params}`;

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Google rechazó la exportación a PDF (código ${response.getResponseCode()}).`);
  }
  return response.getBlob();
}

// =================================================================================================
// DOCUMENTO DE SHEETS PERSISTENTE
// =================================================================================================

/**
 * Devuelve la carpeta donde se guardan las hojas CCL generadas, creándola si hace falta.
 * @return {Folder}
 */
function getCclFolder_() {
  const existentes = DriveApp.getFoldersByName(CCL_OUTPUT_FOLDER_NAME);
  return existentes.hasNext() ? existentes.next() : DriveApp.createFolder(CCL_OUTPUT_FOLDER_NAME);
}

/**
 * Lee el valor de una columna de la hoja 'Cotizaciones' para un folio dado.
 * @param {string} folio
 * @param {string} columnName
 * @return {string} El valor, o cadena vacía si la columna o la fila no existen.
 */
function getQuoteColumnValue_(folio, columnName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COTIZACIONES_SHEET_NAME);
  if (!sheet) return "";

  const data = sheet.getDataRange().getValues();
  const headers = data.shift() || [];
  const folioIdx = headers.indexOf("Folio");
  const colIdx = headers.indexOf(columnName);
  if (folioIdx === -1 || colIdx === -1) return "";

  const row = data.find(r => r[folioIdx] == folio);
  return row ? String(row[colIdx] || "") : "";
}

/**
 * Escribe el valor de una columna de la hoja 'Cotizaciones' para un folio dado,
 * creando la columna si no existe (mismo patrón auto-reparable que 'Formato').
 * @param {string} folio
 * @param {string} columnName
 * @param {string} value
 */
function setQuoteColumnValue_(folio, columnName, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COTIZACIONES_SHEET_NAME);
  if (!sheet) throw new Error(`Hoja "${COTIZACIONES_SHEET_NAME}" no encontrada.`);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const folioIdx = headers.indexOf("Folio");
  if (folioIdx === -1) throw new Error("Columna 'Folio' no encontrada en 'Cotizaciones'.");

  let colIdx = headers.indexOf(columnName);
  if (colIdx === -1) {
    sheet.getRange(1, headers.length + 1).setValue(columnName);
    colIdx = headers.length;
    Logger.log(`Nueva columna '${columnName}' agregada de forma auto-reparable en Cotizaciones.`);
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][folioIdx] == folio) {
      sheet.getRange(i + 1, colIdx + 1).setValue(value);
      return;
    }
  }
  Logger.log(`setQuoteColumnValue_: no se encontró el folio ${folio} para escribir '${columnName}'.`);
}

/**
 * Crea o refresca el documento de Google Sheets con el formato CCL de una cotización
 * y devuelve su URL para abrirlo.
 *
 * El documento es un *reflejo* de los datos del sistema, no una fuente de verdad: cada vez
 * que se abre se vuelve a llenar desde la cotización, así que los cambios hechos a mano
 * dentro de la hoja se pierden. Se reutiliza el mismo archivo entre llamadas para no
 * llenar Drive de copias.
 *
 * @param {string} folio - Folio de la cotización.
 * @return {object} { success, url } o { success: false, message }
 */
function openQuoteInSheets(folio) {
  try {
    if (!folio) throw new Error("El folio es requerido.");

    const availability = checkFormatAvailability_("ccl_liverpool");
    if (!availability.available) throw new Error(availability.reason);

    const quoteResponse = getQuoteDetails(folio);
    if (!quoteResponse.success) throw new Error(quoteResponse.message);

    // Si ya se generó antes, se reutiliza el archivo; si lo borraron, se crea de nuevo.
    let spreadsheet = null;
    const storedUrl = getQuoteColumnValue_(folio, CCL_LINK_COLUMN);
    const storedId = storedUrl ? (storedUrl.match(/\/d\/([a-zA-Z0-9-_]+)/) || [])[1] : null;

    if (storedId) {
      try {
        const archivo = DriveApp.getFileById(storedId);
        if (!archivo.isTrashed()) spreadsheet = SpreadsheetApp.openById(storedId);
      } catch (e) {
        Logger.log(`openQuoteInSheets: el documento guardado del folio ${folio} ya no existe. Se crea uno nuevo.`);
      }
    }

    if (!spreadsheet) {
      const copia = DriveApp.getFileById(CCL_TEMPLATE_SHEET_ID)
                            .makeCopy(`Cotizacion_${folio}`, getCclFolder_());
      spreadsheet = SpreadsheetApp.openById(copia.getId());
    }

    const hoja = spreadsheet.getSheetByName(CCL_SHEET_NAME);
    if (!hoja) throw new Error(`El documento no tiene la pestaña '${CCL_SHEET_NAME}'.`);

    fillCclSheet_(hoja, quoteResponse.quote);
    SpreadsheetApp.flush();

    const url = spreadsheet.getUrl();
    setQuoteColumnValue_(folio, CCL_LINK_COLUMN, url);
    Logger.log(`openQuoteInSheets: documento listo para el folio ${folio}: ${url}`);

    return { success: true, url: url };
  } catch (error) {
    Logger.log(`Error en openQuoteInSheets (folio ${folio}): ${error.message}`);
    return { success: false, message: "No se pudo abrir el documento en Sheets: " + error.message };
  }
}

// =================================================================================================
// DESCARGA DESDE EL CLIENTE
// =================================================================================================

/**
 * Función de diagnóstico: se ejecuta a mano desde el editor de Apps Script.
 * Sirve para dos cosas:
 *   1. Disparar la pantalla de autorización cuando se agregan permisos nuevos.
 *   2. Confirmar que la plantilla CCL es accesible antes de usar la app.
 * El resultado se ve en Ver > Registros (Ctrl+Enter).
 */
function probarAccesoCcl() {
  Logger.log("Plantilla configurada: " + (CCL_TEMPLATE_SHEET_ID || "(vacía)"));

  const archivo = DriveApp.getFileById(CCL_TEMPLATE_SHEET_ID);
  Logger.log("Lectura de Drive OK. Archivo: " + archivo.getName());

  const hoja = SpreadsheetApp.openById(CCL_TEMPLATE_SHEET_ID).getSheetByName(CCL_SHEET_NAME);
  Logger.log(hoja
    ? `Hoja '${CCL_SHEET_NAME}' encontrada (gid ${hoja.getSheetId()}).`
    : `ERROR: no existe una pestaña llamada '${CCL_SHEET_NAME}'.`);

  // Copiar y borrar es lo que realmente hace la app al generar el PDF, y es lo que exige
  // el permiso de escritura en Drive. Leer la plantilla no lo prueba: hay que copiarla.
  let copia = null;
  try {
    copia = archivo.makeCopy("TEMP_prueba_permisos");
    Logger.log("Escritura en Drive OK. Copia creada: " + copia.getId());
  } finally {
    if (copia) {
      copia.setTrashed(true);
      Logger.log("Copia temporal borrada correctamente.");
    }
  }

  Logger.log("Listo: el formato CCL ya puede generarse.");
  return true;
}

/**
 * Genera el PDF y lo devuelve codificado para que el navegador lo descargue.
 * google.script.run no puede transportar Blobs, por eso se manda en base64.
 * @param {string} folio - Folio de la cotización.
 * @param {string} formatId - Formato a usar.
 * @return {object} { success, fileName, mimeType, base64 } o { success: false, message }
 */
function downloadQuotePdf(folio, formatId) {
  try {
    const blob = generateQuotePdfBlob(folio, formatId);
    return {
      success: true,
      fileName: blob.getName(),
      mimeType: "application/pdf",
      base64: Utilities.base64Encode(blob.getBytes())
    };
  } catch (error) {
    Logger.log(`Error en downloadQuotePdf (folio ${folio}, formato ${formatId}): ${error.message}`);
    return { success: false, message: "No se pudo generar el PDF: " + error.message };
  }
}
