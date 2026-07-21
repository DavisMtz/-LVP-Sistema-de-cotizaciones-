/**
 * ===============================================================================
 * NUEVAS/MODIFICADAS FUNCIONES PARA ENVÍO DE CORREO HTML
 * ===============================================================================
 */

// Alias institucional desde el que salen las cotizaciones. Debe estar dado de alta
// como "Enviar como" en la cuenta de Gmail que ejecuta el script; si no lo está,
// el envío cae de vuelta a la cuenta propia (ver sendQuoteByEmail).
const MAIL_ALIAS = 'cotizacion@liverpool.com.mx';

/**
 * Indica al cliente desde qué remitente saldrán los correos, para mostrarlo en la
 * pantalla de composición antes de enviar.
 */
function getMailSenderInfo() {
  try {
    const aliasAvailable = GmailApp.getAliases().indexOf(MAIL_ALIAS) !== -1;
    return {
      success: true,
      alias: MAIL_ALIAS,
      aliasAvailable: aliasAvailable,
      effectiveSender: aliasAvailable ? MAIL_ALIAS : (Session.getActiveUser() ? Session.getActiveUser().getEmail() : '')
    };
  } catch (e) {
    return { success: false, alias: MAIL_ALIAS, aliasAvailable: false, message: e.message };
  }
}

/**
 * Formatea un número como moneda MXN.
 * Función de utilidad para ser usada dentro de Apps Script.
 * @param {number} amount - La cantidad a formatear.
 * @return {string} La cantidad formateada como string (ej. $1,234.50).
 */
function formatCurrencyGS(amount) {
  if (isNaN(parseFloat(amount))) return "$0.00";
  return parseFloat(amount).toLocaleString('es-MX', { style: 'currency', currency: 'MXN' });
}


/**
 * Genera el cuerpo HTML completo de una cotización para ser incrustado en un correo.
 * @param {string} folio - El folio de la cotización.
 * @return {object} - Objeto con { success: true, html: '...' } o { success: false, message: '...' }.
 */
function generateQuoteHtml(folio) {
  try {
    const quoteResponse = getQuoteDetails(folio);
    if (!quoteResponse.success) {
      return { success: false, message: "No se pudieron obtener los detalles de la cotización para generar el HTML." };
    }
    const data = quoteResponse.quote;

    let productsHtml = '';
    if (data.products && data.products.length > 0) {
      data.products.forEach(p => {
        const unitPrice = parseFloat(p.unitPrice) || 0;
        const quantity = parseInt(p.quantity) || 0;
        const priceVolume = unitPrice * quantity;
        let finalPricePerLine;
        let discountDisplayString = "-";
        const costPaymentUnique = parseFloat(p.costPaymentUnique) || 0;
        let discountPublicPercent = parseFloat(p.discountPublicPercent) || 0;
        const additionalDiscountApplied = p.additionalDiscountApplied === 'Si';
        let additionalDiscountPercent = parseFloat(p.additionalDiscountPercent) || 0;

        if (costPaymentUnique > 0 && quantity > 0 && unitPrice > 0) {
            finalPricePerLine = costPaymentUnique;
        } else {
            let priceAfterPublic = priceVolume * (1 - (discountPublicPercent / 100));
            priceAfterPublic = Math.max(0, priceAfterPublic);
            finalPricePerLine = priceAfterPublic;
            if (additionalDiscountApplied && additionalDiscountPercent > 0) {
                finalPricePerLine = priceAfterPublic * (1 - (additionalDiscountPercent / 100));
            }
            finalPricePerLine = Math.max(0, finalPricePerLine);
        }
        const totalMonetaryDiscount = priceVolume - finalPricePerLine;
        const effectiveTotalPercentage = priceVolume > 0 ? (totalMonetaryDiscount / priceVolume) * 100 : 0;

        if (totalMonetaryDiscount > 0.001) {
            let details = [];
            if (discountPublicPercent > 0.001 && !(costPaymentUnique > 0)) {
               details.push(`Púb: ${discountPublicPercent.toFixed(2)}%`);
            }
            if (additionalDiscountApplied && additionalDiscountPercent > 0.001 && !(costPaymentUnique > 0)) {
               details.push(`Adic: ${additionalDiscountPercent.toFixed(2)}%`);
            }
            if (details.length > 0) {
               discountDisplayString = `${details.join(' + ')}. Total: ${formatCurrencyGS(totalMonetaryDiscount)} (${effectiveTotalPercentage.toFixed(2)}% DesTot.)`;
            } else if (costPaymentUnique > 0) {
               discountDisplayString = `${formatCurrencyGS(totalMonetaryDiscount)} (${effectiveTotalPercentage.toFixed(2)}% DesTot.)`;
            } else {
               discountDisplayString = "-";
            }
        }
        
        productsHtml += `
          <tr>
            <td style="border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; vertical-align: top; white-space: nowrap;">${p.sku || ''}</td>
            <td style="border: 1px solid #cbd5e1; padding: 6px 8px; text-align: center; vertical-align: top; white-space: nowrap;">${quantity}</td>
            <td style="border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; vertical-align: top; word-break: break-word; line-height: 1.3;">${p.description || ''}</td>
            <td style="border: 1px solid #cbd5e1; padding: 6px 8px; text-align: right; vertical-align: top; white-space: nowrap;">${formatCurrencyGS(unitPrice)}</td>
            <td style="border: 1px solid #cbd5e1; padding: 6px 8px; text-align: right; vertical-align: top; white-space: nowrap;">${formatCurrencyGS(priceVolume)}</td>
            <td style="border: 1px solid #cbd5e1; padding: 6px 8px; text-align: right; vertical-align: top; word-break: break-word; font-size: 8pt; line-height: 1.2;">${discountDisplayString}</td>
            <td style="border: 1px solid #cbd5e1; padding: 6px 8px; text-align: right; vertical-align: top; white-space: nowrap; font-weight: 500;">${formatCurrencyGS(finalPricePerLine)}</td>
          </tr>
        `;
      });
    } else {
      productsHtml = '<tr><td colspan="7" style="text-align: center; padding: 1rem;">No hay productos en esta cotización.</td></tr>';
    }

    const observationsHtml = (data.observations && data.observations.trim() !== '') ? `
      <div style="margin-top: 15px; margin-bottom: 20px; font-size: 10pt; padding: 10px; background-color: #fdfdfd; border-radius: 4px; border: 1px solid #f0f0f0;">
        <h2 style="font-size: 14px; font-weight: 700; color: #E10098; margin-top:0; margin-bottom: 8px; border-bottom: 1px solid #eeeeee; padding-bottom: 4px;">Observaciones Adicionales</h2>
        <p style="white-space: pre-wrap; margin: 0; line-height: 1.5;">${data.observations}</p>
      </div>` : '';

    // HTML optimizado para renderizado PDF a página completa (Letter)
    const fullHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <style>
          @page {
            size: letter;
            margin: 12mm 15mm 12mm 15mm;
          }
          body {
            font-family: Arial, sans-serif;
            color: #333333;
            margin: 0;
            padding: 0;
            background-color: #ffffff;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
          }
          .pdf-container {
            width: 100%;
            margin: 0;
            padding: 0;
            box-sizing: border-box;
          }
        </style>
      </head>
      <body>
        <div class="pdf-container">
          <table width="100%" cellspacing="0" cellpadding="0" style="border-bottom: 2px solid #E10098; padding-bottom: 15px; margin-bottom: 25px;">
            <tr>
              <td valign="top">
                <h1 style="font-size: 24px; font-weight: 700; color: #E10098; margin: 0 0 5px 0;">COTIZACIÓN</h1>
                <p style="font-size: 11px; margin: 2px 0; color: #4A4A4A;"><strong>Folio:</strong> ${data.folio || 'N/A'}</p>
                <p style="font-size: 11px; margin: 2px 0; color: #4A4A4A;"><strong>Fecha de Emisión:</strong> ${data.timestamp ? new Date(data.timestamp).toLocaleDateString('es-MX', { day: '2-digit', month: 'long', year: 'numeric' }) : 'N/A'}</p>
              </td>
              <td valign="top" align="right">
                <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/3/35/Liverpool_logo.svg/1280px-Liverpool_logo.svg.png" alt="Logo Liverpool" style="max-height: 42px; width: auto; margin-bottom: 8px;">
                <p style="font-size: 10px; margin: 2px 0; color: #666666; text-align: right;">Centro de Contacto Liverpool</p>
                <p style="font-size: 10px; margin: 2px 0; color: #666666; text-align: right;">postventaomnicanal@liverpool.com.mx</p>
                <p style="font-size: 10px; margin: 2px 0; color: #666666; text-align: right;">Tel: 55 5262 9999, opción 3 (ext: ${data.advisorExt || 'N/A'})</p>
              </td>
            </tr>
          </table>
          <table width="100%" cellspacing="0" cellpadding="0" style="margin-bottom: 20px; font-size: 11px;">
            <tr>
              <td width="48%" valign="top" style="background-color: #fdfdfd; padding: 10px; border-radius: 4px; border: 1px solid #f0f0f0;">
                <h2 style="margin-top:0; font-size: 14px; font-weight: 700; color: #E10098; margin-bottom: 8px; border-bottom: 1px solid #eeeeee; padding-bottom: 4px;">Información del Asesor</h2>
                <p style="margin: 2px 0; line-height: 1.5;"><strong>Nombre:</strong> ${data.advisorName || 'N/A'}</p>
                <p style="margin: 2px 0; line-height: 1.5;"><strong>Puesto:</strong> Asesor de Ventas</p>
                <p style="margin: 2px 0; line-height: 1.5;"><strong>Extensión:</strong> ${data.advisorExt || 'N/A'}</p>
              </td>
              <td width="4%">&nbsp;</td>
              <td width="48%" valign="top" style="background-color: #fdfdfd; padding: 10px; border-radius: 4px; border: 1px solid #f0f0f0;">
                <h2 style="margin-top:0; font-size: 14px; font-weight: 700; color: #E10098; margin-bottom: 8px; border-bottom: 1px solid #eeeeee; padding-bottom: 4px;">Información del Cliente</h2>
                <p style="margin: 2px 0; line-height: 1.5;"><strong>Dirigida a:</strong> ${data.clientName || 'N/A'}</p>
                <p style="margin: 2px 0; line-height: 1.5;"><strong>Correo:</strong> ${data.clientEmail || 'N/A'}</p>
                <p style="margin: 2px 0; line-height: 1.5;"><strong>Teléfono:</strong> ${data.clientPhone || 'N/A'}</p>
              </td>
            </tr>
          </table>
          <h2 style="font-size: 14px; font-weight: 700; color: #E10098; margin-top: 20px; margin-bottom: 8px; border-bottom: 1px solid #eeeeee; padding-bottom: 4px;">Detalle de Productos</h2>
          <table width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin-bottom: 20px; font-size: 9pt;">
            <thead style="background-color: #f5f5f5; font-weight: 700;">
              <tr>
                <th style="width: 14%; border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; white-space: nowrap;">SKU</th>
                <th style="width: 7%; text-align: center; border: 1px solid #cbd5e1; padding: 6px 8px; white-space: nowrap;">Cant.</th>
                <th style="width: 26%; border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; word-break: break-word;">Descripción</th>
                <th style="width: 12%; text-align: right; border: 1px solid #cbd5e1; padding: 6px 8px; white-space: nowrap;">P. Unitario</th>
                <th style="width: 13%; text-align: right; border: 1px solid #cbd5e1; padding: 6px 8px; white-space: nowrap;">P. x Volumen</th>
                <th style="width: 14%; text-align: right; border: 1px solid #cbd5e1; padding: 6px 8px; word-break: break-word;">Desc. Aplicado</th>
                <th style="width: 14%; text-align: right; border: 1px solid #cbd5e1; padding: 6px 8px; white-space: nowrap;">Total Fila</th>
              </tr>
            </thead>
            <tbody>${productsHtml}</tbody>
          </table>
          <table width="100%" cellspacing="0" cellpadding="0"><tr><td align="right">
            <table style="width: 45%; font-size: 10pt;">
              <tr><td style="padding: 6px 8px; border-bottom: 1px solid #eeeeee;">SUBTOTAL:</td><td style="padding: 6px 8px; border-bottom: 1px solid #eeeeee; text-align: right; font-weight: 500;">${formatCurrencyGS(data.summarySubtotal)}</td></tr>
              <tr><td style="padding: 6px 8px; border-bottom: 1px solid #eeeeee;">IVA (16%):</td><td style="padding: 6px 8px; border-bottom: 1px solid #eeeeee; text-align: right; font-weight: 500;">${formatCurrencyGS(data.summaryVat)}</td></tr>
              <tr style="font-size: 12pt; font-weight: 700; color: #E10098;"><td style="padding: 8px; border-top: 2px solid #333;">TOTAL A PAGAR:</td><td style="padding: 8px; text-align: right; border-top: 2px solid #333;">${formatCurrencyGS(data.summaryTotal)}</td></tr>
            </table>
          </td></tr></table>
          ${observationsHtml}
          <div style="font-size: 8pt; color: #555555; margin-top: 20px; line-height: 1.3;">
            <p>Precios y promociones sujetos a cambio sin previo aviso. Los precios incluyen IVA. La disponibilidad de los artículos está sujeta a existencias al momento de realizar la compra.</p>
          </div>
          <div style="font-size: 8pt; color: #888888; text-align: center; border-top: 1px solid #eeeeee; padding-top: 10px; margin-top: 25px;">
            <p>Gracias por su preferencia.<br>Liverpool - Es parte de mi vida</p>
          </div>
        </div>
      </body>
      </html>
    `;
    return { success: true, html: fullHtml };
  } catch (error) {
    Logger.log(`Error en generateQuoteHtml para folio ${folio}: ${error.message}`);
    return { success: false, message: `Error al generar el HTML de la cotización: ${error.message}` };
  }
}

/**
 * Obtiene los detalles básicos de una cotización para rellenar el formulario de correo.
 * @param {string} folio - El folio de la cotización a buscar.
 * @return {object} Un objeto con los datos del cliente.
 */
function getQuoteDetailsForEmail(folio) {
  try {
    if (!folio) throw new Error("El folio es requerido.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cotizacionesSheet = ss.getSheetByName(COTIZACIONES_SHEET_NAME);
    if (!cotizacionesSheet) throw new Error(`Hoja "${COTIZACIONES_SHEET_NAME}" no encontrada.`);

    const cotAllData = cotizacionesSheet.getDataRange().getValues();
    if (cotAllData.length <= 1) return { success: false, message: "No hay cotizaciones en la hoja." };

    const cotHeaders = cotAllData.shift();
    const folioColIdx = cotHeaders.indexOf("Folio");
    const clientNameColIdx = cotHeaders.indexOf("ClienteNombre");
    const clientEmailColIdx = cotHeaders.indexOf("CorreoCliente");
    const formatColIdx = cotHeaders.indexOf("Formato");

    if ([folioColIdx, clientNameColIdx, clientEmailColIdx].includes(-1)) {
        throw new Error("Faltan columnas requeridas en la hoja 'Cotizaciones'. Verifica: Folio, ClienteNombre, CorreoCliente.");
    }

    const quoteRow = cotAllData.find(row => row[folioColIdx] == folio);
    if (!quoteRow) return { success: false, message: `No se encontró la cotización con el folio ${folio}.` };

    const details = {
      folio: quoteRow[folioColIdx],
      clientName: quoteRow[clientNameColIdx],
      clientEmail: quoteRow[clientEmailColIdx],
      // 'Formato' es una columna que se auto-crea, así que puede no existir en hojas viejas.
      format: (formatColIdx > -1 && quoteRow[formatColIdx]) ? quoteRow[formatColIdx] : DEFAULT_FORMAT_ID
    };

    Logger.log(`Detalles para formulario de correo recuperados para folio ${folio}`);
    return { success: true, data: details };

  } catch (error) {
    Logger.log(`Error en getQuoteDetailsForEmail para folio ${folio}: ${error.message}`);
    return { success: false, message: `Error del servidor: ${error.message}` };
  }
}

/**
 * Envía un correo electrónico con la cotización en formato PDF adjunta, y el mensaje del usuario en el cuerpo.
 * @param {object} emailData - Objeto con {to, subject, body, folio, format}.
 *                             'format' es opcional; si no viene, se usa el guardado en la cotización.
 * @return {object} Un objeto indicando el resultado del envío.
 */
function sendQuoteByEmail(emailData) {
  try {
    if (!emailData.to || !emailData.subject || !emailData.body || !emailData.folio) {
      throw new Error("Faltan datos para enviar el correo (to, subject, body, folio).");
    }

    // 1. Generar el PDF en el formato elegido (el 'actual' se arma desde HTML; el CCL,
    //    desde la plantilla de Google Sheets).
    const pdfBlob = generateQuotePdfBlob(emailData.folio, emailData.format);

    // 2. Obtener los detalles completos de la cotización para armar la plantilla HTML del correo
    const quoteResponse = getQuoteDetails(emailData.folio);
    if (!quoteResponse.success) {
      throw new Error("No se pudieron obtener los detalles de la cotización para armar la plantilla.");
    }
    const quote = quoteResponse.quote;

    // 3. Procesar los productos para el cuerpo del correo con sus fotos (reconstruidas dinámicamente)
    let productsHtml = '';
    if (quote.products && quote.products.length > 0) {
      quote.products.forEach(p => {
        const unitPrice = parseFloat(p.unitPrice) || 0;
        const quantity = parseInt(p.quantity) || 0;
        const priceVolume = unitPrice * quantity;
        let finalPricePerLine;
        const costPaymentUnique = parseFloat(p.costPaymentUnique) || 0;
        let discountPublicPercent = parseFloat(p.discountPublicPercent) || 0;
        const additionalDiscountApplied = p.additionalDiscountApplied === 'Si';
        let additionalDiscountPercent = parseFloat(p.additionalDiscountPercent) || 0;

        if (costPaymentUnique > 0 && quantity > 0 && unitPrice > 0) {
            finalPricePerLine = costPaymentUnique;
        } else {
            let priceAfterPublic = priceVolume * (1 - (discountPublicPercent / 100));
            priceAfterPublic = Math.max(0, priceAfterPublic);
            finalPricePerLine = priceAfterPublic;
            if (additionalDiscountApplied && additionalDiscountPercent > 0) {
                finalPricePerLine = priceAfterPublic * (1 - (additionalDiscountPercent / 100));
            }
            finalPricePerLine = Math.max(0, finalPricePerLine);
        }
        const totalMonetaryDiscount = priceVolume - finalPricePerLine;
        const effectiveTotalPercentage = priceVolume > 0 ? (totalMonetaryDiscount / priceVolume) * 100 : 0;

        const verifiedImgUrl = getVerifiedImageUrl(p.sku, p.imageUrl);

        productsHtml += `
          <!-- CARD DE PRODUCTO INDIVIDUAL -->
          <table cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width: 650px;margin-top: 10px;margin-bottom:10px;" id="lineItems LIV">
            <tbody style="background:#F7F7F7;">
              <tr align="left" style="background:#F7F7F7; width: 100%;">
                <td style="background:#F7F7F7;width:5%;"></td>
                <td style="background:#fff;border-radius:4px;box-shadow:0 2px 4px 0 rgba(0,0,0,0.15);margin:0 auto;padding:15px;width:90%;">
                  <table cellpadding="0" cellspacing="0" style="width:100%;background:#fff;">
                    <tbody>
                      <tr>
                        <!-- Imagen del producto (Left) -->
                        <td align="center" width="35%" valign="top" style="padding-top:10px;padding-right:15px;text-align:center;">
                          <img style="max-width:140px;width:100%;height:auto;border-radius:4px;border:1px solid #f0f0f0;" alt="Liverpool Product" src="${verifiedImgUrl}">
                        </td>
                        <!-- Detalles del producto (Right) -->
                        <td valign="top" style="font-family:sans-serif;color:#333;">
                          <h2 style="color:#333;font-size:15px;margin:10px 0 6px 0;font-weight:bold;line-height:1.4;">
                            ${p.description || ''}
                          </h2>
                          <p style="color:#666;font-size:12px;margin:0 0 10px 0;">
                            Código de producto: <strong>${p.sku || ''}</strong>
                          </p>
                          
                          <!-- Tabla interna de precios -->
                          <table cellpadding="0" cellspacing="0" style="width:100%;font-size:12px;color:#555;border-top:1px dashed #eee;padding-top:8px;">
                            <tr>
                              <td style="width:50%;padding-bottom:5px;">
                                Precio unitario:<br>
                                <span style="color:#333;font-weight:bold;font-size:13px;">${formatCurrencyGS(unitPrice)}</span>
                              </td>
                              <td style="width:50%;padding-bottom:5px;">
                                Cantidad:<br>
                                <span style="color:#333;font-weight:bold;font-size:13px;">${quantity}</span>
                              </td>
                            </tr>
                            <tr>
                              <td style="padding-bottom:5px;">
                                Promoción:<br>
                                <span style="color:#333;font-weight:bold;">${costPaymentUnique > 0 ? 'PAGO ÚNICO' : 'PRECIO BASE'}</span>
                              </td>
                              <td style="padding-bottom:5px;">
                                Descuento:<br>
                                <span style="color:${totalMonetaryDiscount > 0.001 ? '#ef4444' : '#333'};font-weight:bold;">
                                  ${totalMonetaryDiscount > 0.001 ? `-${formatCurrencyGS(totalMonetaryDiscount)} (${effectiveTotalPercentage.toFixed(0)}%)` : '$0.00'}
                                </span>
                              </td>
                            </tr>
                            <tr>
                              <td colspan="2" style="border-top:1px solid #eee;padding-top:8px;">
                                <p style="color:#666;font-size:12px;margin:0;">Total artículo: <span style="color:#f00;font-weight:bold;font-size:14px;margin-left:5px;">${formatCurrencyGS(finalPricePerLine)}</span></p>
                              </td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </td>
                <td style="background:#F7F7F7;width:5%;"></td>
              </tr>
            </tbody>
          </table>
        `;
      });
    } else {
      productsHtml = `
        <table cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width: 650px;" id="noProducts">
          <tr align="center" style="background:#F7F7F7;">
            <td style="width:5%;"></td>
            <td style="background:#fff;border-radius:4px;box-shadow:0 2px 4px 0 rgba(0,0,0,0.15);padding:20px;width:90%;color:#666;font-family:sans-serif;">
              No hay productos en esta cotización.
            </td>
            <td style="width:5%;"></td>
          </tr>
        </table>
      `;
    }

    const userMessageHtml = emailData.body.replace(/\n/g, '<br>');

    // 4. Armar la plantilla HTML del correo similar al ticket de Liverpool
    const finalHtmlBody = `
      <!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
      <html>
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Tu Cotización está Lista</title>
        <style type="text/css">
          body {
            font-family: sans-serif;
            margin: 0 auto !important;
            padding: 0 !important;
            background: #F7F7F7;
            max-width: 650px;
          }
        </style>
      </head>
      <body bgcolor="#F7F7F7" style="background-color: #F7F7F7; margin: 0 auto !important; padding: 0 !important; font-family: sans-serif; max-width: 650px;">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" style="background-color: #F7F7F7;">
          <tbody>
            <tr>
              <td align="center" valign="top" style="padding-top: 20px;">
                
                <!-- TOP HEADER LOGO (Liverpool banner) -->
                <table width="100%" cellspacing="0" cellpadding="0" role="presentation" style="max-width: 650px;">
                  <tbody>
                    <tr>
                      <td align="center" style="background-color: #F7F7F7;">
                        <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/3/35/Liverpool_logo.svg/1280px-Liverpool_logo.svg.png" alt="Liverpool - Es parte de mi vida" style="display: block; padding: 10px 0; text-align: center; height: auto; max-width: 160px; margin: 0 auto; border: 0;">
                      </td>
                    </tr>
                  </tbody>
                </table>

                <!-- HEADER: ¡TU COTIZACIÓN ESTÁ LISTA! -->
                <table align="center" cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width: 650px;margin-top: 10px;margin-bottom:10px;" id="LIV Order">
                  <tbody style="background:#F7F7F7;">
                    <tr align="center" style="background:#F7F7F7;">
                      <td style="width:5%;"></td>
                      <td style="background:#FFF;border-radius:4px;box-shadow:0 2px 4px 0 rgba(0, 0, 0, 0.15);margin:0 auto;padding:25px;width:90%;text-align:center;">
                        <h1 style="margin: 0; color:#333;font-size:24px;font-weight:bold;font-family:sans-serif;">
                          ¡Tu cotización está lista!
                        </h1>
                        <p style="color:#666; margin:15px 0 0 0; font-size:14px; line-height:1.5; font-family:sans-serif; text-align:center;">
                          Te compartimos los detalles de la cotización que solicitaste. Los precios e indicaciones se detallan a continuación.
                        </p>
                      </td>
                      <td style="width:5%;"></td>
                    </tr>
                  </tbody>
                </table>

                <!-- ADVISOR'S CUSTOM MESSAGE CARD -->
                <table cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width: 650px;margin-top: 10px;margin-bottom:10px;" id="advisorMessageCard">
                  <tbody style="background:#F7F7F7;">
                    <tr align="left" style="background:#F7F7F7;">
                      <td style="width:5%;"></td>
                      <td style="background:#fff;border-radius:4px;box-shadow:0 2px 4px 0 rgba(0,0,0,0.15);margin:0 auto;padding:18px;width:90%;font-family:sans-serif;font-size:14px;color:#333;line-height:1.5;">
                        <p style="margin:0 0 10px 0;font-weight:bold;color:#e10098;font-size:14px;">Mensaje de tu Asesor:</p>
                        <div style="background:#fdf2f8;border-left:4px solid #e10098;padding:12px 16px;border-radius:0 4px 4px 0;color:#4c4c4c;line-height:1.5;">
                          ${userMessageHtml}
                        </div>
                      </td>
                      <td style="width:5%;"></td>
                    </tr>
                  </tbody>
                </table>

                <!-- NOTICE BANNER (Yellow alert box) -->
                <table style="width:100%;max-width: 650px;margin-top: 10px;margin-bottom:10px;" cellpadding="0" cellspacing="0" border="0" id="noticeAlert">
                  <tbody style="background:#f7f7f7;">
                    <tr style="background:#f7f7f7;">
                      <td style="width:5%;"></td>
                      <td style="background:#F7F7F7; width:90%;">
                        <table cellpadding="0" cellspacing="0" style="width:100%;">
                          <tr>
                            <td style="border-left:5px solid #ffd457; background:#fff4d4; padding:12px 15px; border-radius: 0 4px 4px 0; font-size:13px; color:#665c40; font-family:sans-serif; text-align:left; line-height:1.4;">
                              <strong>Nota importante:</strong> Los precios y promociones están sujetos a cambios sin previo aviso. Esta cotización tiene fines informativos y la disponibilidad de los artículos se garantiza al concretar la compra.
                            </td>
                          </tr>
                        </table>
                      </td>
                      <td style="width:5%;"></td>
                    </tr>
                  </tbody>
                </table>

                <!-- GENERAL DATES & TOTALS CARD -->
                <table cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width: 650px;margin-top: 10px;margin-bottom:10px;" id="datesAndTotals">
                  <tbody style="background:#F7F7F7;">
                    <tr align="left" style="background:#F7F7F7;">
                      <td style="width:5%;"></td>
                      <td style="background:#fff;border-radius:4px;box-shadow:0 2px 4px 0 rgba(0,0,0,0.15);margin:0 auto;padding:15px;width:90%;">
                        <table cellpadding="0" cellspacing="0" style="background:#fff;width:100%;font-size:13px;font-family:sans-serif;">
                          <tbody>
                            <tr>
                              <td width="50%" align="left" style="color:#333;">
                                Fecha de emisión: <span style="font-weight:700;">${new Date(quote.timestamp).toLocaleDateString('es-MX', { day: '2-digit', month: 'long', year: 'numeric' })}</span>
                              </td>
                              <td width="50%" align="right" style="color:#333;">
                                Total cotizado: <span style="font-weight:700;color:#e10098;font-size:15px;">${formatCurrencyGS(quote.summaryTotal)}</span>
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </td>
                      <td style="width:5%;"></td>
                    </tr>
                  </tbody>
                </table>

                <!-- SUMMARY BLOCK HEADER: CLIENTE Y ASESOR -->
                <table cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width: 650px;margin-top: 15px;margin-bottom:5px;" id="clientAdvisorHeader">
                  <tbody style="background:#F7F7F7;">
                    <tr align="left" style="background:#F7F7F7;">
                      <td style="background:#F7F7F7;width:5%;"></td>
                      <td style="background:#F7F7F7;width:90%;">
                        <h3 style="border-bottom:2px solid #e10098;color:#FFF;font-size:15px;font-weight:normal;margin:0;font-family:sans-serif;">
                          <span style="background:#e10098;display:table-cell;height:30px;line-height:30px;padding:3px 16px 0;border-radius:4px 4px 0 0;">
                            Información del Cliente y Asesor
                          </span>
                        </h3>
                      </td>
                      <td style="background:#F7F7F7;width:5%;"></td>
                    </tr>
                  </tbody>
                </table>

                <!-- SUMMARY BLOCK CONTENT: CLIENTE Y ASESOR -->
                <table cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width: 650px;margin-bottom:10px;" id="clientAdvisorContent">
                  <tbody style="background:#F7F7F7;">
                    <tr align="left" style="background:#F7F7F7; width: 100%;">
                      <td style="background:#F7F7F7;width:5%;"></td>
                      <td style="background:#FFF;border-radius:4px;box-shadow:0 2px 4px 0 rgba(0, 0, 0, 0.15);margin:0 auto;padding:15px;width:90%;font-family:sans-serif;font-size:13px;line-height:1.5;color:#333;">
                        <table cellpadding="0" cellspacing="0" style="width:100%;">
                          <tr>
                            <td width="48%" valign="top" style="border-right:1px solid #eee;padding-right:10px;">
                              <p style="margin:2px 0;color:#e10098;font-weight:bold;font-size:13px;">Dirigido a:</p>
                              <p style="margin:2px 0;"><strong>Cliente:</strong> ${quote.clientName || 'N/A'}</p>
                              <p style="margin:2px 0;"><strong>Correo:</strong> ${quote.clientEmail || 'N/A'}</p>
                              <p style="margin:2px 0;"><strong>Teléfono:</strong> ${quote.clientPhone || 'N/A'}</p>
                            </td>
                            <td width="4%">&nbsp;</td>
                            <td width="48%" valign="top" style="padding-left:10px;">
                              <p style="margin:2px 0;color:#e10098;font-weight:bold;font-size:13px;">Atendido por:</p>
                              <p style="margin:2px 0;"><strong>Asesor:</strong> ${quote.advisorName || 'N/A'}</p>
                              <p style="margin:2px 0;"><strong>Extensión:</strong> ${quote.advisorExt || 'N/A'}</p>
                              <p style="margin:2px 0;"><strong>Folio:</strong> ${quote.folio || 'N/A'}</p>
                            </td>
                          </tr>
                        </table>
                      </td>
                      <td style="background:#F7F7F7;width:5%;"></td>
                    </tr>
                  </tbody>
                </table>

                <!-- SECTION HEADER: TUS PRODUCTOS -->
                <table cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width: 650px;margin-top: 15px;margin-bottom:5px;" id="productsHeader">
                  <tbody style="background:#F7F7F7;">
                    <tr align="left" style="background:#F7F7F7;">
                      <td style="background:#F7F7F7;width:5%;"></td>
                      <td style="background:#F7F7F7;width:90%;">
                        <h3 style="border-bottom:2px solid #e10098;color:#FFF;font-size:15px;font-weight:normal;margin:0;font-family:sans-serif;">
                          <span style="background:#e10098;display:table-cell;height:30px;line-height:30px;padding:3px 16px 0;border-radius:4px 4px 0 0;">
                            Detalle de Artículos
                          </span>
                        </h3>
                      </td>
                      <td style="background:#F7F7F7;width:5%;"></td>
                    </tr>
                  </tbody>
                </table>

                <!-- PRODUCTS LIST LOOP -->
                ${productsHtml}

                <!-- TOTALS SUMMARY CARD -->
                <table cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width: 650px;margin-top: 10px;margin-bottom:10px;" id="totalsSummary">
                  <tbody style="background:#F7F7F7;">
                    <tr align="left" style="background:#F7F7F7;">
                      <td style="width:5%;"></td>
                      <td style="background:#fff;border-radius:4px;box-shadow:0 2px 4px 0 rgba(0,0,0,0.15);margin:0 auto;padding:15px;width:90%;">
                        <table cellpadding="0" cellspacing="0" style="background:#fff;width:100%;font-size:13px;font-family:sans-serif;color:#333;">
                          <tbody>
                            <tr>
                              <td align="right" style="padding: 4px 0;color:#666;">Subtotal:</td>
                              <td align="right" width="30%" style="padding: 4px 0;font-weight:700;">${formatCurrencyGS(quote.summarySubtotal)}</td>
                            </tr>
                            <tr>
                              <td align="right" style="padding: 4px 0;color:#666;">IVA (16%):</td>
                              <td align="right" style="padding: 4px 0;font-weight:700;">${formatCurrencyGS(quote.summaryVat)}</td>
                            </tr>
                            <tr style="font-size:15px;font-weight:bold;color:#e10098;">
                              <td align="right" style="border-top:1.5px solid #e10098;padding-top:10px;margin-top:5px;">TOTAL GENERAL:</td>
                              <td align="right" style="border-top:1.5px solid #e10098;padding-top:10px;margin-top:5px;color:#e10098;font-size:16px;">${formatCurrencyGS(quote.summaryTotal)}</td>
                            </tr>
                          </tbody>
                        </table>
                      </td>
                      <td style="width:5%;"></td>
                    </tr>
                  </tbody>
                </table>

                <!-- EMAIL FOOTER INFO -->
                <table cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width: 650px;margin-top: 20px;margin-bottom:20px;" id="footerInfo">
                  <tbody style="background:#F7F7F7;">
                    <tr align="center" style="background:#F7F7F7;">
                      <td style="width:5%;"></td>
                      <td style="font-family:sans-serif;font-size:11px;color:#888;line-height:1.5;text-align:center;width:90%;">
                        <p style="margin: 0 0 10px 0;"><strong>Nota:</strong> Se adjunta a este correo el archivo PDF oficial con la cotización formal detallada para su descarga o impresión.</p>
                        <p style="margin: 0 0 15px 0;font-weight:bold;color:#e10098;font-size:12px;">Liverpool - Es parte de mi vida</p>
                      </td>
                      <td style="width:5%;"></td>
                    </tr>
                  </tbody>
                </table>

              </td>
            </tr>
          </tbody>
        </table>
      </body>
      </html>
    `;

    // 5. Enviar el correo con el archivo PDF adjunto y la plantilla HTML.
    //    Se intenta enviar desde el alias institucional; si el alias no está dado de
    //    alta en la cuenta que ejecuta el script (Gmail > Configuración > Cuentas >
    //    "Enviar como"), se envía desde la cuenta propia para no bloquear el envío.
    //    Las respuestas del cliente llegan siempre al asesor dueño de la cotización.
    const options = {
      htmlBody: finalHtmlBody,
      name: 'Cotizaciones Ventel Liverpool', // Nombre del remitente que verá el cliente
      attachments: [pdfBlob]
    };
    if (quote.advisorEmail) options.replyTo = quote.advisorEmail;

    let sentFrom = '';
    let aliasAvailable = false;
    try {
      // Requiere el permiso de Gmail (https://mail.google.com/ en appsscript.json).
      // Si el permiso o el alias no están, se registra y se envía por la vía clásica.
      aliasAvailable = GmailApp.getAliases().indexOf(MAIL_ALIAS) !== -1;
    } catch (e) {
      Logger.log('Sin acceso a los alias de Gmail (falta permiso o alias): ' + e.message);
    }

    if (aliasAvailable) {
      try {
        GmailApp.sendEmail(emailData.to, emailData.subject, '', Object.assign({}, options, { from: MAIL_ALIAS }));
        sentFrom = MAIL_ALIAS;
      } catch (e) {
        Logger.log('Fallo el envío con alias, se reintenta por la vía clásica: ' + e.message);
        aliasAvailable = false;
      }
    }
    if (!aliasAvailable) {
      // MailApp usa el permiso de envío que el script ya tenía autorizado desde siempre:
      // el correo SIEMPRE sale, aunque el alias o el permiso de Gmail no estén listos.
      MailApp.sendEmail(emailData.to, emailData.subject, '', options);
      sentFrom = Session.getActiveUser() ? Session.getActiveUser().getEmail() : '';
    }
    Logger.log(`Correo con PDF adjunto enviado a ${emailData.to} desde ${sentFrom || 'cuenta del script'} para el folio ${emailData.folio}`);

    // 6. Actualizar el estatus de la cotización a "Enviada por Correo"
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cotizacionesSheet = ss.getSheetByName(COTIZACIONES_SHEET_NAME);
    const cotHeaders = cotizacionesSheet.getRange(1, 1, 1, cotizacionesSheet.getLastColumn()).getValues()[0];
    const folioColIdx = cotHeaders.indexOf("Folio");
    const statusColIdx = cotHeaders.indexOf("Estatus");
    
    if (folioColIdx > -1 && statusColIdx > -1) {
      const cotDataValues = cotizacionesSheet.getDataRange().getValues();
      for (let i = 1; i < cotDataValues.length; i++) {
          if (cotDataValues[i][folioColIdx] == emailData.folio) {
              cotizacionesSheet.getRange(i + 1, statusColIdx + 1).setValue("Enviada por Correo");
              break;
          }
      }
    }

    return {
      success: true,
      sentFrom: sentFrom,
      aliasUsed: aliasAvailable,
      message: aliasAvailable
        ? `Correo enviado desde ${MAIL_ALIAS}.`
        : `Correo enviado (el alias ${MAIL_ALIAS} aún no está configurado en la cuenta; se usó la cuenta del sistema).`
    };

  } catch (error) {
    Logger.log(`Error al enviar correo para folio ${emailData.folio}: ${error.message} Stack: ${error.stack}`);
    return { success: false, message: `No se pudo enviar el correo: ${error.message}` };
  }
}

/**
 * Verifica si la URL de la imagen preferida responde con estatus 200 (OK).
 * Si no responde 200, escanea secuencialmente otros servidores conocidos de imágenes de Liverpool
 * hasta encontrar uno que contenga el SKU activo, garantizando que el cliente no reciba imágenes rotas.
 *
 * @param {string} sku - El identificador del producto.
 * @param {string} preferredUrl - La URL que recolectó originalmente la extensión.
 * @return {string} Una URL válida verificada o el fallback por defecto de ss628.
 */
function getVerifiedImageUrl(sku, preferredUrl) {
  if (preferredUrl && preferredUrl.indexOf('http') === 0) {
    try {
      const res = UrlFetchApp.fetch(preferredUrl, { muteHttpExceptions: true, followRedirects: true });
      if (res.getResponseCode() === 200) {
        Logger.log(`Imagen preferida verificada con éxito: ${preferredUrl}`);
        return preferredUrl;
      }
    } catch (e) {
      Logger.log("Error al verificar preferredUrl: " + e.message);
    }
  }

  if (!sku) return "https://assets.liverpool.com.mx/assets/images/placeholder.gif";

  // Lista de subdominios de servidores de imágenes comunes de Liverpool
  const subdomains = ["ss628", "ss224", "ss318", "ss414", "ss512", "ss101", "ss202", "ss303", "ss404"];
  
  for (let i = 0; i < subdomains.length; i++) {
    const testUrl = `https://${subdomains[i]}.liverpool.com.mx/xl/${sku}.jpg`;
    try {
      const response = UrlFetchApp.fetch(testUrl, { muteHttpExceptions: true, followRedirects: true });
      if (response.getResponseCode() === 200) {
        Logger.log(`Imagen encontrada con éxito para SKU ${sku} en: ${testUrl}`);
        return testUrl;
      }
    } catch (e) {
      // Continuar si la petición falla
    }
  }
  
  // Fallback por defecto si todo lo demás falla
  return `https://ss628.liverpool.com.mx/xl/${sku}.jpg`;
}
