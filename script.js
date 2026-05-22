// ============================================================
// GALAXIA — Apps Script receptor de formulario
// ============================================================
// INSTRUCCIONES DE INSTALACIÓN:
//
// 1. Abrí tu Google Sheet (creá uno nuevo si no tenés)
// 2. Menú: Extensiones → Apps Script
// 3. Borrá todo el código que aparece y pegá este script completo
// 4. Guardá con Ctrl+S (nombre del proyecto: "Galaxia Forms")
// 5. Click en "Implementar" → "Nueva implementación"
//    - Tipo: Aplicación web
//    - Ejecutar como: Yo (tu cuenta)
//    - Acceso: Cualquier persona
// 6. Click "Implementar" → copiá la URL que te da
// 7. Abrí el archivo forms_galaxia.html y buscá la línea:
//    const APPS_SCRIPT_URL = "PEGAR_URL_AQUI";
//    Reemplazá PEGAR_URL_AQUI con la URL copiada
// ============================================================
 
const SHEET_NAME = "Respuestas"; // Nombre de la hoja donde se guardan los datos
 
// ------------------------------------------------------------
// Función principal — recibe el POST del formulario
// ------------------------------------------------------------
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
 
    // Si la hoja no existe, la crea con encabezados
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      crearEncabezados(sheet);
    }
 
    // Si la hoja existe pero está vacía, agrega encabezados
    if (sheet.getLastRow() === 0) {
      crearEncabezados(sheet);
    }
 
    // Parsear los datos recibidos
    const datos = JSON.parse(e.postData.contents);
 
    // Armar la fila
    const fila = [
      new Date(),                                    // Timestamp
      datos.vendedor         || "",                  // Vendedor
      datos.objetivo         || "",                  // Objetivo
      // --- DATOS GENERALES ---
      datos.contacto         || "",                  // Contacto cliente
      datos.horario          || "",                  // Horario atención
      datos.alcohol          || "",                  // Vende alcohol
      // --- ALTA ---
      datos.excliente        || "",                  // Fue cliente antes
      datos.contribuyente    || "",                  // Tipo contribuyente
      datos.cuit             || "",                  // CUIT/DNI
      datos.razonsocial      || "",                  // Razón social
      datos.fantasia         || "",                  // Nombre fantasía
      datos.localidad        || "",                  // Localidad
      datos.domicilio        || "",                  // Domicilio
      datos.canal            || "",                  // Canal ON/OFF
      datos.subcanal         || "",                  // Subcanal (bloqueado)
      datos.subcanal_cliente || "",                  // Subcanal según cliente
      datos.dias_visita      || "",                  // Días de visita
      // --- BAJA ---
      datos.baja_codigo      || "",                  // Código baja
      datos.baja_motivo      || "",                  // Motivo baja
      // --- POTENCIAL ---
      datos.pot_nombre       || "",                  // Nombre potencial
      datos.pot_motivo       || "",                  // Motivo potencial
      // --- ACTIVO ---
      datos.activo_codigo    || "",                  // Código activo
      datos.activo_comentario|| "",                  // Comentario activo
      // --- RAZÓN SOCIAL ---
      datos.rs_codigo        || "",                  // Código razón social      → col X
      datos.rs_nueva         || "",                  // Nueva razón social       → col Y
      datos.rs_motivo        || "",                  // Motivo cambio RS         → col Z
      datos.rs_contribuyente || "",                  // Nuevo tipo contrib       → col AA
      datos.rs_cuit          || "",                  // Nuevo CUIT               → col AB
      // --- COMENTARIO FINAL ---
      datos.comentario       || "",                  // Comentario extra
    ];
 
    // Buscar la última fila con datos mirando SOLO columna A (ignora fórmulas en AC+)
    const colA = sheet.getRange("A:A").getValues();
    let ultimaFila = 1; // fila 1 = encabezados
    for (let i = 1; i < colA.length; i++) {
      if (colA[i][0] !== "") ultimaFila = i + 1;
      else break;
    }
    const filaDestino = ultimaFila + 1;
 
    // Escribir la fila en las columnas A hasta AB (columna 28)
    sheet.getRange(filaDestino, 1, 1, fila.length).setValues([fila]);
 
    // Respuesta exitosa con CORS
    return ContentService
      .createTextOutput(JSON.stringify({ resultado: "ok", fila: sheet.getLastRow() }))
      .setMimeType(ContentService.MimeType.JSON);
 
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ resultado: "error", mensaje: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
 
// ------------------------------------------------------------
// Función GET — para testear que el script esté activo
// ------------------------------------------------------------
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ estado: "activo", mensaje: "Galaxia Forms receptor funcionando ✓" }))
    .setMimeType(ContentService.MimeType.JSON);
}
 
// ------------------------------------------------------------
// Crea los encabezados en la primera fila
// ------------------------------------------------------------
function crearEncabezados(sheet) {
  const encabezados = [
    "FECHA/HORA",
    "VENDEDOR",
    "OBJETIVO",
    "CONTACTO CLIENTE",
    "HORARIO ATENCIÓN",
    "VENDE ALCOHOL",
    "FUE CLIENTE ANTES",
    "TIPO CONTRIBUYENTE",
    "CUIT/DNI",
    "RAZÓN SOCIAL",
    "NOMBRE FANTASÍA",
    "LOCALIDAD",
    "DOMICILIO",
    "CANAL (ON/OFF)",
    "SUBCANAL (SISTEMA)",
    "SUBCANAL (CLIENTE DIJO)",
    "DÍAS DE VISITA",
    "CÓDIGO BAJA",
    "MOTIVO BAJA",
    "NOMBRE POTENCIAL",
    "MOTIVO POTENCIAL",
    "CÓDIGO ACTIVO",
    "COMENTARIO ACTIVO",
    "CÓDIGO CAMBIO RAZÓN SOCIAL",    // X
    "NUEVA RAZÓN SOCIAL",             // Y
    "MOTIVO CAMBIO RAZÓN SOCIAL",     // Z
    "NUEVO TIPO CONTRIBUYENTE",       // AA
    "NUEVO CUIT/DNI",                 // AB
    "COMENTARIO EXTRA",               // AC
  ];
 
  sheet.appendRow(encabezados);
 
  // Estilo para los encabezados
  const rango = sheet.getRange(1, 1, 1, encabezados.length);
  rango.setBackground("#1A1A2E");
  rango.setFontColor("#FFFFFF");
  rango.setFontWeight("bold");
  rango.setFontSize(10);
 
  // Congelar primera fila
  sheet.setFrozenRows(1);
 
  // Ancho automático de columnas
  sheet.autoResizeColumns(1, encabezados.length);
}
