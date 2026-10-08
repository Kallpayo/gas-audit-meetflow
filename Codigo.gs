function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ MeetFlow CRM')
    .addItem('Abrir Panel Principal', 'abrirPanelWeb')
    .addToUi();
}

function abrirPanelWeb() {
  // Esto abre tu interfaz HTML en una ventana modal dentro de la misma hoja
  var html = HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('MeetFlow CRM')
      .setWidth(1200)
      .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'MeetFlow CRM');
}
const SHEETS_CONFIG = {
  USUARIO: {
    sheetName: 'USUARIO', idPrefix: 'USU',
    headerMap: { 
      'UsuarioID': 'usuarioID', 
      'NumeroEmpleado': 'numeroEmpleado',
      'NombreCompleto': 'nombreCompleto', 
      'CorreoCorporativo': 'correoCorporativo', 
      'TelefonoContacto': 'telefonoContacto', 
      'Genero': 'genero', 
      'Cargo': 'cargo', 
      'Departamento': 'departamento', 
      'TipoContrato': 'tipoContrato',
      'ModalidadTrabajo': 'modalidadTrabajo',
      'ConsentimientoGDPR': 'consentimientoGDPR',
      'FechaRegistro': 'fechaRegistro', 
      'UltimaActualizacion': 'ultimaActualizacion' 
    }
  },
  CONSULTA: {
    sheetName: 'CONSULTA', idPrefix: 'CON',
    headerMap: { 
      'ConsultaID': 'consultaID', 
      'Departamento': 'departamento',
      'MedioAtencion': 'medioAtencion', 
      'FechaConsulta': 'fechaConsulta', 
      'HoraVisita': 'horaVisita', 
      'HoraFin': 'horaFin', 
      'CodigoTicket': 'codigoTicket', 
      'TipoConsulta': 'tipoConsulta', 
      'TemaOrdenDelDia': 'temaOrdenDelDia', 
      'TipoSolucion': 'tipoSolucion', 
      'AcuerdosCompromisos': 'acuerdosCompromisos', 
      'NombreFuncionarioQueAtendio': 'nombreFuncionarioQueAtendio', 
      'AgendadaPreviamente': 'agendadaPreviamente' 
    }
  },
  PARTICIPANTES: {
    sheetName: 'PARTICIPANTES',
    headerMap: { 'ConsultaID': 'consultaID', 'UsuarioID': 'usuarioID', 'Rol': 'rol' }
  },
  CONFIG: {
    sheetName: 'CONFIG',
    headerMap: { 'TipoLista': 'tipoLista', 'ValorItem': 'valorItem' }
  }
};
function logActivity(level, message) {
  Logger.log(`[${level}] ${new Date().toISOString()} - ${message}`);
}

function doGet(e) {
  const page = e.parameter.page || 'index';
  return HtmlService.createTemplateFromFile(page)
    .evaluate()
    .setTitle('MeetFlow CRM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getTemplateContent(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return `<div class="error">Error: Módulo ${filename}.html no encontrado.</div>`;
  }
}

function parseSheetValue(value) {
  if (value == null || value === '') return '';
  if (value instanceof Date) {
    const isTimeOnly = value.getFullYear() === 1899;
    const format = isTimeOnly ? "HH:mm" : "yyyy-MM-dd";
    return Utilities.formatDate(value, Session.getScriptTimeZone(), format);
  }
  return String(value).trim();
}

function getSheetData(sheetNameFromConfig) {
  Logger.log("Intentando acceder a la configuración: " + sheetNameFromConfig);
  try {
    const config = SHEETS_CONFIG[sheetNameFromConfig];
    if (!config) throw new Error('Configuración de hoja no válida.');
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    if (!sheet) throw new Error(`La hoja ${config.sheetName} no existe.`);
    
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return JSON.stringify([]);
    const headers = values[0].map(h => String(h).toLowerCase().replace(/\s+/g, ''));
    Logger.log("Encabezados detectados: " + headers);
    const data = values.slice(1).map(row => {
      const obj = {};
      Object.entries(config.headerMap).forEach(([sheetHeader, jsKey]) => {
        const normalizedHeader = sheetHeader.toLowerCase().replace(/\s+/g, '');
        const index = headers.indexOf(normalizedHeader);
        obj[jsKey] = index !== -1 ? parseSheetValue(row[index]) : '';
      });
      return obj;
    });
    return JSON.stringify(data);
  } catch (e) {
    logActivity('ERROR', `getSheetData: ${e.message}`);
    return JSON.stringify([]);
  }
}

function generateNextId(sheetNameFromConfig, idPrefix) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const config = SHEETS_CONFIG[sheetNameFromConfig];
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return `${idPrefix}001`;
    
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const numericIds = ids
      .filter(id => id && String(id).startsWith(idPrefix))
      .map(id => parseInt(String(id).substring(idPrefix.length), 10))
      .filter(num => !isNaN(num));
    const max = numericIds.length > 0 ? Math.max(...numericIds) : 0;
    return `${idPrefix}${String(max + 1).padStart(3, '0')}`;
  } finally {
    lock.releaseLock();
  }
}


function getDropdownData(listType) {
  try {
    const data = JSON.parse(getSheetData('CONFIG'));
    Logger.log("Buscando listType: " + listType);
    Logger.log("Datos recibidos: " + JSON.stringify(data));
    if (!data || data.length === 0) return JSON.stringify([]);
    
    const unicos = [...new Set(data.filter(d => d.tipoLista === listType).map(d => d.valorItem).filter(Boolean))];
    return JSON.stringify(unicos);
  } catch (e) {
    return JSON.stringify([]);
  }
}

function validateUsuarioData(data) {
  if (!data) throw new Error("Datos vacíos.");
  if (!data.nombreCompleto || data.nombreCompleto.trim().length < 3) throw new Error("Nombre muy corto.");
  if (!data.correoCorporativo || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.correoCorporativo)) throw new Error("Correo inválido.");
  if (!data.genero || !data.cargo || !data.departamento) throw new Error("Género, Cargo y Departamento son requeridos.");
  return true;
}

function validateConsultaData(data) {
  if (!data) throw new Error("Datos vacíos.");
  if (!data.medioAtencion) throw new Error("Medio requerido.");
  if (!data.fechaConsulta) throw new Error("Fecha requerida.");
  if (!data.tipoConsulta) throw new Error("Tipo de consulta requerido.");
  if (!data.nombreFuncionarioQueAtendio) throw new Error("Funcionario requerido.");
  if (data.medioAtencion === 'Ticket') {
    if (!data.codigoTicket) throw new Error("El Código de Ticket es obligatorio.");
    if (!data.departamento) throw new Error("El departamento es obligatorio para Tickets.");
    return true;
  }

  const timeRegex = /^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/;
  if (!timeRegex.test(data.horaVisita)) throw new Error("Hora de inicio inválida.");
  if (!timeRegex.test(data.horaFin)) throw new Error("Hora de fin inválida.");
  if (data.horaFin <= data.horaVisita) throw new Error("La hora de fin debe ser posterior a la de inicio.");
  if (!data.temaOrdenDelDia) throw new Error("El tema/detalle es requerido.");
  
  if (data.medioAtencion === 'Reunión de Trabajo' || data.medioAtencion === 'Reunión') {
    if (!data.departamento) throw new Error("El departamento es obligatorio para reuniones.");
  }
  return true;
}

function createUsuario(usuarioData) {
  try {
    validateUsuarioData(usuarioData);
    const config = SHEETS_CONFIG.USUARIO;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const newId = generateNextId('USUARIO', config.idPrefix);
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const mappedData = {
      'UsuarioID': newId,
      'NumeroEmpleado': usuarioData.numeroEmpleado || 'N/A',
      'NombreCompleto': usuarioData.nombreCompleto.toUpperCase(),
      'CorreoCorporativo': usuarioData.correoCorporativo.toLowerCase(),
      'TelefonoContacto': usuarioData.telefonoContacto || 'N/A',
      'Genero': usuarioData.genero,
      'Cargo': usuarioData.cargo,
      'Departamento': usuarioData.departamento,
      'TipoContrato': usuarioData.tipoContrato,
      'ModalidadTrabajo': usuarioData.modalidadTrabajo,
      'ConsentimientoGDPR': usuarioData.consentimientoGDPR || 'NO',
      'FechaRegistro': now,
      'UltimaActualizacion': now
    };
    sheet.appendRow(headers.map(h => mappedData[h] || ''));
    return { success: true, message: "Expediente creado correctamente con ID: " + newId };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function searchUsuarios(searchQuery) {
  try {
    if (!searchQuery || searchQuery.trim().length < 3) return JSON.stringify([]);
    const query = searchQuery.toLowerCase().trim();
    const users = JSON.parse(getSheetData('USUARIO'));
    
    return JSON.stringify(users.filter(u => 
      (u.nombreCompleto && u.nombreCompleto.toLowerCase().includes(query)) ||
      (u.correoCorporativo && u.correoCorporativo.toLowerCase().includes(query)) ||
      (u.numeroEmpleado && u.numeroEmpleado.toLowerCase().includes(query))
    ).map(u => ({
      usuarioID: u.usuarioID,
      nombreCompleto: u.nombreCompleto,
      correoCorporativo: u.correoCorporativo,
      numeroEmpleado: u.numeroEmpleado || 'Sin ID'
    })));
  } catch (e) {
    return JSON.stringify([]);
  }
}

function createConsulta(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    validateConsultaData(payload.consultaData);
    if (!payload.participantes || payload.participantes.length === 0) {
      throw new Error("Debe agregar al menos un participante/solicitante.");
    }

    const configCons = SHEETS_CONFIG.CONSULTA;
    const sheetCons = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configCons.sheetName);
    const headersCons = sheetCons.getRange(1, 1, 1, sheetCons.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const newId = generateNextId('CONSULTA', configCons.idPrefix);
    const data = payload.consultaData;

    const mappedConsulta = {
      'ConsultaID': newId,
      'Departamento': data.departamento || '',
      'MedioAtencion': data.medioAtencion || '',
      'FechaConsulta': data.fechaConsulta || '',
      'HoraVisita': data.horaVisita || '',
      'HoraFin': data.horaFin || '',
      'CodigoTicket': data.codigoTicket || '',
      'TemaOrdenDelDia': data.temaOrdenDelDia || '',
      'TipoSolucion': data.tipoSolucion || '',
      'TipoConsulta': data.tipoConsulta || '',
      'AcuerdosCompromisos': data.acuerdosCompromisos || '',
      'NombreFuncionarioQueAtendio': data.nombreFuncionarioQueAtendio || '',
      'AgendadaPreviamente': data.agendadaPreviamente || ''
    };

    sheetCons.appendRow(headersCons.map(h => mappedConsulta[h] || ''));
    const configPart = SHEETS_CONFIG.PARTICIPANTES;
    const sheetPart = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configPart.sheetName);
    const headersPart = sheetPart.getRange(1, 1, 1, sheetPart.getLastColumn()).getValues()[0].map(h => String(h).trim());

    payload.participantes.forEach(part => {
      const mappedPart = {
        'ConsultaID': newId,
        'UsuarioID': part.usuarioID || '',
        'Rol': part.rol || ''
      };
      sheetPart.appendRow(headersPart.map(h => mappedPart[h] || ''));
    });
    return { success: true, message: "Registro exitoso." };
  } catch (e) {
    logActivity('ERROR', `createConsulta: ${e.message}`);
    return { success: false, message: "Error al registrar: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function getDatosDashboard() {
  try {
    const consultas = JSON.parse(getSheetData('CONSULTA'));
    const usuarios = JSON.parse(getSheetData('USUARIO'));
    const participantes = JSON.parse(getSheetData('PARTICIPANTES'));

    const usuariosMap = usuarios.reduce((map, u) => { 
      map[u.usuarioID] = u.nombreCompleto; 
      return map; 
    }, {});
    const conteoPorTipo = {}; 
    const conteoPorMedio = {};

    consultas.forEach(c => { 
      const tipo = c.tipoConsulta || 'General'; 
      const medio = c.medioAtencion || 'Oficina'; 
      conteoPorTipo[tipo] = (conteoPorTipo[tipo] || 0) + 1; 
      conteoPorMedio[medio] = (conteoPorMedio[medio] || 0) + 1; 
    });
    const ultimas = consultas.slice(-5).reverse().map(c => { 
      const parts = participantes.filter(p => p.consultaID === c.consultaID); 
      const nombresPartes = parts.map(p => `${usuariosMap[p.usuarioID] || 'Empleado corporativo'} (${p.rol})`).join(', ');
      return { 
        fecha: c.fechaConsulta, 
        funcionario: c.nombreFuncionarioQueAtendio,
        tipo: c.tipoConsulta, 
        medio: c.medioAtencion, 
        participantes: nombresPartes || 'Registro directo',
        resolucion: c.tipoSolucion
      }; 
    });

    const totalConsultas = consultas.length;
    const resueltas = consultas.filter(c => c.tipoSolucion && (c.tipoSolucion.toUpperCase().includes('RESUELTO'))).length;
    const pendientes = consultas.filter(c => c.tipoSolucion && (c.tipoSolucion.toUpperCase().includes('PENDIENTE'))).length;
    const cerradas = consultas.filter(c => c.tipoSolucion && (c.tipoSolucion.toUpperCase().includes('CERRADO'))).length;
    return { 
      success: true,
      data: { 
        kpis: { 
          totalConsultas: totalConsultas, 
          totalResueltas: resueltas,
          totalPendientes: pendientes,
          totalCerradas: cerradas, 
          totalUsuarios: usuarios.length,
          promedioDuracion: calcularPromedioDuracion(consultas) 
        }, 
        conteoPorTipo: conteoPorTipo, 
        conteoPorMedio: conteoPorMedio, 
        ultimasActividades: ultimas 
      } 
    };
  } catch (e) { 
    logActivity('ERROR', `getDatosDashboard: ${e.message}`);
    return { success: false, message: "Excepción en el cálculo de métricas: " + e.message };
  } 
}

function calcularPromedioDuracion(consultas) {
  const validas = consultas.filter(c => 
    c.horaVisita && c.horaFin && c.horaVisita !== '00:00' && 
    /^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/.test(c.horaVisita) && 
    /^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$/.test(c.horaFin)
  );
  if (!validas.length) return 0;
  
  const total = validas.reduce((sum, c) => {
    const [h1, m1] = c.horaVisita.split(':').map(Number);
    const [h2, m2] = c.horaFin.split(':').map(Number);
    return sum + ((h2 * 60 + m2) - (h1 * 60 + m1));
  }, 0);
  return Math.round(total / validas.length);
}

function generarReporteConsultasSheet(funcionarioFiltro) {
  try {
    const carpetaDestino = getMeetFlowFolder("MeetFlow_Reportes");
    let consultas = JSON.parse(getSheetData('CONSULTA'));
    if (consultas.length === 0) return { success: false, message: "No hay datos." };
    if (funcionarioFiltro && funcionarioFiltro !== 'Todos') {
      consultas = consultas.filter(c => c.nombreFuncionarioQueAtendio === funcionarioFiltro);
      if (consultas.length === 0) return { success: false, message: `Sin datos para ${funcionarioFiltro}.` };
    }

    const usuarios = JSON.parse(getSheetData('USUARIO'));
    const participantes = JSON.parse(getSheetData('PARTICIPANTES'));
    const usuariosMap = usuarios.reduce((map, user) => { map[user.usuarioID] = user.nombreCompleto; return map; }, {});
    const sufix = funcionarioFiltro && funcionarioFiltro !== 'Todos' ? `_${funcionarioFiltro.replace(/\s/g, '')}` : '_General';
    const nombreArchivo = `Reporte_MeetFlow${sufix}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")}`;
    const nuevaHoja = SpreadsheetApp.create(nombreArchivo);
    
    const hoja = nuevaHoja.getActiveSheet();
    const headers = ["ID", "Departamento", "Medio", "Fecha", "Inicio", "Fin", "Codigo Ticket", "Tipo", "Tipo Solución", "Tema", "Acuerdos", "Responsable", "Participantes", "Agendado"];
    const datos = consultas.map(c => {
      const parts = participantes.filter(p => p.consultaID === c.consultaID);
      const partNames = parts.map(p => `${usuariosMap[p.usuarioID] || 'N/A'} (${p.rol})`).join('; ');
      return [c.consultaID, c.departamento, c.medioAtencion, c.fechaConsulta, c.horaVisita, c.horaFin, c.codigoTicket, c.tipoConsulta, c.tipoSolucion, c.temaOrdenDelDia, c.acuerdosCompromisos, c.nombreFuncionarioQueAtendio, partNames, c.agendadaPreviamente];
    });
    const outputData = [headers, ...datos];
    hoja.getRange(1, 1, outputData.length, headers.length).setValues(outputData);
    hoja.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
    hoja.autoResizeColumns(1, headers.length);
    hoja.setFrozenRows(1);
    
    const archivo = DriveApp.getFileById(nuevaHoja.getId());
    archivo.moveTo(carpetaDestino);

    return { success: true, fileUrl: archivo.getUrl(), message: "Reporte generado correctamente." };
  } catch (error) {
    return { success: false, message: `Error: ${error.message}` };
  }
}

function getReunionesPorDepartamento(departamento) {
  try {
    const consultas = JSON.parse(getSheetData('CONSULTA'));
    const reuniones = consultas.filter(c => 
      (c.medioAtencion === 'Reunión de Trabajo' || c.medioAtencion === 'Reunión' || c.medioAtencion === 'Ticket' || c.medioAtencion === 'Presencial' || c.medioAtencion === 'Virtual') && 
      (departamento === 'Todas' || c.departamento === departamento)
    );
    return JSON.stringify(reuniones.map(r => ({ 
      id: r.consultaID, 
      fecha: r.fechaConsulta, 
      hora: r.horaVisita, 
      tema: r.temaOrdenDelDia, 
      funcionario: r.nombreFuncionarioQueAtendio,
      canal: r.medioAtencion,
      ticket: r.codigoTicket
    })));
  } catch (e) {
    logActivity('ERROR', `getReunionesPorDepartamento: ${e.message}`);
    return JSON.stringify([]);
  }
}

function generarActaReunionPDF(consultaID) {
  try {
    const consultas = JSON.parse(getSheetData('CONSULTA'));
    const consulta = consultas.find(c => c.consultaID === consultaID);
    if (!consulta) throw new Error("Consulta no encontrada.");
    const nombreEmpresa = getEmpresaNombre();
    const docName = `Acta_${consulta.consultaID}`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    
    body.appendParagraph(nombreEmpresa).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true).setFontSize(14);
    body.appendParagraph("REGISTRO EJECUTIVO DE SESIÓN").setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true).setFontSize(12);
    body.appendParagraph(`Acta No. ${consulta.consultaID}`).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setBold(true).setFontSize(10);
    body.appendParagraph("____________________________________________________________________");
    body.appendParagraph(`Fecha: ${consulta.fechaConsulta} | Hora: ${consulta.horaVisita} - ${consulta.horaFin}`).setBold(true);
    body.appendParagraph(`Departamento: ${consulta.departamento} | Responsable: ${consulta.nombreFuncionarioQueAtendio}`).setBold(true);
    
    body.appendParagraph("\nASISTENTES:").setBold(true).setFontSize(11);
    const usuarios = JSON.parse(getSheetData('USUARIO'));
    const participantes = JSON.parse(getSheetData('PARTICIPANTES'));
    const usuariosMap = usuarios.reduce((map, u) => { map[u.usuarioID] = u.nombreCompleto; return map; }, {});
    const parts = participantes.filter(p => p.consultaID === consultaID);
    
    parts.forEach(p => body.appendParagraph(`• ${usuariosMap[p.usuarioID] || 'N/A'} - ${p.rol}`));
    
    body.appendParagraph("\nASUNTO:").setBold(true).setFontSize(11);
    body.appendParagraph(consulta.temaOrdenDelDia);
    body.appendParagraph("\nACUERDOS Y COMPROMISOS:").setBold(true).setFontSize(11);
    body.appendParagraph(consulta.acuerdosCompromisos);
    
    let tableData = [];
    let row = [];
    parts.forEach((p, i) => {
        const nombre = usuariosMap[p.usuarioID] || 'Invitado';
        const rol = p.rol || 'Asistente';
        row.push(`${nombre}\n${rol}`);
        if ((i + 1) % 3 === 0) { tableData.push(row); row = []; }
    });
    if (row.length > 0) {
        while (row.length < 3) { row.push(""); }
        tableData.push(row);
    }
    
    body.appendParagraph("\n\n\n\n");
    const tablaFirmas = body.appendTable(tableData);
    tablaFirmas.setBorderWidth(0);
    
    tableData.forEach((r, rowIndex) => {
        const rowObj = tablaFirmas.getRow(rowIndex);
        for(let colIndex = 0; colIndex < 3; colIndex++) {
            const cell = rowObj.getCell(colIndex);
            if(cell && cell.getText().trim() !== "") {
                cell.appendParagraph("\n____________________\nFirma").setAlignment(DocumentApp.HorizontalAlignment.CENTER);
                cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            }
        }
    });
    doc.saveAndClose();
    
    const archivoDoc = DriveApp.getFileById(doc.getId());
    const carpetaDestino = getMeetFlowFolder("MeetFlow_Actas");
    const url = `https://docs.google.com/document/d/${doc.getId()}/export?format=pdf`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
    const blob = response.getBlob().setName(`${docName}.pdf`);
    const pdfFile = carpetaDestino.createFile(blob);
    
    archivoDoc.setTrashed(true);
    
    return { 
      success: true, 
      fileUrl: pdfFile.getUrl(), 
      message: "Acta generada con éxito."
    };
  } catch (error) {
    return { success: false, message: "Error: " + error.message };
  }
}

function getEmpresaNombre() {
  try {
    const config = JSON.parse(getSheetData('CONFIG'));
    const empConfig = config.find(c => c.tipoLista === 'Empresa');
    return empConfig ? empConfig.valorItem : "NOMBRE DE TU EMPRESA";
  } catch (e) {
    return "NOMBRE DE TU EMPRESA";
  }
}

function getMeetFlowFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}
function getFileForDownload(fileUrl) {
  try {
    const match = fileUrl.match(/[-\w]{25,}/);
    if (!match) throw new Error("URL no válida.");
    
    const fileId = match[0];
    const file = DriveApp.getFileById(fileId);
    const mimeType = file.getMimeType();
    
    let blob;
    let fileName = file.getName();
    let finalMimeType = mimeType;
    
    // Si el archivo es un Google Sheet, FORZAMOS la conversión a Excel (.xlsx)
    if (mimeType === MimeType.GOOGLE_SHEETS) {
      const url = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(url, { 
        headers: { 'Authorization': 'Bearer ' + token } 
      });
      
      blob = response.getBlob();
      fileName += '.xlsx'; // Le añadimos la extensión correcta
      finalMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    } 
    // Si es un archivo normal o el Acta PDF que ya generamos
    else {
      blob = file.getBlob();
    }
    
    // Convertimos a Base64 para enviarlo limpio al navegador del usuario
    const bytes = blob.getBytes();
    const base64 = Utilities.base64Encode(bytes);
    
    return { 
      success: true, 
      base64: base64, 
      mimeType: finalMimeType, 
      fileName: fileName 
    };
  } catch (error) {
    return { success: false, error: "Error preparando la descarga: " + error.message };
  }
}
// --- INICIO MÓDULO DE CONFIGURACIÓN DINÁMICA ---

function getAllConfigData() {
  try {
    return getSheetData('CONFIG');
  } catch (e) {
    logActivity('ERROR', 'getAllConfigData: ' + e.message);
    return JSON.stringify([]);
  }
}

function updateConfigParameter(tipo, valorNuevo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const config = SHEETS_CONFIG.CONFIG;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    const data = sheet.getDataRange().getValues();
    
    // Recorremos la hoja buscando el parámetro (ej: 'Empresa')
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === tipo) {
        sheet.getRange(i + 1, 2).setValue(valorNuevo);
        return { success: true, message: "El parámetro '" + tipo + "' fue actualizado correctamente." };
      }
    }
    
    // Si no lo encuentra, lo crea al final
    sheet.appendRow([tipo, valorNuevo]);
    return { success: true, message: "El parámetro '" + tipo + "' fue creado y guardado." };
  } catch (e) {
    logActivity('ERROR', 'updateConfigParameter: ' + e.message);
    return { success: false, message: "Error al actualizar la configuración: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function addConfigItem(tipo, valor) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!tipo || !valor) throw new Error("El tipo de lista y el valor son obligatorios.");
    const config = SHEETS_CONFIG.CONFIG;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    
    sheet.appendRow([tipo, valor]);
    return { success: true, message: "Se ha agregado '" + valor + "' a la lista de " + tipo + "." };
  } catch (e) {
    logActivity('ERROR', 'addConfigItem: ' + e.message);
    return { success: false, message: "Error al agregar ítem: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteConfigItem(tipo, valor) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const config = SHEETS_CONFIG.CONFIG;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sheetName);
    const data = sheet.getDataRange().getValues();
    
    let deleted = false;
    // Buscamos desde abajo hacia arriba para que los índices no cambien al eliminar filas
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === String(tipo).trim() && String(data[i][1]).trim() === String(valor).trim()) {
        sheet.deleteRow(i + 1);
        deleted = true;
        break; // Solo elimina la primera coincidencia exacta
      }
    }
    
    if (deleted) {
      return { success: true, message: "El ítem '" + valor + "' fue eliminado correctamente." };
    } else {
      return { success: false, message: "No se encontró el ítem en la base de datos." };
    }
  } catch (e) {
    logActivity('ERROR', 'deleteConfigItem: ' + e.message);
    return { success: false, message: "Error al eliminar ítem: " + e.message };
  } finally {
    lock.releaseLock();
  }
}