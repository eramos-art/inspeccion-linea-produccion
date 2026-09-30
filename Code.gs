/**
 * Grupo Sole · Backend de Inspección de Línea de Producción
 * -----------------------------------------------------------
 * Este script va PEGADO dentro de Extensiones > Apps Script
 * de tu Google Sheet (el que ya tienes creado). Al publicarlo
 * como "Aplicación web" obtienes una URL que se pega en el
 * archivo index.html (constante SCRIPT_URL).
 *
 * Crea automáticamente, si no existen:
 *  - Hoja "Inspecciones"
 *  - Hoja "Catalogo"
 *  - Carpeta de Drive "Inspecciones Grupo Sole - Evidencias" (fotos/videos/firmas)
 */

const SHEET_INSP = 'Inspecciones';
const SHEET_CAT  = 'Catalogo';
const FOLDER_NAME = 'Inspecciones Grupo Sole - Evidencias';

const COLS_INSP = ['ID','Fecha','Linea','Codigo','Lote','Muestra','HoraInicio','HoraFin','HorasHombre',
                    'Inspector','Encargado','Pruebas','Defectos','Detalle','Fotos','Firma','CreadoEn'];
const COLS_CAT  = ['ID','Tipo','Linea','Nombre','DefTipo'];

function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }

function sheetInsp(){
  let sh = ss().getSheetByName(SHEET_INSP);
  if(!sh){ sh = ss().insertSheet(SHEET_INSP); sh.appendRow(COLS_INSP); sh.setFrozenRows(1); sh.hideColumns(1); }
  return sh;
}
function sheetCat(){
  let sh = ss().getSheetByName(SHEET_CAT);
  if(!sh){ sh = ss().insertSheet(SHEET_CAT); sh.appendRow(COLS_CAT); sh.setFrozenRows(1); sh.hideColumns(1); }
  return sh;
}
function evidenceFolder(){
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function rowsToObjects(sheet, cols){
  const data = sheet.getDataRange().getValues();
  data.shift(); // encabezado
  return data.filter(r=>r[0]).map(r=>{
    const o={}; cols.forEach((c,i)=>o[c]=r[i]); return o;
  });
}

// ---- Lectura: GET -> {ok, inspecciones, catalogo} ----
function doGet(e){
  try{
    const inspecciones = rowsToObjects(sheetInsp(), COLS_INSP);
    const catalogo = rowsToObjects(sheetCat(), COLS_CAT);
    return out({ok:true, inspecciones, catalogo});
  }catch(err){ return out({ok:false, error:String(err)}); }
}

// ---- Escritura: POST {action, ...} ----
function doPost(e){
  try{
    const body = JSON.parse(e.postData.contents);
    switch(body.action){
      case 'add':        return out(addInspeccion(body.data));
      case 'delete':      return out(deleteInspeccion(body.id));
      case 'addCatalog':  return out(addCatalogo(body.data));
      case 'uploadFile':  return out(uploadFile(body));
      default: return out({ok:false, error:'Acción no reconocida: '+body.action});
    }
  }catch(err){ return out({ok:false, error:String(err)}); }
}

function addInspeccion(d){
  const id = Utilities.getUuid();
  sheetInsp().appendRow([
    id, d.fecha||'', d.linea||'', d.codigo||'', d.lote||'', d.muestra||0,
    d.horaInicio||'', d.horaFin||'', d.horasHombre||0, d.inspector||'', d.encargado||'',
    JSON.stringify(d.pruebas||[]), JSON.stringify(d.defectos||[]), JSON.stringify(d.detalle||null),
    JSON.stringify(d.media||[]), d.firma||'', new Date().toISOString()
  ]);
  return {ok:true, id};
}

function deleteInspeccion(id){
  const sh = sheetInsp();
  const data = sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(data[i][0]===id){ sh.deleteRow(i+1); return {ok:true}; }
  }
  return {ok:false, error:'No se encontró el registro'};
}

function addCatalogo(d){
  const existentes = sheetCat().getDataRange().getValues().slice(1);
  const dup = existentes.some(r=> r[1]===d.tipo && r[2]===d.linea && String(r[3]).toLowerCase()===String(d.nombre).toLowerCase());
  if(dup) return {ok:true, skipped:true};
  sheetCat().appendRow([Utilities.getUuid(), d.tipo||'', d.linea||'', d.nombre||'', d.defTipo||'']);
  return {ok:true};
}

// Recibe una imagen/video en base64 (dataURL), lo guarda en Drive y devuelve su URL pública de solo lectura
function uploadFile(body){
  const base64 = body.data.split(',').pop();
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, body.mimeType, body.filename);
  const file = evidenceFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = file.getId();
  return {ok:true, id:id, url:'https://drive.google.com/file/d/'+id+'/view', thumb:'https://lh3.googleusercontent.com/d/'+id};
}

function out(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
