/**
 * Painel de Reembolsos servido pelo Google Apps Script.
 *
 * doGet() entrega o painel (Index.html = cópia do index.html do repositório) e
 * obterPlanilha() — chamada pela página via google.script.run — devolve em base64
 * a planilha modificada mais recentemente na pasta abaixo: .xlsx enviado ou
 * Planilha Google (exportada como .xlsx). O acesso é limitado às contas do
 * domínio na implantação (appsscript.json → webapp.access = DOMAIN).
 */

var FOLDER_ID = '1XY4OOaDlrH7_d4BmsNmrMIoxrLx8UJXv';

var MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
var MIME_SHEETS = 'application/vnd.google-apps.spreadsheet';

function doGet(){
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Painel de Reembolsos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// Planilha Google tem prioridade: ao converter um .xlsx, o original costuma ficar na pasta
// e não deve voltar a ser usado só porque alguém mexeu nele. Sem Planilha Google, usa o .xlsx mais recente.
function arquivoMaisRecente(){
  var maisRecente = {};
  var arquivos = DriveApp.getFolderById(FOLDER_ID).getFiles();
  while(arquivos.hasNext()){
    var f = arquivos.next();
    var tipo = f.getMimeType();
    if(tipo !== MIME_XLSX && tipo !== MIME_SHEETS) continue;
    if(!maisRecente[tipo] || f.getLastUpdated() > maisRecente[tipo].getLastUpdated()) maisRecente[tipo] = f;
  }
  return maisRecente[MIME_SHEETS] || maisRecente[MIME_XLSX] || null;
}

// versao: modificadoEm da planilha que a página já tem; se não mudou, evita reenviar o arquivo.
function obterPlanilha(versao){
  var arquivo = arquivoMaisRecente();
  if(!arquivo) return {erro:'pasta_vazia'};
  var modificadoEm = arquivo.getLastUpdated().toISOString();
  if(versao && versao === modificadoEm) return {mudou:false, modificadoEm:modificadoEm};

  var nome = arquivo.getName();
  var bytes;
  if(arquivo.getMimeType() === MIME_SHEETS){
    var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + arquivo.getId() + '/export?format=xlsx', {
      headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()},
      muteHttpExceptions:true
    });
    if(resp.getResponseCode() !== 200) return {erro:'exportacao_falhou', detalhe:'HTTP ' + resp.getResponseCode()};
    bytes = resp.getBlob().getBytes();
    if(!/\.xlsx$/i.test(nome)) nome += '.xlsx';
  } else {
    bytes = arquivo.getBlob().getBytes();
  }

  return {
    mudou: true,
    nome: nome,
    modificadoEm: modificadoEm,
    tamanho: bytes.length,
    tipo: arquivo.getMimeType() === MIME_SHEETS ? 'planilha_google' : 'xlsx',
    base64: Utilities.base64Encode(bytes)
  };
}
