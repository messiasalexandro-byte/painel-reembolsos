/**
 * Painel de Reembolsos — fonte automática pelo Google Drive.
 *
 * Publicado como app da web, devolve em JSON (base64) a planilha modificada
 * mais recentemente na pasta abaixo: .xlsx enviado ou Planilha Google
 * (exportada como .xlsx). Passo a passo no README do repositório.
 */

var FOLDER_ID = '1XY4OOaDlrH7_d4BmsNmrMIoxrLx8UJXv';

var MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
var MIME_SHEETS = 'application/vnd.google-apps.spreadsheet';

// Execute uma vez pelo editor (Executar > configurarChave) e copie a chave do registro de execução.
function configurarChave(){
  var props = PropertiesService.getScriptProperties();
  var chave = props.getProperty('TOKEN');
  if(!chave){
    chave = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty('TOKEN', chave);
  }
  Logger.log('Chave de acesso: ' + chave);
  return chave;
}

function doGet(e){
  var esperado = PropertiesService.getScriptProperties().getProperty('TOKEN');
  var recebido = e && e.parameter ? e.parameter.token : '';
  if(!esperado || recebido !== esperado) return responder({erro:'nao_autorizado'});

  var escolhido = null;
  var arquivos = DriveApp.getFolderById(FOLDER_ID).getFiles();
  while(arquivos.hasNext()){
    var f = arquivos.next();
    var tipo = f.getMimeType();
    if(tipo !== MIME_XLSX && tipo !== MIME_SHEETS) continue;
    if(!escolhido || f.getLastUpdated() > escolhido.getLastUpdated()) escolhido = f;
  }
  if(!escolhido) return responder({erro:'pasta_vazia'});

  var nome = escolhido.getName();
  var bytes;
  if(escolhido.getMimeType() === MIME_SHEETS){
    var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + escolhido.getId() + '/export?format=xlsx', {
      headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()},
      muteHttpExceptions:true
    });
    if(resp.getResponseCode() !== 200) return responder({erro:'exportacao_falhou', detalhe:'HTTP ' + resp.getResponseCode()});
    bytes = resp.getBlob().getBytes();
    if(!/\.xlsx$/i.test(nome)) nome += '.xlsx';
  } else {
    bytes = escolhido.getBlob().getBytes();
  }

  return responder({
    nome: nome,
    modificadoEm: escolhido.getLastUpdated().toISOString(),
    tamanho: bytes.length,
    tipo: escolhido.getMimeType() === MIME_SHEETS ? 'planilha_google' : 'xlsx',
    base64: Utilities.base64Encode(bytes)
  });
}

function responder(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
