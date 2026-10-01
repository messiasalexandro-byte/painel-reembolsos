/**
 * Painel de Reembolsos servido pelo Google Apps Script.
 *
 * doGet() entrega o painel (Index.html = cópia do index.html do repositório) e
 * obterPlanilha() — chamada pela página via google.script.run — devolve em base64
 * a planilha modificada mais recentemente na pasta abaixo: .xlsx enviado ou
 * Planilha Google (exportada como .xlsx). O acesso é limitado às contas do
 * domínio na implantação (appsscript.json → webapp.access = DOMAIN).
 *
 * Acesso: além do domínio, só entram os e-mails de PAINEL_EMAILS (vírgulas) e com a senha
 * PAINEL_SENHA — as duas em Configurações do projeto → Propriedades do script, nunca no código.
 * entrar(senha) devolve um token (HMAC do e-mail com a senha) que a página guarda e manda em
 * obterPlanilha; trocar a senha invalida todos os tokens de uma vez.
 */

var FOLDER_ID = '1XY4OOaDlrH7_d4BmsNmrMIoxrLx8UJXv';

var MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
var MIME_SHEETS = 'application/vnd.google-apps.spreadsheet';

var MAX_TENTATIVAS = 5;
var BLOQUEIO_SEG = 15 * 60;

function doGet(){
  var acesso = verificarConta();
  if(acesso.erro){
    var msg = acesso.erro === 'nao_configurado'
      ? 'O acesso a este painel ainda não foi configurado.'
      : 'Você não tem acesso a este painel' + (acesso.email ? ' com a conta ' + acesso.email : '') + '. Peça a liberação ao responsável.';
    return HtmlService.createHtmlOutput(
      '<div style="font-family:system-ui,sans-serif;max-width:480px;margin:15vh auto;padding:0 16px;color:#1f2933;">' +
      '<h1 style="font-size:20px;">Painel de Reembolsos</h1><p style="line-height:1.5;">' + escaparHtml(msg) + '</p></div>'
    ).setTitle('Painel de Reembolsos').addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Painel de Reembolsos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function escaparHtml(s){
  return String(s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; });
}

/* ===================== Acesso (lista de e-mails + senha) ===================== */

function configAcesso(){
  var props = PropertiesService.getScriptProperties();
  var senha = props.getProperty('PAINEL_SENHA') || '';
  var emails = (props.getProperty('PAINEL_EMAILS') || '').split(/[,;\s]+/)
    .map(function(e){ return e.trim().toLowerCase(); })
    .filter(function(e){ return e; });
  return {senha: senha, emails: emails};
}

function emailAtual(){
  return (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
}

// {email} se a conta está na lista; {erro} se não está ou se falta configurar.
function verificarConta(){
  var cfg = configAcesso();
  var email = emailAtual();
  if(!cfg.senha || !cfg.emails.length) return {erro:'nao_configurado', email:email};
  if(!email || cfg.emails.indexOf(email) < 0) return {erro:'sem_acesso', email:email};
  return {email:email, senha:cfg.senha};
}

function tokenPara(email, senha){
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(email, senha));
}

// comparação sem sair no primeiro caractere diferente
function iguais(a, b){
  a = String(a || ''); b = String(b || '');
  var dif = a.length ^ b.length;
  for(var i=0; i<Math.max(a.length, b.length); i++) dif |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return dif === 0;
}

function entrar(senha){
  var conta = verificarConta();
  if(conta.erro) return {erro:conta.erro};
  var cache = CacheService.getScriptCache();
  var chave = 'tentativas_' + conta.email;
  var erros = Number(cache.get(chave) || 0);
  if(erros >= MAX_TENTATIVAS) return {erro:'bloqueado'};
  if(!iguais(senha, conta.senha)){
    erros++;
    cache.put(chave, String(erros), BLOQUEIO_SEG);
    return erros >= MAX_TENTATIVAS ? {erro:'bloqueado'} : {erro:'senha_incorreta', restantes: MAX_TENTATIVAS - erros};
  }
  cache.remove(chave);
  return {token: tokenPara(conta.email, conta.senha)};
}

function tokenValido(token){
  var conta = verificarConta();
  return !conta.erro && iguais(token, tokenPara(conta.email, conta.senha));
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

// versao: modificadoEm do .xlsx que a página já tem; se não mudou, evita reenviar o arquivo.
// Planilha Google é sempre reenviada: a data de modificação dela no Drive só muda minutos depois
// da edição, e a página compara o conteúdo para decidir se redesenha.
function obterPlanilha(versao, token){
  if(!tokenValido(token)) return {erro:'senha_necessaria'};
  var arquivo = arquivoMaisRecente();
  if(!arquivo) return {erro:'pasta_vazia'};
  var modificadoEm = arquivo.getLastUpdated().toISOString();
  if(versao && versao === modificadoEm && arquivo.getMimeType() === MIME_XLSX) return {mudou:false, modificadoEm:modificadoEm};

  var nome = arquivo.getName();
  var bytes;
  if(arquivo.getMimeType() === MIME_SHEETS){
    var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + arquivo.getId() + '/export?format=xlsx', {
      headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()},
      muteHttpExceptions:true
    });
    if(resp.getResponseCode() !== 200) return {erro:'exportacao_falhou', detalhe:'HTTP ' + resp.getResponseCode()};
    bytes = resp.getBlob().getBytes();
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
