/**
 * Painel de Reembolsos servido pelo Google Apps Script.
 *
 * doGet() entrega o painel (Index.html = cópia do index.html do repositório) e
 * obterPlanilha() — chamada pela página via google.script.run — devolve em base64
 * a planilha modificada mais recentemente na pasta abaixo: .xlsx enviado ou
 * Planilha Google (exportada como .xlsx). O acesso é limitado às contas do
 * domínio na implantação (appsscript.json → webapp.access = DOMAIN).
 *
 * Acesso: além do domínio, só entram os e-mails cadastrados (propriedade "usuario:<e-mail>").
 * O administrador (dono do script e quem estiver em PAINEL_ADMINS) adiciona o e-mail pela tela
 * Acesso do painel; no primeiro acesso a pessoa cria a própria senha (guardada só como hash).
 * entrar()/criarSenha() devolvem um token (HMAC do e-mail + hash com PAINEL_SEGREDO) que a página
 * guarda e manda em cada chamada; trocar ou liberar de novo a senha invalida o token daquela pessoa.
 */

var FOLDER_ID = '1XY4OOaDlrH7_d4BmsNmrMIoxrLx8UJXv';

var MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
var MIME_SHEETS = 'application/vnd.google-apps.spreadsheet';

var MAX_TENTATIVAS = 5;
var BLOQUEIO_SEG = 15 * 60;
var SENHA_MIN = 8;
var HASH_RODADAS = 1000;
var DOMINIO = 'talgui.com.br';
var PREFIXO_USUARIO = 'usuario:';

function doGet(){
  var conta = verificarConta();
  if(conta.erro){
    var msg = 'Você não tem acesso a este painel' + (conta.email ? ' com a conta ' + conta.email : '') + '. Peça a liberação ao responsável.';
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

/* ===================== Acesso (cadastro por e-mail + senha de cada um) ===================== */

function props(){ return PropertiesService.getScriptProperties(); }

function normalizarEmail(e){ return String(e || '').trim().toLowerCase(); }

function emailAtual(){ return normalizarEmail(Session.getActiveUser().getEmail()); }

function emailDono(){ return normalizarEmail(Session.getEffectiveUser().getEmail()); }

function ehAdmin(email){
  if(!email) return false;
  if(email === emailDono()) return true;
  return (props().getProperty('PAINEL_ADMINS') || '').split(/[,;\s]+/).map(normalizarEmail).indexOf(email) >= 0;
}

// Lista antiga (PAINEL_EMAILS, da senha única): vira, uma vez só, cadastro pendente para cada e-mail.
// Não apaga PAINEL_EMAILS/PAINEL_SENHA (podem ser removidas à mão depois de publicar a versão nova).
function migrarListaAntiga(){
  var p = props();
  if(p.getProperty('PAINEL_MIGRADO') !== null) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    if(p.getProperty('PAINEL_MIGRADO') !== null) return;
    (p.getProperty('PAINEL_EMAILS') || '').split(/[,;\s]+/).map(normalizarEmail).forEach(function(e){
      if(e && p.getProperty(PREFIXO_USUARIO + e) === null) p.setProperty(PREFIXO_USUARIO + e, JSON.stringify({adicionadoEm: new Date().toISOString()}));
    });
    p.setProperty('PAINEL_MIGRADO', new Date().toISOString());
  } finally { lock.releaseLock(); }
}

// cadastro da pessoa: {hash, sal, senhaEm} depois de criar a senha; {adicionadoEm} enquanto aguarda.
// foto (opcional): dataURL JPEG 64×64 (até 4000 caracteres) enviado pela página Minha conta; cadastro sem ela continua valendo.
// O administrador sem cadastro conta como pendente, para nunca ficar trancado do lado de fora.
function lerUsuario(email){
  var raw = props().getProperty(PREFIXO_USUARIO + email);
  if(raw === null) return ehAdmin(email) ? {} : null;
  try{ return JSON.parse(raw); } catch(e){ return {}; }
}

function gravarUsuario(email, u){ props().setProperty(PREFIXO_USUARIO + email, JSON.stringify(u)); }

// {email, usuario, admin} se a conta tem cadastro; {erro} se não tem.
function verificarConta(){
  migrarListaAntiga();
  var email = emailAtual();
  var u = email ? lerUsuario(email) : null;
  if(!u) return {erro:'sem_acesso', email:email};
  return {email:email, usuario:u, admin:ehAdmin(email)};
}

function segredo(){
  var p = props();
  var s = p.getProperty('PAINEL_SEGREDO');
  if(!s){ s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('PAINEL_SEGREDO', s); }
  return s;
}

function bytes(s){ return Utilities.newBlob(String(s)).getBytes(); }

function hashSenha(senha, sal){
  var chave = bytes(sal);
  var d = Utilities.computeHmacSha256Signature(bytes(senha), chave);
  for(var i=1; i<HASH_RODADAS; i++) d = Utilities.computeHmacSha256Signature(d, chave);
  return Utilities.base64Encode(d);
}

function tokenPara(email, u){
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(email + '|' + u.hash, segredo()));
}

// comparação sem sair no primeiro caractere diferente
function iguais(a, b){
  a = String(a || ''); b = String(b || '');
  var dif = a.length ^ b.length;
  for(var i=0; i<Math.max(a.length, b.length); i++) dif |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return dif === 0;
}

// conta já conferida e com token válido; null caso contrário
function contaComToken(token){
  var conta = verificarConta();
  if(conta.erro || !conta.usuario.hash || !iguais(token, tokenPara(conta.email, conta.usuario))) return null;
  return conta;
}

// o que a tela de entrada deve mostrar: criar a senha (primeiro acesso) ou digitá-la
function estadoAcesso(){
  var conta = verificarConta();
  if(conta.erro) return {erro:conta.erro};
  return {email:conta.email, cadastro: conta.usuario.hash ? 'entrar' : 'criar'};
}

function criarSenha(senha){
  var conta = verificarConta();
  if(conta.erro) return {erro:conta.erro};
  if(conta.usuario.hash) return {erro:'ja_cadastrado'};
  senha = String(senha || '');
  if(senha.length < SENHA_MIN) return {erro:'senha_curta', minimo:SENHA_MIN};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    var atual = lerUsuario(conta.email);
    if(!atual) return {erro:'sem_acesso'};
    if(atual.hash) return {erro:'ja_cadastrado'};
    var sal = Utilities.getUuid();
    atual.sal = sal;
    atual.hash = hashSenha(senha, sal);
    atual.senhaEm = new Date().toISOString();
    gravarUsuario(conta.email, atual);
    return {token: tokenPara(conta.email, atual), admin: conta.admin, foto: atual.foto || ''};
  } finally { lock.releaseLock(); }
}

function entrar(senha){
  var conta = verificarConta();
  if(conta.erro) return {erro:conta.erro};
  if(!conta.usuario.hash) return {erro:'criar_senha'};
  var cache = CacheService.getScriptCache();
  var chave = 'tentativas_' + conta.email;
  var erros = Number(cache.get(chave) || 0);
  if(erros >= MAX_TENTATIVAS) return {erro:'bloqueado'};
  if(!iguais(hashSenha(String(senha || ''), conta.usuario.sal), conta.usuario.hash)){
    erros++;
    cache.put(chave, String(erros), BLOQUEIO_SEG);
    return erros >= MAX_TENTATIVAS ? {erro:'bloqueado'} : {erro:'senha_incorreta', restantes: MAX_TENTATIVAS - erros};
  }
  cache.remove(chave);
  return {token: tokenPara(conta.email, conta.usuario), admin: conta.admin, foto: conta.usuario.foto || ''};
}

// a página confere, ao abrir com um token guardado, se ele ainda vale e se é admin (e pega a foto do menu)
function sessao(token){
  var conta = contaComToken(token);
  return conta ? {email:conta.email, admin:conta.admin, foto:conta.usuario.foto || ''} : {erro:'senha_necessaria'};
}

/* ----- Minha conta (cada pessoa, com a própria sessão) ----- */

var FOTO_PREFIXO = 'data:image/jpeg;base64,';
var FOTO_MAX = 4000; // cada propriedade aceita ~9 KB e todas juntas no máximo 500 KB (~100 pessoas com foto)

// dataUrl '' remove a foto; senão, só JPEG em base64 até FOTO_MAX caracteres
function salvarFoto(token, dataUrl){
  var conta = contaComToken(token);
  if(!conta) return {erro:'senha_necessaria'};
  var foto = String(dataUrl || '');
  if(foto && (foto.length > FOTO_MAX || foto.indexOf(FOTO_PREFIXO) !== 0 ||
      !/^[A-Za-z0-9+\/]+={0,2}$/.test(foto.slice(FOTO_PREFIXO.length)))) return {erro:'foto_invalida'};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    var atual = lerUsuario(conta.email);
    // a senha pode ter sido trocada ou liberada enquanto a página estava aberta
    if(!atual || !atual.hash || atual.hash !== conta.usuario.hash) return {erro:'senha_necessaria'};
    if(foto) atual.foto = foto; else delete atual.foto;
    gravarUsuario(conta.email, atual);
    return {ok:true, foto:foto};
  } finally { lock.releaseLock(); }
}

// confere a senha atual como entrar() (mesmo limite de tentativas) e grava o hash novo com sal novo.
// O token é HMAC do e-mail + hash: o novo vale neste aparelho e os outros passam a pedir a senha.
function trocarSenha(token, atualSenha, novaSenha){
  var conta = contaComToken(token);
  if(!conta) return {erro:'senha_necessaria'};
  var cache = CacheService.getScriptCache();
  var chave = 'tentativas_' + conta.email;
  var erros = Number(cache.get(chave) || 0);
  if(erros >= MAX_TENTATIVAS) return {erro:'bloqueado'};
  if(!iguais(hashSenha(String(atualSenha || ''), conta.usuario.sal), conta.usuario.hash)){
    erros++;
    cache.put(chave, String(erros), BLOQUEIO_SEG);
    return erros >= MAX_TENTATIVAS ? {erro:'bloqueado'} : {erro:'senha_incorreta', restantes: MAX_TENTATIVAS - erros};
  }
  cache.remove(chave);
  novaSenha = String(novaSenha || '');
  if(novaSenha.length < SENHA_MIN) return {erro:'senha_curta', minimo:SENHA_MIN};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    var atual = lerUsuario(conta.email);
    if(!atual || !atual.hash || atual.hash !== conta.usuario.hash) return {erro:'senha_necessaria'};
    var sal = Utilities.getUuid();
    atual.sal = sal;
    atual.hash = hashSenha(novaSenha, sal);
    atual.senhaEm = new Date().toISOString();
    gravarUsuario(conta.email, atual);
    return {token: tokenPara(conta.email, atual)};
  } finally { lock.releaseLock(); }
}

/* ----- Tela Acesso (só administradores) ----- */

function contaAdmin(token){
  var conta = contaComToken(token);
  return conta && conta.admin ? conta : null;
}

function adminListar(token){
  if(!contaAdmin(token)) return {erro:'senha_necessaria'};
  var todas = props().getProperties();
  var dono = emailDono();
  var lista = [];
  Object.keys(todas).forEach(function(k){
    if(k.indexOf(PREFIXO_USUARIO) !== 0) return;
    var email = k.slice(PREFIXO_USUARIO.length), u = {};
    try{ u = JSON.parse(todas[k]); } catch(e){}
    lista.push({email:email, situacao: u.hash ? 'ativo' : 'aguardando', desde: u.hash ? u.senhaEm : u.adicionadoEm, admin: ehAdmin(email), dono: email === dono});
  });
  lista.sort(function(a, b){ return a.email < b.email ? -1 : 1; });
  return {usuarios:lista};
}

function adminAdicionar(token, email){
  if(!contaAdmin(token)) return {erro:'senha_necessaria'};
  email = normalizarEmail(email);
  if(!/^[^@\s]+@[^@\s]+$/.test(email) || email.split('@')[1] !== DOMINIO) return {erro:'email_invalido'};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    if(props().getProperty(PREFIXO_USUARIO + email) !== null) return {erro:'ja_existe'};
    gravarUsuario(email, {adicionadoEm: new Date().toISOString()});
  } finally { lock.releaseLock(); }
  return adminListar(token);
}

// apaga a senha da pessoa: no próximo acesso ela cria outra (e o token antigo deixa de valer)
function adminLiberarCadastro(token, email){
  var conta = contaAdmin(token);
  if(!conta) return {erro:'senha_necessaria'};
  email = normalizarEmail(email);
  if(email === conta.email) return {erro:'proprio_usuario'};
  if(props().getProperty(PREFIXO_USUARIO + email) === null) return {erro:'nao_encontrado'};
  var novo = {adicionadoEm: new Date().toISOString()};
  var antigo = lerUsuario(email);
  if(antigo && antigo.foto) novo.foto = antigo.foto; // a foto da pessoa sobrevive ao novo cadastro
  gravarUsuario(email, novo);
  CacheService.getScriptCache().remove('tentativas_' + email);
  return adminListar(token);
}

function adminRemover(token, email){
  var conta = contaAdmin(token);
  if(!conta) return {erro:'senha_necessaria'};
  email = normalizarEmail(email);
  if(email === conta.email) return {erro:'proprio_usuario'};
  if(email === emailDono()) return {erro:'dono'};
  props().deleteProperty(PREFIXO_USUARIO + email);
  return adminListar(token);
}

/* ===================== Planilha do Drive ===================== */

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
  if(!contaComToken(token)) return {erro:'senha_necessaria'};
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
