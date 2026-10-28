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
    ).setTitle('Painel Financeiro · Talgui').addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Painel Financeiro · Talgui')
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

// Recebimentos do Nuvem Pago (saldo e transferências bancárias): só o dono do script e quem estiver em
// NUVEMPAGO_LEITORES (e-mails separados por vírgula). Os demais seguem vendo as vendas completadas com as taxas.
function podeVerFinanceiro(email){
  if(!email) return false;
  if(email === emailDono()) return true;
  return (props().getProperty('NUVEMPAGO_LEITORES') || '').split(/[,;\s]+/).map(normalizarEmail).indexOf(email) >= 0;
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
    return {token: tokenPara(conta.email, atual), admin: conta.admin, financeiro: podeVerFinanceiro(conta.email), foto: atual.foto || ''};
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
  return {token: tokenPara(conta.email, conta.usuario), admin: conta.admin, financeiro: podeVerFinanceiro(conta.email), foto: conta.usuario.foto || ''};
}

// a página confere, ao abrir com um token guardado, se ele ainda vale e se é admin (e pega a foto do menu)
function sessao(token){
  var conta = contaComToken(token);
  return conta ? {email:conta.email, admin:conta.admin, financeiro:podeVerFinanceiro(conta.email), foto:conta.usuario.foto || ''} : {erro:'senha_necessaria'};
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

/* ===================== Vendas (exportações da Nuvemshop) ===================== */

// A pasta "nuvemshop" do Drive guarda os CSVs exportados no admin (Vendas → Exportar), na raiz ou em
// subpastas (ex.: "historico"). Cada CSV é lido UMA vez e entra numa base compacta por ano, gravada em
// nuvemshop/_base/vendas-AAAA.json. Os pedidos são juntados pelo número: vale o do arquivo modificado
// por último (o status muda com o tempo e as exportações se sobrepõem em datas).
// Só as colunas abaixo entram na base e saem do servidor — nada de nome, e-mail, CPF, telefone ou endereço.
// O CSV vem com uma linha por produto: os dados do pedido ficam na primeira linha dele.
var VENDAS_PASTA_NOME = 'nuvemshop';
var VENDAS_BASE_PASTA = '_base';
var VENDAS_COLUNAS = ['Número do Pedido','Data','Status do Pedido','Status do Pagamento','Status do Envio',
  'Subtotal','Desconto','Valor do Frete','Total','Estado','Forma de Entrega','Forma de Pagamento','Cupom de Desconto',
  'Data de pagamento','Canal','Data e hora do cancelamento','Motivo do cancelamento','Estorno','Data de estorno',
  'Taxas','Parcelas','Juros','Devolução dos custos','Total líquido','Meio de pagamento'];
var VENDAS_ITENS = VENDAS_COLUNAS.length; // posição da soma de "Quantidade Comprada" na linha do pedido
var VENDAS_SKUS = VENDAS_ITENS + 1;       // "sku:qtd,sku:qtd" do pedido (só na base: alimenta a taxa de devolução por peça)
var VENDAS_TS = VENDAS_SKUS + 1;          // posição da data de modificação do CSV de origem (só na base)
var VENDAS_FORMATO = 2;                   // muda quando a linha da base muda: força reler todos os CSVs uma vez
var VENDAS_PASTAS_FORA = ['_base', 'AJD']; // subpastas de "nuvemshop" que não têm CSV de vendas (AJD = devoluções)
// na resposta, textos repetidos viram índice de um dicionário e valores viram número (resposta ~3x menor)
var VENDAS_CODIFICADAS = [2, 3, 4, 9, 10, 11, 14, 16, 24];
var VENDAS_NUMERICAS = [5, 6, 7, 8, 17, 19, 21, 22, 23];
var VENDAS_TEMPO_MAX = 4 * 60 * 1000; // o Apps Script para em 6 min: guarda o progresso antes disso

// pasta pela propriedade NUVEMSHOP_PASTA (id) ou pelo nome; o id achado fica guardado
function pastaVendas(){
  var p = props();
  var id = p.getProperty('NUVEMSHOP_PASTA');
  if(id){ try{ return DriveApp.getFolderById(id); } catch(e){} }
  var achada = null;
  var it = DriveApp.searchFolders('title contains "' + VENDAS_PASTA_NOME + '" and trashed = false');
  while(it.hasNext()){
    var f = it.next();
    if(f.getName().trim().toLowerCase() !== VENDAS_PASTA_NOME) continue;
    if(!achada || f.getLastUpdated() > achada.getLastUpdated()) achada = f;
  }
  if(achada) p.setProperty('NUVEMSHOP_PASTA', achada.getId());
  return achada;
}

function pastaBaseVendas(pasta){
  var it = pasta.getFoldersByName(VENDAS_BASE_PASTA);
  return it.hasNext() ? it.next() : pasta.createFolder(VENDAS_BASE_PASTA);
}

// CSVs da pasta e das subpastas (menos _base e AJD), do mais antigo para o mais novo
function arquivosVendas(pasta){
  var lista = [];
  (function varrer(p){
    var it = p.getFiles();
    while(it.hasNext()){
      var f = it.next();
      if(/\.csv$/i.test(f.getName()) || /csv/i.test(f.getMimeType())) lista.push(f);
    }
    var subs = p.getFolders();
    while(subs.hasNext()){ var s = subs.next(); if(VENDAS_PASTAS_FORA.indexOf(s.getName()) < 0) varrer(s); }
  })(pasta);
  lista.sort(function(a, b){ return a.getLastUpdated() - b.getLastUpdated(); });
  return lista;
}

// lê um CSV da Nuvemshop (ISO-8859-1, separado por ";") e devolve uma linha por pedido
// com as VENDAS_COLUNAS + a soma de "Quantidade Comprada" (itens) + os SKUs; e o nome de cada SKU
function lerCsvVendas(arquivo){
  var blob = arquivo.getBlob();
  var texto = blob.getDataAsString('ISO-8859-1');
  if(texto.indexOf('Número do Pedido') < 0) texto = blob.getDataAsString('UTF-8').replace(/^﻿/, '');
  var linhas = Utilities.parseCsv(texto, ';');
  if(!linhas.length) return {erro:'csv_vazio'};
  var cab = linhas[0].map(function(c){ return String(c).trim(); });
  var idx = VENDAS_COLUNAS.map(function(c){ return cab.indexOf(c); });
  if(idx[0] < 0 || idx[1] < 0 || idx[8] < 0) return {erro:'csv_formato'};
  var iQtd = cab.indexOf('Quantidade Comprada'), iSku = cab.indexOf('SKU'), iNome = cab.indexOf('Nome do Produto');
  var porNum = {}, ordem = [], nomes = {};
  for(var i = 1; i < linhas.length; i++){
    var l = linhas[i];
    var num = String(l[idx[0]] || '').trim();
    if(!num) continue;
    var qtd = iQtd >= 0 ? Number(l[iQtd]) || 0 : 0;
    var sku = iSku >= 0 ? String(l[iSku] || '').trim().replace(/[,:]/g, '') : '';
    if(sku && iNome >= 0 && l[iNome]) nomes[sku] = String(l[iNome]).trim();
    var item = sku ? sku + ':' + qtd : '';
    if(porNum[num]){
      porNum[num][VENDAS_ITENS] += qtd;
      if(item) porNum[num][VENDAS_SKUS] += (porNum[num][VENDAS_SKUS] ? ',' : '') + item;
      continue;
    }
    var r = idx.map(function(j){ return j >= 0 ? String(l[j] == null ? '' : l[j]).trim() : ''; });
    r.push(qtd, item);
    porNum[num] = r;
    ordem.push(num);
  }
  return {linhas: ordem.map(function(n){ return porNum[n]; }), nomes: nomes};
}

// "16/08/2026 22:09:01" → "2026-08-16" ('' se não for data)
function diaIsoVendas(s){
  var m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(String(s || ''));
  return m ? m[3] + '-' + m[2] + '-' + m[1] : '';
}

// junta as linhas de um CSV (ts = modificação dele) ou da API (ts = updated_at) na base de um ano: o pedido
// só é trocado se a fonte for tão nova quanto o que já está lá — assim a ordem de leitura não importa.
// Campo vazio na fonte nova não apaga o da antiga (a API não traz Taxas/Juros/Líquido, que vêm do CSV).
function juntarPedidos(pedidos, linhas, ts){
  linhas.forEach(function(r){
    var atual = pedidos[r[0]];
    if(atual && atual[VENDAS_TS] > ts) return;
    var x = r.slice(0, VENDAS_TS);
    if(atual) x = x.map(function(v, i){ return (v === '' || v == null) && atual[i] !== '' && atual[i] != null ? atual[i] : v; });
    x[VENDAS_TS] = ts;
    pedidos[r[0]] = x;
  });
}

function estadoVendas(){
  var e = null;
  try{ e = JSON.parse(props().getProperty('VENDAS_BASE') || 'null'); } catch(err){}
  e = e || {};
  e.arqs = e.arqs || {}; e.info = e.info || {}; e.avisos = e.avisos || {}; e.anos = e.anos || {};
  return e;
}

function lerBaseAno(estado, ano){
  var id = estado.anos[ano];
  if(!id) return {};
  try{ return JSON.parse(DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8')).pedidos || {}; }
  catch(e){ return {}; }
}

// grava a base do ano como arquivo novo (blob: sem o limite de 10 MB do createFile com texto) e
// manda o anterior para a lixeira; ano sem pedidos sai da base
function gravarBaseAno(estado, pastaBase, ano, pedidos){
  var antigo = estado.anos[ano];
  if(Object.keys(pedidos).length){
    var blob = Utilities.newBlob(JSON.stringify({ano:Number(ano), pedidos:pedidos}), 'application/json', 'vendas-' + ano + '.json');
    estado.anos[ano] = pastaBase.createFile(blob).getId();
  } else delete estado.anos[ano];
  if(antigo){ try{ DriveApp.getFileById(antigo).setTrashed(true); } catch(e){} }
}

// lê os CSVs ainda não consolidados. Devolve {estado} pronto ou {erro:'consolidando', feitos, total}
// quando o tempo acabou (a próxima chamada continua de onde parou).
function consolidarVendas(pasta){
  var arquivos = arquivosVendas(pasta);
  var estado = estadoVendas();
  var atuais = {};
  arquivos.forEach(function(f){ atuais[f.getId()] = f.getLastUpdated().getTime(); });
  var removido = estado.formato !== VENDAS_FORMATO || Object.keys(estado.arqs).some(function(id){ return !(id in atuais); });
  var falta = arquivos.filter(function(f){ return estado.arqs[f.getId()] !== atuais[f.getId()]; });
  if(!removido && !falta.length) return {estado:estado, arquivos:arquivos.length};

  var lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return {erro:'consolidando', feitos:null, total:arquivos.length};
  try{
    estado = estadoVendas(); // outra execução pode ter avançado enquanto esta esperava
    var inicio = Date.now(), pastaBase = pastaBaseVendas(pasta), bases = {};
    removido = estado.formato !== VENDAS_FORMATO || Object.keys(estado.arqs).some(function(id){ return !(id in atuais); });
    if(removido){
      // um CSV saiu da pasta (ou a base é de um formato antigo): refaz do zero, para não sobrar pedido de arquivo apagado
      Object.keys(estado.anos).forEach(function(ano){ bases[ano] = {}; });
      estado.arqs = {}; estado.info = {}; estado.avisos = {}; estado.formato = VENDAS_FORMATO;
      Object.keys(bases).forEach(function(ano){ gravarBaseAno(estado, pastaBase, ano, bases[ano]); });
      props().setProperty('VENDAS_BASE', JSON.stringify(estado));
      props().deleteProperty('VENDAS_API'); // a base zerou: a API recarrega tudo desde o início da loja
    }
    falta = arquivos.filter(function(f){ return estado.arqs[f.getId()] !== atuais[f.getId()]; });
    var feitos = 0, nomes = null;
    for(var i = 0; i < falta.length; i++){
      if(feitos && Date.now() - inicio > VENDAS_TEMPO_MAX) break;
      var f = falta[i], id = f.getId(), ts = atuais[id];
      var lido = lerCsvVendas(f);
      delete estado.avisos[id];
      if(lido.erro){
        estado.avisos[id] = {nome:f.getName(), erro:lido.erro};
        delete estado.info[id];
      } else {
        var porAno = {}, de = '', ate = '';
        lido.linhas.forEach(function(r){
          var dia = diaIsoVendas(r[1]);
          if(!dia) return;
          if(!de || dia < de) de = dia;
          if(!ate || dia > ate) ate = dia;
          (porAno[dia.slice(0, 4)] = porAno[dia.slice(0, 4)] || []).push(r);
        });
        Object.keys(porAno).forEach(function(ano){
          if(!bases[ano]) bases[ano] = lerBaseAno(estado, ano);
          juntarPedidos(bases[ano], porAno[ano], ts);
          gravarBaseAno(estado, pastaBase, ano, bases[ano]);
        });
        estado.info[id] = {nome:f.getName(), modificadoEm:new Date(ts).toISOString(), pedidos:lido.linhas.length, de:de, ate:ate};
        if(!nomes) nomes = lerJsonBase(pastaBase, PRODUTOS_ARQ) || {};
        Object.keys(lido.nomes).forEach(function(sku){ nomes[sku] = lido.nomes[sku]; });
        gravarJsonBase(pastaBase, PRODUTOS_ARQ, nomes);
      }
      estado.arqs[id] = ts;
      // progresso salvo arquivo a arquivo: se a execução cair no meio, nada do que já foi lido se perde
      props().setProperty('VENDAS_BASE', JSON.stringify(estado));
      feitos++;
    }
    var restam = falta.length - feitos;
    if(restam > 0) return {erro:'consolidando', feitos:arquivos.length - restam, total:arquivos.length};
    return {estado:estado, arquivos:arquivos.length};
  } finally { lock.releaseLock(); }
}

// pedidos da base → formato da resposta (sem os SKUs e a coluna ts, textos repetidos num dicionário)
function codificarVendas(pedidos){
  var dic = {}, idx = {};
  VENDAS_CODIFICADAS.forEach(function(c){ dic[c] = []; idx[c] = Object.create(null); });
  var linhas = Object.keys(pedidos).map(function(k){
    var x = pedidos[k].slice(0, VENDAS_SKUS);
    VENDAS_CODIFICADAS.forEach(function(c){
      var v = x[c] == null ? '' : String(x[c]);
      var i = idx[c][v];
      if(i === undefined){ i = idx[c][v] = dic[c].length; dic[c].push(v); }
      x[c] = i;
    });
    VENDAS_NUMERICAS.forEach(function(c){
      if(x[c] === '' || x[c] == null) return;
      var n = Number(x[c]);
      if(isFinite(n)) x[c] = n;
    });
    return x;
  });
  return {dic:dic, linhas:linhas};
}

// ano: o que a página quer ver (padrão: o mais recente). versao: a que a página já tem; se nenhum CSV
// mudou, responde {mudou:false} sem reenviar os pedidos.
function obterVendas(token, versao, ano){
  var inicioChamada = Date.now();
  if(!contaComToken(token)) return {erro:'senha_necessaria'};
  var pasta = pastaVendas();
  if(!pasta) return {erro:'pasta_vendas'};
  var c = consolidarVendas(pasta);
  if(c.erro) return c;
  var api = null;
  if(apiNuvemshopLigada()){
    var s = estadoVendasApi();
    if(!s.ultimo || s.pendente || Date.now() - (s.verificadoEm || 0) > VENDAS_API_INTERVALO) s = sincronizarVendasApi(inicioChamada);
    if(s.erro === 'consolidando') return {erro:'consolidando', feitos:null, total:null, api:true};
    api = {ultimo:s.ultimo || null, verificadoEm:s.verificadoEm || null, erro:s.erro || null};
  }
  if(!c.arquivos && !api) return {erro:'sem_csv'};
  var estado = estadoVendas();
  var anos = Object.keys(estado.anos).map(Number).sort(function(a, b){ return a - b; });
  var avisos = Object.keys(estado.avisos).map(function(id){ return estado.avisos[id]; });
  if(!anos.length) return {erro:'csv_formato', avisos:avisos};
  ano = Number(ano);
  if(anos.indexOf(ano) < 0) ano = anos[anos.length - 1];
  // npVersao: extratos do Nuvem Pago completam Taxas/Juros/Líquido das vendas que vieram só da API
  var npVersao = props().getProperty('NUVEMPAGO_VERSAO') || '';
  var assinatura = Object.keys(estado.arqs).sort().map(function(id){ return id + '@' + estado.arqs[id]; }).join('|') + '|' + ano +
    (api ? '|api@' + (estado.apiMudou || '') : '') + (npVersao ? '|np@' + npVersao : '');
  var novaVersao = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, assinatura));
  if(versao && versao === novaVersao) return {mudou:false, versao:novaVersao, ano:ano, api:api};
  var pedidos = lerBaseAno(estado, ano);
  if(npVersao){
    try{ completarComNuvemPago(pedidos, financeiroNuvemPago(lerBaseNuvemPago(pastaBaseVendas(pasta)))); }
    catch(e){ console.warn('Nuvem Pago nas vendas: ' + (e && e.message || e)); }
  }
  var cod = codificarVendas(pedidos);
  var arquivos = Object.keys(estado.info).map(function(id){ return estado.info[id]; })
    .sort(function(a, b){ return a.modificadoEm < b.modificadoEm ? -1 : 1; });
  return {mudou:true, versao:novaVersao, ano:ano, anos:anos, colunas:VENDAS_COLUNAS.concat(['Itens']),
    dic:cod.dic, linhas:cod.linhas, arquivos:arquivos, avisos:avisos, api:api};
}

/* ===================== Vendas pela API da Nuvemshop ===================== */

// Com as propriedades NUVEMSHOP_TOKEN (access_token) e NUVEMSHOP_LOJA (store_id) preenchidas no editor,
// o servidor busca os pedidos alterados desde a última vez (updated_at_min) e os junta na mesma base dos CSVs.
// Roda pelo gatilho de 15 min (instalarSincronizacaoVendas) e, se a última busca tiver mais de 5 min,
// também quando alguém abre as vendas. A primeira carga vai desde o início da loja e continua de onde
// parou se o tempo acabar (VENDAS_API.pendente). Os CSVs seguem aceitos: completam Taxas, Juros e Líquido.
// Estado em VENDAS_API: {ultimo, verificadoEm, pendente:{campo, min, max, pagina, inicio}, erro}.
// Busca sempre em janelas de até 7 dias: um updated_at_min desde jan/2026 faz a API cair em 500 depois de
// ~10 s (diagnosticoApiNuvemshop, 20/10/2026), enquanto 7 dias respondem em 2–4 s. A primeira carga vai por
// created_at; depois, por updated_at desde a última busca (traz cancelamentos e estornos de pedidos antigos).
var VENDAS_API_URL = 'https://api.nuvemshop.com.br/v1/';
var VENDAS_API_INICIO = '2026-01-01T00:00:00-03:00'; // antes do primeiro pedido (a loja abriu em fev/2026)
var VENDAS_API_INTERVALO = 5 * 60 * 1000;
var VENDAS_API_MARGEM = 10 * 60 * 1000;               // volta 10 min no updated_at, para não perder nada na virada
var VENDAS_API_POR_PAGINA = 200;
var VENDAS_API_JANELA = 7 * 24 * 3600 * 1000;
var VENDAS_API_CAMPOS = ['id','number','created_at','updated_at','status','payment_status','shipping_status',
  'subtotal','discount','shipping_cost_customer','total','shipping_address','shipping_option','gateway','gateway_name',
  'payment_details','coupon','paid_at','completed_at','storefront','cancelled_at','cancel_reason','products'].join(',');
// tradução para os rótulos do CSV (o front classifica por esses textos: Index.html, vSituacao).
// Confirmado, Vencido, Mobile e Cancelamento automático conferidos na api-amostra de 20/10/2026.
var VENDAS_API_STATUS = {open:'Aberto', closed:'Fechado', cancelled:'Cancelado'};
var VENDAS_API_PAGTO = {pending:'Pendente', authorized:'Autorizado', paid:'Confirmado', abandoned:'Abandonado',
  refunded:'Estornado', partially_refunded:'Parcialmente estornado', voided:'Vencido'};
// pagamento que chegou a ser confirmado (só esses têm "Data de pagamento" no CSV)
var VENDAS_API_PAGOS = {authorized:1, paid:1, refunded:1, partially_refunded:1};
var VENDAS_API_ENVIO = {unpacked:'Não está embalado', unfulfilled:'Não enviado', unshipped:'Não enviado',
  packed:'Embalado', fulfilled:'Enviado', shipped:'Enviado', delivered:'Entregue'};
var VENDAS_API_MEIO = {pix:'Pix', credit_card:'Cartão de crédito', debit_card:'Cartão de débito', boleto:'Boleto',
  wire_transfer:'Transferência', cash:'Dinheiro', other:'Outro'};
var VENDAS_API_CANAL = {store:'Loja virtual', mobile:'Mobile', meli:'Mercado Livre', api:'API', form:'Formulário', pos:'Loja física'};
var VENDAS_API_MOTIVO = {customer:'Cliente', inventory:'Sem estoque', fraud:'Fraude', other:'Outro',
  automatic:'Cancelamento automático', refund:'Reembolso'};

function apiNuvemshopConfigurada(){
  var p = props();
  return !!(p.getProperty('NUVEMSHOP_TOKEN') && p.getProperty('NUVEMSHOP_LOJA'));
}

// a base só passa a receber a API depois de instalarSincronizacaoVendas (antes dá para rodar a amostra
// e comparar com o CSV sem misturar nada)
function apiNuvemshopLigada(){ return apiNuvemshopConfigurada() && props().getProperty('VENDAS_API_LIGADA') === '1'; }

function estadoVendasApi(){
  var e = null;
  try{ e = JSON.parse(props().getProperty('VENDAS_API') || 'null'); } catch(err){}
  return e || {};
}

// GET na API. Devolve {lista} (array do corpo), {fim:true} na página depois da última (a API responde 404)
// ou {erro}. Espera e tenta de novo quando bate no limite de 2 chamadas por segundo (429).
function chamarApiNuvemshop(caminho, params){
  var p = props();
  var q = Object.keys(params || {}).map(function(k){ return k + '=' + encodeURIComponent(params[k]); }).join('&');
  var url = VENDAS_API_URL + p.getProperty('NUVEMSHOP_LOJA') + '/' + caminho + (q ? '?' + q : '');
  for(var tentativa = 0; tentativa < 5; tentativa++){
    var r = UrlFetchApp.fetch(url, {
      headers:{'Authentication':'bearer ' + p.getProperty('NUVEMSHOP_TOKEN'),
        'User-Agent':'Talgui Painel (' + RELATORIO_PARA + ')'},
      muteHttpExceptions:true
    });
    var cod = r.getResponseCode();
    if(cod === 429){ Utilities.sleep(1000 * (tentativa + 1)); continue; }
    if(cod === 404) return {fim:true};
    if(cod === 401 || cod === 403) return {erro:'api_acesso'};
    if(cod >= 300) return {erro:'api_' + cod};
    try{ var corpo = JSON.parse(r.getContentText('UTF-8')); }
    catch(e){ return {erro:'api_resposta'}; }
    return Array.isArray(corpo) ? {lista:corpo} : {erro:'api_resposta'};
  }
  return {erro:'api_429'};
}

// "2026-08-16T22:09:01+0000" → "16/08/2026 19:09:01" (horário de Brasília, como no CSV); '' se vazio
function dataApiParaCsv(s){
  if(!s) return '';
  var d = new Date(String(s).replace(/([+-]\d{2})(\d{2})$/, '$1:$2'));
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
}

function tsApi(s){
  var d = new Date(String(s || '').replace(/([+-]\d{2})(\d{2})$/, '$1:$2'));
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function textoApi(v){ return v == null ? '' : String(v).trim(); }

// "Data de pagamento" do CSV (só o dia). paid_at vem sempre null do Nuvem Pago; completed_at chega como
// {date:"2026-10-01 02:57:27.000000", timezone:"UTC"} e o dia dele bate com o CSV
function dataPagamentoApi(o){
  if(!VENDAS_API_PAGOS[o.payment_status]) return '';
  var s = o.paid_at;
  if(!s && o.completed_at) s = typeof o.completed_at === 'object' ? String(o.completed_at.date || '').replace(' ', 'T').replace(/\.\d+$/, '') + '+0000' : o.completed_at;
  return dataApiParaCsv(s).slice(0, 10);
}

// pedido da API → linha da base (mesmo layout de lerCsvVendas: VENDAS_COLUNAS + itens + SKUs).
// Nada de nome, e-mail, CPF, telefone ou endereço: do endereço só sai a UF.
function pedidoApiParaLinha(o, nomes, fotos){
  var pd = o.payment_details || {};
  var cupom = Array.isArray(o.coupon) ? o.coupon.map(function(c){ return c && c.code; }).filter(String).join(', ') : textoApi(o.coupon);
  var envio = o.shipping_option;
  if(envio && typeof envio === 'object') envio = envio.name || envio.code || '';
  var col = {
    'Número do Pedido':textoApi(o.number),
    'Data':dataApiParaCsv(o.created_at),
    'Status do Pedido':VENDAS_API_STATUS[o.status] || textoApi(o.status),
    'Status do Pagamento':VENDAS_API_PAGTO[o.payment_status] || textoApi(o.payment_status),
    'Status do Envio':VENDAS_API_ENVIO[o.shipping_status] || textoApi(o.shipping_status),
    'Subtotal':textoApi(o.subtotal),
    'Desconto':textoApi(o.discount),
    'Valor do Frete':textoApi(o.shipping_cost_customer),
    'Total':textoApi(o.total),
    'Estado':textoApi((o.shipping_address || {}).province),
    'Forma de Entrega':textoApi(envio),
    'Forma de Pagamento':textoApi(o.gateway_name || o.gateway),
    'Cupom de Desconto':cupom,
    'Data de pagamento':dataPagamentoApi(o),
    'Canal':VENDAS_API_CANAL[o.storefront] || textoApi(o.storefront),
    'Data e hora do cancelamento':dataApiParaCsv(o.cancelled_at),
    'Motivo do cancelamento':VENDAS_API_MOTIVO[o.cancel_reason] || textoApi(o.cancel_reason),
    'Parcelas':pd.method === 'credit_card' && pd.installments ? String(pd.installments) : '', // no CSV, Pix fica vazio
    'Meio de pagamento':VENDAS_API_MEIO[pd.method] || textoApi(pd.method)
  };
  var r = VENDAS_COLUNAS.map(function(c){ return col[c] || ''; });
  var qtd = 0, itens = [];
  (o.products || []).forEach(function(it){
    var q = Number(it.quantity) || 0, sku = textoApi(it.sku).replace(/[,:]/g, '');
    qtd += q;
    if(sku){
      itens.push(sku + ':' + q); if(it.name) nomes[sku] = textoApi(it.name);
      // foto do item (só o endereço da imagem, nunca dado do cliente); vazio não sobrescreve
      if(fotos){ var u = urlFotoValida(it.image && it.image.src); if(u) fotos[sku] = u; }
    }
  });
  r.push(qtd, itens.join(','));
  return r;
}

// busca os pedidos alterados desde a última vez e junta na base. Devolve o VENDAS_API atualizado;
// {erro:'consolidando'} se outra execução estiver com a base ou se o tempo acabou no meio da carga.
// relogio = início da execução (o limite de 6 min do Apps Script conta desde ali).
function sincronizarVendasApi(relogio){
  if(!apiNuvemshopConfigurada()) return {erro:'api_sem_token'};
  var pasta = pastaVendas();
  if(!pasta) return {erro:'pasta_vendas'};
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return {erro:'consolidando'};
  try{
    var s = estadoVendasApi(), estado = estadoVendas(), pastaBase = pastaBaseVendas(pasta);
    var inicio = Date.now(), relogioIni = relogio || inicio, bases = {}, nomes = {}, fotosNovas = {}, mudou = false, erro = null;
    // janela de até 7 dias a partir de minMs, sem passar do início deste ciclo
    function janela(campo, minMs, ini){
      return {campo:campo, min:new Date(minMs).toISOString(), max:new Date(Math.min(minMs + VENDAS_API_JANELA, ini)).toISOString(),
        pagina:1, inicio:ini};
    }
    var guardada = s.pendente || s.retomar;
    var pend = guardada && guardada.campo ? guardada
      : s.ultimo ? janela('updated_at', s.ultimo - VENDAS_API_MARGEM, inicio)
      : janela('created_at', Date.parse(VENDAS_API_INICIO), inicio);
    var cicloIni = pend.inicio;
    // a API não tem a data de entrega (delivered_at): guarda a primeira vez em que o pedido apareceu "delivered".
    // entregas[num] = [ms em que ainda não estava entregue (0 = desconhecido), ms da primeira vez visto entregue]
    var entregas = null, entregasMudou = false;
    function marcarEntrega(ant, num, ts){
      if(!entregas) entregas = lerJsonBase(pastaBase, ENTREGAS_ARQ) || {};
      if(entregas[num]) return;
      var jaEntregue = ant && /entreg/i.test(ant[4]);
      var antes = ant && !jaEntregue ? Number(ant[VENDAS_TS]) || 0 : 0;
      var ate = jaEntregue && Number(ant[VENDAS_TS]) ? Math.min(Number(ant[VENDAS_TS]), ts) : ts;
      entregas[num] = [antes, ate];
      entregasMudou = true;
    }
    while(true){
      if(Date.now() - relogioIni > VENDAS_TEMPO_MAX) break;
      var q = {status:'any', per_page:VENDAS_API_POR_PAGINA, page:pend.pagina, fields:VENDAS_API_CAMPOS};
      q[pend.campo + '_min'] = pend.min; q[pend.campo + '_max'] = pend.max;
      var r = chamarApiNuvemshop('orders', q);
      if(r.erro){ erro = r.erro; break; }
      var lista = r.fim ? [] : r.lista;
      var porAno = {};
      lista.forEach(function(o){
        var linha = pedidoApiParaLinha(o, nomes, fotosNovas), dia = diaIsoVendas(linha[1]);
        if(!dia || !linha[0]) return;
        (porAno[dia.slice(0, 4)] = porAno[dia.slice(0, 4)] || []).push({linha:linha, ts:tsApi(o.updated_at), entregue:o.shipping_status === 'delivered'});
      });
      Object.keys(porAno).forEach(function(ano){
        if(!bases[ano]) bases[ano] = lerBaseAno(estado, ano);
        porAno[ano].forEach(function(x){
          // antes de juntar: a linha anterior diz se o pedido ainda não estava entregue (limite inferior da entrega)
          if(x.entregue) marcarEntrega(bases[ano][x.linha[0]], x.linha[0], x.ts);
          juntarPedidos(bases[ano], [x.linha], x.ts);
        });
        mudou = true;
      });
      if(lista.length < VENDAS_API_POR_PAGINA){
        // janela terminada: a próxima começa onde esta acabou, até chegar ao início do ciclo
        var fimJanela = Date.parse(pend.max);
        if(fimJanela >= pend.inicio){ pend = null; break; }
        pend = janela(pend.campo, fimJanela, pend.inicio);
      } else pend.pagina++;
    }
    // grava uma vez por execução (cada gravação cria um arquivo novo da base do ano)
    Object.keys(bases).forEach(function(ano){ gravarBaseAno(estado, pastaBase, ano, bases[ano]); });
    if(entregasMudou) gravarJsonBase(pastaBase, ENTREGAS_ARQ, entregas);
    if(mudou){
      estado.apiMudou = Date.now();
      var todos = lerJsonBase(pastaBase, PRODUTOS_ARQ) || {};
      Object.keys(nomes).forEach(function(sku){ todos[sku] = nomes[sku]; });
      gravarJsonBase(pastaBase, PRODUTOS_ARQ, todos);
    }
    // fotos: erro aqui nunca derruba as vendas
    try{ if(Object.keys(fotosNovas).length) juntarFotos(pastaBase, fotosNovas); }
    catch(e){ console.warn('Fotos (captura nos pedidos): ' + (e && e.message || e)); }
    props().setProperty('VENDAS_BASE', JSON.stringify(estado));
    s.verificadoEm = Date.now();
    // com erro, guarda em "retomar" a janela e a página em que parou: a próxima verificação (gatilho ou
    // painel, depois do intervalo) tenta de novo dali, sem o painel insistir a cada abertura
    delete s.pendente; delete s.retomar; delete s.erro;
    if(erro){ s.erro = erro; s.retomar = pend; }
    else if(pend) s.pendente = pend;
    else s.ultimo = cicloIni;
    props().setProperty('VENDAS_API', JSON.stringify(s));
    return s.pendente && !erro ? {erro:'consolidando'} : s;
  } finally { lock.releaseLock(); }
}

// gatilho de 15 min: mantém a base em dia sem ninguém abrir o painel
function sincronizarVendasGatilho(){
  var inicio = Date.now();
  if(apiNuvemshopLigada()){
    var s = sincronizarVendasApi(inicio);
    if(s.erro && s.erro !== 'consolidando') console.warn('Vendas pela API: ' + s.erro);
  }
  // extratos novos do Nuvem Pago entram na base mesmo sem ninguém abrir a página de Recebimentos
  // (só com folga de tempo: a API vem primeiro e o Apps Script para em 6 min)
  if(Date.now() - inicio > 90 * 1000) return;
  try{
    var np = consolidarNuvemPago(inicio);
    if(np.erro && np.erro !== 'consolidando' && np.erro !== 'pasta_nuvempago') console.warn('Nuvem Pago: ' + np.erro);
  } catch(e){ console.warn('Nuvem Pago: ' + (e && e.message || e)); }
  // fotos das peças: só enquanto a carga não terminou (ou na renovação semanal) e se ainda sobrou tempo
  try{
    if(apiNuvemshopLigada() && fotosPrecisaRodar() && Date.now() - inicio < FOTOS_INICIO_MAX){
      var ft = atualizarFotosProdutos(inicio);
      if(ft.erro && ft.erro !== 'consolidando') console.warn('Fotos: ' + ft.erro);
    }
  } catch(e){ console.warn('Fotos: ' + (e && e.message || e)); }
}

// rodar uma vez no editor (conta dona do script), depois de preencher NUVEMSHOP_TOKEN e NUVEMSHOP_LOJA
function instalarSincronizacaoVendas(){
  ScriptApp.getProjectTriggers().forEach(function(g){
    if(g.getHandlerFunction() === 'sincronizarVendasGatilho') ScriptApp.deleteTrigger(g);
  });
  if(!apiNuvemshopConfigurada()) throw new Error('Preencha NUVEMSHOP_TOKEN e NUVEMSHOP_LOJA nas Propriedades do script.');
  props().setProperty('VENDAS_API_LIGADA', '1');
  ScriptApp.newTrigger('sincronizarVendasGatilho').timeBased().everyMinutes(15).create();
  sincronizarVendasGatilho();
}

// rodar uma vez no editor: grava nuvemshop/_base/api-amostra.json com alguns pedidos de setembro/2026 como
// vêm da API (sem dados pessoais), lado a lado com a linha da base (CSV) e a tradução feita aqui — para
// conferir o mapeamento antes de confiar nos números.
function amostraApiNuvemshop(){
  var r = chamarApiNuvemshop('orders', {status:'any', created_at_min:'2026-09-01T00:00:00-03:00',
    created_at_max:'2026-09-30T23:59:59-03:00', per_page:100});
  if(r.erro || r.fim){ console.log('Erro: ' + (r.erro || 'sem pedidos')); return; }
  var vistos = {}, escolhidos = [];
  r.lista.forEach(function(o){
    var chave = [o.status, o.payment_status, o.shipping_status, (o.payment_details || {}).method, o.storefront].join('|');
    if(!vistos[chave] && escolhidos.length < 15){ vistos[chave] = true; escolhidos.push(o); }
  });
  var estado = estadoVendas(), base = lerBaseAno(estado, '2026'), nomes = {};
  var limpos = escolhidos.map(function(o){
    var x = {};
    Object.keys(o).forEach(function(k){
      if(/^(customer|contact_|billing_|client_details|note|owner_note|fulfillments|shipping_address|token|landing_url|checkout)/.test(k)) return;
      x[k] = o[k];
    });
    x.shipping_address = {province:(o.shipping_address || {}).province};
    var linha = pedidoApiParaLinha(o, nomes), traduzido = {}, csv = {};
    VENDAS_COLUNAS.forEach(function(c, i){ traduzido[c] = linha[i]; });
    var b = base[String(o.number)];
    if(b) VENDAS_COLUNAS.forEach(function(c, i){ csv[c] = b[i]; });
    return {api:x, traduzido:traduzido, base_csv:b ? csv : null};
  });
  var pastaBase = pastaBaseVendas(pastaVendas());
  gravarJsonBase(pastaBase, 'api-amostra.json', {geradoEm:new Date().toISOString(), totalNoMes:r.lista.length, pedidos:limpos});
  console.log('Amostra gravada: ' + limpos.length + ' pedidos (de ' + r.lista.length + ' lidos).');
}

// rodar no editor quando a sincronização avisar api_5xx: testa, uma a uma, as diferenças entre a busca da
// amostra (que funciona) e a da sincronização, e mostra no registro qual delas a Nuvemshop recusa
function diagnosticoApiNuvemshop(){
  // 20/10: updated_at_min desde jan/2026 dá 500 em ~10 s (timeout deles); created_at de um mês responde em 4 s.
  // Agora: que janela de updated_at a API aguenta, e se created_at aguenta 200 por página e páginas adiante
  var agora = new Date(), h = 3600e3;
  function iso(ms){ return new Date(ms).toISOString(); }
  var testes = [
    ['1 alterados na última hora', {status:'any', updated_at_min:iso(agora - h), per_page:100, page:1}],
    ['2 alterados no último dia', {status:'any', updated_at_min:iso(agora - 24 * h), per_page:100, page:1}],
    ['3 alterados nos últimos 7 dias', {status:'any', updated_at_min:iso(agora - 7 * 24 * h), per_page:100, page:1}],
    ['4 alterados entre 01 e 08/09', {status:'any', updated_at_min:'2026-09-01T00:00:00-03:00', updated_at_max:'2026-09-08T00:00:00-03:00', per_page:100, page:1}],
    ['5 criados em set, 200/pág, campos', {status:'any', created_at_min:'2026-09-01T00:00:00-03:00', created_at_max:'2026-09-30T23:59:59-03:00', per_page:200, page:1, fields:VENDAS_API_CAMPOS}],
    ['6 criados em set, página 10', {status:'any', created_at_min:'2026-09-01T00:00:00-03:00', created_at_max:'2026-09-30T23:59:59-03:00', per_page:200, page:10, fields:VENDAS_API_CAMPOS}],
    ['7 criados em mar, 200/pág', {status:'any', created_at_min:'2026-03-01T00:00:00-03:00', created_at_max:'2026-03-31T23:59:59-03:00', per_page:200, page:1, fields:VENDAS_API_CAMPOS}]
  ];
  testes.forEach(function(t){
    var ini = Date.now(), r = chamarApiNuvemshop('orders', t[1]);
    console.log(t[0] + ': ' + (r.erro ? 'ERRO ' + r.erro : r.fim ? 'sem pedidos' : 'ok, ' + r.lista.length + ' pedidos') + ' em ' + (Date.now() - ini) + ' ms');
    Utilities.sleep(600);
  });
}

/* ===================== Fotos das peças (fotos.json) ===================== */

// fotos.json (na pasta _base) = {SKU: "https://…imagem"}. Guarda só o ENDEREÇO da imagem hospedada na Nuvemshop;
// nenhum arquivo é copiado e nenhum dado de cliente entra. Duas fontes:
//  1) captura incremental: cada pedido lido pela API (item.image.src), em pedidoApiParaLinha/sincronizarVendasApi;
//  2) carga do catálogo: atualizarFotosProdutos (GET products, paginado), retomável e refeita 1x por semana.
// Propriedades do script: FOTOS_PAGINA (próxima página a ler), FOTOS_COMPLETO (AAAA-MM-DD da última carga completa),
// FOTOS_VERSAO (muda quando o conteúdo de fotos.json muda; entra na versão de obterDevolucoes), FOTOS_ERRO.
var FOTOS_ARQ = 'fotos.json';
var FOTOS_TEMPO_MAX = 3 * 60 * 1000;  // orçamento da carga, contado desde o início da execução
var FOTOS_INICIO_MAX = 120 * 1000;    // no gatilho, só começa se a execução ainda tem menos que isso
var FOTOS_RENOVAR_DIAS = 7;           // refaz a carga completa a cada semana (produtos novos/alterados)

function urlFotoValida(u){
  u = textoApi(u);
  return /^https:\/\/[^\s"'<>]+$/i.test(u) ? u : '';
}

function lerFotos(pastaBase){ return lerJsonBase(pastaBase, FOTOS_ARQ) || {}; }

function gravarFotos(pastaBase, fotos){
  gravarJsonBase(pastaBase, FOTOS_ARQ, fotos);
  props().setProperty('FOTOS_VERSAO', String(Date.now()));
}

// junta as novas no fotos.json; só regrava se algo mudou. Nunca troca uma URL por vazio. Devolve quantas mudaram.
function juntarFotos(pastaBase, novas){
  var todas = lerFotos(pastaBase), mudou = 0;
  Object.keys(novas).forEach(function(sku){
    var u = urlFotoValida(novas[sku]);
    if(u && todas[sku] !== u){ todas[sku] = u; mudou++; }
  });
  if(mudou) gravarFotos(pastaBase, todas);
  return mudou;
}

function fotosPrecisaRodar(){
  var p = props(), ok = p.getProperty('FOTOS_COMPLETO');
  if(p.getProperty('FOTOS_PAGINA') || !ok) return true; // carga em andamento ou nunca feita
  var t = Date.parse(ok + 'T12:00:00Z');
  return isNaN(t) || Date.now() - t > FOTOS_RENOVAR_DIAS * 24 * 3600 * 1000;
}

// produto da API → {sku: url}. Imagem da variante (variant.image_id → images[].id); sem ela, a de menor position.
function fotosDoProduto(prod){
  var imgs = (prod.images || []).filter(function(i){ return i && urlFotoValida(i.src); });
  if(!imgs.length) return {};
  var porId = {};
  imgs.forEach(function(i){ porId[String(i.id)] = i; });
  var primeira = imgs.slice().sort(function(a, b){ return (Number(a.position) || 0) - (Number(b.position) || 0); })[0];
  var out = {};
  (prod.variants || []).forEach(function(v){
    var sku = textoApi(v && v.sku).replace(/[,:]/g, '');
    if(!sku) return;
    var img = (v.image_id != null && porId[String(v.image_id)]) || primeira;
    out[sku] = urlFotoValida(img.src);
  });
  return out;
}

// carga histórica do catálogo. Retoma da página guardada em FOTOS_PAGINA; ao terminar, grava FOTOS_COMPLETO.
// Pode ser rodada à mão no editor (o gatilho de 15 min também a chama enquanto não terminar).
// relogio = início da execução (o limite do Apps Script conta desde ali).
function atualizarFotosProdutos(relogio){
  if(!apiNuvemshopConfigurada()) return {erro:'api_sem_token'};
  var pasta = pastaVendas();
  if(!pasta) return {erro:'pasta_vendas'};
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return {erro:'consolidando'};
  try{
    var p = props(), pastaBase = pastaBaseVendas(pasta), ini = typeof relogio === 'number' ? relogio : Date.now();
    var pagina = Number(p.getProperty('FOTOS_PAGINA')) || 1;
    var acumuladas = {}, erro = null, terminou = false, lidos = 0;
    while(true){
      if(Date.now() - ini > FOTOS_TEMPO_MAX) break;
      var r = chamarApiNuvemshop('products', {per_page:200, page:pagina, fields:'id,images,variants'});
      if(r.erro){ erro = r.erro; break; }
      var lista = r.fim ? [] : r.lista;
      lista.forEach(function(prod){
        var f = fotosDoProduto(prod);
        Object.keys(f).forEach(function(sku){ acumuladas[sku] = f[sku]; });
      });
      lidos += lista.length;
      if(lista.length < 200){ terminou = true; break; }
      pagina++;
      Utilities.sleep(500); // a API aceita ~2 chamadas por segundo
    }
    // grava o que veio, mesmo se parou no meio; a página só avança depois de gravar
    if(Object.keys(acumuladas).length) juntarFotos(pastaBase, acumuladas);
    if(erro){ p.setProperty('FOTOS_ERRO', erro); p.setProperty('FOTOS_PAGINA', String(pagina)); }
    else {
      p.deleteProperty('FOTOS_ERRO');
      if(terminou){
        p.deleteProperty('FOTOS_PAGINA');
        p.setProperty('FOTOS_COMPLETO', Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd'));
      } else p.setProperty('FOTOS_PAGINA', String(pagina));
    }
    console.log('Fotos: ' + lidos + ' produtos lidos' + (terminou ? ', carga completa.' : erro ? ', parou com erro ' + erro + '.' : ', continua na página ' + pagina + ' (rode de novo).'));
    return erro ? {erro:erro} : {terminou:terminou, pagina:pagina, produtos:lidos};
  } finally { lock.releaseLock(); }
}

/* ===================== Arquivos JSON da pasta _base ===================== */

var PRODUTOS_ARQ = 'produtos.json'; // SKU → "Nome do Produto" da Nuvemshop ("Blusa Monica Marrom (Marrom, 36)")
var ENTREGAS_ARQ = 'entregas.json'; // nº do pedido → [ms antes da entrega (0 = desconhecido), ms da 1ª vez visto "delivered"] (sincronizarVendasApi)

// o mais novo com esse nome (gravar cria um arquivo novo e manda o anterior para a lixeira)
function arquivoBase(pastaBase, nome){
  var it = pastaBase.getFilesByName(nome), achado = null;
  while(it.hasNext()){ var f = it.next(); if(!achado || f.getDateCreated() > achado.getDateCreated()) achado = f; }
  return achado;
}

function lerJsonBase(pastaBase, nome){
  var f = arquivoBase(pastaBase, nome);
  if(!f) return null;
  try{ return JSON.parse(f.getBlob().getDataAsString('UTF-8')); } catch(e){ return null; }
}

function gravarJsonBase(pastaBase, nome, obj){
  var antigo = arquivoBase(pastaBase, nome);
  pastaBase.createFile(Utilities.newBlob(JSON.stringify(obj), 'application/json', nome));
  if(antigo){ try{ antigo.setTrashed(true); } catch(e){} }
}

/* ===================== Devoluções (exportações da AfterClick) ===================== */

// A subpasta "AJD" de "nuvemshop" guarda os exports da AfterClick (Trocas e devoluções → exportar), em .xlsx ou
// .csv, uma linha por peça. Cada arquivo é lido UMA vez e entra na base nuvemshop/_base/devolucoes.json, por
// solicitação (vale o arquivo modificado por último). Como o export não traz a data de conclusão, a base guarda
// em "hist" quando cada solicitação foi vista pela primeira vez em cada status (data de modificação do arquivo).
// Só as colunas abaixo entram: nada de cliente, documento, e-mail ou nota fiscal.
var DEVOL_PASTA_NOME = 'AJD';
var DEVOL_BASE_ARQ = 'devolucoes.json';
var DEVOL_ESTADO_ARQ = 'devolucoes-estado.json'; // num arquivo, não numa propriedade: um export por dia estoura os 9 KB
var DEVOL_COLUNAS = ['Nº do Pedido','Nº Pedido de Origem','Processo','Data de criação','Valor Total','Desconto','Frete',
  'Status','Motivo','Solução Desejada','Tipo de Solicitação','Status da Devolução','Data da Devolução'];
var DEVOL_ITEM = ['SKU','Preço Por','Quantidade'];
var DEVOL_ITENS = DEVOL_COLUNAS.length; // [[sku, preço, qtd], ...] da solicitação
var DEVOL_TS = DEVOL_ITENS + 1;         // data de modificação do arquivo de origem (só na base)
var DEVOL_CODIFICADAS = [2, 7, 8, 9, 10, 11];
var DEVOL_TMP = '_tmp_devolucoes';

function pastaDevolucoes(pastaV){
  var it = pastaV.getFoldersByName(DEVOL_PASTA_NOME);
  return it.hasNext() ? it.next() : null;
}

function arquivosDevolucoes(pasta){
  var lista = [], it = pasta.getFiles();
  while(it.hasNext()){
    var f = it.next(), nome = f.getName(), tipo = f.getMimeType();
    if(nome.indexOf(DEVOL_TMP) === 0) continue;
    if(/\.(csv|xlsx)$/i.test(nome) || tipo === MIME_XLSX || tipo === MIME_SHEETS || /csv/i.test(tipo)) lista.push(f);
  }
  lista.sort(function(a, b){ return a.getLastUpdated() - b.getLastUpdated(); });
  return lista;
}

// planilha Google (ou .xlsx convertido numa cópia temporária pela API do Drive) → linhas do CSV exportado
function exportarCsvPlanilha(id){
  var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + id + '/export?format=csv', {
    headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions:true
  });
  if(resp.getResponseCode() !== 200) return {erro:'exportacao_falhou', detalhe:'HTTP ' + resp.getResponseCode()};
  return Utilities.parseCsv(resp.getContentText('UTF-8'), ',');
}

function linhasDevolucoes(f){
  var tipo = f.getMimeType();
  if(tipo === MIME_SHEETS) return exportarCsvPlanilha(f.getId());
  if(tipo === MIME_XLSX || /\.xlsx$/i.test(f.getName())){
    var resp = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + f.getId() + '/copy?supportsAllDrives=true&fields=id', {
      method:'post', contentType:'application/json', muteHttpExceptions:true,
      payload:JSON.stringify({name:DEVOL_TMP, mimeType:MIME_SHEETS}),
      headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()}
    });
    if(resp.getResponseCode() !== 200) return {erro:'conversao_falhou', detalhe:'HTTP ' + resp.getResponseCode()};
    var tmp = JSON.parse(resp.getContentText()).id;
    try{ return exportarCsvPlanilha(tmp); }
    finally { try{ DriveApp.getFileById(tmp).setTrashed(true); } catch(e){} }
  }
  var blob = f.getBlob(), texto = blob.getDataAsString('UTF-8').replace(/^﻿/, '');
  // CSV salvo em ISO-8859-1 (Excel no Windows): em UTF-8 os acentos viram o caractere inválido
  if(texto.indexOf('Pedido de Origem') < 0 || texto.indexOf('\uFFFD') >= 0) texto = blob.getDataAsString('ISO-8859-1');
  var primeira = texto.slice(0, texto.indexOf('\n'));
  return Utilities.parseCsv(texto, primeira.split(';').length > primeira.split(',').length ? ';' : ',');
}

function textoDevol(v){ v = String(v == null ? '' : v).trim(); return v === 'undefined' || v === 'null' ? '' : v; }

// "R$ 1.234,56" → 1234.56
function numeroBR(v){
  var s = textoDevol(v).replace(/[^\d,.\-]/g, '');
  if(s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  var n = Number(s);
  return isFinite(n) ? n : 0;
}

// "01/10/2026 - 11:40" ou "2026-10-05T11:12:08" → "2026-10-01 11:40" ('' se não for data)
function dataHoraDevol(v){
  var s = textoDevol(v), m = /^(\d{2})\/(\d{2})\/(\d{4})(?:\D+(\d{2}):(\d{2}))?/.exec(s);
  if(m) return m[3] + '-' + m[2] + '-' + m[1] + (m[4] ? ' ' + m[4] + ':' + m[5] : '');
  m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?/.exec(s);
  return m ? m[1] + '-' + m[2] + '-' + m[3] + (m[4] ? ' ' + m[4] + ':' + m[5] : '') : '';
}

// arquivo → uma linha por solicitação (DEVOL_COLUNAS + itens)
function lerDevolucoes(f){
  var linhas = linhasDevolucoes(f);
  if(linhas.erro) return linhas;
  if(!linhas.length) return {erro:'csv_vazio'};
  var cab = linhas[0].map(function(c){ return String(c).replace(/^﻿/, '').trim(); });
  var ix = DEVOL_COLUNAS.map(function(c){ return cab.indexOf(c); });
  var it = DEVOL_ITEM.map(function(c){ return cab.indexOf(c); });
  if(ix[0] < 0 || ix[1] < 0 || ix[7] < 0) return {erro:'formato_devolucoes'};
  var sols = {}, ordem = [];
  for(var i = 1; i < linhas.length; i++){
    var l = linhas[i], id = textoDevol(l[ix[0]]);
    if(!id) continue;
    var s = sols[id];
    if(!s){
      s = ix.map(function(j){ return j >= 0 ? textoDevol(l[j]) : ''; });
      s[3] = dataHoraDevol(s[3]); s[12] = dataHoraDevol(s[12]);
      s[4] = numeroBR(s[4]); s[5] = numeroBR(s[5]); s[6] = numeroBR(s[6]);
      s.push([]);
      sols[id] = s; ordem.push(id);
    }
    var sku = it[0] >= 0 ? textoDevol(l[it[0]]) : '';
    if(sku) s[DEVOL_ITENS].push([sku, it[1] >= 0 ? numeroBR(l[it[1]]) : 0, it[2] >= 0 ? Number(textoDevol(l[it[2]])) || 1 : 1]);
  }
  return {sols:ordem.map(function(id){ return sols[id]; })};
}

// lê os arquivos ainda não consolidados; mesma lógica de tempo e de lock das vendas
function consolidarDevolucoes(pastaV, pasta){
  var arquivos = arquivosDevolucoes(pasta);
  var pastaBase = pastaBaseVendas(pastaV);
  var estado = lerJsonBase(pastaBase, DEVOL_ESTADO_ARQ) || {};
  estado.arqs = estado.arqs || {}; estado.info = estado.info || {}; estado.avisos = estado.avisos || {};
  var atuais = {};
  arquivos.forEach(function(f){ atuais[f.getId()] = f.getLastUpdated().getTime(); });
  var removido = Object.keys(estado.arqs).some(function(id){ return !(id in atuais); });
  var falta = arquivos.filter(function(f){ return estado.arqs[f.getId()] !== atuais[f.getId()]; });
  if(!removido && !falta.length) return {estado:estado, pastaBase:pastaBase, arquivos:arquivos.length};

  var lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return {erro:'consolidando', feitos:null, total:arquivos.length};
  try{
    estado = lerJsonBase(pastaBase, DEVOL_ESTADO_ARQ) || estado;
    estado.arqs = estado.arqs || {}; estado.info = estado.info || {}; estado.avisos = estado.avisos || {};
    var base = lerJsonBase(pastaBase, DEVOL_BASE_ARQ) || {};
    base.sols = base.sols || {}; base.hist = base.hist || {};
    var inicio = Date.now();
    removido = Object.keys(estado.arqs).some(function(id){ return !(id in atuais); });
    if(removido){
      // um arquivo saiu da pasta: refaz as solicitações do zero (o histórico de status fica: é o que já foi visto)
      base.sols = {}; estado.arqs = {}; estado.info = {}; estado.avisos = {};
    }
    falta = arquivos.filter(function(f){ return estado.arqs[f.getId()] !== atuais[f.getId()]; });
    var feitos = 0;
    for(var i = 0; i < falta.length; i++){
      if(feitos && Date.now() - inicio > VENDAS_TEMPO_MAX) break;
      var f = falta[i], id = f.getId(), ts = atuais[id];
      var lido = lerDevolucoes(f);
      delete estado.avisos[id];
      if(lido.erro){
        estado.avisos[id] = {nome:f.getName(), erro:lido.erro, detalhe:lido.detalhe || ''};
        delete estado.info[id];
      } else {
        var de = '', ate = '';
        lido.sols.forEach(function(s){
          var atual = base.sols[s[0]];
          if(!atual || atual[DEVOL_TS] <= ts){ s[DEVOL_TS] = ts; base.sols[s[0]] = s; }
          var h = base.hist[s[0]] = base.hist[s[0]] || {};
          if(s[7] && (!(s[7] in h) || h[s[7]] > ts)) h[s[7]] = ts;
          var dia = s[3].slice(0, 10);
          if(dia && (!de || dia < de)) de = dia;
          if(dia && (!ate || dia > ate)) ate = dia;
        });
        if(!base.inicio || ts < base.inicio) base.inicio = ts;
        estado.info[id] = {nome:f.getName(), modificadoEm:new Date(ts).toISOString(), solicitacoes:lido.sols.length, de:de, ate:ate};
      }
      estado.arqs[id] = ts;
      // base antes do estado: se cair entre os dois, o arquivo é lido de novo (e a junção é idempotente)
      gravarJsonBase(pastaBase, DEVOL_BASE_ARQ, base);
      gravarJsonBase(pastaBase, DEVOL_ESTADO_ARQ, estado);
      feitos++;
    }
    if(removido && !feitos){
      gravarJsonBase(pastaBase, DEVOL_BASE_ARQ, base);
      gravarJsonBase(pastaBase, DEVOL_ESTADO_ARQ, estado);
    }
    var restam = falta.length - feitos;
    if(restam > 0) return {erro:'consolidando', feitos:arquivos.length - restam, total:arquivos.length};
    return {estado:estado, pastaBase:pastaBase, arquivos:arquivos.length, leu:feitos > 0};
  } finally { lock.releaseLock(); }
}

/* ===================== NF de devolução (exportação do quadro Financeiro) ===================== */

// A exportação diária dos cards FINA ("AAAA-MM-DD - Exportação do Painel - Financeiro", planilha Google) traz na
// Descrição ou nas Observações "NF emitida: Nº 13933 - Emitida em 12/10/2026" (ou só "NF emitida em 15/09").
// Vale a planilha mais nova da pasta. Só as colunas abaixo entram: nada de nome, telefone, PIX ou e-mail.
var NF_PASTA_ID = '1Cy5a2oVeecFEo6_Z8na99gZUv5rP6v2w'; // a propriedade NF_PASTA troca a pasta
var NF_NOME = 'Exportação do Painel - Financeiro';
var NF_ARQ = 'nf-financeiro.json';
var NF_COLUNAS = ['Código','Data criação','Fase','Descrição','Pedido Nuvem','Tratativa Financeira - Ação',
  'Tratativa Financeira - Observações','Data de processamento/Financeiro','Data NF emitida'];

function arquivoNfMaisRecente(){
  var pasta;
  try{ pasta = DriveApp.getFolderById(props().getProperty('NF_PASTA') || NF_PASTA_ID); } catch(e){ return null; }
  var it = pasta.getFilesByType(MIME_SHEETS), achado = null;
  while(it.hasNext()){
    var f = it.next();
    if(f.getName().indexOf(NF_NOME) >= 0 && (!achado || f.getLastUpdated() > achado.getLastUpdated())) achado = f;
  }
  return achado;
}

// "12/10/2026" ou "15/09" (sem ano: o do card; o anterior se cair mais de 6 meses depois da criação) → "2026-10-12"
function diaNf(d, m, a, criado){
  if(!a){
    a = criado ? Number(criado.slice(0, 4)) : new Date().getFullYear();
    if(criado && Number(m) - Number(criado.slice(5, 7)) > 6) a--;
  } else if(a.length === 2) a = '20' + a;
  var iso = a + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
  return /^\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/.test(iso) ? iso : '';
}

// texto do card → {num, data} ou null
function nfDoTexto(t, criado){
  var i = t.search(/NF\s+emitida/i);
  if(i < 0) return null;
  var s = t.slice(i, i + 160);
  var n = /N\s*[º°o.]\s*(\d+)/i.exec(s), d = /(\d{1,2})\/(\d{1,2})(?:\/(\d{4}|\d{2}))?/.exec(s);
  var data = d ? diaNf(('0' + d[1]).slice(-2), ('0' + d[2]).slice(-2), d[3], criado) : '';
  return n || data ? {num:n ? n[1] : '', data:data} : null;
}

// planilha → cards [codigo, fase, pedido, acao, nfNum, nfData, criado, processado]
function lerNotasFinanceiro(f){
  var linhas = exportarCsvPlanilha(f.getId());
  if(linhas.erro) return linhas;
  if(!linhas.length) return {erro:'csv_vazio'};
  var cab = linhas[0].map(function(c){ return String(c).replace(/^﻿/, '').trim(); });
  var ix = NF_COLUNAS.map(function(c){ return cab.indexOf(c); });
  if(ix[0] < 0 || ix[4] < 0) return {erro:'formato_nf'};
  var cards = [];
  for(var i = 1; i < linhas.length; i++){
    var l = linhas[i], v = ix.map(function(j){ return j >= 0 ? textoDevol(l[j]) : ''; });
    var pedido = v[4].replace(/\s+/g, '').toUpperCase();
    if(!v[0] || !pedido) continue;
    var criado = dataHoraDevol(v[1]);
    var nf = nfDoTexto(v[3] + '\n' + v[6], criado) || {num:'', data:''};
    if(!nf.data) nf.data = dataHoraDevol(v[8]).slice(0, 10);
    cards.push([v[0], v[2], pedido, v[5], nf.num, nf.data, criado, dataHoraDevol(v[7]).slice(0, 10)]);
  }
  return {cards:cards};
}

// lê a planilha mais nova só quando ela muda; o resultado fica em _base/nf-financeiro.json
function notasFinanceiro(pastaBase){
  var f = arquivoNfMaisRecente();
  if(!f) return {erro:'sem_nf'};
  var ts = f.getLastUpdated().getTime(), salvo = lerJsonBase(pastaBase, NF_ARQ);
  if(salvo && salvo.arquivo === f.getId() && salvo.modificado === ts) return salvo;
  var lido = lerNotasFinanceiro(f);
  if(lido.erro) return {erro:lido.erro, nome:f.getName()};
  var novo = {arquivo:f.getId(), nome:f.getName(), modificado:ts, cards:lido.cards};
  gravarJsonBase(pastaBase, NF_ARQ, novo);
  return novo;
}

// mesma regra de "venda" da página (vSituacao): pagamento confirmado, sem cancelamento, total > 0
function ehVendaBase(r){
  if(/cancel/i.test(r[2]) || r[15] || /cancel/i.test(r[3])) return false;
  if(/recus/i.test(r[3]) || /100%/.test(r[11]) || !(Number(r[8]) > 0)) return false;
  return /confirm|estorn|pago|aprovad|autoriz/i.test(r[3]);
}

// "Blusa Monica Marrom (Marrom, 36)" → {modelo:'Blusa Monica Marrom', cor:'Marrom', tam:'36'}
function partesProduto(nome, sku){
  var m = /^(.*?)\s*\(([^()]*)\)\s*$/.exec(nome || '');
  if(!m) return {modelo: nome || 'SKU ' + sku, cor:'', tam:''};
  var p = m[2].split(',').map(function(x){ return x.trim(); });
  return {modelo:m[1], cor:p.length > 1 ? p.slice(0, -1).join(', ') : '', tam:p[p.length - 1]};
}

// pedido de origem de uma solicitação → [data do pedido, status do pagamento, meio de pagamento, parcelas, status do envio,
// estorno, data do estorno, total, venda (1/0), entrega depois de (dia ou ''), entrega até (dia ou ''), data de pagamento].
// Textos e valores vão como estão na base (o front lê com vData e vNum). "Entrega até" vem de entregas.json ou, sem ele,
// do momento em que a linha "Entregue" foi gravada (VENDAS_TS): a entrega aconteceu antes disso.
function resumoPedidoDevol(r, venda, ent, dia){
  var ate = ent ? ent[1] : (/entreg/i.test(r[4]) && Number(r[VENDAS_TS]) ? Number(r[VENDAS_TS]) : 0);
  return [r[1] || '', r[3] || '', r[24] || r[11] || '', r[20] || '', r[4] || '', r[17] || '', r[18] || '', r[8] || '',
    venda ? 1 : 0, ent && ent[0] ? dia(ent[0]) : '', ate ? dia(ate) : '', r[13] || ''];
}

// versao: a que a página já tem; se nada mudou, responde {mudou:false}. A taxa de devolução por peça usa as
// peças vendidas (SKUs da base de vendas), agrupadas por modelo e por modelo|tamanho, por mês do pedido.
function obterDevolucoes(token, versao){
  if(!contaComToken(token)) return {erro:'senha_necessaria'};
  var pastaV = pastaVendas();
  if(!pastaV) return {erro:'pasta_vendas'};
  var pasta = pastaDevolucoes(pastaV);
  if(!pasta) return {erro:'pasta_devolucoes'};
  var c = consolidarDevolucoes(pastaV, pasta);
  if(c.erro) return c;
  if(!c.arquivos) return {erro:'sem_arquivo'};
  // leu arquivo novo nesta chamada: as vendas ficam para a próxima (as duas juntas podem passar dos 6 min)
  if(c.leu) return {erro:'consolidando', feitos:c.arquivos, total:c.arquivos};
  var cv = consolidarVendas(pastaV);
  if(cv.erro === 'consolidando') return cv;
  var ev = cv.estado || {arqs:{}, anos:{}};
  var nf;
  try{ nf = notasFinanceiro(c.pastaBase); } catch(e){ nf = {erro:'nf_falhou', detalhe:String(e && e.message || e)}; }
  // apiMudou: pedidos atualizados pela API (status de pagamento, envio e entregas dos pedidos de origem)
  var assinatura = [c.estado.arqs, ev.arqs].map(function(a){
    return Object.keys(a).sort().map(function(id){ return id + '@' + a[id]; }).join('|');
  }).join('#') + '#nf:' + (nf.erro || nf.arquivo + '@' + nf.modificado) + '#api:' + (ev.apiMudou || '') + '#fotos:' + (props().getProperty('FOTOS_VERSAO') || '');
  var novaVersao = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, assinatura));
  if(versao && versao === novaVersao) return {mudou:false, versao:novaVersao};

  var base = lerJsonBase(c.pastaBase, DEVOL_BASE_ARQ) || {sols:{}, hist:{}};
  var nomes = lerJsonBase(c.pastaBase, PRODUTOS_ARQ) || {};
  var dic = {}, idxDic = {};
  DEVOL_CODIFICADAS.forEach(function(k){ dic[k] = []; idxDic[k] = Object.create(null); });
  function cod(k, v){ var i = idxDic[k][v]; if(i === undefined){ i = idxDic[k][v] = dic[k].length; dic[k].push(v); } return i; }
  var skus = [], idxSku = Object.create(null), modelos = [], idxModelo = Object.create(null), partes = Object.create(null);
  function parte(sku){ return partes[sku] || (partes[sku] = partesProduto(nomes[sku], sku)); }
  function modeloDe(nome){ var i = idxModelo[nome]; if(i === undefined){ i = idxModelo[nome] = modelos.length; modelos.push(nome); } return i; }
  function skuIdx(sku){
    var i = idxSku[sku];
    if(i === undefined){ var p = parte(sku); i = idxSku[sku] = skus.length; skus.push([sku, modeloDe(p.modelo), p.cor, p.tam]); }
    return i;
  }
  var dia = function(ts){ return Utilities.formatDate(new Date(ts), 'America/Sao_Paulo', 'yyyy-MM-dd'); };
  var origens = Object.create(null);
  // "Visto em": dia do arquivo mais novo que trouxe a solicitação (DEVOL_TS). Se ficar para trás das outras,
  // a solicitação saiu da janela de 60 dias do export e o status dela pode estar desatualizado
  var linhas = Object.keys(base.sols).map(function(id){
    var s = base.sols[id], x = s.slice(0, DEVOL_ITENS);
    x[1] = (String(x[1]).match(/\d+/) || [''])[0];
    if(x[1]) origens[x[1]] = true;
    DEVOL_CODIFICADAS.forEach(function(k){ x[k] = cod(k, x[k] == null ? '' : String(x[k])); });
    x.push((s[DEVOL_ITENS] || []).map(function(it){ return [skuIdx(it[0]), it[1], it[2]]; }));
    x.push(s[DEVOL_TS] ? dia(s[DEVOL_TS]) : '');
    return x;
  });
  var hist = {};
  Object.keys(base.hist).forEach(function(id){
    if(!base.sols[id]) return;
    var h = base.hist[id];
    hist[id] = Object.keys(h).map(function(st){ return [cod(7, st), dia(h[st])]; });
  });
  // peças vendidas dos modelos devolvidos (por modelo e por modelo|tamanho), por mês do pedido;
  // mesesVenda (todos os anos da base) = denominador da taxa de devolução: [pedidos de venda, itens, entregues, faturamento];
  // pedidos = resumo dos pedidos de origem das solicitações (resumoPedidoDevol), sem dados de cliente
  var vendidas = {}, mesesVenda = {}, pedidosOrigem = {};
  var entregas = lerJsonBase(c.pastaBase, ENTREGAS_ARQ) || {};
  Object.keys(ev.anos || {}).forEach(function(ano){
    var pedidos = lerBaseAno(ev, ano);
    Object.keys(pedidos).forEach(function(n){
      var r = pedidos[n], venda = ehVendaBase(r);
      var mes = diaIsoVendas(r[1]).slice(0, 7);
      if(venda && mes){
        var mv = mesesVenda[mes] = mesesVenda[mes] || [0, 0, 0, 0];
        mv[0]++; mv[1] += Number(r[VENDAS_ITENS]) || 0; if(/entreg/i.test(r[4])) mv[2]++; mv[3] += Number(r[8]) || 0;
      }
      if(origens[n]) pedidosOrigem[n] = resumoPedidoDevol(r, venda, entregas[n], dia);
      if(!r[VENDAS_SKUS] || !venda) return;
      String(r[VENDAS_SKUS]).split(',').forEach(function(item){
        var k = item.lastIndexOf(':'), sku = item.slice(0, k), q = Number(item.slice(k + 1)) || 0;
        var p = parte(sku), m = idxModelo[p.modelo];
        if(m === undefined) return;
        [String(m), m + '|' + p.tam].forEach(function(chave){
          var v = vendidas[chave] = vendidas[chave] || {};
          v[mes] = (v[mes] || 0) + q;
        });
      });
    });
  });
  var arquivos = Object.keys(c.estado.info).map(function(id){ return c.estado.info[id]; })
    .sort(function(a, b){ return a.modificadoEm < b.modificadoEm ? -1 : 1; });
  var avisos = Object.keys(c.estado.avisos).map(function(id){ return c.estado.avisos[id]; });
  // fotos só dos SKUs destas devoluções (payload pequeno); falha de fotos nunca derruba as devoluções
  var fotos = {};
  try{
    var todasFotos = lerFotos(c.pastaBase);
    skus.forEach(function(x){ if(todasFotos[x[0]]) fotos[x[0]] = todasFotos[x[0]]; });
  } catch(e){ fotos = {}; }
  return {mudou:true, versao:novaVersao, colunas:DEVOL_COLUNAS.concat(['Itens', 'Visto em']), dic:dic, linhas:linhas,
    skus:skus, fotos:fotos, modelos:modelos, vendidas:vendidas, hist:hist, inicio:base.inicio ? dia(base.inicio) : '',
    pedidos:pedidosOrigem, mesesVenda:mesesVenda,
    arquivos:arquivos, avisos:avisos, auto:estadoDevolExt(),
    nf:nf.erro ? {erro:nf.erro, nome:nf.nome || ''} : {cards:nf.cards, nome:nf.nome, modificadoEm:new Date(nf.modificado).toISOString()}};
}

/* ===================== Devoluções enviadas pela extensão do Chrome ===================== */

// A extensão (pasta Relatorios/extensao-afterclick) lê as solicitações na API interna da AfterClick com a sessão
// aberta no Chrome e manda, a cada hora, um CSV no mesmo layout do export manual (DEVOL_COLUNAS + DEVOL_ITEM).
// Aqui ele vira UM arquivo por dia em nuvemshop/AJD (afterclick-auto-AAAA-MM-DD.csv, atualizado no lugar: apagar
// arquivo faz consolidarDevolucoes refazer tudo) e é consolidado na hora. A chave fica em DEVOL_EXT_CHAVE
// (gerarChaveExtensao) e o último envio em DEVOL_EXT: {ultimo, solicitacoes, arquivo, erro}.
var DEVOL_EXT_PREFIXO = 'afterclick-auto-';
var DEVOL_EXT_MAX = 20 * 1024 * 1024;

function estadoDevolExt(){
  try{ return JSON.parse(props().getProperty('DEVOL_EXT') || 'null'); } catch(e){ return null; }
}

function respostaJson(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  var corpo = null;
  try{ corpo = JSON.parse(e.postData.contents); } catch(err){ return respostaJson({erro:'pedido_invalido'}); }
  if(!corpo || corpo.acao !== 'devolucoes') return respostaJson({erro:'pedido_invalido'});
  var chave = props().getProperty('DEVOL_EXT_CHAVE');
  if(!chave || corpo.chave !== chave) return respostaJson({erro:'chave_invalida'});
  var r = receberDevolucoesExtensao(String(corpo.conteudo || ''), Number(corpo.solicitacoes) || 0);
  var s = estadoDevolExt() || {};
  if(r.erro && r.erro !== 'consolidando') s.erro = r.erro;
  else { s.ultimo = Date.now(); s.solicitacoes = r.solicitacoes; s.arquivo = r.arquivo; delete s.erro; }
  props().setProperty('DEVOL_EXT', JSON.stringify(s));
  return respostaJson(r);
}

function receberDevolucoesExtensao(conteudo, solicitacoes){
  if(!conteudo || conteudo.length > DEVOL_EXT_MAX) return {erro:'tamanho'};
  if(conteudo.indexOf('Pedido de Origem') < 0) return {erro:'formato_devolucoes'};
  var pastaV = pastaVendas();
  if(!pastaV) return {erro:'pasta_vendas'};
  var pasta = pastaDevolucoes(pastaV);
  if(!pasta) return {erro:'pasta_devolucoes'};
  var nome = DEVOL_EXT_PREFIXO + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd') + '.csv';
  var it = pasta.getFilesByName(nome), f = it.hasNext() ? it.next() : null;
  if(f){
    // conteúdo novo no mesmo arquivo (a data de modificação muda e consolidarDevolucoes relê)
    var resp = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files/' + f.getId() + '?uploadType=media&supportsAllDrives=true', {
      method:'patch', contentType:'text/csv; charset=utf-8', payload:Utilities.newBlob(conteudo, 'text/csv').getBytes(),
      headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions:true
    });
    if(resp.getResponseCode() !== 200) return {erro:'gravacao_falhou', detalhe:'HTTP ' + resp.getResponseCode()};
  } else f = pasta.createFile(Utilities.newBlob(conteudo, 'text/csv', nome));
  var c = consolidarDevolucoes(pastaV, pasta);
  return {ok:true, arquivo:nome, solicitacoes:solicitacoes, erro:c.erro || undefined};
}

// rodar uma vez no editor: cria a chave que a extensão manda junto (copiar do registro para as opções dela)
function gerarChaveExtensao(){
  var chave = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  props().setProperty('DEVOL_EXT_CHAVE', chave);
  console.log('Chave da extensão (cole nas opções da extensão): ' + chave);
}

/* ===================== Nuvem Pago (extratos soltos numa pasta do Drive) ===================== */

// O extrato "saldo disponível" do Nuvem Pago (CSV, UTF-8 com BOM, separado por ";") fica numa pasta do Drive
// (propriedade NUVEMPAGO_PASTA ou a pasta abaixo). Uma linha por movimento: Venda, Estorno, Contestação e
// Transferência para o banco, com bruto, taxas, juros e líquido. "Nº transação" é o número do pedido da Nuvemshop.
// Os extratos se sobrepõem e não têm id de movimento: a base guarda os movimentos por dia de recebimento + Entrada/Saída
// + tipo de transação, e cada grupo vem de UM extrato só (o que tem mais linhas nele; no empate, o mais novo). Assim o
// mesmo período em dois arquivos não conta em dobro, extratos separados por entrada e saída se completam e uma correção
// num export novo substitui a antiga. O nome do cliente não é guardado. Base em nuvemshop/_base/nuvempago.json.
var NP_PASTA_ID = '1erxY_gqUBNfkSroQ1v0xxY_g0IkxNe2E';
var NP_BASE_ARQ = 'nuvempago.json';
var NP_ESTADO_ARQ = 'nuvempago-estado.json';
var NP_FORMATO = 2; // muda quando a base muda de formato: força reler todos os extratos uma vez
var NP_COLUNAS = ['Data do pagamento','Data de recebimento','Tipo de movimentação','Tipo de transação','Nº transação',
  'Forma de pagamento','Bandeira','Nº Parcelas','Valor Bruto (R$)','Taxas (R$)','Juros (R$)','Custos Totais (R$)','Valor Liquido (R$)'];
var NP_CODIFICADAS = [2, 3, 5, 6];
var NP_NUMERICAS = [8, 9, 10, 11, 12];

function pastaNuvemPago(){
  try{ return DriveApp.getFolderById(props().getProperty('NUVEMPAGO_PASTA') || NP_PASTA_ID); } catch(e){ return null; }
}

function arquivosNuvemPago(pasta){
  var lista = [], it = pasta.getFiles();
  while(it.hasNext()){
    var f = it.next();
    if(/\.csv$/i.test(f.getName()) || /csv/i.test(f.getMimeType())) lista.push(f);
  }
  lista.sort(function(a, b){ return a.getLastUpdated() - b.getLastUpdated(); });
  return lista;
}

// "Valor Liquido (R$)" ou "Valor Líquido (R$)": cabeçalho sem acento, sem caixa e sem espaço sobrando
function chaveCabecalho(c){
  return String(c == null ? '' : c).replace(/^﻿/, '').trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ');
}

// "01/10/2026" → "2026-10-01" ('' se não for data)
function diaNuvemPago(v){
  var m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(textoDevol(v));
  return m ? m[3] + '-' + m[2] + '-' + m[1] : '';
}

// arquivo → {grupos: {"dia|movimentação|tipo": [linhas]}, de, ate, n} (linha no layout de NP_COLUNAS, sem o nome)
function lerNuvemPago(f){
  var blob = f.getBlob(), texto = blob.getDataAsString('UTF-8').replace(/^﻿/, '');
  if(texto.indexOf('�') >= 0) texto = blob.getDataAsString('ISO-8859-1');
  var linhas = Utilities.parseCsv(texto, ';');
  if(!linhas.length) return {erro:'csv_vazio'};
  var cab = linhas[0].map(chaveCabecalho);
  var ix = NP_COLUNAS.map(function(c){ return cab.indexOf(chaveCabecalho(c)); });
  if(ix[0] < 0 || ix[1] < 0 || ix[3] < 0 || ix[12] < 0) return {erro:'formato_nuvempago'};
  var grupos = {}, de = '', ate = '', n = 0;
  for(var i = 1; i < linhas.length; i++){
    var l = linhas[i];
    var x = ix.map(function(j){ var v = j >= 0 ? textoDevol(l[j]) : ''; return v === '-' ? '' : v; });
    x[0] = diaNuvemPago(x[0]); x[1] = diaNuvemPago(x[1]);
    if(!x[1] || !x[3]) continue;
    x[7] = x[7] ? Number(x[7]) || '' : '';
    NP_NUMERICAS.forEach(function(k){ x[k] = numeroBR(x[k]); });
    var k = x[1] + '|' + chaveCabecalho(x[2]) + '|' + chaveCabecalho(x[3]);
    (grupos[k] = grupos[k] || []).push(x);
    n++;
    if(!de || x[1] < de) de = x[1];
    if(!ate || x[1] > ate) ate = x[1];
  }
  return {grupos:grupos, de:de, ate:ate, n:n};
}

// lê os extratos ainda não consolidados; mesma lógica de lock das devoluções. relogio = início da execução
// (no gatilho, a API da Nuvemshop vem antes e o Apps Script para em 6 min). NUVEMPAGO_VERSAO (propriedade) muda
// a cada base completa nova: entra na versão de obterVendas sem ler arquivo.
function consolidarNuvemPago(relogio){
  var pasta = pastaNuvemPago();
  if(!pasta) return {erro:'pasta_nuvempago'};
  var pastaV = pastaVendas();
  if(!pastaV) return {erro:'pasta_vendas'};
  var pastaBase = pastaBaseVendas(pastaV);
  var arquivos = arquivosNuvemPago(pasta);
  var estado = lerJsonBase(pastaBase, NP_ESTADO_ARQ) || {};
  estado.arqs = estado.arqs || {}; estado.info = estado.info || {}; estado.avisos = estado.avisos || {};
  var atuais = {};
  arquivos.forEach(function(f){ atuais[f.getId()] = f.getLastUpdated().getTime(); });
  var removido = estado.formato !== NP_FORMATO || Object.keys(estado.arqs).some(function(id){ return !(id in atuais); });
  var falta = arquivos.filter(function(f){ return estado.arqs[f.getId()] !== atuais[f.getId()]; });
  if(!removido && !falta.length) return {estado:estado, pastaBase:pastaBase, arquivos:arquivos.length};

  var lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return {erro:'consolidando', feitos:null, total:arquivos.length};
  try{
    estado = lerJsonBase(pastaBase, NP_ESTADO_ARQ) || estado;
    estado.arqs = estado.arqs || {}; estado.info = estado.info || {}; estado.avisos = estado.avisos || {};
    var base = lerJsonBase(pastaBase, NP_BASE_ARQ) || {};
    var inicio = relogio || Date.now();
    removido = estado.formato !== NP_FORMATO || Object.keys(estado.arqs).some(function(id){ return !(id in atuais); });
    // um extrato saiu da pasta, foi trocado no lugar ou a base é de um formato antigo: refaz do zero
    if(removido || falta.some(function(f){ return f.getId() in estado.arqs; })){
      base = {}; estado.arqs = {}; estado.info = {}; estado.avisos = {}; estado.formato = NP_FORMATO;
    }
    base.grupos = base.grupos || {};
    falta = arquivos.filter(function(f){ return estado.arqs[f.getId()] !== atuais[f.getId()]; });
    var feitos = 0;
    for(var i = 0; i < falta.length; i++){
      if(feitos && Date.now() - inicio > VENDAS_TEMPO_MAX) break;
      var f = falta[i], id = f.getId(), ts = atuais[id];
      var lido = lerNuvemPago(f);
      delete estado.avisos[id];
      if(lido.erro){
        estado.avisos[id] = {nome:f.getName(), erro:lido.erro};
        delete estado.info[id];
      } else {
        // o mesmo dia (e o mesmo tipo) em mais de um extrato: vale a versão com mais linhas (um export feito no meio
        // do dia fica com menos) e, no empate, o arquivo mais novo (correção feita pelo Nuvem Pago num export novo)
        Object.keys(lido.grupos).forEach(function(k){
          var novo = lido.grupos[k], atual = base.grupos[k];
          if(!atual || novo.length > atual.l.length || (novo.length === atual.l.length && ts >= atual.ts)) base.grupos[k] = {ts:ts, l:novo};
        });
        estado.info[id] = {nome:f.getName(), modificadoEm:new Date(ts).toISOString(), movimentos:lido.n, de:lido.de, ate:lido.ate};
      }
      estado.arqs[id] = ts;
      // base antes do estado: se cair entre os dois, o arquivo é lido de novo (e a junção é idempotente)
      gravarJsonBase(pastaBase, NP_BASE_ARQ, base);
      gravarJsonBase(pastaBase, NP_ESTADO_ARQ, estado);
      feitos++;
    }
    if(!feitos){
      gravarJsonBase(pastaBase, NP_BASE_ARQ, base);
      gravarJsonBase(pastaBase, NP_ESTADO_ARQ, estado);
    }
    var restam = falta.length - feitos;
    if(restam > 0) return {erro:'consolidando', feitos:arquivos.length - restam, total:arquivos.length};
    // só com a base completa: no meio de uma reconstrução as vendas não são completadas com dados parciais
    props().setProperty('NUVEMPAGO_VERSAO', String(Date.now()));
    return {estado:estado, pastaBase:pastaBase, arquivos:arquivos.length};
  } finally { lock.releaseLock(); }
}

// todos os movimentos da base, um por linha do extrato
function movimentosNuvemPago(base){
  var out = [];
  Object.keys(base.grupos || {}).forEach(function(k){ base.grupos[k].l.forEach(function(x){ out.push(x); }); });
  return out;
}

// pedido → soma das linhas de Venda e de Estorno no Nuvem Pago (para completar as vendas que vieram só da API)
function financeiroNuvemPago(base){
  var porPedido = {};
  movimentosNuvemPago(base).forEach(function(x){
    var num = x[4], tipo = chaveCabecalho(x[3]);
    if(!num || (tipo !== 'venda' && tipo !== 'estorno')) return;
    var p = porPedido[num] = porPedido[num] || {venda:false, taxas:0, juros:0, estorno:0, devCustos:0, liquido:0, dataEstorno:''};
    if(tipo === 'venda'){ p.venda = true; p.taxas -= x[9]; p.juros -= x[10]; }
    else {
      p.estorno -= x[8]; p.devCustos += x[11];
      if(!p.dataEstorno || x[0] > p.dataEstorno) p.dataEstorno = x[0];
    }
    p.liquido += x[12];
  });
  return porPedido;
}

// pedido sem Total líquido (veio só da API): Taxas, Estorno, Devolução dos custos e Total líquido saem do Nuvem Pago,
// no mesmo sentido e formato do CSV da Nuvemshop (Líquido = Total − Taxas − Juros − Estorno + Devolução dos custos;
// lá "Taxas" já soma taxa e juros e "Juros" fica vazio). Campo já preenchido não muda. Pedido estornado na Nuvemshop
// sem a linha de Estorno no extrato fica sem completar ("aguardando Nuvem Pago"): o líquido cheio seria maior que o real.
// Só na resposta: a base de vendas fica como veio.
function completarComNuvemPago(pedidos, fin){
  var r2 = function(n){ return String(Math.round(n * 100) / 100); };
  Object.keys(pedidos).forEach(function(num){
    var r = pedidos[num], p = fin[num];
    if(!p || !p.venda || (r[23] !== '' && r[23] != null)) return;
    if(!p.estorno && (/estorn/i.test(r[3]) || Number(r[17]) > 0)) return;
    r[19] = r2(p.taxas + p.juros); r[23] = r2(p.liquido);
    if(p.estorno){
      if(r[17] === '' || r[17] == null) r[17] = r2(p.estorno);
      if(r[22] === '' || r[22] == null) r[22] = r2(p.devCustos);
      if((r[18] === '' || r[18] == null) && p.dataEstorno) r[18] = p.dataEstorno.split('-').reverse().join('/');
    }
  });
  return pedidos;
}

function lerBaseNuvemPago(pastaBase){
  return lerJsonBase(pastaBase, NP_BASE_ARQ) || {grupos:{}};
}

// versao: a que a página já tem; se nada mudou, responde {mudou:false}
function obterNuvemPago(token, versao){
  var conta = contaComToken(token);
  if(!conta) return {erro:'senha_necessaria'};
  if(!podeVerFinanceiro(conta.email)) return {erro:'sem_permissao'};
  var c = consolidarNuvemPago();
  if(c.erro) return c;
  if(!c.arquivos) return {erro:'sem_arquivo'};
  var assinatura = Object.keys(c.estado.arqs).sort().map(function(id){ return id + '@' + c.estado.arqs[id]; }).join('|');
  var novaVersao = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, 'np' + NP_FORMATO + '#' + assinatura));
  if(versao && versao === novaVersao) return {mudou:false, versao:novaVersao};
  var dic = {}, idxDic = {};
  NP_CODIFICADAS.forEach(function(k){ dic[k] = []; idxDic[k] = Object.create(null); });
  var linhas = movimentosNuvemPago(lerBaseNuvemPago(c.pastaBase)).map(function(x){
    x = x.slice();
    NP_CODIFICADAS.forEach(function(c){
      var v = x[c] == null ? '' : String(x[c]), i = idxDic[c][v];
      if(i === undefined){ i = idxDic[c][v] = dic[c].length; dic[c].push(v); }
      x[c] = i;
    });
    return x;
  });
  var arquivos = Object.keys(c.estado.info).map(function(id){ return c.estado.info[id]; })
    .sort(function(a, b){ return a.modificadoEm < b.modificadoEm ? -1 : 1; });
  var avisos = Object.keys(c.estado.avisos).map(function(id){ return c.estado.avisos[id]; });
  return {mudou:true, versao:novaVersao, colunas:NP_COLUNAS, dic:dic, linhas:linhas,
    arquivos:arquivos, avisos:avisos, pasta:props().getProperty('NUVEMPAGO_PASTA') || NP_PASTA_ID};
}

/* ===================== Saldo conhecido do Nuvem Pago (compartilhado no servidor) ===================== */

// O dono informa o saldo do Nuvem Pago no fim de um dia (tela de Saldo do Nuvem Pago). O valor fica numa propriedade do script,
// com data, autor e hora, e vale para todo o perfil financeiro (antes ficava só no navegador de cada pessoa). Cada gravação
// também entra num histórico dos 20 últimos. Quem não tem acesso ao financeiro não lê nem grava.
var NP_SALDO_PROP = 'NP_SALDO_CONHECIDO';
var NP_SALDO_HIST = 20;

function lerSaldoNP_(){
  var x = null;
  try{ x = JSON.parse(props().getProperty(NP_SALDO_PROP) || 'null'); } catch(e){}
  return x && typeof x === 'object' ? x : {data:'', valor:null, autor:'', em:'', hist:[]};
}

// {saldo:{data,valor,autor,em}|null, hist:[...]} ou {erro}
function obterSaldoNP(token){
  var conta = contaComToken(token);
  if(!conta) return {erro:'senha_necessaria'};
  if(!podeVerFinanceiro(conta.email)) return {erro:'sem_permissao'};
  var x = lerSaldoNP_();
  var tem = x.data && typeof x.valor === 'number' && isFinite(x.valor);
  return {saldo: tem ? {data:x.data, valor:x.valor, autor:x.autor || '', em:x.em || ''} : null, hist:x.hist || []};
}

// dados: {data:'AAAA-MM-DD', valor:número} grava; {limpar:true} limpa o saldo atual (o histórico fica)
function salvarSaldoNP(token, dados){
  var conta = contaComToken(token);
  if(!conta) return {erro:'senha_necessaria'};
  if(!podeVerFinanceiro(conta.email)) return {erro:'sem_permissao'};
  dados = dados || {};
  var limpar = dados.limpar === true;
  var data = String(dados.data || ''), valor = Number(dados.valor);
  if(!limpar){
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(data);
    var d = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
    if(!d || d.getMonth() !== Number(m[2]) - 1 || d.getDate() !== Number(m[3])) return {erro:'data_invalida'};
    if(dados.valor === '' || dados.valor == null || !isFinite(valor) || Math.abs(valor) > 1e9) return {erro:'valor_invalido'};
    valor = Math.round(valor * 100) / 100;
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    var x = lerSaldoNP_(), agora = new Date().toISOString();
    var hist = (x.hist || []).slice();
    if(limpar){
      hist.unshift({data:'', valor:null, autor:conta.email, em:agora, acao:'limpou'});
      x = {data:'', valor:null, autor:conta.email, em:agora, hist:hist.slice(0, NP_SALDO_HIST)};
    } else {
      hist.unshift({data:data, valor:valor, autor:conta.email, em:agora});
      x = {data:data, valor:valor, autor:conta.email, em:agora, hist:hist.slice(0, NP_SALDO_HIST)};
    }
    props().setProperty(NP_SALDO_PROP, JSON.stringify(x));
  } finally { lock.releaseLock(); }
  return obterSaldoNP(token);
}

/* ===================== Relatório da manhã por e-mail ===================== */

// Uma rotina do Claude (na nuvem) monta o relatório da manhã a partir da Nuvemshop e grava um .html na
// pasta abaixo. enviarRelatoriosDiarios() roda por gatilho a cada 15 min, manda cada arquivo novo por
// e-mail (assunto = <title> do arquivo) e marca o arquivo como enviado na descrição dele.
// Para ligar: rodar instalarRelatorioDiario() uma vez no editor, com a conta dona do script.
var RELATORIO_PASTA_ID = '11w2NTecfVU82w5qUiCnn7VAPe9kpU_5K';
var RELATORIO_PARA = 'messias.alexandro@talgui.com.br';

function enviarRelatoriosDiarios(){
  var it = DriveApp.getFolderById(RELATORIO_PASTA_ID).getFiles();
  while(it.hasNext()){
    var f = it.next();
    if(!/\.html?$/i.test(f.getName()) || /^enviado/.test(f.getDescription() || '')) continue;
    var html = f.getBlob().getDataAsString('UTF-8');
    var t = /<title>([\s\S]*?)<\/title>/i.exec(html);
    MailApp.sendEmail({to:RELATORIO_PARA, subject:t ? t[1].trim() : f.getName(), htmlBody:html, name:'Painel Talgui'});
    f.setDescription('enviado ' + new Date().toISOString());
  }
}

function instalarRelatorioDiario(){
  ScriptApp.getProjectTriggers().forEach(function(g){
    if(g.getHandlerFunction() === 'enviarRelatoriosDiarios') ScriptApp.deleteTrigger(g);
  });
  ScriptApp.newTrigger('enviarRelatoriosDiarios').timeBased().everyMinutes(15).create();
  enviarRelatoriosDiarios();
}
