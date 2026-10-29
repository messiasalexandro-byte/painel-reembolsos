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
  var conta = verificarConta_();
  if(conta.erro){
    var msg = 'Você não tem acesso a este painel' + (conta.email ? ' com a conta ' + conta.email : '') + '. Peça a liberação ao responsável.';
    return HtmlService.createHtmlOutput(
      '<div style="font-family:system-ui,sans-serif;max-width:480px;margin:15vh auto;padding:0 16px;color:#1f2933;">' +
      '<h1 style="font-size:20px;">Painel de Reembolsos</h1><p style="line-height:1.5;">' + escaparHtml_(msg) + '</p></div>'
    ).setTitle('Painel Financeiro · Talgui').addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Painel Financeiro · Talgui')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function escaparHtml_(s){
  return String(s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; });
}

/* ===================== Acesso (cadastro por e-mail + senha de cada um) ===================== */

function props_(){ return PropertiesService.getScriptProperties(); }

function normalizarEmail_(e){ return String(e || '').trim().toLowerCase(); }

function emailAtual_(){ return normalizarEmail_(Session.getActiveUser().getEmail()); }

function emailDono_(){ return normalizarEmail_(Session.getEffectiveUser().getEmail()); }

/* ===================== Proteção das funções públicas (etapa 27) ===================== */

// Qualquer função de topo SEM "_" no fim pode ser chamada por qualquer pessoa do domínio via google.script.run.
// Por isso: (1) tudo o que é interno termina em "_"; (2) as que a página chama exigem o token do painel
// (contaComToken_) e, conforme o caso, perfil financeiro ou admin; (3) as que precisam manter o nome sem "_"
// (handlers de gatilho já instalados e funções que o dono roda à mão no editor) chamam exigirDono_().
//
// exigirDono_ recusa só quando quem executa é OUTRA pessoa: Session.getActiveUser() devolve o e-mail do visitante
// numa chamada vinda do web app; no editor devolve o dono; em gatilho de tempo devolve o dono (ou vazio, conforme o
// tipo de conta) e Session.getEffectiveUser() é o dono em todos os casos. Por isso "vazio ou igual ao dono" passa
// (gatilho e editor continuam funcionando) e "e-mail de outra pessoa" é bloqueado (visitante do web app).
function exigirDono_(){
  var atual = emailAtual_();
  if(atual && atual !== emailDono_()) throw new Error('Esta função é restrita ao dono do painel.');
}

// Para funções que só o dono roda à mão no editor (nenhum gatilho as chama): ali o e-mail vem sempre preenchido,
// então "vazio" também é recusado (fecha o caso de um visitante do web app com e-mail oculto).
function exigirDonoEditor_(){
  var atual = emailAtual_();
  if(!atual || atual !== emailDono_()) throw new Error('Esta função é restrita ao dono do painel (rodar no editor).');
}

// Versão deste código. Só mude ao publicar uma versão nova (veja gatilhoLiberado_ e marcarProducaoVersao).
var CODIGO_VERSAO = 'etapa30-2026-10-29';

// Os gatilhos de 15 min rodam o código HEAD (o último push, mesmo que só para teste) sobre as bases reais.
// Depois que o dono roda marcarProducaoVersao() (com a versão publicada), os gatilhos só trabalham se
// CODIGO_VERSAO == PRODUCAO_VERSAO: um push de teste com CODIGO_VERSAO diferente deixa os gatilhos parados até publicar.
// Sem PRODUCAO_VERSAO gravada (antes da primeira marcação) os gatilhos rodam como sempre.
function gatilhoLiberado_(){
  var prod = props_().getProperty('PRODUCAO_VERSAO');
  if(!prod || prod === CODIGO_VERSAO) return true;
  console.warn('Gatilho parado: o código (' + CODIGO_VERSAO + ') não é o publicado (' + prod + '). Publique e rode marcarProducaoVersao().');
  return false;
}

// rodar no editor logo DEPOIS de publicar a versão nova no /exec (e sem ter dado push de outra versão antes)
function marcarProducaoVersao(){
  exigirDonoEditor_();
  props_().setProperty('PRODUCAO_VERSAO', CODIGO_VERSAO);
  console.log('Versão de produção marcada: ' + CODIGO_VERSAO);
}

/* ----- Leitura de base com falha não pode virar "base vazia" ----- */

// Erro de leitura/interpretação de um arquivo da base (ou propriedade de estado). "Arquivo não existe" NÃO é erro
// (base nova): lerJsonBase_ devolve null. Já arquivo que existe e não abre/não é JSON válido interrompe a rodada
// antes de qualquer gravação, para o painel não gravar uma base vazia por cima da boa.
function erroBase_(onde, detalhe){
  var e = new Error('Não consegui ler ' + onde + (detalhe ? ' (' + detalhe + ')' : '') + '. Nada foi gravado.');
  e.baseIlegivel = true; e.onde = String(onde); e.detalhe = String(detalhe || '');
  return e;
}

function textoErro_(e){ return String(e && e.message || e); }

// guarda o último erro de leitura por área na propriedade BASE_ERRO ({área:{em, onde, detalhe}})
function registrarErroBase_(area, e){
  try{
    var p = props_(), todos = {};
    try{ todos = JSON.parse(p.getProperty('BASE_ERRO') || '{}') || {}; } catch(x){ todos = {}; }
    todos[area] = {em:new Date().toISOString(), onde:String(e.onde || '').slice(0, 120), detalhe:String(e.detalhe || '').slice(0, 300)};
    p.setProperty('BASE_ERRO', JSON.stringify(todos));
  } catch(x){}
  console.error('Base ilegível (' + area + '): ' + textoErro_(e));
}

function limparErroBase_(area){
  try{
    var p = props_(), raw = p.getProperty('BASE_ERRO');
    if(raw === null) return;
    var todos = JSON.parse(raw) || {};
    if(!(area in todos)) return;
    delete todos[area];
    if(Object.keys(todos).length) p.setProperty('BASE_ERRO', JSON.stringify(todos)); else p.deleteProperty('BASE_ERRO');
  } catch(x){}
}

// roda fn(); se a leitura de uma base falhar, registra e devolve {erro:'base_ilegivel'} em vez de seguir com base vazia
function comErroBase_(area, fn){
  var r;
  try{ r = fn(); }
  catch(e){
    if(!e || !e.baseIlegivel) throw e;
    registrarErroBase_(area, e);
    return {erro:'base_ilegivel', detalhe:e.message};
  }
  // um erro base_ilegivel que veio "de dentro" (devolvido por outra função já protegida) continua registrado
  if(!(r && r.erro === 'base_ilegivel')) limparErroBase_(area);
  return r;
}

function ehAdmin_(email){
  if(!email) return false;
  if(email === emailDono_()) return true;
  return (props_().getProperty('PAINEL_ADMINS') || '').split(/[,;\s]+/).map(normalizarEmail_).indexOf(email) >= 0;
}

// Recebimentos do Nuvem Pago (saldo e transferências bancárias): só o dono do script e quem estiver em
// NUVEMPAGO_LEITORES (e-mails separados por vírgula). Os demais seguem vendo as vendas completadas com as taxas.
function podeVerFinanceiro_(email){
  if(!email) return false;
  if(email === emailDono_()) return true;
  return (props_().getProperty('NUVEMPAGO_LEITORES') || '').split(/[,;\s]+/).map(normalizarEmail_).indexOf(email) >= 0;
}

// Lista antiga (PAINEL_EMAILS, da senha única): vira, uma vez só, cadastro pendente para cada e-mail.
// Não apaga PAINEL_EMAILS/PAINEL_SENHA (podem ser removidas à mão depois de publicar a versão nova).
function migrarListaAntiga_(){
  var p = props_();
  if(p.getProperty('PAINEL_MIGRADO') !== null) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    if(p.getProperty('PAINEL_MIGRADO') !== null) return;
    (p.getProperty('PAINEL_EMAILS') || '').split(/[,;\s]+/).map(normalizarEmail_).forEach(function(e){
      if(e && p.getProperty(PREFIXO_USUARIO + e) === null) p.setProperty(PREFIXO_USUARIO + e, JSON.stringify({adicionadoEm: new Date().toISOString()}));
    });
    p.setProperty('PAINEL_MIGRADO', new Date().toISOString());
  } finally { lock.releaseLock(); }
}

// cadastro da pessoa: {hash, sal, senhaEm} depois de criar a senha; {adicionadoEm} enquanto aguarda.
// foto (opcional): dataURL JPEG 64×64 (até 4000 caracteres) enviado pela página Minha conta; cadastro sem ela continua valendo.
// O administrador sem cadastro conta como pendente, para nunca ficar trancado do lado de fora.
function lerUsuario_(email){
  var raw = props_().getProperty(PREFIXO_USUARIO + email);
  if(raw === null) return ehAdmin_(email) ? {} : null;
  try{ return JSON.parse(raw); } catch(e){ return {}; }
}

function gravarUsuario_(email, u){ props_().setProperty(PREFIXO_USUARIO + email, JSON.stringify(u)); }

// {email, usuario, admin} se a conta tem cadastro; {erro} se não tem.
function verificarConta_(){
  migrarListaAntiga_();
  var email = emailAtual_();
  var u = email ? lerUsuario_(email) : null;
  if(!u) return {erro:'sem_acesso', email:email};
  return {email:email, usuario:u, admin:ehAdmin_(email)};
}

function segredo_(){
  var p = props_();
  var s = p.getProperty('PAINEL_SEGREDO');
  if(!s){ s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('PAINEL_SEGREDO', s); }
  return s;
}

function bytes_(s){ return Utilities.newBlob(String(s)).getBytes(); }

function hashSenha_(senha, sal){
  var chave = bytes_(sal);
  var d = Utilities.computeHmacSha256Signature(bytes_(senha), chave);
  for(var i=1; i<HASH_RODADAS; i++) d = Utilities.computeHmacSha256Signature(d, chave);
  return Utilities.base64Encode(d);
}

function tokenPara_(email, u){
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(email + '|' + u.hash, segredo_()));
}

// comparação sem sair no primeiro caractere diferente
function iguais_(a, b){
  a = String(a || ''); b = String(b || '');
  var dif = a.length ^ b.length;
  for(var i=0; i<Math.max(a.length, b.length); i++) dif |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return dif === 0;
}

// conta já conferida e com token válido; null caso contrário
function contaComToken_(token){
  var conta = verificarConta_();
  if(conta.erro || !conta.usuario.hash || !iguais_(token, tokenPara_(conta.email, conta.usuario))) return null;
  return conta;
}

// o que a tela de entrada deve mostrar: criar a senha (primeiro acesso) ou digitá-la
function estadoAcesso(){
  var conta = verificarConta_();
  if(conta.erro) return {erro:conta.erro};
  return {email:conta.email, cadastro: conta.usuario.hash ? 'entrar' : 'criar'};
}

function criarSenha(senha){
  var conta = verificarConta_();
  if(conta.erro) return {erro:conta.erro};
  if(conta.usuario.hash) return {erro:'ja_cadastrado'};
  senha = String(senha || '');
  if(senha.length < SENHA_MIN) return {erro:'senha_curta', minimo:SENHA_MIN};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    var atual = lerUsuario_(conta.email);
    if(!atual) return {erro:'sem_acesso'};
    if(atual.hash) return {erro:'ja_cadastrado'};
    var sal = Utilities.getUuid();
    atual.sal = sal;
    atual.hash = hashSenha_(senha, sal);
    atual.senhaEm = new Date().toISOString();
    gravarUsuario_(conta.email, atual);
    return {token: tokenPara_(conta.email, atual), admin: conta.admin, financeiro: podeVerFinanceiro_(conta.email), foto: atual.foto || ''};
  } finally { lock.releaseLock(); }
}

function entrar(senha){
  var conta = verificarConta_();
  if(conta.erro) return {erro:conta.erro};
  if(!conta.usuario.hash) return {erro:'criar_senha'};
  var cache = CacheService.getScriptCache();
  var chave = 'tentativas_' + conta.email;
  var erros = Number(cache.get(chave) || 0);
  if(erros >= MAX_TENTATIVAS) return {erro:'bloqueado'};
  if(!iguais_(hashSenha_(String(senha || ''), conta.usuario.sal), conta.usuario.hash)){
    erros++;
    cache.put(chave, String(erros), BLOQUEIO_SEG);
    return erros >= MAX_TENTATIVAS ? {erro:'bloqueado'} : {erro:'senha_incorreta', restantes: MAX_TENTATIVAS - erros};
  }
  cache.remove(chave);
  return {token: tokenPara_(conta.email, conta.usuario), admin: conta.admin, financeiro: podeVerFinanceiro_(conta.email), foto: conta.usuario.foto || ''};
}

// a página confere, ao abrir com um token guardado, se ele ainda vale e se é admin (e pega a foto do menu)
function sessao(token){
  var conta = contaComToken_(token);
  return conta ? {email:conta.email, admin:conta.admin, financeiro:podeVerFinanceiro_(conta.email), foto:conta.usuario.foto || ''} : {erro:'senha_necessaria'};
}

/* ----- Minha conta (cada pessoa, com a própria sessão) ----- */

var FOTO_PREFIXO = 'data:image/jpeg;base64,';
var FOTO_MAX = 4000; // cada propriedade aceita ~9 KB e todas juntas no máximo 500 KB (~100 pessoas com foto)

// dataUrl '' remove a foto; senão, só JPEG em base64 até FOTO_MAX caracteres
function salvarFoto(token, dataUrl){
  var conta = contaComToken_(token);
  if(!conta) return {erro:'senha_necessaria'};
  var foto = String(dataUrl || '');
  if(foto && (foto.length > FOTO_MAX || foto.indexOf(FOTO_PREFIXO) !== 0 ||
      !/^[A-Za-z0-9+\/]+={0,2}$/.test(foto.slice(FOTO_PREFIXO.length)))) return {erro:'foto_invalida'};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    var atual = lerUsuario_(conta.email);
    // a senha pode ter sido trocada ou liberada enquanto a página estava aberta
    if(!atual || !atual.hash || atual.hash !== conta.usuario.hash) return {erro:'senha_necessaria'};
    if(foto) atual.foto = foto; else delete atual.foto;
    gravarUsuario_(conta.email, atual);
    return {ok:true, foto:foto};
  } finally { lock.releaseLock(); }
}

// confere a senha atual como entrar() (mesmo limite de tentativas) e grava o hash novo com sal novo.
// O token é HMAC do e-mail + hash: o novo vale neste aparelho e os outros passam a pedir a senha.
function trocarSenha(token, atualSenha, novaSenha){
  var conta = contaComToken_(token);
  if(!conta) return {erro:'senha_necessaria'};
  var cache = CacheService.getScriptCache();
  var chave = 'tentativas_' + conta.email;
  var erros = Number(cache.get(chave) || 0);
  if(erros >= MAX_TENTATIVAS) return {erro:'bloqueado'};
  if(!iguais_(hashSenha_(String(atualSenha || ''), conta.usuario.sal), conta.usuario.hash)){
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
    var atual = lerUsuario_(conta.email);
    if(!atual || !atual.hash || atual.hash !== conta.usuario.hash) return {erro:'senha_necessaria'};
    var sal = Utilities.getUuid();
    atual.sal = sal;
    atual.hash = hashSenha_(novaSenha, sal);
    atual.senhaEm = new Date().toISOString();
    gravarUsuario_(conta.email, atual);
    return {token: tokenPara_(conta.email, atual)};
  } finally { lock.releaseLock(); }
}

/* ----- Tela Acesso (só administradores) ----- */

function contaAdmin_(token){
  var conta = contaComToken_(token);
  return conta && conta.admin ? conta : null;
}

function adminListar(token){
  if(!contaAdmin_(token)) return {erro:'senha_necessaria'};
  var todas = props_().getProperties();
  var dono = emailDono_();
  var lista = [];
  Object.keys(todas).forEach(function(k){
    if(k.indexOf(PREFIXO_USUARIO) !== 0) return;
    var email = k.slice(PREFIXO_USUARIO.length), u = {};
    try{ u = JSON.parse(todas[k]); } catch(e){}
    lista.push({email:email, situacao: u.hash ? 'ativo' : 'aguardando', desde: u.hash ? u.senhaEm : u.adicionadoEm, admin: ehAdmin_(email), dono: email === dono});
  });
  lista.sort(function(a, b){ return a.email < b.email ? -1 : 1; });
  return {usuarios:lista};
}

function adminAdicionar(token, email){
  if(!contaAdmin_(token)) return {erro:'senha_necessaria'};
  email = normalizarEmail_(email);
  if(!/^[^@\s]+@[^@\s]+$/.test(email) || email.split('@')[1] !== DOMINIO) return {erro:'email_invalido'};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    if(props_().getProperty(PREFIXO_USUARIO + email) !== null) return {erro:'ja_existe'};
    gravarUsuario_(email, {adicionadoEm: new Date().toISOString()});
  } finally { lock.releaseLock(); }
  return adminListar(token);
}

// apaga a senha da pessoa: no próximo acesso ela cria outra (e o token antigo deixa de valer)
function adminLiberarCadastro(token, email){
  var conta = contaAdmin_(token);
  if(!conta) return {erro:'senha_necessaria'};
  email = normalizarEmail_(email);
  if(email === conta.email) return {erro:'proprio_usuario'};
  if(props_().getProperty(PREFIXO_USUARIO + email) === null) return {erro:'nao_encontrado'};
  var novo = {adicionadoEm: new Date().toISOString()};
  var antigo = lerUsuario_(email);
  if(antigo && antigo.foto) novo.foto = antigo.foto; // a foto da pessoa sobrevive ao novo cadastro
  gravarUsuario_(email, novo);
  CacheService.getScriptCache().remove('tentativas_' + email);
  return adminListar(token);
}

function adminRemover(token, email){
  var conta = contaAdmin_(token);
  if(!conta) return {erro:'senha_necessaria'};
  email = normalizarEmail_(email);
  if(email === conta.email) return {erro:'proprio_usuario'};
  if(email === emailDono_()) return {erro:'dono'};
  props_().deleteProperty(PREFIXO_USUARIO + email);
  return adminListar(token);
}

/* ===================== Planilha do Drive ===================== */

// Planilha Google tem prioridade: ao converter um .xlsx, o original costuma ficar na pasta
// e não deve voltar a ser usado só porque alguém mexeu nele. Sem Planilha Google, usa o .xlsx mais recente.
function arquivoMaisRecente_(){
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
  if(!contaComToken_(token)) return {erro:'senha_necessaria'};
  var arquivo = arquivoMaisRecente_();
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
// Só as colunas abaixo entram na base e saem do servidor. Do cliente entram o nome do comprador e a cidade de entrega
// (decisão do usuário em 29/10/2026, para os rankings de clientes e de cidades); o e-mail só como "Cliente", uma chave
// (resumo) que junta os pedidos da mesma pessoa vindos do CSV e da API. Nada de e-mail, CPF, telefone ou endereço.
// O CSV vem com uma linha por produto: os dados do pedido ficam na primeira linha dele.
var VENDAS_PASTA_NOME = 'nuvemshop';
var VENDAS_BASE_PASTA = '_base';
var VENDAS_COLUNAS = ['Número do Pedido','Data','Status do Pedido','Status do Pagamento','Status do Envio',
  'Subtotal','Desconto','Valor do Frete','Total','Estado','Forma de Entrega','Forma de Pagamento','Cupom de Desconto',
  'Data de pagamento','Canal','Data e hora do cancelamento','Motivo do cancelamento','Estorno','Data de estorno',
  'Taxas','Parcelas','Juros','Devolução dos custos','Total líquido','Meio de pagamento',
  'Cidade','Nome do comprador','Cliente'];
var VENDAS_COL_CLIENTE = 27;              // "Cliente": chaveCliente_ do e-mail (não vem pronto no CSV nem na API)
var VENDAS_ITENS = VENDAS_COLUNAS.length; // posição da soma de "Quantidade Comprada" na linha do pedido
var VENDAS_SKUS = VENDAS_ITENS + 1;       // "sku:qtd,sku:qtd" do pedido (só na base: alimenta a taxa de devolução por peça)
var VENDAS_TS = VENDAS_SKUS + 1;          // posição da data de modificação do CSV de origem (só na base)
var VENDAS_FORMATO = 3;                   // muda quando a linha da base muda: força reler todos os CSVs uma vez
var VENDAS_PASTAS_FORA = ['_base', 'AJD']; // subpastas de "nuvemshop" que não têm CSV de vendas (AJD = devoluções)
// na resposta, textos repetidos viram índice de um dicionário e valores viram número (resposta ~3x menor)
var VENDAS_CODIFICADAS = [2, 3, 4, 9, 10, 11, 14, 16, 24, 25, 26, 27];
var VENDAS_NUMERICAS = [5, 6, 7, 8, 17, 19, 21, 22, 23];
var VENDAS_TEMPO_MAX = 4 * 60 * 1000; // o Apps Script para em 6 min: guarda o progresso antes disso

// pasta pela propriedade NUVEMSHOP_PASTA (id) ou pelo nome; o id achado fica guardado
function pastaVendas_(){
  var p = props_();
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

var PASTA_BASE_CACHE_ = {}; // vale só nesta execução do Apps Script
function pastaBaseVendas_(pasta){
  var chave = pasta.getId();
  if(PASTA_BASE_CACHE_[chave]) return PASTA_BASE_CACHE_[chave];
  var it = pasta.getFoldersByName(VENDAS_BASE_PASTA);
  return PASTA_BASE_CACHE_[chave] = it.hasNext() ? it.next() : pasta.createFolder(VENDAS_BASE_PASTA);
}

// CSVs da pasta e das subpastas (menos _base e AJD), do mais antigo para o mais novo
function arquivosVendas_(pasta){
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
function lerCsvVendas_(arquivo){
  var blob = arquivo.getBlob();
  var texto = blob.getDataAsString('ISO-8859-1');
  if(texto.indexOf('Número do Pedido') < 0) texto = blob.getDataAsString('UTF-8').replace(/^﻿/, '');
  var linhas = Utilities.parseCsv(texto, ';');
  if(!linhas.length) return {erro:'csv_vazio'};
  var cab = linhas[0].map(function(c){ return String(c).trim(); });
  var idx = VENDAS_COLUNAS.map(function(c){ return cab.indexOf(c); });
  if(idx[0] < 0 || idx[1] < 0 || idx[8] < 0) return {erro:'csv_formato'};
  var iQtd = cab.indexOf('Quantidade Comprada'), iSku = cab.indexOf('SKU'), iNome = cab.indexOf('Nome do Produto');
  var iEmail = cab.indexOf('E-mail'), iPreco = cab.indexOf('Valor do Produto');
  var porNum = {}, ordem = [], nomes = {};
  for(var i = 1; i < linhas.length; i++){
    var l = linhas[i];
    var num = String(l[idx[0]] || '').trim();
    if(!num) continue;
    var qtd = iQtd >= 0 ? Number(l[iQtd]) || 0 : 0;
    var sku = iSku >= 0 ? String(l[iSku] || '').trim().replace(/[,:]/g, '') : '';
    if(sku && iNome >= 0 && l[iNome]) nomes[sku] = String(l[iNome]).trim();
    var item = sku ? itemSkuTexto_(sku, qtd, iPreco >= 0 ? l[iPreco] : '') : '';
    if(porNum[num]){
      porNum[num][VENDAS_ITENS] += qtd;
      if(item) porNum[num][VENDAS_SKUS] += (porNum[num][VENDAS_SKUS] ? ',' : '') + item;
      continue;
    }
    var r = idx.map(function(j){ return j >= 0 ? String(l[j] == null ? '' : l[j]).trim() : ''; });
    r[VENDAS_COL_CLIENTE] = iEmail >= 0 ? chaveCliente_(l[iEmail]) : '';
    r.push(qtd, item);
    porNum[num] = r;
    ordem.push(num);
  }
  return {linhas: ordem.map(function(n){ return porNum[n]; }), nomes: nomes};
}

// e-mail → chave do cliente: 12 caracteres do HMAC-SHA256 do e-mail em minúsculas, com o segredo da propriedade
// CLIENTE_SEGREDO (criado na primeira vez; '' sem e-mail). O e-mail nunca é guardado, e sem o segredo a chave não volta
// ao e-mail nem por tentativa com uma lista de e-mails. Só serve para juntar os pedidos da mesma pessoa.
// Trocar ou apagar CLIENTE_SEGREDO muda todas as chaves: depois disso a base de vendas precisa ser refeita.
var CLIENTE_SEGREDO_ = null; // vale só nesta execução
function chaveCliente_(email){
  var e = String(email == null ? '' : email).trim().toLowerCase();
  if(!e || e.indexOf('@') < 0) return '';
  if(!CLIENTE_SEGREDO_){
    var p = props_();
    CLIENTE_SEGREDO_ = p.getProperty('CLIENTE_SEGREDO');
    if(!CLIENTE_SEGREDO_){ CLIENTE_SEGREDO_ = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('CLIENTE_SEGREDO', CLIENTE_SEGREDO_); }
  }
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(e, CLIENTE_SEGREDO_)).slice(0, 12);
}

// item do pedido na base: "sku:qtd:preço unitário" (preço com ponto; desde o formato 3 — antes era só "sku:qtd").
// O preço é o do produto no pedido, antes do desconto do pedido (CSV "Valor do Produto"; API products[].price).
function itemSkuTexto_(sku, qtd, preco){
  var t = String(preco == null ? '' : preco).trim(), n;
  if(/,\d{1,2}$/.test(t)) t = t.replace(/\./g, '').replace(',', '.');
  n = Number(t);
  return sku + ':' + qtd + (t !== '' && isFinite(n) ? ':' + Math.round(n * 100) / 100 : '');
}
function itensSku_(texto){
  return String(texto || '').split(',').map(function(item){
    var p = item.split(':');
    return {sku:p[0], q:Number(p[1]) || 0, preco:p.length > 2 && p[2] !== '' ? Number(p[2]) || 0 : null};
  }).filter(function(it){ return it.sku; });
}

// "16/08/2026 22:09:01" → "2026-08-16" ('' se não for data)
function diaIsoVendas_(s){
  var m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(String(s || ''));
  return m ? m[3] + '-' + m[2] + '-' + m[1] : '';
}

// junta as linhas de um CSV (ts = modificação dele) ou da API (ts = updated_at) na base de um ano: o pedido
// só é trocado se a fonte for tão nova quanto o que já está lá — assim a ordem de leitura não importa.
// Campo vazio na fonte nova não apaga o da antiga (a API não traz Taxas/Juros/Líquido, que vêm do CSV).
function juntarPedidos_(pedidos, linhas, ts){
  linhas.forEach(function(r){
    var atual = pedidos[r[0]];
    if(atual && atual[VENDAS_TS] > ts) return;
    var x = r.slice(0, VENDAS_TS);
    if(atual) x = x.map(function(v, i){ return (v === '' || v == null) && atual[i] !== '' && atual[i] != null ? atual[i] : v; });
    x[VENDAS_TS] = ts;
    pedidos[r[0]] = x;
  });
}

// O controle da base de vendas ({formato, arqs, info, avisos, anos, apiMudou}) mora em _base/vendas3-estado.json.
// Etapa 30 (formato 3, com cidade e cliente): a base nova é PARALELA à antiga — vendas3-estado.json, vendas3-AAAA.json e
// a propriedade VENDAS_API3. Enquanto o /exec roda a versão anterior, ela segue lendo e gravando vendas-estado.json,
// vendas-AAAA.json, VENDAS_API e VENDAS_BASE (o espelho da etapa 27); se as duas versões usassem os mesmos arquivos,
// uma zeraria a base da outra a cada chamada (o formato diferente força refazer do zero). A base nova começa vazia e é
// montada pelos CSVs da pasta e pela API desde o início da loja. Depois de publicar e rodar marcarProducaoVersao(),
// limparVendasBaseAntiga() apaga o que era só da versão anterior.
var VENDAS_ESTADO_ARQ = 'vendas3-estado.json';
var VENDAS_ANO_PREFIXO = 'vendas3-';
var VENDAS_API_PROP = 'VENDAS_API3';
var VENDAS_ANTIGA = {estado:'vendas-estado.json', prefixo:'vendas-', api:'VENDAS_API', prop:'VENDAS_BASE', espelho:'VENDAS_BASE_ESPELHO'};

function normalizarEstadoVendas_(e){
  e = e && typeof e === 'object' ? e : {};
  e.arqs = e.arqs || {}; e.info = e.info || {}; e.avisos = e.avisos || {}; e.anos = e.anos || {};
  return e;
}

// sem a pasta ou sem o arquivo = base nova (vazia); arquivo que não abre lança erro (lerJsonBase_)
function estadoVendas_(){
  var pasta = pastaVendas_();
  if(!pasta) return normalizarEstadoVendas_(null);
  return normalizarEstadoVendas_(lerJsonBase_(pastaBaseVendas_(pasta), VENDAS_ESTADO_ARQ));
}

function gravarEstadoVendas_(estado, pastaBase){ gravarJsonBase_(pastaBase, VENDAS_ESTADO_ARQ, estado); }

// rodar no editor DEPOIS de publicar a versão nova no /exec e de rodar marcarProducaoVersao(): apaga a base da versão
// anterior (vendas-AAAA.json e vendas-estado.json vão para a lixeira) e as propriedades VENDAS_API, VENDAS_BASE e
// VENDAS_BASE_ESPELHO. Antes disso a versão antiga publicada ainda as usa.
function limparVendasBaseAntiga(){
  exigirDonoEditor_();
  var p = props_();
  if(p.getProperty('PRODUCAO_VERSAO') !== CODIGO_VERSAO) throw new Error('Publique a versão nova e rode marcarProducaoVersao() antes de limpar a base antiga.');
  var pasta = pastaVendas_();
  if(!pasta) throw new Error('Pasta nuvemshop não encontrada.');
  var pastaBase = pastaBaseVendas_(pasta);
  var nova = lerJsonBase_(pastaBase, VENDAS_ESTADO_ARQ);
  if(!nova || !Object.keys(nova.anos || {}).length) throw new Error('A base nova (' + VENDAS_ESTADO_ARQ + ') ainda não existe: não apaguei a antiga.');
  // só depois de a base nova estar completa: todos os CSVs lidos e a carga da API terminada
  var cv = consolidarVendas_(pasta);
  if(cv.erro) throw new Error('A base nova ainda está lendo as exportações (' + cv.erro + '): não apaguei a antiga. Tente de novo mais tarde.');
  if(apiNuvemshopLigada_()){
    var s = estadoVendasApi_();
    if(!s.ultimo || s.pendente || s.retomar) throw new Error('A carga da API na base nova ainda não terminou: não apaguei a antiga. Tente de novo mais tarde.');
  }
  var it = pastaBase.getFiles(), lixo = 0, rx = /^vendas-(\d{4}|estado)\.json$/;
  while(it.hasNext()){ var f = it.next(); if(rx.test(f.getName())){ f.setTrashed(true); lixo++; } }
  [VENDAS_ANTIGA.api, VENDAS_ANTIGA.prop, VENDAS_ANTIGA.espelho].forEach(function(k){ p.deleteProperty(k); });
  console.log('Base antiga das vendas apagada: ' + lixo + ' arquivo(s) na lixeira e as propriedades ' + VENDAS_ANTIGA.api + ', ' +
    VENDAS_ANTIGA.prop + ' e ' + VENDAS_ANTIGA.espelho + ' removidas.');
}

// base de um ano. Ano sem base ainda = vazio (base nova). Arquivo que existe no estado mas não abre, não é JSON
// ou não tem "pedidos" lança erro: seguir com {} faria o painel gravar uma base vazia por cima da boa.
function lerBaseAno_(estado, ano){
  var id = estado.anos[ano];
  if(!id) return {};
  var nome = VENDAS_ANO_PREFIXO + ano + '.json', dados;
  try{ dados = JSON.parse(DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8')); }
  catch(e){ throw erroBase_(nome, textoErro_(e)); }
  if(!dados || !dados.pedidos || typeof dados.pedidos !== 'object') throw erroBase_(nome, 'conteúdo inesperado');
  return dados.pedidos;
}

// grava a base do ano como arquivo novo (blob: sem o limite de 10 MB do createFile com texto) e
// manda o anterior para a lixeira; ano sem pedidos sai da base
function gravarBaseAno_(estado, pastaBase, ano, pedidos){
  var antigo = estado.anos[ano];
  if(Object.keys(pedidos).length){
    var blob = Utilities.newBlob(JSON.stringify({ano:Number(ano), pedidos:pedidos}), 'application/json', VENDAS_ANO_PREFIXO + ano + '.json');
    estado.anos[ano] = pastaBase.createFile(blob).getId();
  } else delete estado.anos[ano];
  if(antigo){ try{ DriveApp.getFileById(antigo).setTrashed(true); } catch(e){} }
}

// lê os CSVs ainda não consolidados. Devolve {estado} pronto ou {erro:'consolidando', feitos, total}
// quando o tempo acabou (a próxima chamada continua de onde parou).
function consolidarVendas_(pasta){ return comErroBase_('vendas', function(){ return consolidarVendasNucleo_(pasta); }); }

function consolidarVendasNucleo_(pasta){
  var arquivos = arquivosVendas_(pasta);
  var estado = estadoVendas_();
  var atuais = {};
  arquivos.forEach(function(f){ atuais[f.getId()] = f.getLastUpdated().getTime(); });
  var removido = estado.formato !== VENDAS_FORMATO || Object.keys(estado.arqs).some(function(id){ return !(id in atuais); });
  var falta = arquivos.filter(function(f){ return estado.arqs[f.getId()] !== atuais[f.getId()]; });
  if(!removido && !falta.length) return {estado:estado, arquivos:arquivos.length};

  var lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return {erro:'consolidando', feitos:null, total:arquivos.length};
  try{
    estado = estadoVendas_(); // outra execução pode ter avançado enquanto esta esperava
    var inicio = Date.now(), pastaBase = pastaBaseVendas_(pasta), bases = {};
    removido = estado.formato !== VENDAS_FORMATO || Object.keys(estado.arqs).some(function(id){ return !(id in atuais); });
    if(removido){
      // um CSV saiu da pasta (ou a base é de um formato antigo): refaz do zero, para não sobrar pedido de arquivo apagado
      Object.keys(estado.anos).forEach(function(ano){ bases[ano] = {}; });
      estado.arqs = {}; estado.info = {}; estado.avisos = {}; estado.formato = VENDAS_FORMATO;
      Object.keys(bases).forEach(function(ano){ gravarBaseAno_(estado, pastaBase, ano, bases[ano]); });
      gravarEstadoVendas_(estado, pastaBase);
      props_().deleteProperty(VENDAS_API_PROP); // a base zerou: a API recarrega tudo desde o início da loja
    }
    falta = arquivos.filter(function(f){ return estado.arqs[f.getId()] !== atuais[f.getId()]; });
    var feitos = 0, nomes = null;
    for(var i = 0; i < falta.length; i++){
      if(feitos && Date.now() - inicio > VENDAS_TEMPO_MAX) break;
      var f = falta[i], id = f.getId(), ts = atuais[id];
      var lido = lerCsvVendas_(f);
      delete estado.avisos[id];
      if(lido.erro){
        estado.avisos[id] = {nome:f.getName(), erro:lido.erro};
        delete estado.info[id];
      } else {
        // produtos.json é lido ANTES de gravar a base do ano: se a leitura falhar, nada é gravado nesta rodada
        if(!nomes) nomes = lerJsonBase_(pastaBase, PRODUTOS_ARQ) || {};
        var porAno = {}, de = '', ate = '';
        lido.linhas.forEach(function(r){
          var dia = diaIsoVendas_(r[1]);
          if(!dia) return;
          if(!de || dia < de) de = dia;
          if(!ate || dia > ate) ate = dia;
          (porAno[dia.slice(0, 4)] = porAno[dia.slice(0, 4)] || []).push(r);
        });
        // lê a base de TODOS os anos do CSV antes de gravar qualquer uma (leitura falha = nada gravado)
        Object.keys(porAno).forEach(function(ano){ if(!bases[ano]) bases[ano] = lerBaseAno_(estado, ano); });
        Object.keys(porAno).forEach(function(ano){
          juntarPedidos_(bases[ano], porAno[ano], ts);
          gravarBaseAno_(estado, pastaBase, ano, bases[ano]);
        });
        estado.info[id] = {nome:f.getName(), modificadoEm:new Date(ts).toISOString(), pedidos:lido.linhas.length, de:de, ate:ate};
        Object.keys(lido.nomes).forEach(function(sku){ nomes[sku] = lido.nomes[sku]; });
        gravarJsonBase_(pastaBase, PRODUTOS_ARQ, nomes);
      }
      estado.arqs[id] = ts;
      // progresso salvo arquivo a arquivo: se a execução cair no meio, nada do que já foi lido se perde
      gravarEstadoVendas_(estado, pastaBase);
      feitos++;
    }
    var restam = falta.length - feitos;
    if(restam > 0) return {erro:'consolidando', feitos:arquivos.length - restam, total:arquivos.length};
    return {estado:estado, arquivos:arquivos.length};
  } finally { lock.releaseLock(); }
}

// pedidos da base → formato da resposta (sem a coluna ts, textos repetidos num dicionário). Os SKUs do pedido vão
// na última posição como [índice em skus, qtd, preço unitário (ou null), índice, qtd, preço, ...]; skus = [[sku, índice do modelo, cor, tamanho]] e
// modelos = nomes dos modelos (partesProduto_ sobre o nome do produto, o mesmo corte das Devoluções).
function codificarVendas_(pedidos, nomesProduto){
  var dic = {}, idx = {};
  var skus = [], idxSku = Object.create(null), modelos = [], idxModelo = Object.create(null);
  function skuIdx(sku){
    var i = idxSku[sku];
    if(i === undefined){
      var p = partesProduto_(nomesProduto[sku], sku), m = idxModelo[p.modelo];
      if(m === undefined){ m = idxModelo[p.modelo] = modelos.length; modelos.push(p.modelo); }
      i = idxSku[sku] = skus.length; skus.push([sku, m, p.cor, p.tam]);
    }
    return i;
  }
  VENDAS_CODIFICADAS.forEach(function(c){ dic[c] = []; idx[c] = Object.create(null); });
  var linhas = Object.keys(pedidos).map(function(k){
    var x = pedidos[k].slice(0, VENDAS_SKUS), pecas = [];
    itensSku_(pedidos[k][VENDAS_SKUS]).forEach(function(it){ pecas.push(skuIdx(it.sku), it.q, it.preco); });
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
    x.push(pecas);
    return x;
  });
  return {dic:dic, linhas:linhas, skus:skus, modelos:modelos};
}

// ano: o que a página quer ver (padrão: o mais recente). versao: a que a página já tem; se nenhum CSV
// mudou, responde {mudou:false} sem reenviar os pedidos.
function obterVendas(token, versao, ano){ return comErroBase_('vendas', function(){ return obterVendas_(token, versao, ano); }); }

function obterVendas_(token, versao, ano){
  var inicioChamada = Date.now();
  if(!contaComToken_(token)) return {erro:'senha_necessaria'};
  var pasta = pastaVendas_();
  if(!pasta) return {erro:'pasta_vendas'};
  var c = consolidarVendas_(pasta);
  if(c.erro) return c;
  var api = null;
  if(apiNuvemshopLigada_()){
    var s = estadoVendasApi_();
    if(!s.ultimo || s.pendente || Date.now() - (s.verificadoEm || 0) > VENDAS_API_INTERVALO) s = sincronizarVendasApi_(inicioChamada);
    if(s.erro === 'consolidando') return {erro:'consolidando', feitos:null, total:null, api:true};
    api = {ultimo:s.ultimo || null, verificadoEm:s.verificadoEm || null, erro:s.erro || null};
  }
  if(!c.arquivos && !api) return {erro:'sem_csv'};
  var estado = estadoVendas_();
  var anos = Object.keys(estado.anos).map(Number).sort(function(a, b){ return a - b; });
  var avisos = Object.keys(estado.avisos).map(function(id){ return estado.avisos[id]; });
  if(!anos.length) return {erro:'csv_formato', avisos:avisos};
  ano = Number(ano);
  if(anos.indexOf(ano) < 0) ano = anos[anos.length - 1];
  // npVersao: extratos do Nuvem Pago completam Taxas/Juros/Líquido das vendas que vieram só da API
  var npVersao = props_().getProperty('NUVEMPAGO_VERSAO') || '';
  var assinatura = Object.keys(estado.arqs).sort().map(function(id){ return id + '@' + estado.arqs[id]; }).join('|') + '|' + ano +
    (api ? '|api@' + (estado.apiMudou || '') : '') + (npVersao ? '|np@' + npVersao : '') +
    '|fotos@' + (props_().getProperty('FOTOS_VERSAO') || '');
  var novaVersao = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, assinatura));
  if(versao && versao === novaVersao) return {mudou:false, versao:novaVersao, ano:ano, api:api};
  var pedidos = lerBaseAno_(estado, ano);
  if(npVersao){
    try{ completarComNuvemPago_(pedidos, financeiroNuvemPago_(lerBaseNuvemPago_(pastaBaseVendas_(pasta)))); }
    catch(e){ console.warn('Nuvem Pago nas vendas: ' + (e && e.message || e)); }
  }
  var pastaBaseV = pastaBaseVendas_(pasta);
  var cod = codificarVendas_(pedidos, lerJsonBase_(pastaBaseV, PRODUTOS_ARQ) || {});
  // fotos só dos SKUs vendidos no ano (endereço da imagem); falha das fotos nunca derruba as vendas
  var fotos = {};
  try{ var todasFotos = lerFotos_(pastaBaseV); cod.skus.forEach(function(x){ if(todasFotos[x[0]]) fotos[x[0]] = todasFotos[x[0]]; }); }
  catch(e){ fotos = {}; }
  var arquivos = Object.keys(estado.info).map(function(id){ return estado.info[id]; })
    .sort(function(a, b){ return a.modificadoEm < b.modificadoEm ? -1 : 1; });
  return {mudou:true, versao:novaVersao, ano:ano, anos:anos, colunas:VENDAS_COLUNAS.concat(['Itens', 'Peças']),
    dic:cod.dic, linhas:cod.linhas, skus:cod.skus, modelos:cod.modelos, fotos:fotos, arquivos:arquivos, avisos:avisos, api:api};
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
  'payment_details','coupon','paid_at','completed_at','storefront','cancelled_at','cancel_reason','products',
  'contact_name','contact_email','customer'].join(',');
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

function apiNuvemshopConfigurada_(){
  var p = props_();
  return !!(p.getProperty('NUVEMSHOP_TOKEN') && p.getProperty('NUVEMSHOP_LOJA'));
}

// a base só passa a receber a API depois de instalarSincronizacaoVendas (antes dá para rodar a amostra
// e comparar com o CSV sem misturar nada)
function apiNuvemshopLigada_(){ return apiNuvemshopConfigurada_() && props_().getProperty('VENDAS_API_LIGADA') === '1'; }

function estadoVendasApi_(){
  var e = null;
  try{ e = JSON.parse(props_().getProperty(VENDAS_API_PROP) || 'null'); } catch(err){}
  return e || {};
}

// GET na API. Devolve {lista} (array do corpo), {fim:true} na página depois da última (a API responde 404)
// ou {erro}. Espera e tenta de novo quando bate no limite de 2 chamadas por segundo (429).
function chamarApiNuvemshop_(caminho, params){
  var p = props_();
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
function dataApiParaCsv_(s){
  if(!s) return '';
  var d = new Date(String(s).replace(/([+-]\d{2})(\d{2})$/, '$1:$2'));
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
}

function tsApi_(s){
  var d = new Date(String(s || '').replace(/([+-]\d{2})(\d{2})$/, '$1:$2'));
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function textoApi_(v){ return v == null ? '' : String(v).trim(); }

// "Data de pagamento" do CSV (só o dia). paid_at vem sempre null do Nuvem Pago; completed_at chega como
// {date:"2026-10-01 02:57:27.000000", timezone:"UTC"} e o dia dele bate com o CSV
function dataPagamentoApi_(o){
  if(!VENDAS_API_PAGOS[o.payment_status]) return '';
  var s = o.paid_at;
  if(!s && o.completed_at) s = typeof o.completed_at === 'object' ? String(o.completed_at.date || '').replace(' ', 'T').replace(/\.\d+$/, '') + '+0000' : o.completed_at;
  return dataApiParaCsv_(s).slice(0, 10);
}

// pedido da API → linha da base (mesmo layout de lerCsvVendas_: VENDAS_COLUNAS + itens + SKUs).
// Do cliente: o nome (contact_name), a cidade de entrega e a chave do e-mail (chaveCliente_). Nada de e-mail,
// CPF, telefone ou endereço: do endereço só saem a UF e a cidade.
function pedidoApiParaLinha_(o, nomes, fotos){
  var pd = o.payment_details || {};
  var cupom = Array.isArray(o.coupon) ? o.coupon.map(function(c){ return c && c.code; }).filter(String).join(', ') : textoApi_(o.coupon);
  var envio = o.shipping_option;
  if(envio && typeof envio === 'object') envio = envio.name || envio.code || '';
  var col = {
    'Número do Pedido':textoApi_(o.number),
    'Data':dataApiParaCsv_(o.created_at),
    'Status do Pedido':VENDAS_API_STATUS[o.status] || textoApi_(o.status),
    'Status do Pagamento':VENDAS_API_PAGTO[o.payment_status] || textoApi_(o.payment_status),
    'Status do Envio':VENDAS_API_ENVIO[o.shipping_status] || textoApi_(o.shipping_status),
    'Subtotal':textoApi_(o.subtotal),
    'Desconto':textoApi_(o.discount),
    'Valor do Frete':textoApi_(o.shipping_cost_customer),
    'Total':textoApi_(o.total),
    'Estado':textoApi_((o.shipping_address || {}).province),
    'Forma de Entrega':textoApi_(envio),
    'Forma de Pagamento':textoApi_(o.gateway_name || o.gateway),
    'Cupom de Desconto':cupom,
    'Data de pagamento':dataPagamentoApi_(o),
    'Canal':VENDAS_API_CANAL[o.storefront] || textoApi_(o.storefront),
    'Data e hora do cancelamento':dataApiParaCsv_(o.cancelled_at),
    'Motivo do cancelamento':VENDAS_API_MOTIVO[o.cancel_reason] || textoApi_(o.cancel_reason),
    'Parcelas':pd.method === 'credit_card' && pd.installments ? String(pd.installments) : '', // no CSV, Pix fica vazio
    'Meio de pagamento':VENDAS_API_MEIO[pd.method] || textoApi_(pd.method),
    'Cidade':textoApi_((o.shipping_address || {}).city),
    'Nome do comprador':textoApi_(o.contact_name || (o.customer || {}).name),
    'Cliente':chaveCliente_(o.contact_email || (o.customer || {}).email)
  };
  var r = VENDAS_COLUNAS.map(function(c){ return col[c] || ''; });
  var qtd = 0, itens = [];
  (o.products || []).forEach(function(it){
    var q = Number(it.quantity) || 0, sku = textoApi_(it.sku).replace(/[,:]/g, '');
    qtd += q;
    if(sku){
      itens.push(itemSkuTexto_(sku, q, it.price)); if(it.name) nomes[sku] = textoApi_(it.name);
      // foto do item (só o endereço da imagem, nunca dado do cliente); vazio não sobrescreve
      if(fotos){ var u = urlFotoValida_(it.image && it.image.src); if(u) fotos[sku] = u; }
    }
  });
  r.push(qtd, itens.join(','));
  return r;
}

// busca os pedidos alterados desde a última vez e junta na base. Devolve o VENDAS_API atualizado;
// {erro:'consolidando'} se outra execução estiver com a base ou se o tempo acabou no meio da carga.
// relogio = início da execução (o limite de 6 min do Apps Script conta desde ali).
function sincronizarVendasApi_(relogio){ return comErroBase_('vendas', function(){ return sincronizarVendasApiNucleo_(relogio); }); }

function sincronizarVendasApiNucleo_(relogio){
  if(!apiNuvemshopConfigurada_()) return {erro:'api_sem_token'};
  var pasta = pastaVendas_();
  if(!pasta) return {erro:'pasta_vendas'};
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return {erro:'consolidando'};
  try{
    var s = estadoVendasApi_(), estado = estadoVendas_(), pastaBase = pastaBaseVendas_(pasta);
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
      if(!entregas) entregas = lerJsonBase_(pastaBase, ENTREGAS_ARQ) || {};
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
      var r = chamarApiNuvemshop_('orders', q);
      if(r.erro){ erro = r.erro; break; }
      var lista = r.fim ? [] : r.lista;
      var porAno = {};
      lista.forEach(function(o){
        var linha = pedidoApiParaLinha_(o, nomes, fotosNovas), dia = diaIsoVendas_(linha[1]);
        if(!dia || !linha[0]) return;
        (porAno[dia.slice(0, 4)] = porAno[dia.slice(0, 4)] || []).push({linha:linha, ts:tsApi_(o.updated_at), entregue:o.shipping_status === 'delivered'});
      });
      Object.keys(porAno).forEach(function(ano){
        if(!bases[ano]) bases[ano] = lerBaseAno_(estado, ano);
        porAno[ano].forEach(function(x){
          // antes de juntar: a linha anterior diz se o pedido ainda não estava entregue (limite inferior da entrega)
          if(x.entregue) marcarEntrega(bases[ano][x.linha[0]], x.linha[0], x.ts);
          juntarPedidos_(bases[ano], [x.linha], x.ts);
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
    // produtos.json é lido antes de qualquer gravação: se a leitura falhar, nada foi gravado e o estado não muda
    var todos = mudou ? (lerJsonBase_(pastaBase, PRODUTOS_ARQ) || {}) : null;
    // grava uma vez por execução (cada gravação cria um arquivo novo da base do ano)
    Object.keys(bases).forEach(function(ano){ gravarBaseAno_(estado, pastaBase, ano, bases[ano]); });
    if(entregasMudou) gravarJsonBase_(pastaBase, ENTREGAS_ARQ, entregas);
    if(mudou){
      estado.apiMudou = Date.now();
      Object.keys(nomes).forEach(function(sku){ todos[sku] = nomes[sku]; });
      gravarJsonBase_(pastaBase, PRODUTOS_ARQ, todos);
    }
    // fotos: erro aqui nunca derruba as vendas
    try{ if(Object.keys(fotosNovas).length) juntarFotos_(pastaBase, fotosNovas); }
    catch(e){ if(e && e.baseIlegivel) registrarErroBase_('fotos', e); console.warn('Fotos (captura nos pedidos): ' + (e && e.message || e)); }
    gravarEstadoVendas_(estado, pastaBase);
    s.verificadoEm = Date.now();
    // com erro, guarda em "retomar" a janela e a página em que parou: a próxima verificação (gatilho ou
    // painel, depois do intervalo) tenta de novo dali, sem o painel insistir a cada abertura
    delete s.pendente; delete s.retomar; delete s.erro;
    if(erro){ s.erro = erro; s.retomar = pend; }
    else if(pend) s.pendente = pend;
    else s.ultimo = cicloIni;
    props_().setProperty(VENDAS_API_PROP, JSON.stringify(s));
    return s.pendente && !erro ? {erro:'consolidando'} : s;
  } finally { lock.releaseLock(); }
}

// gatilho de 15 min: mantém a base em dia sem ninguém abrir o painel
function sincronizarVendasGatilho(){
  exigirDono_();               // gatilho (dono) e editor passam; visitante do web app não
  if(!gatilhoLiberado_()) return; // código de teste (HEAD) diferente do publicado: não mexe nas bases reais
  var inicio = Date.now();
  if(apiNuvemshopLigada_()){
    var s = sincronizarVendasApi_(inicio);
    if(s.erro && s.erro !== 'consolidando') console.warn('Vendas pela API: ' + s.erro + (s.detalhe ? ' (' + s.detalhe + ')' : ''));
  }
  // extratos novos do Nuvem Pago entram na base mesmo sem ninguém abrir a página de Recebimentos
  // (só com folga de tempo: a API vem primeiro e o Apps Script para em 6 min)
  if(Date.now() - inicio > 90 * 1000) return;
  try{
    var np = consolidarNuvemPago_(inicio);
    if(np.erro && np.erro !== 'consolidando' && np.erro !== 'pasta_nuvempago') console.warn('Nuvem Pago: ' + np.erro);
  } catch(e){ console.warn('Nuvem Pago: ' + (e && e.message || e)); }
  // fotos das peças: só enquanto a carga não terminou (ou na renovação semanal) e se ainda sobrou tempo
  try{
    if(apiNuvemshopLigada_() && fotosPrecisaRodar_() && Date.now() - inicio < FOTOS_INICIO_MAX){
      var ft = atualizarFotosProdutos(inicio);
      if(ft.erro && ft.erro !== 'consolidando') console.warn('Fotos: ' + ft.erro);
      else if(!ft.erro) limparErroBase_('fotos');
    }
  } catch(e){ if(e && e.baseIlegivel) registrarErroBase_('fotos', e); console.warn('Fotos: ' + (e && e.message || e)); }
}

// rodar uma vez no editor (conta dona do script), depois de preencher NUVEMSHOP_TOKEN e NUVEMSHOP_LOJA
function instalarSincronizacaoVendas(){
  exigirDonoEditor_();
  ScriptApp.getProjectTriggers().forEach(function(g){
    if(g.getHandlerFunction() === 'sincronizarVendasGatilho') ScriptApp.deleteTrigger(g);
  });
  if(!apiNuvemshopConfigurada_()) throw new Error('Preencha NUVEMSHOP_TOKEN e NUVEMSHOP_LOJA nas Propriedades do script.');
  props_().setProperty('VENDAS_API_LIGADA', '1');
  ScriptApp.newTrigger('sincronizarVendasGatilho').timeBased().everyMinutes(15).create();
  sincronizarVendasGatilho();
}

// rodar uma vez no editor: grava nuvemshop/_base/api-amostra.json com alguns pedidos de setembro/2026 como
// vêm da API (sem dados pessoais), lado a lado com a linha da base (CSV) e a tradução feita aqui — para
// conferir o mapeamento antes de confiar nos números.
function amostraApiNuvemshop(){
  exigirDonoEditor_();
  var r = chamarApiNuvemshop_('orders', {status:'any', created_at_min:'2026-09-01T00:00:00-03:00',
    created_at_max:'2026-09-30T23:59:59-03:00', per_page:100});
  if(r.erro || r.fim){ console.log('Erro: ' + (r.erro || 'sem pedidos')); return; }
  var vistos = {}, escolhidos = [];
  r.lista.forEach(function(o){
    var chave = [o.status, o.payment_status, o.shipping_status, (o.payment_details || {}).method, o.storefront].join('|');
    if(!vistos[chave] && escolhidos.length < 15){ vistos[chave] = true; escolhidos.push(o); }
  });
  var estado = estadoVendas_(), base = lerBaseAno_(estado, '2026'), nomes = {};
  var limpos = escolhidos.map(function(o){
    var x = {};
    Object.keys(o).forEach(function(k){
      if(/^(customer|contact_|billing_|client_details|note|owner_note|fulfillments|shipping_address|token|landing_url|checkout)/.test(k)) return;
      x[k] = o[k];
    });
    x.shipping_address = {province:(o.shipping_address || {}).province};
    var linha = pedidoApiParaLinha_(o, nomes), traduzido = {}, csv = {};
    // a amostra é para conferir o mapeamento: nome e cidade do cliente ficam de fora (só "preenchido" ou vazio)
    var semCliente = function(c, v){ return c === 'Nome do comprador' || c === 'Cidade' ? (v ? '(preenchido)' : '') : v; };
    VENDAS_COLUNAS.forEach(function(c, i){ traduzido[c] = semCliente(c, linha[i]); });
    var b = base[String(o.number)];
    if(b) VENDAS_COLUNAS.forEach(function(c, i){ csv[c] = semCliente(c, b[i]); });
    return {api:x, traduzido:traduzido, base_csv:b ? csv : null};
  });
  var pastaBase = pastaBaseVendas_(pastaVendas_());
  gravarJsonBase_(pastaBase, 'api-amostra.json', {geradoEm:new Date().toISOString(), totalNoMes:r.lista.length, pedidos:limpos});
  console.log('Amostra gravada: ' + limpos.length + ' pedidos (de ' + r.lista.length + ' lidos).');
}

// rodar no editor quando a sincronização avisar api_5xx: testa, uma a uma, as diferenças entre a busca da
// amostra (que funciona) e a da sincronização, e mostra no registro qual delas a Nuvemshop recusa
function diagnosticoApiNuvemshop(){
  exigirDonoEditor_();
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
    var ini = Date.now(), r = chamarApiNuvemshop_('orders', t[1]);
    console.log(t[0] + ': ' + (r.erro ? 'ERRO ' + r.erro : r.fim ? 'sem pedidos' : 'ok, ' + r.lista.length + ' pedidos') + ' em ' + (Date.now() - ini) + ' ms');
    Utilities.sleep(600);
  });
}

/* ===================== Fotos das peças (fotos.json) ===================== */

// fotos.json (na pasta _base) = {SKU: "https://…imagem"}. Guarda só o ENDEREÇO da imagem hospedada na Nuvemshop;
// nenhum arquivo é copiado e nenhum dado de cliente entra. Duas fontes:
//  1) captura incremental: cada pedido lido pela API (item.image.src), em pedidoApiParaLinha_/sincronizarVendasApi_;
//  2) carga do catálogo: atualizarFotosProdutos (GET products, paginado), retomável e refeita 1x por semana.
// Propriedades do script: FOTOS_PAGINA (próxima página a ler), FOTOS_COMPLETO (AAAA-MM-DD da última carga completa),
// FOTOS_VERSAO (muda quando o conteúdo de fotos.json muda; entra na versão de obterDevolucoes), FOTOS_ERRO.
var FOTOS_ARQ = 'fotos.json';
var FOTOS_TEMPO_MAX = 3 * 60 * 1000;  // orçamento da carga, contado desde o início da execução
var FOTOS_INICIO_MAX = 120 * 1000;    // no gatilho, só começa se a execução ainda tem menos que isso
var FOTOS_RENOVAR_DIAS = 7;           // refaz a carga completa a cada semana (produtos novos/alterados)

function urlFotoValida_(u){
  u = textoApi_(u);
  return /^https:\/\/[^\s"'<>]+$/i.test(u) ? u : '';
}

function lerFotos_(pastaBase){ return lerJsonBase_(pastaBase, FOTOS_ARQ) || {}; }

function gravarFotos_(pastaBase, fotos){
  gravarJsonBase_(pastaBase, FOTOS_ARQ, fotos);
  props_().setProperty('FOTOS_VERSAO', String(Date.now()));
}

// junta as novas no fotos.json; só regrava se algo mudou. Nunca troca uma URL por vazio. Devolve quantas mudaram.
function juntarFotos_(pastaBase, novas){
  var todas = lerFotos_(pastaBase), mudou = 0;
  Object.keys(novas).forEach(function(sku){
    var u = urlFotoValida_(novas[sku]);
    if(u && todas[sku] !== u){ todas[sku] = u; mudou++; }
  });
  if(mudou) gravarFotos_(pastaBase, todas);
  return mudou;
}

function fotosPrecisaRodar_(){
  var p = props_(), ok = p.getProperty('FOTOS_COMPLETO');
  if(p.getProperty('FOTOS_PAGINA') || !ok) return true; // carga em andamento ou nunca feita
  var t = Date.parse(ok + 'T12:00:00Z');
  return isNaN(t) || Date.now() - t > FOTOS_RENOVAR_DIAS * 24 * 3600 * 1000;
}

// produto da API → {sku: url}. Imagem da variante (variant.image_id → images[].id); sem ela, a de menor position.
function fotosDoProduto_(prod){
  var imgs = (prod.images || []).filter(function(i){ return i && urlFotoValida_(i.src); });
  if(!imgs.length) return {};
  var porId = {};
  imgs.forEach(function(i){ porId[String(i.id)] = i; });
  var primeira = imgs.slice().sort(function(a, b){ return (Number(a.position) || 0) - (Number(b.position) || 0); })[0];
  var out = {};
  (prod.variants || []).forEach(function(v){
    var sku = textoApi_(v && v.sku).replace(/[,:]/g, '');
    if(!sku) return;
    var img = (v.image_id != null && porId[String(v.image_id)]) || primeira;
    out[sku] = urlFotoValida_(img.src);
  });
  return out;
}

// carga histórica do catálogo. Retoma da página guardada em FOTOS_PAGINA; ao terminar, grava FOTOS_COMPLETO.
// Pode ser rodada à mão no editor (o gatilho de 15 min também a chama enquanto não terminar).
// relogio = início da execução (o limite do Apps Script conta desde ali).
function atualizarFotosProdutos(relogio){
  exigirDono_();
  if(!apiNuvemshopConfigurada_()) return {erro:'api_sem_token'};
  var pasta = pastaVendas_();
  if(!pasta) return {erro:'pasta_vendas'};
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return {erro:'consolidando'};
  try{
    var p = props_(), pastaBase = pastaBaseVendas_(pasta), ini = typeof relogio === 'number' ? relogio : Date.now();
    var pagina = Number(p.getProperty('FOTOS_PAGINA')) || 1;
    var acumuladas = {}, erro = null, terminou = false, lidos = 0;
    while(true){
      if(Date.now() - ini > FOTOS_TEMPO_MAX) break;
      var r = chamarApiNuvemshop_('products', {per_page:200, page:pagina, fields:'id,images,variants'});
      if(r.erro){ erro = r.erro; break; }
      var lista = r.fim ? [] : r.lista;
      lista.forEach(function(prod){
        var f = fotosDoProduto_(prod);
        Object.keys(f).forEach(function(sku){ acumuladas[sku] = f[sku]; });
      });
      lidos += lista.length;
      if(lista.length < 200){ terminou = true; break; }
      pagina++;
      Utilities.sleep(500); // a API aceita ~2 chamadas por segundo
    }
    // grava o que veio, mesmo se parou no meio; a página só avança depois de gravar
    if(Object.keys(acumuladas).length) juntarFotos_(pastaBase, acumuladas);
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
var ENTREGAS_ARQ = 'entregas.json'; // nº do pedido → [ms antes da entrega (0 = desconhecido), ms da 1ª vez visto "delivered"] (sincronizarVendasApi_)

// o mais novo com esse nome (gravar cria um arquivo novo e manda o anterior para a lixeira)
function arquivoBase_(pastaBase, nome){
  var it = pastaBase.getFilesByName(nome), achado = null;
  while(it.hasNext()){ var f = it.next(); if(!achado || f.getDateCreated() > achado.getDateCreated()) achado = f; }
  return achado;
}

// null SÓ quando o arquivo não existe (base nova). Arquivo que existe mas não abre ou não é JSON de objeto lança
// erroBase_: quem chama não deve gravar nada depois disso (antes devolvia null e a base era regravada vazia).
function lerJsonBase_(pastaBase, nome){
  var texto;
  try{
    var f = arquivoBase_(pastaBase, nome);
    if(!f) return null;
    texto = f.getBlob().getDataAsString('UTF-8');
  } catch(e){ throw erroBase_(nome, textoErro_(e)); }
  var obj;
  try{ obj = JSON.parse(texto); } catch(e){ throw erroBase_(nome, 'conteúdo inválido'); }
  if(!obj || typeof obj !== 'object') throw erroBase_(nome, 'conteúdo inesperado');
  return obj;
}

function gravarJsonBase_(pastaBase, nome, obj){
  var antigo = arquivoBase_(pastaBase, nome);
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

function pastaDevolucoes_(pastaV){
  var it = pastaV.getFoldersByName(DEVOL_PASTA_NOME);
  return it.hasNext() ? it.next() : null;
}

function arquivosDevolucoes_(pasta){
  var lista = [], it = pasta.getFiles();
  while(it.hasNext()){
    var f = it.next(), nome = f.getName(), tipo = f.getMimeType();
    if(nome.indexOf(DEVOL_TMP) === 0) continue;
    if(ehArquivoOp_(nome)) continue; // NFs e pedidos enviados pela extensão (JSON, etapa 28) não são exports de devolução
    if(/\.(csv|xlsx)$/i.test(nome) || tipo === MIME_XLSX || tipo === MIME_SHEETS || /csv/i.test(tipo)) lista.push(f);
  }
  lista.sort(function(a, b){ return a.getLastUpdated() - b.getLastUpdated(); });
  return lista;
}

// planilha Google (ou .xlsx convertido numa cópia temporária pela API do Drive) → linhas do CSV exportado
function exportarCsvPlanilha_(id){
  var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + id + '/export?format=csv', {
    headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions:true
  });
  if(resp.getResponseCode() !== 200) return {erro:'exportacao_falhou', detalhe:'HTTP ' + resp.getResponseCode()};
  return Utilities.parseCsv(resp.getContentText('UTF-8'), ',');
}

function linhasDevolucoes_(f){
  var tipo = f.getMimeType();
  if(tipo === MIME_SHEETS) return exportarCsvPlanilha_(f.getId());
  if(tipo === MIME_XLSX || /\.xlsx$/i.test(f.getName())){
    var resp = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + f.getId() + '/copy?supportsAllDrives=true&fields=id', {
      method:'post', contentType:'application/json', muteHttpExceptions:true,
      payload:JSON.stringify({name:DEVOL_TMP, mimeType:MIME_SHEETS}),
      headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()}
    });
    if(resp.getResponseCode() !== 200) return {erro:'conversao_falhou', detalhe:'HTTP ' + resp.getResponseCode()};
    var tmp = JSON.parse(resp.getContentText()).id;
    try{ return exportarCsvPlanilha_(tmp); }
    finally { try{ DriveApp.getFileById(tmp).setTrashed(true); } catch(e){} }
  }
  var blob = f.getBlob(), texto = blob.getDataAsString('UTF-8').replace(/^﻿/, '');
  // CSV salvo em ISO-8859-1 (Excel no Windows): em UTF-8 os acentos viram o caractere inválido
  if(texto.indexOf('Pedido de Origem') < 0 || texto.indexOf('\uFFFD') >= 0) texto = blob.getDataAsString('ISO-8859-1');
  var primeira = texto.slice(0, texto.indexOf('\n'));
  return Utilities.parseCsv(texto, primeira.split(';').length > primeira.split(',').length ? ';' : ',');
}

function textoDevol_(v){ v = String(v == null ? '' : v).trim(); return v === 'undefined' || v === 'null' ? '' : v; }

// "R$ 1.234,56" → 1234.56
function numeroBR_(v){
  var s = textoDevol_(v).replace(/[^\d,.\-]/g, '');
  if(s.indexOf(',') >= 0) s = s.replace(/\./g, '').replace(',', '.');
  var n = Number(s);
  return isFinite(n) ? n : 0;
}

// "01/10/2026 - 11:40" ou "2026-10-05T11:12:08" → "2026-10-01 11:40" ('' se não for data)
function dataHoraDevol_(v){
  var s = textoDevol_(v), m = /^(\d{2})\/(\d{2})\/(\d{4})(?:\D+(\d{2}):(\d{2}))?/.exec(s);
  if(m) return m[3] + '-' + m[2] + '-' + m[1] + (m[4] ? ' ' + m[4] + ':' + m[5] : '');
  m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?/.exec(s);
  return m ? m[1] + '-' + m[2] + '-' + m[3] + (m[4] ? ' ' + m[4] + ':' + m[5] : '') : '';
}

// arquivo → uma linha por solicitação (DEVOL_COLUNAS + itens)
function lerDevolucoes_(f){
  var linhas = linhasDevolucoes_(f);
  if(linhas.erro) return linhas;
  if(!linhas.length) return {erro:'csv_vazio'};
  var cab = linhas[0].map(function(c){ return String(c).replace(/^﻿/, '').trim(); });
  var ix = DEVOL_COLUNAS.map(function(c){ return cab.indexOf(c); });
  var it = DEVOL_ITEM.map(function(c){ return cab.indexOf(c); });
  if(ix[0] < 0 || ix[1] < 0 || ix[7] < 0) return {erro:'formato_devolucoes'};
  var sols = {}, ordem = [];
  for(var i = 1; i < linhas.length; i++){
    var l = linhas[i], id = textoDevol_(l[ix[0]]);
    if(!id) continue;
    var s = sols[id];
    if(!s){
      s = ix.map(function(j){ return j >= 0 ? textoDevol_(l[j]) : ''; });
      s[3] = dataHoraDevol_(s[3]); s[12] = dataHoraDevol_(s[12]);
      s[4] = numeroBR_(s[4]); s[5] = numeroBR_(s[5]); s[6] = numeroBR_(s[6]);
      s.push([]);
      sols[id] = s; ordem.push(id);
    }
    var sku = it[0] >= 0 ? textoDevol_(l[it[0]]) : '';
    if(sku) s[DEVOL_ITENS].push([sku, it[1] >= 0 ? numeroBR_(l[it[1]]) : 0, it[2] >= 0 ? Number(textoDevol_(l[it[2]])) || 1 : 1]);
  }
  return {sols:ordem.map(function(id){ return sols[id]; })};
}

// lê os arquivos ainda não consolidados; mesma lógica de tempo e de lock das vendas
function consolidarDevolucoes_(pastaV, pasta){ return comErroBase_('devolucoes', function(){ return consolidarDevolucoesNucleo_(pastaV, pasta); }); }

function consolidarDevolucoesNucleo_(pastaV, pasta){
  var arquivos = arquivosDevolucoes_(pasta);
  var pastaBase = pastaBaseVendas_(pastaV);
  var estado = lerJsonBase_(pastaBase, DEVOL_ESTADO_ARQ) || {};
  estado.arqs = estado.arqs || {}; estado.info = estado.info || {}; estado.avisos = estado.avisos || {};
  var atuais = {};
  arquivos.forEach(function(f){ atuais[f.getId()] = f.getLastUpdated().getTime(); });
  var removido = Object.keys(estado.arqs).some(function(id){ return !(id in atuais); });
  var falta = arquivos.filter(function(f){ return estado.arqs[f.getId()] !== atuais[f.getId()]; });
  if(!removido && !falta.length) return {estado:estado, pastaBase:pastaBase, arquivos:arquivos.length};

  var lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return {erro:'consolidando', feitos:null, total:arquivos.length};
  try{
    estado = lerJsonBase_(pastaBase, DEVOL_ESTADO_ARQ) || estado;
    estado.arqs = estado.arqs || {}; estado.info = estado.info || {}; estado.avisos = estado.avisos || {};
    var baseLida = lerJsonBase_(pastaBase, DEVOL_BASE_ARQ);
    // base ausente com arquivos já consolidados no estado = anomalia (o histórico "hist" não se refaz): não segue
    if(!baseLida && Object.keys(estado.arqs).length) throw erroBase_(DEVOL_BASE_ARQ, 'não encontrado, mas o estado indica arquivos já lidos');
    var base = baseLida || {};
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
      var lido = lerDevolucoes_(f);
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
      gravarJsonBase_(pastaBase, DEVOL_BASE_ARQ, base);
      gravarJsonBase_(pastaBase, DEVOL_ESTADO_ARQ, estado);
      feitos++;
    }
    if(removido && !feitos){
      gravarJsonBase_(pastaBase, DEVOL_BASE_ARQ, base);
      gravarJsonBase_(pastaBase, DEVOL_ESTADO_ARQ, estado);
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

function arquivoNfMaisRecente_(){
  var pasta;
  try{ pasta = DriveApp.getFolderById(props_().getProperty('NF_PASTA') || NF_PASTA_ID); } catch(e){ return null; }
  var it = pasta.getFilesByType(MIME_SHEETS), achado = null;
  while(it.hasNext()){
    var f = it.next();
    if(f.getName().indexOf(NF_NOME) >= 0 && (!achado || f.getLastUpdated() > achado.getLastUpdated())) achado = f;
  }
  return achado;
}

// "12/10/2026" ou "15/09" (sem ano: o do card; o anterior se cair mais de 6 meses depois da criação) → "2026-10-12"
function diaNf_(d, m, a, criado){
  if(!a){
    a = criado ? Number(criado.slice(0, 4)) : new Date().getFullYear();
    if(criado && Number(m) - Number(criado.slice(5, 7)) > 6) a--;
  } else if(a.length === 2) a = '20' + a;
  var iso = a + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
  return /^\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/.test(iso) ? iso : '';
}

// texto do card → {num, data} ou null
function nfDoTexto_(t, criado){
  var i = t.search(/NF\s+emitida/i);
  if(i < 0) return null;
  var s = t.slice(i, i + 160);
  var n = /N\s*[º°o.]\s*(\d+)/i.exec(s), d = /(\d{1,2})\/(\d{1,2})(?:\/(\d{4}|\d{2}))?/.exec(s);
  var data = d ? diaNf_(('0' + d[1]).slice(-2), ('0' + d[2]).slice(-2), d[3], criado) : '';
  return n || data ? {num:n ? n[1] : '', data:data} : null;
}

// planilha → cards [codigo, fase, pedido, acao, nfNum, nfData, criado, processado]
function lerNotasFinanceiro_(f){
  var linhas = exportarCsvPlanilha_(f.getId());
  if(linhas.erro) return linhas;
  if(!linhas.length) return {erro:'csv_vazio'};
  var cab = linhas[0].map(function(c){ return String(c).replace(/^﻿/, '').trim(); });
  var ix = NF_COLUNAS.map(function(c){ return cab.indexOf(c); });
  if(ix[0] < 0 || ix[4] < 0) return {erro:'formato_nf'};
  var cards = [];
  for(var i = 1; i < linhas.length; i++){
    var l = linhas[i], v = ix.map(function(j){ return j >= 0 ? textoDevol_(l[j]) : ''; });
    var pedido = v[4].replace(/\s+/g, '').toUpperCase();
    if(!v[0] || !pedido) continue;
    var criado = dataHoraDevol_(v[1]);
    var nf = nfDoTexto_(v[3] + '\n' + v[6], criado) || {num:'', data:''};
    if(!nf.data) nf.data = dataHoraDevol_(v[8]).slice(0, 10);
    cards.push([v[0], v[2], pedido, v[5], nf.num, nf.data, criado, dataHoraDevol_(v[7]).slice(0, 10)]);
  }
  return {cards:cards};
}

// lê a planilha mais nova só quando ela muda; o resultado fica em _base/nf-financeiro.json
function notasFinanceiro_(pastaBase){
  var f = arquivoNfMaisRecente_();
  if(!f) return {erro:'sem_nf'};
  var ts = f.getLastUpdated().getTime(), salvo = null;
  try{ salvo = lerJsonBase_(pastaBase, NF_ARQ); } catch(e){ if(!e || !e.baseIlegivel) throw e; } // é só cache: refaz da planilha
  if(salvo && salvo.arquivo === f.getId() && salvo.modificado === ts) return salvo;
  var lido = lerNotasFinanceiro_(f);
  if(lido.erro) return {erro:lido.erro, nome:f.getName()};
  var novo = {arquivo:f.getId(), nome:f.getName(), modificado:ts, cards:lido.cards};
  gravarJsonBase_(pastaBase, NF_ARQ, novo);
  return novo;
}

// mesma regra de "venda" da página (vSituacao): pagamento confirmado, sem cancelamento, total > 0
function ehVendaBase_(r){
  if(/cancel/i.test(r[2]) || r[15] || /cancel/i.test(r[3])) return false;
  if(/recus/i.test(r[3]) || /100%/.test(r[11]) || !(Number(r[8]) > 0)) return false;
  return /confirm|estorn|pago|aprovad|autoriz/i.test(r[3]);
}

// "Blusa Monica Marrom (Marrom, 36)" → {modelo:'Blusa Monica Marrom', cor:'Marrom', tam:'36'}
function partesProduto_(nome, sku){
  var m = /^(.*?)\s*\(([^()]*)\)\s*$/.exec(nome || '');
  if(!m) return {modelo: nome || 'SKU ' + sku, cor:'', tam:''};
  var p = m[2].split(',').map(function(x){ return x.trim(); });
  return {modelo:m[1], cor:p.length > 1 ? p.slice(0, -1).join(', ') : '', tam:p[p.length - 1]};
}

// pedido de origem de uma solicitação → [data do pedido, status do pagamento, meio de pagamento, parcelas, status do envio,
// estorno, data do estorno, total, venda (1/0), entrega depois de (dia ou ''), entrega até (dia ou ''), data de pagamento].
// Textos e valores vão como estão na base (o front lê com vData e vNum). "Entrega até" vem de entregas.json ou, sem ele,
// do momento em que a linha "Entregue" foi gravada (VENDAS_TS): a entrega aconteceu antes disso.
function resumoPedidoDevol_(r, venda, ent, dia){
  var ate = ent ? ent[1] : (/entreg/i.test(r[4]) && Number(r[VENDAS_TS]) ? Number(r[VENDAS_TS]) : 0);
  return [r[1] || '', r[3] || '', r[24] || r[11] || '', r[20] || '', r[4] || '', r[17] || '', r[18] || '', r[8] || '',
    venda ? 1 : 0, ent && ent[0] ? dia(ent[0]) : '', ate ? dia(ate) : '', r[13] || ''];
}

// versao: a que a página já tem; se nada mudou, responde {mudou:false}. A taxa de devolução por peça usa as
// peças vendidas (SKUs da base de vendas), agrupadas por modelo e por modelo|tamanho, por mês do pedido.
function obterDevolucoes(token, versao){ return comErroBase_('devolucoes', function(){ return obterDevolucoes_(token, versao); }); }

function obterDevolucoes_(token, versao){
  if(!contaComToken_(token)) return {erro:'senha_necessaria'};
  var pastaV = pastaVendas_();
  if(!pastaV) return {erro:'pasta_vendas'};
  var pasta = pastaDevolucoes_(pastaV);
  if(!pasta) return {erro:'pasta_devolucoes'};
  var c = consolidarDevolucoes_(pastaV, pasta);
  if(c.erro) return c;
  if(!c.arquivos) return {erro:'sem_arquivo'};
  // leu arquivo novo nesta chamada: as vendas ficam para a próxima (as duas juntas podem passar dos 6 min)
  if(c.leu) return {erro:'consolidando', feitos:c.arquivos, total:c.arquivos};
  var cv = consolidarVendas_(pastaV);
  if(cv.erro === 'consolidando') return cv;
  var ev = cv.estado || {arqs:{}, anos:{}};
  var nf;
  try{ nf = notasFinanceiro_(c.pastaBase); } catch(e){ nf = {erro:'nf_falhou', detalhe:String(e && e.message || e)}; }
  // apiMudou: pedidos atualizados pela API (status de pagamento, envio e entregas dos pedidos de origem)
  var assinatura = [c.estado.arqs, ev.arqs].map(function(a){
    return Object.keys(a).sort().map(function(id){ return id + '@' + a[id]; }).join('|');
  }).join('#') + '#nf:' + (nf.erro || nf.arquivo + '@' + nf.modificado) + '#api:' + (ev.apiMudou || '') + '#fotos:' + (props_().getProperty('FOTOS_VERSAO') || '');
  var novaVersao = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, assinatura));
  if(versao && versao === novaVersao) return {mudou:false, versao:novaVersao};

  var baseArq = lerJsonBase_(c.pastaBase, DEVOL_BASE_ARQ);
  if(!baseArq && Object.keys(c.estado.arqs).length) throw erroBase_(DEVOL_BASE_ARQ, 'não encontrado, mas o estado indica arquivos já lidos');
  var base = baseArq || {sols:{}, hist:{}};
  var nomes = lerJsonBase_(c.pastaBase, PRODUTOS_ARQ) || {};
  var dic = {}, idxDic = {};
  DEVOL_CODIFICADAS.forEach(function(k){ dic[k] = []; idxDic[k] = Object.create(null); });
  function cod(k, v){ var i = idxDic[k][v]; if(i === undefined){ i = idxDic[k][v] = dic[k].length; dic[k].push(v); } return i; }
  var skus = [], idxSku = Object.create(null), modelos = [], idxModelo = Object.create(null), partes = Object.create(null);
  function parte(sku){ return partes[sku] || (partes[sku] = partesProduto_(nomes[sku], sku)); }
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
  // pedidos = resumo dos pedidos de origem das solicitações (resumoPedidoDevol_), sem dados de cliente
  var vendidas = {}, mesesVenda = {}, pedidosOrigem = {};
  var entregas = lerJsonBase_(c.pastaBase, ENTREGAS_ARQ) || {};
  Object.keys(ev.anos || {}).forEach(function(ano){
    var pedidos = lerBaseAno_(ev, ano);
    Object.keys(pedidos).forEach(function(n){
      var r = pedidos[n], venda = ehVendaBase_(r);
      var mes = diaIsoVendas_(r[1]).slice(0, 7);
      if(venda && mes){
        var mv = mesesVenda[mes] = mesesVenda[mes] || [0, 0, 0, 0];
        mv[0]++; mv[1] += Number(r[VENDAS_ITENS]) || 0; if(/entreg/i.test(r[4])) mv[2]++; mv[3] += Number(r[8]) || 0;
      }
      if(origens[n]) pedidosOrigem[n] = resumoPedidoDevol_(r, venda, entregas[n], dia);
      if(!r[VENDAS_SKUS] || !venda) return;
      itensSku_(r[VENDAS_SKUS]).forEach(function(it){
        var sku = it.sku, q = it.q;
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
    var todasFotos = lerFotos_(c.pastaBase);
    skus.forEach(function(x){ if(todasFotos[x[0]]) fotos[x[0]] = todasFotos[x[0]]; });
  } catch(e){ fotos = {}; }
  return {mudou:true, versao:novaVersao, colunas:DEVOL_COLUNAS.concat(['Itens', 'Visto em']), dic:dic, linhas:linhas,
    skus:skus, fotos:fotos, modelos:modelos, vendidas:vendidas, hist:hist, inicio:base.inicio ? dia(base.inicio) : '',
    pedidos:pedidosOrigem, mesesVenda:mesesVenda,
    arquivos:arquivos, avisos:avisos, auto:estadoDevolExt_(),
    nf:nf.erro ? {erro:nf.erro, nome:nf.nome || ''} : {cards:nf.cards, nome:nf.nome, modificadoEm:new Date(nf.modificado).toISOString()}};
}

/* ===================== Devoluções enviadas pela extensão do Chrome ===================== */

// A extensão (pasta Relatorios/extensao-afterclick) lê as solicitações na API interna da AfterClick com a sessão
// aberta no Chrome e manda, a cada hora, um CSV no mesmo layout do export manual (DEVOL_COLUNAS + DEVOL_ITEM).
// Aqui ele vira UM arquivo por dia em nuvemshop/AJD (afterclick-auto-AAAA-MM-DD.csv, atualizado no lugar: apagar
// arquivo faz consolidarDevolucoes_ refazer tudo) e é consolidado na hora. A chave fica em DEVOL_EXT_CHAVE
// (gerarChaveExtensao) e o último envio em DEVOL_EXT: {ultimo, solicitacoes, arquivo, erro}.
var DEVOL_EXT_PREFIXO = 'afterclick-auto-';
var DEVOL_EXT_MAX = 20 * 1024 * 1024;

function estadoDevolExt_(){
  try{ return JSON.parse(props_().getProperty('DEVOL_EXT') || 'null'); } catch(e){ return null; }
}

function respostaJson_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Público de propósito (a extensão AfterClick não tem login do Google): a barreira é a chave DEVOL_EXT_CHAVE.
// Qualquer falha vira uma resposta JSON com erro claro (nunca a página de erro do Google) e fica no registro.
function doPost(e){
  try{
    var corpo = null;
    try{ corpo = JSON.parse(e.postData.contents); } catch(err){ return respostaJson_({erro:'pedido_invalido'}); }
    var acao = corpo && corpo.acao;
    if(acao !== 'devolucoes' && acao !== 'nfs' && acao !== 'pedidos') return respostaJson_({erro:'pedido_invalido'});
    var chave = props_().getProperty('DEVOL_EXT_CHAVE');
    if(!chave || typeof corpo.chave !== 'string' || !iguais_(corpo.chave, chave)) return respostaJson_({erro:'chave_invalida'});
    // NFs e pedidos têm tratamento e estado próprios: uma falha deles nunca marca erro nas devoluções
    if(acao !== 'devolucoes') return respostaJson_(processarOperacaoPost_(acao, corpo));
    var r = receberDevolucoesExtensao_(String(corpo.conteudo || ''), Number(corpo.solicitacoes) || 0);
    var s = estadoDevolExt_() || {};
    if(r.erro && r.erro !== 'consolidando') s.erro = r.erro;
    else { s.ultimo = Date.now(); s.solicitacoes = r.solicitacoes; s.arquivo = r.arquivo; delete s.erro; }
    props_().setProperty('DEVOL_EXT', JSON.stringify(s));
    return respostaJson_(r);
  } catch(err){
    console.error('doPost: ' + textoErro_(err));
    try{
      var s2 = estadoDevolExt_() || {}; s2.erro = 'erro_interno';
      props_().setProperty('DEVOL_EXT', JSON.stringify(s2));
    } catch(x){}
    return respostaJson_({erro:'erro_interno', detalhe:'O painel não conseguiu processar o envio. Tente de novo; se repetir, avise o responsável.'});
  }
}

function receberDevolucoesExtensao_(conteudo, solicitacoes){
  if(!conteudo || conteudo.length > DEVOL_EXT_MAX) return {erro:'tamanho'};
  if(conteudo.indexOf('Pedido de Origem') < 0) return {erro:'formato_devolucoes'};
  var pastaV = pastaVendas_();
  if(!pastaV) return {erro:'pasta_vendas'};
  var pasta = pastaDevolucoes_(pastaV);
  if(!pasta) return {erro:'pasta_devolucoes'};
  var nome = DEVOL_EXT_PREFIXO + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd') + '.csv';
  var it = pasta.getFilesByName(nome), f = it.hasNext() ? it.next() : null;
  if(f){
    // conteúdo novo no mesmo arquivo (a data de modificação muda e consolidarDevolucoes_ relê)
    var resp = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files/' + f.getId() + '?uploadType=media&supportsAllDrives=true', {
      method:'patch', contentType:'text/csv; charset=utf-8', payload:Utilities.newBlob(conteudo, 'text/csv').getBytes(),
      headers:{Authorization:'Bearer ' + ScriptApp.getOAuthToken()}, muteHttpExceptions:true
    });
    if(resp.getResponseCode() !== 200) return {erro:'gravacao_falhou', detalhe:'HTTP ' + resp.getResponseCode()};
  } else f = pasta.createFile(Utilities.newBlob(conteudo, 'text/csv', nome));
  var c = consolidarDevolucoes_(pastaV, pasta);
  return {ok:true, arquivo:nome, solicitacoes:solicitacoes, erro:c.erro || undefined};
}

// rodar uma vez no editor: cria a chave que a extensão manda junto (copiar do registro para as opções dela)
function gerarChaveExtensao(){
  exigirDonoEditor_();
  var chave = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  props_().setProperty('DEVOL_EXT_CHAVE', chave);
  console.log('Chave da extensão (cole nas opções da extensão): ' + chave);
}

/* ===================== Operação AfterClick: NFs de venda e status dos pedidos (extensão) ===================== */

// A mesma extensão que manda as devoluções manda também (acao 'nfs' e 'pedidos') as NFs emitidas e a lista de pedidos
// da AfterClick (AJD). Só vão campos operacionais (nada de nome, documento, e-mail ou protocolo da NF).
// Fluxo de cada envio: (1) mescla as linhas no arquivo do dia em nuvemshop/AJD (afterclick-nfs-AAAA-MM-DD.json e
// afterclick-pedidos-AAAA-MM-DD.json: o JSON guarda a última versão de cada linha enviada naquele dia);
// (2) consolida os arquivos ainda não lidos em nuvemshop/_base/operacao.json, indexado pelo número do pedido da
// Nuvemshop (o número depois de "TLG-"). Arquivos .json não entram em consolidarDevolucoes_ (ehArquivoOp_).
// Base: {formato, ped:{n:[...OP_PED_COLS, lidoEm]}, pedRev:{"D-TLG-n-1":[...]}, nfs:{n:[[...OP_NF_COLS, vistaEm]]},
//        nfRev:{n:[[...]]} (NFs "input" das devoluções D-TLG-n-k, ligadas ao pedido n), hist:{n:{"s:status":ms,"u:sub":ms}}, inicio}
// Datas viram "AAAA-MM-DD HH:MM" (horário de Brasília, como a AfterClick devolve); carimbos "lidoEm/vistaEm/hist" são ms.
var OP_BASE_ARQ = 'operacao.json';
var OP_ESTADO_ARQ = 'operacao-estado.json';
var OP_FORMATO = 1;
var OP_PREFIXOS = {nfs:'afterclick-nfs-', pedidos:'afterclick-pedidos-'};
var OP_EXT_PROP = 'OPER_EXT'; // uma propriedade por tipo (OPER_EXT_nfs, OPER_EXT_pedidos): {ultimo, recebidas, invalidas, ignoradas, arquivo, erro}
var OP_RETENCAO_DIAS = 400;    // pedidos com data do pedido mais antiga que isso saem da base (poda no máx. 1x por dia)
var OP_ARQ_DIAS = 90;          // estado.arqs/info só guardam arquivos dos últimos 90 dias (pelo dia no nome do arquivo)
var OP_MAX_LINHAS = 100000;
var OP_NF_COLS = ['invoiceId','orderId','orderNumber','typeInvoice','invoiceNumber','invoiceSeries','invoiceKey','creationDate',
  'issuanceDate','authorizationDate','invoiceValue','invoiceStatus','isCanceled','cancelationDate','dispatchedDate'];
var OP_PED_COLS = ['orderId','orderNumber','creationDate','orderStatus','orderStatusSub','totalValue','origin','orderType','uf','lastMessage'];
var OP_NF_VISTO = OP_NF_COLS.length;   // posição de "vistaEm": 1ª vez que a extensão viu a NF (só na base)
var OP_PED_VISTO = OP_PED_COLS.length; // posição de "lidoEm": última leitura em que o pedido veio (só na base)

function ehArquivoOp_(nome){ return /^afterclick-(nfs|pedidos)-\d{4}-\d{2}-\d{2}\.json$/.test(String(nome)); }

// {nfs:{...}, pedidos:{...}} (cada tipo numa propriedade própria: dois envios não disputam a mesma propriedade)
function estadoOperExt_(){
  var r = {};
  ['nfs', 'pedidos'].forEach(function(t){
    try{ var x = JSON.parse(props_().getProperty(OP_EXT_PROP + '_' + t) || 'null'); if(x) r[t] = x; } catch(e){}
  });
  return r;
}

// lastMessage da AfterClick é texto livre e pode trazer dado pessoal: só um código fechado é guardado. O texto nunca sai da extensão;
// se uma extensão antiga o mandar, é classificado aqui e descartado.
function codigoMensagemOp_(v){
  var t = textoDevol_(v);
  if(!t) return '';
  if(t === 'abacos_ok' || t === 'autorizado' || t === 'erro' || t === 'outro') return t;
  if(/^aprovado com sucesso no abacos!?$/i.test(t)) return 'abacos_ok';
  if(/^autorizado!?$/i.test(t)) return 'autorizado';
  if(/erro|falha|rejei/i.test(t)) return 'erro';
  return 'outro';
}

// dia (AAAA-MM-DD, Brasília) de agora menos n dias
function diaOp_(dias){ return Utilities.formatDate(new Date(Date.now() - (dias || 0) * 864e5), 'America/Sao_Paulo', 'yyyy-MM-dd'); }

// só "TLG-57160" → {n:'57160', rev:false} e "D-TLG-55624-1" → {n:'55624', rev:true}; qualquer outro formato/prefixo → null (ignorada)
function numeroPedidoOp_(texto){
  var s = String(texto == null ? '' : texto).trim(), m;
  if((m = /^TLG-(\d+)$/.exec(s))) return {n:m[1], rev:false};
  if((m = /^D-TLG-(\d+)-\d+$/.exec(s))) return {n:m[1], rev:true};
  return null;
}

// linha enviada pela extensão (colunas por nome) → linha no layout OP_NF_COLS; null se não tem id
function normalizarNfOp_(r, ix){
  function v(c){ return ix[c] >= 0 ? r[ix[c]] : null; }
  var id = Number(v('invoiceId'));
  if(!isFinite(id) || id <= 0 || !textoDevol_(v('orderNumber'))) return null;
  var valor = Number(v('invoiceValue')), cancelada = v('isCanceled');
  return [id, Number(v('orderId')) || 0, textoDevol_(v('orderNumber')).slice(0, 40), textoDevol_(v('typeInvoice')).toLowerCase().slice(0, 12),
    textoDevol_(v('invoiceNumber')).slice(0, 20), textoDevol_(v('invoiceSeries')).slice(0, 10), textoDevol_(v('invoiceKey')).replace(/\D/g, '').slice(0, 44),
    dataHoraDevol_(v('creationDate')), dataHoraDevol_(v('issuanceDate')), dataHoraDevol_(v('authorizationDate')),
    isFinite(valor) ? Math.round(valor * 100) / 100 : 0, Number(v('invoiceStatus')) || 0, (cancelada === true || Number(cancelada) === 1) ? 1 : 0,
    dataHoraDevol_(v('cancelationDate')), dataHoraDevol_(v('dispatchedDate'))];
}

function normalizarPedidoOp_(r, ix){
  function v(c){ return ix[c] >= 0 ? r[ix[c]] : null; }
  var id = Number(v('orderId'));
  if(!isFinite(id) || id <= 0 || !textoDevol_(v('orderNumber'))) return null;
  var total = Number(v('totalValue'));
  return [id, textoDevol_(v('orderNumber')).slice(0, 40), dataHoraDevol_(v('creationDate')), textoDevol_(v('orderStatus')).slice(0, 40),
    textoDevol_(v('orderStatusSub')).slice(0, 40), isFinite(total) ? Math.round(total * 100) / 100 : 0, textoDevol_(v('origin')).slice(0, 30),
    Number(v('orderType')) || 0, textoDevol_(v('uf')).slice(0, 2), codigoMensagemOp_(v('lastMessage'))];
}

// aplica linhas já normalizadas na base; cont conta novas/atualizadas/ignoradas
function aplicarNfsOp_(base, linhas, ts, cont){
  linhas.forEach(function(r){
    var p = numeroPedidoOp_(r[2]);
    if(!p){ cont.ignoradas++; return; }
    var lista = (r[3] === 'input' || p.rev) ? base.nfRev : base.nfs;
    var arr = lista[p.n] = lista[p.n] || [], i = -1;
    for(var k = 0; k < arr.length; k++) if(arr[k][0] === r[0]){ i = k; break; }
    var nova = r.concat([i >= 0 ? arr[i][OP_NF_VISTO] : ts]); // a vistaEm é a da 1ª vez
    if(i < 0){ arr.push(nova); cont.novas++; }
    else { if(arr[i].join('|') !== nova.join('|')) cont.atualizadas++; arr[i] = nova; }
    if(arr.length > 1) arr.sort(function(a, b){ return a[8] < b[8] ? -1 : a[8] > b[8] ? 1 : a[0] - b[0]; });
  });
}

function aplicarPedidosOp_(base, linhas, ts, cont){
  linhas.forEach(function(r){
    var p = numeroPedidoOp_(r[1]);
    if(!p){ cont.ignoradas++; return; }
    var nova = r.concat([ts]);
    if(p.rev){ // pedido de troca/devolução (D-TLG-n-k): fica à parte, já vem pelas devoluções
      var antR = base.pedRev[r[1]];
      if(!antR) cont.novas++; else if(antR.slice(0, OP_PED_VISTO).join('|') !== r.join('|')) cont.atualizadas++;
      base.pedRev[r[1]] = nova;
      return;
    }
    var ant = base.ped[p.n];
    if(!ant) cont.novas++; else if(ant.slice(0, OP_PED_VISTO).join('|') !== r.join('|')) cont.atualizadas++;
    base.ped[p.n] = nova;
    // hist: quando cada status ("s:") e sub-status ("u:") foi visto pela 1ª vez (carimbo do arquivo)
    var h = base.hist[p.n] = base.hist[p.n] || {};
    if(r[3] && !(('s:' + r[3]) in h)) h['s:' + r[3]] = ts;
    if(r[4] && !(('u:' + r[4]) in h)) h['u:' + r[4]] = ts;
  });
}

function baseOpVazia_(){ return {formato:OP_FORMATO, ped:{}, pedRev:{}, nfs:{}, nfRev:{}, hist:{}, inicio:0}; }

// estado + base. Base ausente com arquivos já consolidados = anomalia (o hist não se refaz): erro, nada é gravado.
function carregarOperacao_(pastaBase){
  var estado = lerJsonBase_(pastaBase, OP_ESTADO_ARQ) || {};
  estado.arqs = estado.arqs || {}; estado.info = estado.info || {}; estado.avisos = estado.avisos || {};
  var baseLida = lerJsonBase_(pastaBase, OP_BASE_ARQ);
  if(!baseLida && Object.keys(estado.arqs).length) throw erroBase_(OP_BASE_ARQ, 'não encontrado, mas o estado indica arquivos já lidos');
  var base = baseLida || baseOpVazia_();
  ['ped', 'pedRev', 'nfs', 'nfRev', 'hist'].forEach(function(k){ base[k] = base[k] || {}; });
  return {estado:estado, base:base};
}

// arquivos de NFs/pedidos da pasta AJD (o mais novo de cada nome), do mais antigo para o mais novo
function listarArquivosOp_(pasta){
  var porNome = {}, it = pasta.getFiles(), corte = diaOp_(OP_ARQ_DIAS);
  while(it.hasNext()){
    var f = it.next(), nome = f.getName();
    if(!ehArquivoOp_(nome)) continue;
    // arquivo com mais de 90 dias (dia no nome) nunca é lido de novo: o estado esquece esses nomes, mas a base já os tem
    if(/(\d{4}-\d{2}-\d{2})\.json$/.exec(nome)[1] < corte) continue;
    var ts = f.getLastUpdated().getTime();
    if(!porNome[nome] || ts > porNome[nome].ts) porNome[nome] = {nome:nome, ts:ts};
  }
  return Object.keys(porNome).map(function(n){ return porNome[n]; }).sort(function(a, b){ return a.ts - b.ts || (a.nome < b.nome ? -1 : 1); });
}

function lerArquivoOp_(pasta, nome){
  var obj = lerJsonBase_(pasta, nome);
  if(!obj) return null;
  if(obj.formato !== OP_FORMATO || !Array.isArray(obj.linhas) || (obj.tipo !== 'nfs' && obj.tipo !== 'pedidos')) throw erroBase_(nome, 'formato desconhecido');
  return obj;
}

// Lê TUDO o que precisa antes de gravar qualquer coisa (leitura que falha não deixa nada pela metade) e depois grava:
// arquivo do dia → base → estado (se cair no meio, o arquivo é lido de novo e a mescla é idempotente).
// extra (opcional) = {tipo, dia, linhas}: linhas novas a mesclar no arquivo do dia antes de consolidar. Chamar com o lock.
function consolidarOperacaoSemLock_(pasta, pastaBase, extra){
  var cur = carregarOperacao_(pastaBase), estado = cur.estado, base = cur.base;
  var nomeExtra = extra ? OP_PREFIXOS[extra.tipo] + extra.dia + '.json' : '';
  var lidos = [];
  listarArquivosOp_(pasta).forEach(function(a){
    if(a.nome === nomeExtra || estado.arqs[a.nome] === a.ts) return;
    var obj = lerArquivoOp_(pasta, a.nome);
    if(obj) lidos.push({nome:a.nome, ts:a.ts, obj:obj});
  });
  var diaObj = null;
  if(extra){
    var atual = lerArquivoOp_(pasta, nomeExtra), porId = {};
    ((atual && atual.linhas) || []).forEach(function(r){ porId[r[0]] = r; });
    extra.linhas.forEach(function(r){ porId[r[0]] = r; });
    diaObj = {formato:OP_FORMATO, tipo:extra.tipo, dia:extra.dia, colunas:extra.tipo === 'nfs' ? OP_NF_COLS : OP_PED_COLS,
      linhas:Object.keys(porId).map(function(k){ return porId[k]; })};
  }
  // ---- gravações
  if(diaObj){
    gravarJsonBase_(pasta, nomeExtra, diaObj);
    lidos.push({nome:nomeExtra, ts:arquivoBase_(pasta, nomeExtra).getLastUpdated().getTime(), obj:diaObj});
  }
  var cont = {novas:0, atualizadas:0, ignoradas:0};
  lidos.sort(function(a, b){ return a.ts - b.ts; });
  lidos.forEach(function(x){
    var c = {novas:0, atualizadas:0, ignoradas:0};
    if(x.obj.tipo === 'nfs') aplicarNfsOp_(base, x.obj.linhas, x.ts, c); else aplicarPedidosOp_(base, x.obj.linhas, x.ts, c);
    cont.novas += c.novas; cont.atualizadas += c.atualizadas; cont.ignoradas += c.ignoradas;
    estado.arqs[x.nome] = x.ts;
    estado.info[x.nome] = {tipo:x.obj.tipo, dia:x.obj.dia, modificadoEm:new Date(x.ts).toISOString(), linhas:x.obj.linhas.length};
    if(c.ignoradas) estado.avisos[x.nome] = {nome:x.nome, ignoradas:c.ignoradas}; else delete estado.avisos[x.nome];
    if(!base.inicio || x.ts < base.inicio) base.inicio = x.ts;
  });
  if(lidos.length){
    podarOperacao_(base, estado);
    if(!estado.epoca) estado.epoca = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    base.formato = OP_FORMATO; estado.formato = OP_FORMATO;
    gravarJsonBase_(pastaBase, OP_BASE_ARQ, base);
    gravarJsonBase_(pastaBase, OP_ESTADO_ARQ, estado);
  }
  return {estado:estado, base:base, lidos:lidos.length, contagem:cont};
}

// Retenção (no máx. 1x por dia, só quando há gravação): tira da base os pedidos com data do pedido > OP_RETENCAO_DIAS e esquece
// no estado os arquivos com mais de OP_ARQ_DIAS. Data vazia nunca é apagada.
function podarOperacao_(base, estado){
  var hoje = diaOp_(0);
  if(base.podadoEm === hoje) return;
  var corte = diaOp_(OP_RETENCAO_DIAS), corteArq = diaOp_(OP_ARQ_DIAS);
  var velha = function(d){ d = String(d || '').slice(0, 10); return !!d && d < corte; };
  Object.keys(base.ped).forEach(function(n){ if(velha(base.ped[n][2])){ delete base.ped[n]; delete base.hist[n]; } });
  Object.keys(base.pedRev).forEach(function(k){ if(velha(base.pedRev[k][2])) delete base.pedRev[k]; });
  ['nfs', 'nfRev'].forEach(function(k){
    Object.keys(base[k]).forEach(function(n){ if(base[k][n].every(function(r){ return velha(r[7] || r[8]); })) delete base[k][n]; });
  });
  Object.keys(base.hist).forEach(function(n){ if(!base.ped[n]) delete base.hist[n]; });
  Object.keys(estado.arqs).forEach(function(nome){
    var m = /(\d{4}-\d{2}-\d{2})\.json$/.exec(nome);
    if(m && m[1] < corteArq){ delete estado.arqs[nome]; delete estado.info[nome]; delete estado.avisos[nome]; }
  });
  base.podadoEm = hoje;
}

// consolida o que estiver pendente (nada pendente = só lê o estado, sem lock nem gravação)
function consolidarOperacao_(pasta, pastaBase){
  var estado = lerJsonBase_(pastaBase, OP_ESTADO_ARQ) || {arqs:{}};
  var arqs = estado.arqs || {};
  var pendente = listarArquivosOp_(pasta).some(function(a){ return arqs[a.nome] !== a.ts; });
  if(!pendente) return {estado:{arqs:arqs, info:estado.info || {}, avisos:estado.avisos || {}, epoca:estado.epoca || ''}};
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return {erro:'consolidando'};
  try{ return consolidarOperacaoSemLock_(pasta, pastaBase, null); }
  finally { lock.releaseLock(); }
}

// recebe {acao, chave, colunas:[nomes], linhas:[[...]], epoca}. epoca = a que o painel devolveu no último envio: se não
// bate com a da base (base refeita/apagada), responde reenviar_tudo e a extensão manda tudo de novo (ela só manda o que mudou).
function receberOperacaoExtensao_(tipo, corpo){
  var cols = tipo === 'nfs' ? OP_NF_COLS : OP_PED_COLS;
  if(!Array.isArray(corpo.colunas) || !Array.isArray(corpo.linhas)) return {erro:'formato_' + tipo};
  if(corpo.linhas.length > OP_MAX_LINHAS) return {erro:'tamanho'};
  var ix = {};
  cols.forEach(function(c){ ix[c] = corpo.colunas.indexOf(c); });
  var chaves = tipo === 'nfs' ? ['invoiceId', 'orderNumber'] : ['orderId', 'orderNumber'];
  if(ix[chaves[0]] < 0 || ix[chaves[1]] < 0) return {erro:'formato_' + tipo};
  var normalizar = tipo === 'nfs' ? normalizarNfOp_ : normalizarPedidoOp_, linhas = [], invalidas = 0;
  corpo.linhas.forEach(function(r){
    var n = Array.isArray(r) ? normalizar(r, ix) : null;
    if(n) linhas.push(n); else invalidas++;
  });
  var pastaV = pastaVendas_();
  if(!pastaV) return {erro:'pasta_vendas'};
  var pasta = pastaDevolucoes_(pastaV);
  if(!pasta) return {erro:'pasta_devolucoes'};
  var pastaBase = pastaBaseVendas_(pastaV);
  var lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return {erro:'ocupado'};
  try{
    var estado = lerJsonBase_(pastaBase, OP_ESTADO_ARQ) || {};
    var epocaBase = String(estado.epoca || '');
    if(String(corpo.epoca || '') !== epocaBase) return {erro:'reenviar_tudo', epoca:epocaBase};
    var dia = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
    var nome = OP_PREFIXOS[tipo] + dia + '.json';
    if(!linhas.length) return {ok:true, tipo:tipo, recebidas:0, invalidas:invalidas, novas:0, atualizadas:0, ignoradas:0, arquivo:nome, epoca:epocaBase};
    var c = consolidarOperacaoSemLock_(pasta, pastaBase, {tipo:tipo, dia:dia, linhas:linhas});
    return {ok:true, tipo:tipo, recebidas:linhas.length, invalidas:invalidas, novas:c.contagem.novas, atualizadas:c.contagem.atualizadas,
      ignoradas:c.contagem.ignoradas, arquivo:nome, epoca:c.estado.epoca};
  } finally { lock.releaseLock(); }
}

// doPost de 'nfs' e 'pedidos': qualquer falha vira JSON com erro e fica em OPER_EXT (sem tocar no estado das devoluções)
function processarOperacaoPost_(tipo, corpo){
  var r;
  try{ r = comErroBase_('operacao', function(){ return receberOperacaoExtensao_(tipo, corpo); }); }
  catch(err){
    console.error('doPost ' + tipo + ': ' + textoErro_(err));
    r = {erro:'erro_interno', detalhe:'O painel não conseguiu processar o envio. Tente de novo; se repetir, avise o responsável.'};
  }
  try{
    var prop = OP_EXT_PROP + '_' + tipo, x = {};
    try{ x = JSON.parse(props_().getProperty(prop) || '{}') || {}; } catch(e2){ x = {}; }
    if(r.erro && r.erro !== 'reenviar_tudo') x.erro = r.erro;
    else if(!r.erro){
      x.ultimo = Date.now(); x.recebidas = r.recebidas; x.invalidas = r.invalidas || 0; x.ignoradas = r.ignoradas || 0; x.arquivo = r.arquivo; delete x.erro;
    }
    props_().setProperty(prop, JSON.stringify(x));
  } catch(e){}
  return r;
}

// Base de operação para o painel (parte B). versao = a que a página já tem; sem mudança → {mudou:false}.
function obterOperacao(token, versao){ return comErroBase_('operacao', function(){ return obterOperacao_(token, versao); }); }

function obterOperacao_(token, versao){
  if(!contaComToken_(token)) return {erro:'senha_necessaria'};
  var pastaV = pastaVendas_();
  if(!pastaV) return {erro:'pasta_vendas'};
  var pasta = pastaDevolucoes_(pastaV);
  if(!pasta) return {erro:'pasta_devolucoes'};
  var pastaBase = pastaBaseVendas_(pastaV);
  var c = consolidarOperacao_(pasta, pastaBase);
  if(c.erro) return c;
  var estado = c.estado;
  if(!Object.keys(estado.arqs).length) return {erro:'sem_arquivo'};
  var assinatura = Object.keys(estado.arqs).sort().map(function(n){ return n + '@' + estado.arqs[n]; }).join('|') + '#' + (estado.epoca || '');
  var novaVersao = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, assinatura));
  if(versao && versao === novaVersao) return {mudou:false, versao:novaVersao};
  var base = carregarOperacao_(pastaBase).base;
  var arquivos = Object.keys(estado.info).map(function(n){ return estado.info[n]; })
    .sort(function(a, b){ return a.modificadoEm < b.modificadoEm ? -1 : 1; });
  return {mudou:true, versao:novaVersao, formato:OP_FORMATO, colunasNf:OP_NF_COLS.concat(['vistaEm']), colunasPed:OP_PED_COLS.concat(['lidoEm']),
    ped:base.ped, pedRev:base.pedRev, nfs:base.nfs, nfRev:base.nfRev, hist:base.hist, inicio:base.inicio || 0,
    arquivos:arquivos, avisos:Object.keys(estado.avisos).map(function(n){ return estado.avisos[n]; }), auto:estadoOperExt_()};
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

function pastaNuvemPago_(){
  try{ return DriveApp.getFolderById(props_().getProperty('NUVEMPAGO_PASTA') || NP_PASTA_ID); } catch(e){ return null; }
}

function arquivosNuvemPago_(pasta){
  var lista = [], it = pasta.getFiles();
  while(it.hasNext()){
    var f = it.next();
    if(/\.csv$/i.test(f.getName()) || /csv/i.test(f.getMimeType())) lista.push(f);
  }
  lista.sort(function(a, b){ return a.getLastUpdated() - b.getLastUpdated(); });
  return lista;
}

// "Valor Liquido (R$)" ou "Valor Líquido (R$)": cabeçalho sem acento, sem caixa e sem espaço sobrando
function chaveCabecalho_(c){
  return String(c == null ? '' : c).replace(/^﻿/, '').trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ');
}

// "01/10/2026" → "2026-10-01" ('' se não for data)
function diaNuvemPago_(v){
  var m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(textoDevol_(v));
  return m ? m[3] + '-' + m[2] + '-' + m[1] : '';
}

// arquivo → {grupos: {"dia|movimentação|tipo": [linhas]}, de, ate, n} (linha no layout de NP_COLUNAS, sem o nome)
function lerNuvemPago_(f){
  var blob = f.getBlob(), texto = blob.getDataAsString('UTF-8').replace(/^﻿/, '');
  if(texto.indexOf('�') >= 0) texto = blob.getDataAsString('ISO-8859-1');
  var linhas = Utilities.parseCsv(texto, ';');
  if(!linhas.length) return {erro:'csv_vazio'};
  var cab = linhas[0].map(chaveCabecalho_);
  var ix = NP_COLUNAS.map(function(c){ return cab.indexOf(chaveCabecalho_(c)); });
  if(ix[0] < 0 || ix[1] < 0 || ix[3] < 0 || ix[12] < 0) return {erro:'formato_nuvempago'};
  var grupos = {}, de = '', ate = '', n = 0;
  for(var i = 1; i < linhas.length; i++){
    var l = linhas[i];
    var x = ix.map(function(j){ var v = j >= 0 ? textoDevol_(l[j]) : ''; return v === '-' ? '' : v; });
    x[0] = diaNuvemPago_(x[0]); x[1] = diaNuvemPago_(x[1]);
    if(!x[1] || !x[3]) continue;
    x[7] = x[7] ? Number(x[7]) || '' : '';
    NP_NUMERICAS.forEach(function(k){ x[k] = numeroBR_(x[k]); });
    var k = x[1] + '|' + chaveCabecalho_(x[2]) + '|' + chaveCabecalho_(x[3]);
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
function consolidarNuvemPago_(relogio){ return comErroBase_('nuvempago', function(){ return consolidarNuvemPagoNucleo_(relogio); }); }

function consolidarNuvemPagoNucleo_(relogio){
  var pasta = pastaNuvemPago_();
  if(!pasta) return {erro:'pasta_nuvempago'};
  var pastaV = pastaVendas_();
  if(!pastaV) return {erro:'pasta_vendas'};
  var pastaBase = pastaBaseVendas_(pastaV);
  var arquivos = arquivosNuvemPago_(pasta);
  var estado = lerJsonBase_(pastaBase, NP_ESTADO_ARQ) || {};
  estado.arqs = estado.arqs || {}; estado.info = estado.info || {}; estado.avisos = estado.avisos || {};
  var atuais = {};
  arquivos.forEach(function(f){ atuais[f.getId()] = f.getLastUpdated().getTime(); });
  var removido = estado.formato !== NP_FORMATO || Object.keys(estado.arqs).some(function(id){ return !(id in atuais); });
  var falta = arquivos.filter(function(f){ return estado.arqs[f.getId()] !== atuais[f.getId()]; });
  if(!removido && !falta.length) return {estado:estado, pastaBase:pastaBase, arquivos:arquivos.length};

  var lock = LockService.getScriptLock();
  if(!lock.tryLock(30000)) return {erro:'consolidando', feitos:null, total:arquivos.length};
  try{
    estado = lerJsonBase_(pastaBase, NP_ESTADO_ARQ) || estado;
    estado.arqs = estado.arqs || {}; estado.info = estado.info || {}; estado.avisos = estado.avisos || {};
    var baseLida = lerJsonBase_(pastaBase, NP_BASE_ARQ);
    var base = baseLida || {};
    var inicio = relogio || Date.now();
    removido = estado.formato !== NP_FORMATO || Object.keys(estado.arqs).some(function(id){ return !(id in atuais); });
    // um extrato saiu da pasta, foi trocado no lugar ou a base é de um formato antigo: refaz do zero
    if(removido || falta.some(function(f){ return f.getId() in estado.arqs; })){
      base = {}; estado.arqs = {}; estado.info = {}; estado.avisos = {}; estado.formato = NP_FORMATO;
    } else if(!baseLida && Object.keys(estado.arqs).length){
      // sem reconstrução prevista, base ausente com extratos já lidos no estado = anomalia: não grava base vazia
      throw erroBase_(NP_BASE_ARQ, 'não encontrado, mas o estado indica extratos já lidos');
    }
    base.grupos = base.grupos || {};
    falta = arquivos.filter(function(f){ return estado.arqs[f.getId()] !== atuais[f.getId()]; });
    var feitos = 0;
    for(var i = 0; i < falta.length; i++){
      if(feitos && Date.now() - inicio > VENDAS_TEMPO_MAX) break;
      var f = falta[i], id = f.getId(), ts = atuais[id];
      var lido = lerNuvemPago_(f);
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
      gravarJsonBase_(pastaBase, NP_BASE_ARQ, base);
      gravarJsonBase_(pastaBase, NP_ESTADO_ARQ, estado);
      feitos++;
    }
    if(!feitos){
      gravarJsonBase_(pastaBase, NP_BASE_ARQ, base);
      gravarJsonBase_(pastaBase, NP_ESTADO_ARQ, estado);
    }
    var restam = falta.length - feitos;
    if(restam > 0) return {erro:'consolidando', feitos:arquivos.length - restam, total:arquivos.length};
    // só com a base completa: no meio de uma reconstrução as vendas não são completadas com dados parciais
    props_().setProperty('NUVEMPAGO_VERSAO', String(Date.now()));
    return {estado:estado, pastaBase:pastaBase, arquivos:arquivos.length};
  } finally { lock.releaseLock(); }
}

// todos os movimentos da base, um por linha do extrato
function movimentosNuvemPago_(base){
  var out = [];
  Object.keys(base.grupos || {}).forEach(function(k){ base.grupos[k].l.forEach(function(x){ out.push(x); }); });
  return out;
}

// pedido → soma das linhas de Venda e de Estorno no Nuvem Pago (para completar as vendas que vieram só da API)
function financeiroNuvemPago_(base){
  var porPedido = {};
  movimentosNuvemPago_(base).forEach(function(x){
    var num = x[4], tipo = chaveCabecalho_(x[3]);
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
function completarComNuvemPago_(pedidos, fin){
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

function lerBaseNuvemPago_(pastaBase){
  return lerJsonBase_(pastaBase, NP_BASE_ARQ) || {grupos:{}};
}

// versao: a que a página já tem; se nada mudou, responde {mudou:false}
function obterNuvemPago(token, versao){ return comErroBase_('nuvempago', function(){ return obterNuvemPago_(token, versao); }); }

function obterNuvemPago_(token, versao){
  var conta = contaComToken_(token);
  if(!conta) return {erro:'senha_necessaria'};
  if(!podeVerFinanceiro_(conta.email)) return {erro:'sem_permissao'};
  var c = consolidarNuvemPago_();
  if(c.erro) return c;
  if(!c.arquivos) return {erro:'sem_arquivo'};
  var assinatura = Object.keys(c.estado.arqs).sort().map(function(id){ return id + '@' + c.estado.arqs[id]; }).join('|');
  var novaVersao = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, 'np' + NP_FORMATO + '#' + assinatura));
  if(versao && versao === novaVersao) return {mudou:false, versao:novaVersao};
  var dic = {}, idxDic = {};
  NP_CODIFICADAS.forEach(function(k){ dic[k] = []; idxDic[k] = Object.create(null); });
  var linhas = movimentosNuvemPago_(lerBaseNuvemPago_(c.pastaBase)).map(function(x){
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
    arquivos:arquivos, avisos:avisos, pasta:props_().getProperty('NUVEMPAGO_PASTA') || NP_PASTA_ID};
}

/* ===================== Saldo conhecido do Nuvem Pago (compartilhado no servidor) ===================== */

// O dono informa o saldo do Nuvem Pago no fim de um dia (tela de Saldo do Nuvem Pago). O valor fica numa propriedade do script,
// com data, autor e hora, e vale para todo o perfil financeiro (antes ficava só no navegador de cada pessoa). Cada gravação
// também entra num histórico dos 20 últimos. Quem não tem acesso ao financeiro não lê nem grava.
var NP_SALDO_PROP = 'NP_SALDO_CONHECIDO';
var NP_SALDO_HIST = 20;

function lerSaldoNP_(){
  var x = null;
  try{ x = JSON.parse(props_().getProperty(NP_SALDO_PROP) || 'null'); } catch(e){}
  return x && typeof x === 'object' ? x : {data:'', valor:null, autor:'', em:'', hist:[]};
}

// {saldo:{data,valor,autor,em}|null, hist:[...]} ou {erro}
function obterSaldoNP(token){
  var conta = contaComToken_(token);
  if(!conta) return {erro:'senha_necessaria'};
  if(!podeVerFinanceiro_(conta.email)) return {erro:'sem_permissao'};
  var x = lerSaldoNP_();
  var tem = x.data && typeof x.valor === 'number' && isFinite(x.valor);
  return {saldo: tem ? {data:x.data, valor:x.valor, autor:x.autor || '', em:x.em || ''} : null, hist:x.hist || []};
}

// dados: {data:'AAAA-MM-DD', valor:número} grava; {limpar:true} limpa o saldo atual (o histórico fica)
function salvarSaldoNP(token, dados){
  var conta = contaComToken_(token);
  if(!conta) return {erro:'senha_necessaria'};
  if(!podeVerFinanceiro_(conta.email)) return {erro:'sem_permissao'};
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
    props_().setProperty(NP_SALDO_PROP, JSON.stringify(x));
  } finally { lock.releaseLock(); }
  return obterSaldoNP(token);
}

/* ===================== Metas e prazos de Vendas (etapa 30) ===================== */

// Metas mensais aprovadas pelo usuário em 29/10/2026 (proposta do vendas-ecommerce sobre o histórico de mar–out/2026).
// verde = a partir de quanto é verde; amarelo = a partir de quanto é amarelo (abaixo disso, vermelho). Em "pct" o valor
// é percentual (38 = 38%). Os prazos (dias) são das filas de Vendas. Ficam na propriedade VENDAS_METAS; só o
// administrador grava; todo cadastrado lê. Sem propriedade gravada valem os padrões abaixo.
var VENDAS_METAS_PROP = 'VENDAS_METAS';
var VENDAS_METAS_PADRAO = {
  faturamento:{verde:2180000, amarelo:1490000},
  pedidos:{verde:5000, amarelo:3500},
  ticket:{verde:400, amarelo:340},
  recorrentes:{verde:38, amarelo:30},
  prazos:{pendente:2, boleto:4, cancelPago:5, topProblema:3, recusadoRecorrente:1, emRisco:90}
};

function metasVendas_(){
  var x = null;
  try{ x = JSON.parse(props_().getProperty(VENDAS_METAS_PROP) || 'null'); } catch(e){}
  var m = JSON.parse(JSON.stringify(VENDAS_METAS_PADRAO));
  if(x && typeof x === 'object'){
    Object.keys(m).forEach(function(k){
      if(!x[k] || typeof x[k] !== 'object') return;
      Object.keys(m[k]).forEach(function(c){ var v = Number(x[k][c]); if(isFinite(v) && v >= 0) m[k][c] = v; });
    });
    m.autor = x.autor || ''; m.em = x.em || '';
  }
  return m;
}

function obterMetasVendas(token){
  if(!contaComToken_(token)) return {erro:'senha_necessaria'};
  return {metas:metasVendas_(), padrao:VENDAS_METAS_PADRAO};
}

// metas: mesmo formato de VENDAS_METAS_PADRAO (só os campos conhecidos são lidos); {padrao:true} volta aos padrões
function salvarMetasVendas(token, metas){
  var conta = contaAdmin_(token);
  if(!conta) return {erro:'sem_permissao'};
  metas = metas || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    if(metas.padrao === true){ props_().deleteProperty(VENDAS_METAS_PROP); return obterMetasVendas(token); }
    var nova = {};
    for(var k in VENDAS_METAS_PADRAO){
      nova[k] = {};
      for(var c in VENDAS_METAS_PADRAO[k]){
        var v = metas[k] && metas[k][c] !== '' && metas[k][c] != null ? Number(metas[k][c]) : VENDAS_METAS_PADRAO[k][c];
        if(!isFinite(v) || v < 0 || v > (k === 'recorrentes' ? 100 : 1e9)) return {erro:'valor_invalido', campo:k + '.' + c};
        nova[k][c] = Math.round(v * 100) / 100;
      }
      if(k !== 'prazos' && nova[k].amarelo > nova[k].verde) return {erro:'faixa_invalida', campo:k};
    }
    nova.autor = conta.email; nova.em = new Date().toISOString();
    props_().setProperty(VENDAS_METAS_PROP, JSON.stringify(nova));
  } finally { lock.releaseLock(); }
  return obterMetasVendas(token);
}

/* ===================== Relatório da manhã por e-mail ===================== */

// Uma rotina do Claude (na nuvem) monta o relatório da manhã a partir da Nuvemshop e grava um .html na
// pasta abaixo. enviarRelatoriosDiarios() roda por gatilho a cada 15 min, manda cada arquivo novo por
// e-mail (assunto = <title> do arquivo) e marca o arquivo como enviado na descrição dele.
// Para ligar: rodar instalarRelatorioDiario() uma vez no editor, com a conta dona do script.
var RELATORIO_PASTA_ID = '11w2NTecfVU82w5qUiCnn7VAPe9kpU_5K';
var RELATORIO_PARA = 'messias.alexandro@talgui.com.br';

function enviarRelatoriosDiarios(){
  exigirDono_();
  if(!gatilhoLiberado_()) return;
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
  exigirDonoEditor_();
  ScriptApp.getProjectTriggers().forEach(function(g){
    if(g.getHandlerFunction() === 'enviarRelatoriosDiarios') ScriptApp.deleteTrigger(g);
  });
  ScriptApp.newTrigger('enviarRelatoriosDiarios').timeBased().everyMinutes(15).create();
  enviarRelatoriosDiarios();
}
