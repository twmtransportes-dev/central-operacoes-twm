/**
 * Ranking de Economia de Combustível — backend do TWM DataLink
 *
 * INSTALAÇÃO
 * 1. Cole este arquivo como um novo .gs no projeto Apps Script existente.
 * 2. No doPost, logo depois do JSON.parse, acrescente:
 *
 *      var resp = rkRotear(dados);
 *      if (resp) return rkJson(resp);
 *
 * 3. O RK_ARQUIVO_ID abaixo já aponta para a planilha "Consumo de combustivel".
 * 4. Rode UMA VEZ a função rkInstalar() pelo editor.
 * 5. Rode rkTestarPasta() e confira o log antes de implantar.
 * 6. Implante com "Nova versão" na implantação existente.
 *
 * COMO O ARQUIVO DEVE ESTAR NO DRIVE
 * Uma aba por mês. O nome da aba define a competência:
 *   "08/2026", "AGOSTO", "AGOSTO 2026", "AGO/26", "2026-08" — todos funcionam.
 * O relatório é acumulado do dia 1 até a exportação, então cada leitura
 * substitui o mês inteiro. Abas sem mês reconhecível são ignoradas.
 * Como o arquivo já é planilha do Google, não precisa ativar a Drive API.
 */

var RK_PLANILHA = '1Q1XHvwuho8nNFAUMd3XcZx70njgU7kKatVnbk45X5Xg'; // PLANILHA_ID do DataLink

// Arquivo do Drive com o relatório de consumo — uma aba por mês.
var RK_ARQUIVO_ID = '1VYxdMt2Abx4ljJJYRY5qwLUBqQo-_QtTbG5SP12JR8s';

// Alternativa: deixe RK_ARQUIVO_ID vazio e informe uma pasta;
// aí o script usa o arquivo mais recente que encontrar nela.
var RK_PASTA_ID = '';

var RK_ABA_DADOS  = 'RANKING_CONSUMO';
var RK_ABA_CONFIG = 'RANKING_CONFIG';
var RK_CABECALHO  = ['COMPETENCIA','MOTORISTA','VEICULO','MODELO','KM','LITROS','FAIXA_VERDE','CO2','IMPORTADO_EM'];

var RK_HORA_SYNC = 3; // hora do dia em que o gatilho roda

/* ==========================================================
   INSTALAÇÃO — rode uma vez pelo editor
   ========================================================== */
function rkInstalar(){
  rkAba(RK_ABA_DADOS, RK_CABECALHO);
  rkAba(RK_ABA_CONFIG, ['CHAVE','VALOR']);

  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction() === 'rkSincronizar') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rkSincronizar').timeBased().everyDays(1).atHour(RK_HORA_SYNC).create();

  Logger.log('Pronto. Gatilho diário às ' + RK_HORA_SYNC + 'h. Cada aba do arquivo vira um mês.');
  return 'ok';
}

/* ==========================================================
   ROTEADOR
   ========================================================== */
function rkJson(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Devolve null se a ação não for deste módulo — deixa o roteador seguir. */
function rkRotear(d){
  if(!d || !d.acao) return null;
  try{
    switch(d.acao){
      case 'rkInicio':
        var cfg = rkLerConfig();
        return { config: cfg, periodos: rkPeriodos(), sync: cfg.ultimaSync || null };
      case 'rkDados':          return { linhas: rkLerLinhas(d.competencia) };
      case 'rkHistorico':      return rkHistorico();
      case 'rkSalvarConfig':   return rkSalvarConfig(d.config);
      case 'rkImportar':       return rkGravar(d.competencia, d.linhas);
      case 'rkSincronizarJa':  return rkSincronizar(true);
      default: return null;
    }
  }catch(err){
    return { erro: String(err && err.message || err) };
  }
}

/* ==========================================================
   ABAS INTERNAS E CONFIGURAÇÃO
   ========================================================== */
function rkAba(nome, cabecalho){
  var ss = SpreadsheetApp.openById(RK_PLANILHA);
  var aba = ss.getSheetByName(nome);
  if(!aba){
    aba = ss.insertSheet(nome);
    aba.appendRow(cabecalho);
    aba.setFrozenRows(1);
    aba.getRange(1,1,1,cabecalho.length).setFontWeight('bold');
  }
  return aba;
}

function rkLerConfig(){
  var aba = rkAba(RK_ABA_CONFIG, ['CHAVE','VALOR']);
  var vals = aba.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][0]) === 'config'){
      try{ return JSON.parse(vals[i][1]) || {}; }catch(e){ return {}; }
    }
  }
  return {};
}

function rkSalvarConfig(config){
  var aba   = rkAba(RK_ABA_CONFIG, ['CHAVE','VALOR']);
  var atual = rkLerConfig();
  var novo  = config || {};
  // campos controlados pelo servidor: preserva se o app nao mandou
  if(novo.ultimaSync     === undefined) novo.ultimaSync     = atual.ultimaSync;
  if(novo.carimboArquivo === undefined) novo.carimboArquivo = atual.carimboArquivo;

  var txt  = JSON.stringify(novo);
  var vals = aba.getDataRange().getValues();
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][0]) === 'config'){
      aba.getRange(i+1, 2).setValue(txt);
      return { ok:true };
    }
  }
  aba.appendRow(['config', txt]);
  return { ok:true };
}

/* ==========================================================
   LEITURA DOS DADOS GRAVADOS
   ========================================================== */
function rkPeriodos(){
  var aba = rkAba(RK_ABA_DADOS, RK_CABECALHO);
  var n = aba.getLastRow();
  if(n < 2) return [];
  var col = aba.getRange(2, 1, n-1, 1).getValues();
  var vistos = {}, out = [];
  for(var i=0;i<col.length;i++){
    var c = String(col[i][0] || '').trim();
    if(c && !vistos[c]){ vistos[c] = 1; out.push(c); }
  }
  return out.sort().reverse();
}

function rkLerLinhas(competencia){
  var aba = rkAba(RK_ABA_DADOS, RK_CABECALHO);
  var n = aba.getLastRow();
  if(n < 2) return [];
  var vals = aba.getRange(2, 1, n-1, RK_CABECALHO.length).getValues();
  var out = [];
  for(var i=0;i<vals.length;i++){
    var v = vals[i];
    if(competencia && String(v[0]).trim() !== String(competencia).trim()) continue;
    if(!v[1]) continue;
    out.push({
      motorista: String(v[1]).trim(),
      veiculo:   String(v[2]).trim(),
      modelo:    String(v[3]).trim(),
      km:        Number(v[4]) || 0,
      litros:    Number(v[5]) || 0,
      faixaVerde:Number(v[6]) || 0,
      co2:       Number(v[7]) || 0
    });
  }
  return out;
}

/* ==========================================================
   GRAVACAO — substitui integralmente a competencia
   ========================================================== */
function rkGravar(competencia, linhas){
  if(!competencia) throw new Error('Competência não informada');
  if(!linhas || !linhas.length) throw new Error('Nenhuma linha para gravar');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    var aba  = rkAba(RK_ABA_DADOS, RK_CABECALHO);
    var comp = String(competencia).trim();
    var n    = aba.getLastRow();

    // apaga o que ja existia nesta competencia, em blocos contiguos
    if(n > 1){
      var col = aba.getRange(2, 1, n-1, 1).getValues();
      var i = col.length - 1;
      while(i >= 0){
        if(String(col[i][0]).trim() === comp){
          var fim = i;
          while(i >= 0 && String(col[i][0]).trim() === comp) i--;
          aba.deleteRows(i + 3, fim - i);
        } else { i--; }
      }
    }

    var agora = new Date();
    var bloco = linhas.map(function(l){
      return [comp,
        String(l.motorista || '').trim(),
        String(l.veiculo   || '').trim(),
        String(l.modelo    || '').trim(),
        Number(l.km) || 0, Number(l.litros) || 0,
        Number(l.faixaVerde) || 0, Number(l.co2) || 0,
        agora];
    });
    aba.getRange(aba.getLastRow() + 1, 1, bloco.length, RK_CABECALHO.length).setValues(bloco);
    return { ok:true, gravadas: bloco.length, competencia: comp };
  } finally {
    lock.releaseLock();
  }
}

/* ==========================================================
   SINCRONIZACAO COM O DRIVE
   Le TODAS as abas do arquivo mais recente da pasta.
   Cada aba e um mes e substitui aquele mes por inteiro.
   ========================================================== */
function rkSincronizar(forcar){
  var cfg = rkLerConfig();
  var arq = rkFonte();

  var carimbo = arq.getId() + '|' + arq.getLastUpdated().getTime();
  if(!forcar && cfg.carimboArquivo === carimbo)
    return { ok:true, semMudanca:true, arquivo: arq.getName(), sync: cfg.ultimaSync || null };

  var abas = rkAbasDoArquivo(arq);
  if(!abas.length) throw new Error('O arquivo "' + arq.getName() + '" não tem nenhuma aba legível');

  var anoBase = rkAnoBase(arq);
  var meses = [], ignoradas = [], total = 0;

  for(var i = 0; i < abas.length; i++){
    var ab = abas[i];

    var comp = rkCompetenciaDaAba(ab.nome, anoBase);
    // arquivo de aba unica: aceita o mes vindo do nome do arquivo ou da configuracao
    if(!comp && abas.length === 1) comp = cfg.competenciaForcada || rkCompetenciaDoArquivo(arq);
    if(!comp){ ignoradas.push(ab.nome + ' (mês não identificado)'); continue; }

    var linhas;
    try{
      linhas = rkLinhasDaGrade(ab.valores);
    }catch(e){
      ignoradas.push(ab.nome + ' (' + e.message + ')');
      continue;
    }
    if(!linhas.length){ ignoradas.push(ab.nome + ' (sem linhas de consumo)'); continue; }

    var r = rkGravar(comp, linhas);
    meses.push({ competencia: comp, aba: ab.nome, gravadas: r.gravadas });
    total += r.gravadas;
  }

  if(!meses.length)
    throw new Error('Nenhuma aba pôde ser lida. ' + ignoradas.join('; '));

  meses.sort(function(a,b){ return a.competencia < b.competencia ? -1 : 1; });

  cfg = rkLerConfig();
  cfg.carimboArquivo = carimbo;
  cfg.ultimaSync = {
    quando: Utilities.formatDate(new Date(), 'America/Sao_Paulo', "dd/MM/yyyy 'às' HH:mm"),
    arquivo: arq.getName(),
    linhas: total,
    meses: meses,
    ignoradas: ignoradas,
    competencia: meses[meses.length-1].competencia
  };
  rkSalvarConfig(cfg);

  return { ok:true, arquivo: arq.getName(), linhas: total, meses: meses,
           ignoradas: ignoradas, competencia: cfg.ultimaSync.competencia, sync: cfg.ultimaSync };
}

/** Arquivo de origem: o id fixo, ou o mais recente da pasta. */
function rkFonte(){
  if(RK_ARQUIVO_ID && RK_ARQUIVO_ID.indexOf('COLE_AQUI') !== 0){
    try{
      return DriveApp.getFileById(RK_ARQUIVO_ID);
    }catch(e){
      throw new Error('Não consegui abrir o arquivo do Drive. Confira o RK_ARQUIVO_ID e se a conta do script tem acesso a ele. Detalhe: ' + e);
    }
  }
  if(!RK_PASTA_ID) throw new Error('Informe RK_ARQUIVO_ID ou RK_PASTA_ID no script');
  var arq = rkArquivoMaisRecente(RK_PASTA_ID);
  if(!arq) throw new Error('Nenhum arquivo compatível na pasta do Drive');
  return arq;
}

function rkArquivoMaisRecente(pastaId){
  var it = DriveApp.getFolderById(pastaId).getFiles();
  var melhor = null;
  while(it.hasNext()){
    var f = it.next();
    var nome = f.getName().toLowerCase();
    var mime = f.getMimeType();
    var serve = mime === MimeType.GOOGLE_SHEETS
      || /\.(xlsx|xls|csv)$/.test(nome)
      || mime === MimeType.MICROSOFT_EXCEL
      || mime === MimeType.MICROSOFT_EXCEL_LEGACY
      || mime === MimeType.CSV;
    if(!serve) continue;
    if(!melhor || f.getLastUpdated() > melhor.getLastUpdated()) melhor = f;
  }
  return melhor;
}

/** Devolve [{nome, valores}] de todas as abas, convertendo o Excel se preciso. */
function rkAbasDoArquivo(arq){
  if(arq.getMimeType() === MimeType.CSV){
    return [{ nome: arq.getName(), valores: Utilities.parseCsv(arq.getBlob().getDataAsString('UTF-8')) }];
  }

  var id = arq.getId(), temporario = false;

  if(arq.getMimeType() !== MimeType.GOOGLE_SHEETS){
    var tmp;
    try{
      tmp = Drive.Files.create(
        { name: 'tmp_ranking_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS },
        arq.getBlob()
      );
    }catch(e){
      throw new Error('Não consegui converter o Excel. Ative o serviço avançado "Drive API" (v3) no editor do Apps Script, ou salve o relatório na pasta como planilha do Google. Detalhe: ' + e);
    }
    id = tmp.id; temporario = true;
  }

  try{
    return SpreadsheetApp.openById(id).getSheets().map(function(sh){
      return { nome: sh.getName(), valores: sh.getDataRange().getValues() };
    });
  } finally {
    if(temporario){ try{ DriveApp.getFileById(id).setTrashed(true); }catch(e){} }
  }
}

/* ==========================================================
   COMPETENCIA
   ========================================================== */
var RK_MESES = ['JANEIRO','FEVEREIRO','MARCO','ABRIL','MAIO','JUNHO',
                'JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
var RK_MESES_CURTO = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];

/** Ano de referencia quando a aba traz so o nome do mes. */
function rkAnoBase(arq){
  var m = /(20\d{2})/.exec(String(arq.getName()));
  if(m) return m[1];
  return Utilities.formatDate(arq.getLastUpdated(), 'America/Sao_Paulo', 'yyyy');
}

/** "AGOSTO 2026", "AGO/26", "08-2026", "2026-08" -> "2026-08". Null se nao achar mes. */
function rkCompetenciaDaAba(nomeAba, anoBase){
  var t = rkNorm(nomeAba);
  if(!t) return null;

  var m;

  // 2026-08 / 2026-8
  m = /(20\d{2})[-_.\/ ]?(1[0-2]|0?[1-9])(?!\d)/.exec(t);
  if(m) return m[1] + '-' + rkMes2(m[2]);

  // 08-2026 / 8-2026 / 08/26
  m = /(?:^|[^\d])(1[0-2]|0?[1-9])[-_.\/](20\d{2}|\d{2})(?!\d)/.exec(t);
  if(m) return rkAno4(m[2]) + '-' + rkMes2(m[1]);

  // nome do mes por extenso ou abreviado, com ano opcional
  for(var i = 0; i < 12; i++){
    var re  = new RegExp('(?:^|[^A-Z])' + RK_MESES[i] + '(?![A-Z])');
    var reC = new RegExp('(?:^|[^A-Z])' + RK_MESES_CURTO[i] + '(?![A-Z])');
    if(re.test(t) || reC.test(t)){
      var ano = /(20\d{2})/.exec(t);
      if(ano) return ano[1] + '-' + ('0' + (i+1)).slice(-2);
      var resto = t.replace(RK_MESES[i], ' ').replace(RK_MESES_CURTO[i], ' ');
      var a2 = /(?:^|[^\d])(\d{2})(?!\d)/.exec(resto);
      return (a2 ? rkAno4(a2[1]) : anoBase) + '-' + ('0' + (i+1)).slice(-2);
    }
  }
  return null;
}

function rkAno4(a){
  a = String(a);
  return a.length === 4 ? a : '20' + a;
}

function rkMes2(m){
  return ('0' + String(m)).slice(-2);
}

/** Fallback para arquivo de aba unica: procura o mes no nome do arquivo. */
function rkCompetenciaDoArquivo(arq){
  var c = rkCompetenciaDaAba(arq.getName(), rkAnoBase(arq));
  if(c) return c;
  return Utilities.formatDate(arq.getLastUpdated(), 'America/Sao_Paulo', 'yyyy-MM');
}

/* ==========================================================
   LEITURA DO RELATORIO DE CONSUMO
   ========================================================== */
function rkNorm(s){
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/\s+/g, ' ').trim();
}

function rkNum(v){
  if(typeof v === 'number') return v;
  if(v instanceof Date) return 0;
  var s = String(v).trim();
  if(!s) return 0;
  if(/,\d{1,3}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s.replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function rkLinhasDaGrade(grade){
  if(!grade || !grade.length) return [];

  var iCab = -1;
  for(var i = 0; i < Math.min(15, grade.length); i++){
    var linha = grade[i].map(rkNorm);
    var temMot = linha.some(function(c){ return c.indexOf('MOTORISTA') >= 0 || c.indexOf('CONDUTOR') >= 0; });
    var temCon = linha.some(function(c){ return c.indexOf('CONSUMO') >= 0 || c.indexOf('LITROS') >= 0; });
    if(temMot && temCon){ iCab = i; break; }
  }
  if(iCab < 0) throw new Error('cabeçalho não encontrado');

  var cab = grade[iCab].map(rkNorm);
  var idx = {};
  for(var j = 0; j < cab.length; j++){
    var c = cab[j];
    if(!c) continue;
    if(idx.motorista == null && (c.indexOf('MOTORISTA') >= 0 || c.indexOf('CONDUTOR') >= 0)) { idx.motorista = j; continue; }
    if(idx.modelo    == null && c.indexOf('MODELO') >= 0)   { idx.modelo = j; continue; }
    if(idx.veiculo   == null && (c.indexOf('VEICULO') >= 0 || c.indexOf('PLACA') >= 0)) { idx.veiculo = j; continue; }
    if(idx.km        == null && (c.indexOf('DISTANCIA') >= 0 || c === 'KM')) { idx.km = j; continue; }
    if(idx.litros    == null && (c.indexOf('CONSUMO') >= 0 || c.indexOf('LITROS') >= 0)) { idx.litros = j; continue; }
    if(idx.co2       == null && (c.indexOf('CO2') >= 0 || c.indexOf('EMISSAO') >= 0)) { idx.co2 = j; continue; }
    if(idx.faixaVerde == null && c.indexOf('FAIXA VERDE') >= 0 && c.indexOf('ECONOMICA') < 0) { idx.faixaVerde = j; continue; }
  }
  if(idx.motorista == null || idx.km == null || idx.litros == null)
    throw new Error('colunas de motorista, distância ou consumo não localizadas');

  var out = [];
  for(var r = iCab + 1; r < grade.length; r++){
    var L = grade[r];
    var nome = String(L[idx.motorista] == null ? '' : L[idx.motorista]).trim();
    if(!nome) continue;
    if(rkNorm(nome).indexOf('TOTAL') === 0) continue;
    out.push({
      motorista: nome,
      veiculo:   idx.veiculo != null ? String(L[idx.veiculo] || '').trim() : '',
      modelo:    idx.modelo  != null ? String(L[idx.modelo]  || '').trim() : '',
      km:        rkNum(L[idx.km]),
      litros:    rkNum(L[idx.litros]),
      faixaVerde: idx.faixaVerde != null ? rkNum(L[idx.faixaVerde]) : 0,
      co2:       idx.co2 != null ? rkNum(L[idx.co2]) : 0
    });
  }
  return out;
}

/* ==========================================================
   HISTORICO — todos os meses, agregado por motorista + modelo.
   Aplica aqui os mesmos filtros de linha que o app usa, para
   que o calculo do mes e o do ano batam exatamente.
   ========================================================== */
function rkHistorico(){
  var cfg     = rkLerConfig();
  var kmLinha = Number(cfg.kmLinha) || 0;
  var rendMax = Number(cfg.rendMax) || 99;

  var aba = rkAba(RK_ABA_DADOS, RK_CABECALHO);
  var n = aba.getLastRow();
  if(n < 2) return { comps:[], nomes:[], modelos:[], d:[] };
  var vals = aba.getRange(2, 1, n-1, RK_CABECALHO.length).getValues();

  var comps = [], nomes = [], modelos = [];
  var iC = {}, iN = {}, iM = {}, acc = {};

  function idxDe(valor, lista, mapa, chave){
    if(mapa[chave] == null){ mapa[chave] = lista.length; lista.push(valor); }
    return mapa[chave];
  }

  for(var i = 0; i < vals.length; i++){
    var v = vals[i];
    var comp = String(v[0] || '').trim();
    var nome = String(v[1] || '').trim();
    if(!comp || !nome) continue;

    var km = Number(v[4]) || 0, lt = Number(v[5]) || 0;
    if(km <= 0 || lt <= 0) continue;
    if(km < kmLinha) continue;
    var rend = km / lt;
    if(rend > rendMax || rend < 0.3) continue;

    var modelo = String(v[3] || '').trim();
    var ci = idxDe(comp,   comps,   iC, comp);
    var ni = idxDe(nome,   nomes,   iN, rkNorm(nome));
    var mi = idxDe(modelo, modelos, iM, rkNorm(modelo));

    var k = ci + '|' + ni + '|' + mi;
    var a = acc[k] || (acc[k] = [ci, ni, mi, 0, 0, 0]);
    a[3] += km;
    a[4] += lt;
    a[5] += (Number(v[6]) || 0) * km;
  }

  var d = [];
  for(var k2 in acc){
    var a2 = acc[k2];
    d.push([a2[0], a2[1], a2[2], Math.round(a2[3]*10)/10, Math.round(a2[4]*100)/100, Math.round(a2[5])]);
  }
  return { comps: comps, nomes: nomes, modelos: modelos, d: d };
}

/* ==========================================================
   TESTE — rode pelo editor antes de implantar
   ========================================================== */
function rkTestarPasta(){
  var arq;
  try{ arq = rkFonte(); }
  catch(e){ Logger.log('ERRO: ' + e.message); return; }
  Logger.log('Arquivo: ' + arq.getName() + '  |  ' + arq.getLastUpdated());

  var anoBase = rkAnoBase(arq);
  var abas = rkAbasDoArquivo(arq);
  Logger.log('Abas encontradas: ' + abas.length + '  |  ano de referência: ' + anoBase);

  abas.forEach(function(ab){
    var comp = rkCompetenciaDaAba(ab.nome, anoBase);
    if(!comp){ Logger.log('  [ ' + ab.nome + ' ] -> mês NÃO identificado, será ignorada'); return; }
    var linhas = [];
    try{ linhas = rkLinhasDaGrade(ab.valores); }
    catch(e){ Logger.log('  [ ' + ab.nome + ' ] -> ' + comp + ' -> erro: ' + e.message); return; }
    var km = 0, mot = {};
    linhas.forEach(function(l){ km += l.km; mot[rkNorm(l.motorista)] = 1; });
    Logger.log('  [ ' + ab.nome + ' ] -> ' + comp + ' -> ' + linhas.length + ' linhas, '
             + Object.keys(mot).length + ' motoristas, ' + Math.round(km) + ' km');
  });
}
