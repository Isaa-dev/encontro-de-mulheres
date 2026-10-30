// 1o Encontro Internacional de Mulheres - inscricoes
// Grava as inscricoes do site na planilha e para em 150.

var PLANILHA_ID = '1C3CvtTn1fQJr_-zUMaoD59otWQl-TfxFGubUAxTsErQ';
var LIMITE = 150;
var LIMITE_SEXTA = 50;
var CABECALHO = [
  'N\u00ba', 'Data/Hora', 'Nome completo', 'Tem filhos?',
  'Quantos filhos ir\u00e3o', 'Grupo do sorteio'
];

function json_(obj) {
  var saida = ContentService.createTextOutput(JSON.stringify(obj));
  return saida.setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return json_(getStatus());
}

function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    return json_(registrar(dados));
  } catch (err) {
    return json_({ ok: false, mensagem: 'Erro. Tente novamente.' });
  }
}

function aba_() {
  return SpreadsheetApp.openById(PLANILHA_ID).getSheets()[0];
}

function total_(sh) {
  return Math.max(0, sh.getLastRow() - 1);
}

function getStatus() {
  return { total: total_(aba_()), limite: LIMITE };
}

// nome sem acentos, minusculo e sem espacos extras
function chave_(nome) {
  return String(nome || '').normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

// deixa o cabecalho da planilha certo (sem a coluna de telefone)
function cabecalho_(sh) {
  var atual = sh.getRange(1, 1, 1, 7).getValues()[0];
  if (atual[3] !== CABECALHO[3] || atual[6] !== '') {
    sh.getRange(1, 1, 1, 7).clearContent();
    sh.getRange(1, 1, 1, 6).setValues([CABECALHO]);
  }
}

function registrar(d) {
  d = d || {};
  var nome = String(d.nome || '').trim().replace(/\s+/g, ' ');
  var temFilhos = (d.temFilhos === 'Sim') ? 'Sim' : 'N\u00e3o';
  var qtd = 0;
  if (temFilhos === 'Sim') {
    qtd = parseInt(d.qtdFilhos, 10) || 0;
    qtd = Math.max(0, Math.min(10, qtd));
  }

  if (nome.split(' ').length < 2 || nome.length > 120) {
    return { ok: false, mensagem: 'Informe seu nome completo.' };
  }
  if (/^[=+@-]/.test(nome)) {
    nome = ' ' + nome;
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = aba_();
    cabecalho_(sh);
    var total = total_(sh);
    if (total >= LIMITE) {
      return { ok: false, motivo: 'esgotado' };
    }

    if (total > 0) {
      var nomes = sh.getRange(2, 3, total, 1).getValues();
      for (var i = 0; i < nomes.length; i++) {
        if (chave_(nomes[i][0]) === chave_(nome)) {
          return { ok: false, mensagem: 'Este nome j\u00e1 est\u00e1 inscrito.' };
        }
      }
    }

    var numero = total + 1;
    var grupo = 'Sabado 07/11 (inscricoes 51 a 150)';
    if (numero <= LIMITE_SEXTA) {
      grupo = 'Sexta 06/11 (inscricoes 1 a 50)';
    }

    var linha = sh.getRange(numero + 1, 1, 1, 6);
    linha.setNumberFormats([[
      '0', 'dd/MM/yyyy HH:mm:ss', '@', '@', '0', '@'
    ]]);
    linha.setValues([[
      numero, new Date(), nome, temFilhos, qtd, grupo
    ]]);
    SpreadsheetApp.flush();
    return { ok: true, numero: numero, grupo: grupo };
  } finally {
    lock.releaseLock();
  }
}
