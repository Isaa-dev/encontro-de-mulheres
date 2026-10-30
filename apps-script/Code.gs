// 1o Encontro Internacional de Mulheres - inscricoes
// Grava as inscricoes do site na planilha e para em 150.

var PLANILHA_ID = '1C3CvtTn1fQJr_-zUMaoD59otWQl-TfxFGubUAxTsErQ';
var LIMITE = 150;
var LIMITE_SEXTA = 50;

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

function soNumeros_(txt) {
  return String(txt || '').replace(/\D/g, '');
}

function registrar(d) {
  d = d || {};
  var nome = String(d.nome || '').trim().replace(/\s+/g, ' ');
  var tel = String(d.telefone || '').trim();
  var temFilhos = (d.temFilhos === 'Sim') ? 'Sim' : 'Nao';
  var qtd = 0;
  if (temFilhos === 'Sim') {
    qtd = parseInt(d.qtdFilhos, 10) || 0;
    qtd = Math.max(0, Math.min(10, qtd));
  }
  if (temFilhos === 'Nao') {
    temFilhos = 'N\u00e3o';
  }

  if (nome.split(' ').length < 2 || nome.length > 120) {
    return { ok: false, mensagem: 'Informe seu nome completo.' };
  }
  if (soNumeros_(tel).length < 10) {
    return { ok: false, mensagem: 'Informe um telefone com DDD.' };
  }
  if (/^[=+@-]/.test(nome)) {
    nome = ' ' + nome;
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = aba_();
    var total = total_(sh);
    if (total >= LIMITE) {
      return { ok: false, motivo: 'esgotado' };
    }

    if (total > 0) {
      var tels = sh.getRange(2, 4, total, 1).getValues();
      for (var i = 0; i < tels.length; i++) {
        if (soNumeros_(tels[i][0]) === soNumeros_(tel)) {
          return { ok: false, mensagem: 'Este telefone j\u00e1 est\u00e1 inscrito.' };
        }
      }
    }

    var numero = total + 1;
    var grupo = 'Sabado 07/11 (inscricoes 51 a 150)';
    if (numero <= LIMITE_SEXTA) {
      grupo = 'Sexta 06/11 (inscricoes 1 a 50)';
    }

    var linha = sh.getRange(numero + 1, 1, 1, 7);
    linha.setNumberFormats([[
      '0', 'dd/MM/yyyy HH:mm:ss', '@', '@', '@', '0', '@'
    ]]);
    linha.setValues([[
      numero, new Date(), nome, tel, temFilhos, qtd, grupo
    ]]);
    SpreadsheetApp.flush();
    return { ok: true, numero: numero, grupo: grupo };
  } finally {
    lock.releaseLock();
  }
}
