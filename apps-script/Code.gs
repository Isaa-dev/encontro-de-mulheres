/**
 * 1º Encontro Internacional de Mulheres – inscrições
 * Recebe as inscrições do site (GitHub Pages), grava na planilha e trava em 150 inscritas.
 *
 * Como publicar: Implantar › Nova implantação › App da Web
 *   Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 * Copie o link (termina em /exec) e cole em API_URL no index.html.
 */

const PLANILHA_ID = '1C3CvtTn1fQJr_-zUMaoD59otWQl-TfxFGubUAxTsErQ';
const LIMITE = 150;
const LIMITE_SEXTA = 50; // inscrições 1–50 → sorteio de sexta; 51–150 → sorteio de sábado

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET ?acao=status → quantas inscritas já existem
function doGet(e) {
  return json_(getStatus());
}

// POST com o JSON da inscrição
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);
    return json_(registrar(dados));
  } catch (err) {
    return json_({ ok: false, mensagem: 'Não foi possível concluir. Tente novamente.' });
  }
}

function aba_() {
  return SpreadsheetApp.openById(PLANILHA_ID).getSheets()[0];
}

function total_(sh) {
  return Math.max(0, sh.getLastRow() - 1); // linha 1 = cabeçalho
}

function getStatus() {
  return { total: total_(aba_()), limite: LIMITE };
}

function registrar(d) {
  const nome = String(d && d.nome || '').trim().replace(/\s+/g, ' ');
  const tel = String(d && d.telefone || '').trim();
  const temFilhos = d && d.temFilhos === 'Sim' ? 'Sim' : 'Não';
  const qtd = temFilhos === 'Sim' ? Math.max(0, Math.min(10, parseInt(d.qtdFilhos, 10) || 0)) : 0;

  if (nome.split(' ').length < 2 || nome.length > 120) return { ok: false, mensagem: 'Informe seu nome completo.' };
  if (tel.replace(/\D/g, '').length < 10) return { ok: false, mensagem: 'Informe um telefone com DDD.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // evita duas pessoas pegarem a mesma vaga ao mesmo tempo
  try {
    const sh = aba_();
    const total = total_(sh);
    if (total >= LIMITE) return { ok: false, motivo: 'esgotado' };

    // bloqueia inscrição repetida (mesmo telefone)
    if (total > 0) {
      const tels = sh.getRange(2, 4, total, 1).getValues().map(r => String(r[0]).replace(/\D/g, ''));
      if (tels.indexOf(tel.replace(/\D/g, '')) !== -1) {
        return { ok: false, mensagem: 'Este telefone já está inscrito.' };
      }
    }

    const numero = total + 1;
    const grupo = numero <= LIMITE_SEXTA ? 'Sexta 06/11 (inscrições 1–50)' : 'Sábado 07/11 (inscrições 51–150)';
    const linha = sh.getRange(numero + 1, 1, 1, 7);
    linha.setNumberFormats([['0', 'dd/MM/yyyy HH:mm:ss', '@', '@', '@', '0', '@']]);
    // o "'" impede que a planilha transforme textos em fórmulas
    linha.setValues([[numero, new Date(), nome.replace(/^[=+\-@]/, "'$&"), tel, temFilhos, qtd, grupo]]);
    SpreadsheetApp.flush();
    return { ok: true, numero: numero, grupo: grupo };
  } finally {
    lock.releaseLock();
  }
}
