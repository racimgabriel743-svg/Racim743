// Leitura da planilha "MASTERS ESTOCADO E CROSS" (abas FORT e COMPER, colunas C a I).
// Equivalente ao getMasters() do Apps Script, usando a mesma conta de serviço das outras áreas.
const { JWT } = require('google-auth-library');

const ABA_FORT = 'FORT';
const ABA_COMPER = 'COMPER';
const LINHA_INICIAL = 3; // dados começam na linha 3 (cabeçalho na linha 2)
const INTERVALO = `C${LINHA_INICIAL}:I`; // C, D, E, F, G, H, I

function parsePrivateKey(raw) {
  let key = raw.replace(/^["']+|["']+$/g, '').trim();
  key = key.replace(/\\n/g, '\n');
  const BEGIN = '-----BEGIN PRIVATE KEY-----';
  const END = '-----END PRIVATE KEY-----';
  const bi = key.indexOf(BEGIN);
  const ei = key.indexOf(END);
  if (bi !== -1 && ei !== -1) {
    const body = key.substring(bi + BEGIN.length, ei).replace(/\s+/g, '');
    const lines = body.match(/.{1,64}/g) || [];
    key = `${BEGIN}\n${lines.join('\n')}\n${END}\n`;
  }
  return key;
}

let auth = null;
function getAuth() {
  if (auth) return auth;
  const clientEmail = process.env.GOOGLE_SHEETS_CLIENT_EMAIL;
  const rawKey = process.env.GOOGLE_SHEETS_PRIVATE_KEY;
  if (!clientEmail) throw new Error('Variável GOOGLE_SHEETS_CLIENT_EMAIL não configurada');
  if (!rawKey) throw new Error('Variável GOOGLE_SHEETS_PRIVATE_KEY não configurada');
  auth = new JWT({
    email: clientEmail,
    key: parsePrivateKey(rawKey),
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
  return auth;
}

// ID da planilha "MASTERS ESTOCADO E CROSS" (a variável GOOGLE_SHEETS_ID_MASTERS, se existir, tem prioridade)
const ID_PADRAO_MASTERS = '16NEeuloLDu2Kkd7Tr55Q2xVT--8NuymAgfoyiW0cABg';

async function lerAba(nomeAba) {
  const id = (process.env.GOOGLE_SHEETS_ID_MASTERS || '').trim() || ID_PADRAO_MASTERS;

  const range = encodeURIComponent(`'${nomeAba}'!${INTERVALO}`);
  let valores;
  try {
    const r = await getAuth().request({
      url: `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${range}?valueRenderOption=FORMATTED_VALUE`,
    });
    valores = (r.data && r.data.values) || [];
  } catch (e) {
    const status = e.response && e.response.status;
    // aba inexistente -> lista vazia (mesmo comportamento do Apps Script)
    if (status === 400 && /Unable to parse range/i.test(JSON.stringify(e.response.data || ''))) return [];
    throw new Error(`Falha ao ler a aba ${nomeAba}: ${e.message}`);
  }

  const linhas = [];
  valores.forEach((l) => {
    const c = (i) => (l[i] == null ? '' : String(l[i]));
    const loja = c(1).trim();
    const master = c(2).trim();
    if (!loja && !master) return; // ignora linhas vazias
    const m = loja.match(/\d+/);
    linhas.push({
      data: c(0),
      loja,
      numLoja: m ? m[0] : '',
      master,
      etq: c(3),
      status: c(4),
      tipo: c(5),
      local: c(6),
    });
  });
  return linhas;
}

async function getMasters() {
  const [fort, comper] = await Promise.all([lerAba(ABA_FORT), lerAba(ABA_COMPER)]);
  const atualizadoEm = new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' }).replace(',', '');
  return { fort, comper, atualizadoEm };
}

module.exports = { getMasters };
