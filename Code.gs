// ===== CONFIGURAÇÃO (ajuste os nomes conforme os cabeçalhos da planilha) =====
const ABA = 'Form_Responses';
const COL_EMAIL = 'Endereço de e-mail'; // liga a venda ao usuário logado
const COL_DATA = 'DATA DA VENDA';       // coluna usada no filtro de data
// [rótulo exibido ao usuário, cabeçalho na planilha]
const COLS_USUARIO = [
  ['ID IXC', 'ID IXC'],
  ['Cliente', 'NOME COMPLETO'],
  ['Plano', 'PLANO'],
  ['Cidade', 'CIDADE'],                       // ajuste se o cabeçalho for outro
  ['Data da venda', 'DATA DA VENDA'],
  ['1º pagamento', 'Data 1º Pagamento'],
  ['3º pagamento', 'Data 3º Pagamento'],
  ['Pagamento da venda', 'Data 1º Premio']    // ajuste se for outra coluna
];

// ===== API =====
function doPost(e) {
  let r;
  try {
    const p = JSON.parse(e.postData.contents);
    r = p.acao === 'login' ? login(p.email, p.senha) : vendas(p);
  } catch (err) {
    r = { ok: false, erro: 'Erro: ' + err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(r))
    .setMimeType(ContentService.MimeType.JSON);
}

const norm = s => String(s).normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
  .replace(/\s+/g, ' ').trim().toLowerCase();

function hashSenha(senha, salt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + senha)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function login(email, senha) {
  const cache = CacheService.getScriptCache();
  const chave = 'tent_' + norm(email);
  const tent = Number(cache.get(chave) || 0);
  if (tent >= 5) return { ok: false, erro: 'Muitas tentativas. Aguarde 15 minutos.' };

  const u = SpreadsheetApp.getActive().getSheetByName('Usuarios').getDataRange().getValues();
  for (let i = 1; i < u.length; i++) {
    const [em, hash, salt, perfil, status] = u[i];
    if (norm(em) === norm(email) && status === 'ativo' && hashSenha(senha, salt) === hash) {
      const token = Utilities.getUuid();
      cache.put('s_' + token, JSON.stringify({ email: em, perfil }), 21600); // 6h
      cache.remove(chave);
      return { ok: true, token, perfil };
    }
  }
  cache.put(chave, String(tent + 1), 900);
  return { ok: false, erro: 'Email ou senha inválidos.' };
}

function vendas(p) {
  const s = JSON.parse(CacheService.getScriptCache().get('s_' + p.token) || 'null');
  if (!s) return { ok: false, erro: 'Sessão expirada.', expirada: true };

  const v = SpreadsheetApp.getActive().getSheetByName(ABA).getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const cab = v[0].map(norm);
  const idx = h => {
    const i = cab.indexOf(norm(h));
    if (i < 0) throw new Error('Coluna não encontrada: ' + h);
    return i;
  };
  const iMail = idx(COL_EMAIL), iData = idx(COL_DATA);
  const master = s.perfil === 'master';
  const cols = master
    ? v[0].map((h, i) => [String(h).replace(/\s+/g, ' ').trim(), i])
    : COLS_USUARIO.map(([rot, h]) => [rot, idx(h)]);

  const fmt = x => x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : String(x);
  const linhas = [];
  v.slice(1).forEach(l => {
    if (l.every(c => c === '')) return;
    if (!master && norm(l[iMail]) !== norm(s.email)) return; // filtro por usuário no servidor
    const d = l[iData] instanceof Date ? Utilities.formatDate(l[iData], tz, 'yyyy-MM-dd') : '';
    if (p.ini && d < p.ini) return;
    if (p.fim && d > p.fim) return;
    linhas.push(cols.map(([, i]) => fmt(l[i])));
  });
  return { ok: true, perfil: s.perfil, colunas: cols.map(c => c[0]), linhas };
}

// ===== ADMINISTRAÇÃO (execute manualmente no editor do Apps Script) =====
// Exemplo: criarUsuario('joao@email.com', 'senha123', 'user')
//          criarUsuario('paulo@email.com', 'senhaForte', 'master')
function criarUsuario(email, senha, perfil) {
  const ss = SpreadsheetApp.getActive();
  let aba = ss.getSheetByName('Usuarios');
  if (!aba) {
    aba = ss.insertSheet('Usuarios');
    aba.appendRow(['email', 'hash', 'salt', 'perfil', 'status']);
  }
  const salt = Utilities.getUuid();
  aba.appendRow([email.trim(), hashSenha(senha, salt), salt, perfil || 'user', 'ativo']);
}
