// ===== CONFIGURAÇÃO =====
const VERSAO = 'v7';
const PLANILHA_ID = '1zHydk2TfcPmjaX3AszUKqX38U7G0RHa-GmzrkF5vdeQ';
const ABA = 'Respostas ao formulário 1';
const COL_EMAIL_LETRA = 'S'; // e-mail do vendedor (liga a venda ao usuário)
const COL_DATA = 'F';        // data da venda
// Colunas que o master pode liberar para cada usuário: letra -> rótulo
const CATALOGO = {
  J: 'ID IXC', C: 'Nome completo', G: 'Plano contratado', E: 'Cidade',
  F: 'Data da venda', L: 'Data 1º pagamento', M: 'Data 3º pagamento',
  B: 'Vendedor', H: 'Valor', I: 'Prêmio', K: 'Data da instalação', R: 'Status'
};
const PADRAO = 'J,C,G,E,F,L,M'; // usado quando o usuário não tem colunas definidas

const NEG = { ok: false, erro: 'Sessão expirada.', expirada: true };
const PROIB = { ok: false, erro: 'Acesso negado.' };

const planilha = () => SpreadsheetApp.openById(PLANILHA_ID);
const norm = s => String(s).normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
  .replace(/\s+/g, ' ').trim().toLowerCase();
const L2I = L => L.toUpperCase().split('').reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0) - 1;
const num = x => typeof x === 'number' ? x
  : Number(String(x).replace(/[^\d,-]/g, '').replace(',', '.')) || 0;
const dataISO = (x, tz) => {
  if (x instanceof Date) return Utilities.formatDate(x, tz, 'yyyy-MM-dd');
  const m = String(x).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : '';
};

function abaPorNome(nome) {
  const abas = planilha().getSheets();
  const a = abas.find(x => norm(x.getName()) === norm(nome));
  if (!a) throw new Error('Aba "' + nome + '" não encontrada. Existentes: ' + abas.map(x => x.getName()).join(', '));
  return a;
}

const sessao = t => JSON.parse(CacheService.getScriptCache().get('s_' + t) || 'null');

// ===== API =====
function doGet() { return ContentService.createTextOutput('Portal de vendas ' + VERSAO); }

function doPost(e) {
  let r;
  try {
    const p = JSON.parse(e.postData.contents);
    const acoes = {
      login: () => login(p.email, p.senha), vendas: () => vendas(p),
      dashboard: () => dashboard(p), usuarios: () => usuarios(p),
      salvarUsuario: () => salvarUsuario(p)
    };
    r = acoes[p.acao] ? acoes[p.acao]() : { ok: false, erro: 'Ação inválida.' };
  } catch (err) {
    r = { ok: false, erro: 'Erro: ' + err.message };
  }
  r.versao = VERSAO;
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function hashSenha(senha, salt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + senha)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function login(email, senha) {
  const cache = CacheService.getScriptCache();
  const chave = 'tent_' + norm(email);
  const tent = Number(cache.get(chave) || 0);
  if (tent >= 5) return { ok: false, erro: 'Muitas tentativas. Aguarde 15 minutos.' };
  const u = abaPorNome('Usuarios').getDataRange().getValues();
  for (let i = 1; i < u.length; i++) {
    const [em, hash, salt, perfil, status, colunas] = u[i];
    if (norm(em) === norm(email) && norm(status) === 'ativo' && hashSenha(senha, salt) === hash) {
      const token = Utilities.getUuid();
      cache.put('s_' + token, JSON.stringify({ email: em, perfil: norm(perfil), cols: String(colunas || '') }), 21600);
      cache.remove(chave);
      return { ok: true, token, perfil: norm(perfil) };
    }
  }
  cache.put(chave, String(tent + 1), 900);
  return { ok: false, erro: 'Email ou senha inválidos.' };
}

function vendas(p) {
  const s = sessao(p.token);
  return s ? consultar(s, p.ini, p.fim) : NEG;
}

// Falha FECHADA: qualquer dúvida retorna erro, nunca dados.
function consultar(s, ini, fim) {
  const master = s.perfil === 'master';
  const v = abaPorNome(ABA).getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  let cols;
  if (master) {
    cols = v[0].map((h, i) => [String(h).replace(/\s+/g, ' ').trim(), i]);
  } else {
    const ls = String(s.cols || PADRAO).split(',').map(x => x.trim().toUpperCase()).filter(x => CATALOGO[x]);
    if (s.perfil !== 'user' || !norm(s.email || '') || !ls.length) return { ok: false, erro: 'Perfil de acesso inválido.' };
    cols = ls.map(l => [CATALOGO[l], L2I(l)]);
  }
  const iMail = L2I(COL_EMAIL_LETRA), iData = L2I(COL_DATA);
  const fmt = x => x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : String(x);
  const linhas = [];
  v.slice(1).forEach(l => {
    if (l.every(c => c === '')) return;
    if (!master && norm(l[iMail] || '') !== norm(s.email)) return;
    const d = dataISO(l[iData], tz);
    if (ini && d < ini) return;
    if (fim && d > fim) return;
    linhas.push(cols.map(([, i]) => fmt(l[i])));
  });
  return { ok: true, perfil: s.perfil, colunas: cols.map(c => c[0]), linhas };
}

// ===== SOMENTE MASTER =====
function dashboard(p) {
  const s = sessao(p.token);
  if (!s) return NEG;
  if (s.perfil !== 'master') return PROIB;
  const tz = Session.getScriptTimeZone();
  const R = { total: 0, valor: 0, premio: 0, vendedor: {}, plano: {}, cidade: {}, status: {}, mes: {} };
  const cont = (o, k) => { k = String(k || '').trim() || '—'; o[k] = (o[k] || 0) + 1; };
  abaPorNome(ABA).getDataRange().getValues().slice(1).forEach(l => {
    if (l.every(c => c === '')) return;
    const d = dataISO(l[5], tz);
    if ((p.ini && d < p.ini) || (p.fim && d > p.fim)) return;
    R.total++; R.valor += num(l[7]); R.premio += num(l[8]);
    cont(R.vendedor, l[1]); cont(R.plano, l[6]); cont(R.cidade, l[4]);
    cont(R.status, l[17]); cont(R.mes, d.slice(0, 7));
  });
  return Object.assign({ ok: true }, R);
}

function usuarios(p) {
  const s = sessao(p.token);
  if (!s) return NEG;
  if (s.perfil !== 'master') return PROIB;
  const u = abaPorNome('Usuarios').getDataRange().getValues().slice(1).filter(r => r[0]);
  return {
    ok: true, catalogo: CATALOGO, padrao: PADRAO,
    usuarios: u.map(r => ({ email: r[0], perfil: norm(r[3]), status: norm(r[4]), colunas: String(r[5] || '') }))
  };
}

function salvarUsuario(p) {
  const s = sessao(p.token);
  if (!s) return NEG;
  if (s.perfil !== 'master') return PROIB;
  const email = String(p.email || '').trim();
  const perfil = p.perfil === 'master' ? 'master' : 'user';
  const status = p.status === 'inativo' ? 'inativo' : 'ativo';
  const cols = (p.colunas || []).filter(c => CATALOGO[c]).join(',');
  const senha = String(p.senha || '');
  if (!/^\S+@\S+\.\S+$/.test(email)) return { ok: false, erro: 'Email inválido.' };
  if (perfil === 'user' && !cols) return { ok: false, erro: 'Escolha ao menos uma coluna para este usuário.' };
  if (senha && senha.length < 6) return { ok: false, erro: 'A senha precisa ter 6 ou mais caracteres.' };

  const aba = abaPorNome('Usuarios');
  const d = aba.getDataRange().getValues();
  if (!d[0][5]) aba.getRange(1, 6).setValue('colunas');
  const i = d.findIndex((r, n) => n > 0 && norm(r[0]) === norm(email));
  if (i < 0) {
    if (!senha) return { ok: false, erro: 'Defina uma senha para o novo usuário.' };
    const salt = Utilities.getUuid();
    aba.appendRow([email, hashSenha(senha, salt), salt, perfil, status, cols]);
  } else {
    if (norm(email) === norm(s.email) && (perfil !== 'master' || status !== 'ativo')) {
      return { ok: false, erro: 'Você não pode rebaixar nem desativar o seu próprio acesso.' };
    }
    aba.getRange(i + 1, 4, 1, 3).setValues([[perfil, status, cols]]);
    if (senha) {
      const salt = Utilities.getUuid();
      aba.getRange(i + 1, 2, 1, 2).setValues([[hashSenha(senha, salt), salt]]);
    }
  }
  return { ok: true };
}
