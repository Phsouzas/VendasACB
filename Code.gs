// ===== CONFIGURAÇÃO =====
const VERSAO = 'v6';
const PLANILHA_ID = '1zHydk2TfcPmjaX3AszUKqX38U7G0RHa-GmzrkF5vdeQ';
const ABA = 'Respostas ao formulário 1';
const COL_EMAIL_LETRA = 'S';   // E-mail do vendedor: liga a venda ao usuário logado
const COL_DATA = '#F';         // Data da Venda: usada no filtro de data
// Colunas exibidas ao usuário comum: [rótulo exibido, #LETRA da coluna]
const COLS_USUARIO = [
  ['ID IXC', '#J'],
  ['Nome completo', '#C'],
  ['Plano contratado', '#G'],
  ['Cidade', '#E'],
  ['Data da venda', '#F'],
  ['Data 1º pagamento', '#L'],
  ['Data 3º pagamento', '#M']
];

const planilha = () => PLANILHA_ID
  ? SpreadsheetApp.openById(PLANILHA_ID)
  : SpreadsheetApp.getActive();

const norm = s => String(s).normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
  .replace(/\s+/g, ' ').trim().toLowerCase();

const letraParaIndice = L => L.toUpperCase().split('')
  .reduce((a, c) => a * 26 + c.charCodeAt(0) - 64, 0) - 1;

const indiceParaLetra = i => {
  let n = i + 1, s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
};

function abaPorNome(nome) {
  const abas = planilha().getSheets();
  const achada = abas.find(a => norm(a.getName()) === norm(nome));
  if (!achada) {
    throw new Error('Aba "' + nome + '" não encontrada. Abas existentes: ' +
      abas.map(a => '"' + a.getName() + '"').join(', '));
  }
  return achada;
}

// ===== API =====
// Abra a URL /exec no navegador: deve mostrar a versão do código publicado.
function doGet() {
  return ContentService.createTextOutput('Portal de vendas ' + VERSAO);
}

function doPost(e) {
  let r;
  try {
    const p = JSON.parse(e.postData.contents);
    r = p.acao === 'login' ? login(p.email, p.senha) : vendas(p);
  } catch (err) {
    r = { ok: false, erro: 'Erro: ' + err.message };
  }
  r.versao = VERSAO;
  return ContentService.createTextOutput(JSON.stringify(r))
    .setMimeType(ContentService.MimeType.JSON);
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
    const [em, hash, salt, perfil, status] = u[i];
    if (norm(em) === norm(email) && norm(status) === 'ativo' && hashSenha(senha, salt) === hash) {
      const token = Utilities.getUuid();
      cache.put('s_' + token, JSON.stringify({ email: em, perfil: norm(perfil) }), 21600); // 6h
      cache.remove(chave);
      return { ok: true, token, perfil: norm(perfil) };
    }
  }
  cache.put(chave, String(tent + 1), 900);
  return { ok: false, erro: 'Email ou senha inválidos.' };
}

function vendas(p) {
  const s = JSON.parse(CacheService.getScriptCache().get('s_' + p.token) || 'null');
  if (!s) return { ok: false, erro: 'Sessão expirada.', expirada: true };
  return consultar(s, p.ini, p.fim);
}

// Núcleo da consulta. Falha FECHADA: qualquer dúvida retorna erro, nunca dados.
function consultar(s, ini, fim) {
  const master = s.perfil === 'master';
  if (!master && (s.perfil !== 'user' || !norm(s.email || ''))) {
    return { ok: false, erro: 'Perfil de acesso inválido.' };
  }

  const v = abaPorNome(ABA).getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const cab = v[0].map(norm);
  const achar = h => h.charAt(0) === '#' ? letraParaIndice(h.slice(1)) : cab.indexOf(norm(h));
  const listaCab = () => v[0].map((x, i) =>
    indiceParaLetra(i) + '="' + String(x).replace(/\s+/g, ' ').trim() + '"').join(' | ');

  const iMail = letraParaIndice(COL_EMAIL_LETRA);
  const iData = achar(COL_DATA);
  const faltando = iData < 0 ? [COL_DATA] : [];
  const cols = master
    ? v[0].map((h, i) => [String(h).replace(/\s+/g, ' ').trim(), i])
    : COLS_USUARIO.map(([rot, h]) => {
        const i = achar(h);
        if (i < 0) faltando.push(h);
        return [rot, i];
      });
  if (faltando.length) {
    throw new Error('Colunas não encontradas: ' + faltando.join(', ') + '. Cabeçalhos da aba: ' + listaCab());
  }
  if (!master && cols.length !== COLS_USUARIO.length) {
    return { ok: false, erro: 'Configuração de colunas inválida.' };
  }

  const fmt = x => x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : String(x);
  // converte data (Date ou texto dd/mm/aaaa) para aaaa-mm-dd, usada no filtro
  const dataISO = x => {
    if (x instanceof Date) return Utilities.formatDate(x, tz, 'yyyy-MM-dd');
    const m = String(x).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : '';
  };

  const linhas = [];
  v.slice(1).forEach(l => {
    if (l.every(c => c === '')) return;
    // usuário comum: só as vendas cujo e-mail (coluna S) é o do login
    if (!master && norm(l[iMail] || '') !== norm(s.email)) return;
    const d = dataISO(l[iData]);
    if (ini && d < ini) return;
    if (fim && d > fim) return;
    linhas.push(cols.map(([, i]) => fmt(l[i])));
  });
  return { ok: true, perfil: s.perfil, colunas: cols.map(c => c[0]), linhas };
}

// ===== TESTE NO EDITOR (não depende da implantação) =====
// Troque pelo email do usuário comum, selecione "testeUsuario", Executar,
// e veja o resultado em "Registro de execução".
function testeUsuario() {
  const r = consultar({ email: 'email-do-usuario@exemplo.com', perfil: 'user' }, '', '');
  Logger.log(JSON.stringify({
    ok: r.ok, erro: r.erro, colunas: r.colunas, total: r.linhas && r.linhas.length
  }));
}

// ===== ADMINISTRAÇÃO =====
function criarUsuario(email, senha, perfil) {
  const ss = planilha();
  let aba = ss.getSheets().find(a => norm(a.getName()) === 'usuarios');
  if (!aba) {
    aba = ss.insertSheet('Usuarios');
    aba.appendRow(['email', 'hash', 'salt', 'perfil', 'status']);
  }
  const salt = Utilities.getUuid();
  aba.appendRow([email.trim(), hashSenha(senha, salt), salt, perfil || 'user', 'ativo']);
}

// Para criar um NOVO usuário: preencha abaixo, selecione "criarNovoUsuario"
// no menu de funções, clique em Executar (uma vez) e depois APAGUE a senha daqui.
// ATENÇÃO ao perfil: 'user' = só as próprias vendas / 'master' = vê tudo.
function criarNovoUsuario() {
  criarUsuario('email@exemplo.com', 'senha-aqui', 'user');
}
