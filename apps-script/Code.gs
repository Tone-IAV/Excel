/**
 * Backend em Apps Script para a plataforma gamificada do curso de Excel.
 * Ele utiliza a planilha https://docs.google.com/spreadsheets/d/1DCzIOIcRBaJ3WJVQCOWg6KXyghjRGMJUHekV1fGygJ4
 * como base de dados, persistindo usuários, progresso, ranking e configurações.
 */

const SPREADSHEET_ID = '1DCzIOIcRBaJ3WJVQCOWg6KXyghjRGMJUHekV1fGygJ4';

const SHEETS = {
  USERS: 'Users',
  TOKENS: 'Tokens',
  MODULES: 'Modules',
  QUESTIONS: 'Questions',
  PROGRESS: 'Progress',
  CHECKINS: 'Checkins',
  CONFIG: 'Config',
  EMBEDS: 'Embeds',
  AUDIT: 'AuditLog'
};

const TABLE_HEADERS = {
  Users: ['ID','Nome','Email','SenhaHash','Admin','XP','CriadoEm','AtualizadoEm'],
  Tokens: ['Token','UserID','ExpiresAt','CriadoEm'],
  Modules: ['ModuleID','Ordem','Titulo','Descricao','XP','VideoURL','MaterialURL','Ativo'],
  Questions: ['ModuleID','QuestionID','Tipo','Enunciado','OpcoesJSON','Correta','Peso','MinCaracteres','Feedback'],
  Progress: ['UserID','ModuleID','Score','Done','XP','AnswersJSON','AtualizadoEm'],
  Checkins: ['UserID','Date','XP','RegistradoEm'],
  Config: ['Chave','Valor'],
  Embeds: ['Tipo','URL','AtualizadoPor','AtualizadoEm'],
  AuditLog: ['Timestamp','Action','UserID','Payload']
};

const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

function doGet() {
  return jsonResponse({ success: true, message: 'Use POST com JSON {action, payload}.' });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData?.contents || '{}');
    const action = payload.action;
    const token = payload.token;
    const data = payload.payload || {};
    if (!action) throw new Error('Informe a ação.');

    const context = resolveContext(token);
    let result = {};

    switch (action) {
      case 'bootstrap':
        result = handleBootstrap(context);
        break;
      case 'signup':
        result = handleSignup(data);
        break;
      case 'login':
        result = handleLogin(data);
        break;
      case 'logout':
        result = handleLogout(context.tokenRow);
        break;
      case 'checkin':
        requireUser(context.user);
        result = handleCheckin(context.user);
        break;
      case 'submitProgress':
        requireUser(context.user);
        result = handleSubmitProgress(context.user, data);
        break;
      case 'updateEmbed':
        requireAdmin(context.user);
        result = handleUpdateEmbed(context.user, data);
        break;
      case 'awardXp':
        requireAdmin(context.user);
        result = handleAwardXp(context.user, data);
        break;
      case 'exportData':
        requireAdmin(context.user);
        result = handleExportData();
        break;
      case 'importData':
        requireAdmin(context.user);
        result = handleImportData(data);
        break;
      default:
        throw new Error('Ação não suportada: ' + action);
    }

    return jsonResponse(Object.assign({ success: true }, result));
  } catch (err) {
    return jsonResponse({ success: false, message: err.message || 'Erro inesperado.' });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ===== Helpers básicos =====
function ensureSheet(name) {
  const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  const headers = TABLE_HEADERS[name];
  if (headers && sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function readSheet(name) {
  const sheet = ensureSheet(name);
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(cell => cell === '' || cell === null)) continue;
    const obj = {};
    headers.forEach((h, idx) => obj[h] = row[idx]);
    obj._rowNumber = i + 1;
    rows.push(obj);
  }
  return rows;
}

function writeRow(name, obj, rowNumber) {
  const sheet = ensureSheet(name);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  if (rowNumber) {
    sheet.getRange(rowNumber, 1, 1, headers.length).setValues([data]);
  } else {
    sheet.appendRow(data);
  }
}

function hashPassword(raw) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw));
}

function resolveContext(token) {
  if (!token) return { user: null, tokenRow: null };
  const tokens = readSheet(SHEETS.TOKENS);
  const match = tokens.find(t => String(t.Token) === String(token));
  if (!match) return { user: null, tokenRow: null };
  if (match.ExpiresAt && new Date(match.ExpiresAt) < new Date()) {
    removeToken(match.Token);
    return { user: null, tokenRow: null };
  }
  const user = getUserById(match.UserID);
  if (!user) {
    removeToken(match.Token);
    return { user: null, tokenRow: null };
  }
  return { user: user, tokenRow: match };
}

function removeToken(token) {
  if (!token) return;
  const sheet = ensureSheet(SHEETS.TOKENS);
  const tokens = readSheet(SHEETS.TOKENS);
  const match = tokens.find(t => t.Token === token);
  if (match) {
    sheet.deleteRow(match._rowNumber);
  }
}

function issueToken(userId) {
  const token = Utilities.getUuid();
  const expires = new Date(Date.now() + 1000 * 60 * 60 * 24 * 30); // 30 dias
  writeRow(SHEETS.TOKENS, {
    Token: token,
    UserID: userId,
    ExpiresAt: expires,
    CriadoEm: new Date()
  });
  return token;
}

function requireUser(user) {
  if (!user) throw new Error('Faça login para continuar.');
}

function requireAdmin(user) {
  requireUser(user);
  if (!user.admin) throw new Error('Ação exclusiva para administradores.');
}

function mapUser(row) {
  if (!row) return null;
  return {
    id: row.ID,
    nome: row.Nome,
    email: row.Email,
    admin: String(row.Admin).toLowerCase() === 'true',
    xp: Number(row.XP || 0)
  };
}

function getUserByEmail(email) {
  const rows = readSheet(SHEETS.USERS);
  return rows.map(mapUserRow).find(u => u.Email && String(u.Email).toLowerCase() === String(email).toLowerCase()) || null;
}

function getUserById(id) {
  const rows = readSheet(SHEETS.USERS);
  const match = rows.find(r => r.ID === id);
  if (!match) return null;
  const mapped = mapUserRow(match);
  mapped._rowNumber = match._rowNumber;
  return mapped;
}

function mapUserRow(row) {
  return {
    ID: row.ID,
    Nome: row.Nome,
    Email: row.Email,
    SenhaHash: row.SenhaHash,
    Admin: row.Admin,
    XP: Number(row.XP || 0),
    CriadoEm: row.CriadoEm,
    AtualizadoEm: row.AtualizadoEm,
    _rowNumber: row._rowNumber
  };
}

function setUserXp(userId, newXp) {
  const user = getUserById(userId);
  if (!user) throw new Error('Usuário não encontrado.');
  user.XP = newXp;
  user.AtualizadoEm = new Date();
  writeRow(SHEETS.USERS, user, user._rowNumber);
}

function buildUserProfile(userId) {
  const user = getUserById(userId);
  if (!user) return null;
  const progresso = {};
  const progressRows = readSheet(SHEETS.PROGRESS).filter(p => String(p.UserID) === String(userId));
  progressRows.forEach(p => {
    const modId = Number(p.ModuleID);
    progresso[modId] = {
      done: String(p.Done).toLowerCase() === 'true' || Number(p.Score || 0) >= 70,
      score: Number(p.Score || 0),
      xp: Number(p.XP || 0)
    };
  });
  const checkins = readSheet(SHEETS.CHECKINS)
    .filter(c => String(c.UserID) === String(userId))
    .map(c => c.Date);
  return {
    id: user.ID,
    nome: user.Nome,
    email: user.Email,
    admin: String(user.Admin).toLowerCase() === 'true',
    xp: Number(user.XP || 0),
    progresso: progresso,
    checkins: checkins
  };
}

function loadConfig() {
  const cfg = { xpCheckin: 5, xpPorNivel: 100, ciclo: 'Turma Excel — 2025' };
  readSheet(SHEETS.CONFIG).forEach(item => {
    const key = String(item.Chave || '').trim();
    const value = item.Valor;
    if (!key) return;
    if (key === 'xpCheckin' || key === 'xpPorNivel') {
      cfg[key] = Number(value || 0);
    } else {
      cfg[key] = value;
    }
  });
  return cfg;
}

function loadEmbeds() {
  const embeds = { excel: '', ppt: '' };
  readSheet(SHEETS.EMBEDS).forEach(item => {
    const type = String(item.Tipo || '').toLowerCase();
    if (type && item.URL) {
      embeds[type] = item.URL;
    }
  });
  return embeds;
}

const DEFAULT_MODULES = (function() {
  const modules = Array.from({ length: 24 }).map((_, i) => ({
    id: i + 1,
    titulo: `${i + 1}. Módulo ${i + 1}`,
    desc: 'Atividade prática do módulo.',
    xp: 30,
    perguntas: []
  }));
  modules[0] = {
    id: 1,
    titulo: '1. Boas‑vindas, interface e fundamentos',
    desc: 'Crie sua primeira planilha com cabeçalhos e formate os dados. Responda as perguntas para validar o aprendizado.',
    xp: 40,
    perguntas: [
      {
        id: 'q1',
        tipo: 'mc',
        enunciado: 'Onde fica a Faixa de Opções (Ribbon)?',
        opcoes: ['Na barra lateral esquerda','Na parte superior da janela','Dentro da célula ativa','No menu de contexto'],
        correta: 1
      },
      {
        id: 'q2',
        tipo: 'mc',
        enunciado: 'Qual atalho salva rapidamente?',
        opcoes: ['Ctrl/Cmd + N','Ctrl/Cmd + 1','Ctrl/Cmd + S','Ctrl/Cmd + Enter'],
        correta: 2
      },
      {
        id: 'q3',
        tipo: 'text',
        enunciado: 'Descreva dois cuidados ao criar cabeçalhos da planilha.',
        minCaracteres: 10
      }
    ]
  };
  return modules;
})();

function loadModules() {
  const moduleRows = readSheet(SHEETS.MODULES);
  if (!moduleRows.length) return DEFAULT_MODULES;
  const questionRows = readSheet(SHEETS.QUESTIONS);
  const modules = moduleRows
    .filter(row => String(row.Ativo || '').toLowerCase() !== 'false')
    .map(row => ({
      id: Number(row.ModuleID),
      ordem: Number(row.Ordem || row.ModuleID),
      titulo: row.Titulo,
      desc: row.Descricao,
      xp: Number(row.XP || 0),
      videoUrl: row.VideoURL,
      materialUrl: row.MaterialURL,
      perguntas: []
    }));
  modules.sort((a, b) => a.ordem - b.ordem);
  const moduleMap = {};
  modules.forEach(m => moduleMap[m.id] = m);
  questionRows.forEach(q => {
    const modId = Number(q.ModuleID);
    if (!moduleMap[modId]) return;
    const pergunta = {
      id: q.QuestionID,
      tipo: q.Tipo || 'mc',
      enunciado: q.Enunciado,
      opcoes: [],
      correta: q.Correta !== undefined ? Number(q.Correta) : null,
      peso: q.Peso ? Number(q.Peso) : 1,
      minCaracteres: q.MinCaracteres ? Number(q.MinCaracteres) : 10,
      feedback: q.Feedback || ''
    };
    if (q.OpcoesJSON) {
      try {
        pergunta.opcoes = JSON.parse(q.OpcoesJSON);
      } catch (err) {
        pergunta.opcoes = String(q.OpcoesJSON).split('||');
      }
    }
    moduleMap[modId].perguntas.push(pergunta);
  });
  modules.forEach(m => {
    if (!m.perguntas) m.perguntas = [];
  });
  return modules;
}

function buildRanking() {
  const cfg = loadConfig();
  const xpNivel = Number(cfg.xpPorNivel || 100);
  return readSheet(SHEETS.USERS)
    .filter(row => String(row.Admin).toLowerCase() !== 'true')
    .map(row => ({
      id: row.ID,
      nome: row.Nome,
      email: row.Email,
      xp: Number(row.XP || 0),
      nivel: 1 + Math.floor(Number(row.XP || 0) / xpNivel)
    }))
    .sort((a, b) => b.xp - a.xp);
}

function buildAdminList() {
  const cfg = loadConfig();
  const xpNivel = Number(cfg.xpPorNivel || 100);
  const progress = readSheet(SHEETS.PROGRESS);
  return readSheet(SHEETS.USERS)
    .filter(row => String(row.Admin).toLowerCase() !== 'true')
    .map(row => {
      const concluidos = progress.filter(p => String(p.UserID) === String(row.ID) && (String(p.Done).toLowerCase() === 'true' || Number(p.Score || 0) >= 70)).length;
      return {
        id: row.ID,
        nome: row.Nome,
        email: row.Email,
        xp: Number(row.XP || 0),
        nivel: 1 + Math.floor(Number(row.XP || 0) / xpNivel),
        concluidos: concluidos
      };
    })
    .sort((a, b) => b.xp - a.xp);
}

function logAudit(action, userId, payload) {
  writeRow(SHEETS.AUDIT, {
    Timestamp: new Date(),
    Action: action,
    UserID: userId || '',
    Payload: payload ? JSON.stringify(payload) : ''
  });
}

// ===== Handlers =====
function handleBootstrap(context) {
  const user = context?.user || null;
  const tokenRow = context?.tokenRow || null;
  const modules = loadModules();
  const cfg = loadConfig();
  const embeds = loadEmbeds();
  const ranking = buildRanking();
  const adminUsers = user && user.admin ? buildAdminList() : [];
  const profile = user ? buildUserProfile(user.id) : null;
  let token = '';
  if (user) {
    if (tokenRow) {
      tokenRow.ExpiresAt = new Date(Date.now() + 1000 * 60 * 60 * 24 * 30);
      tokenRow.CriadoEm = tokenRow.CriadoEm || new Date();
      writeRow(SHEETS.TOKENS, tokenRow, tokenRow._rowNumber);
      token = tokenRow.Token;
    } else {
      token = issueToken(user.id);
    }
  }
  return {
    modules: modules,
    cfg: cfg,
    embeds: embeds,
    ranking: ranking,
    adminUsers: adminUsers,
    user: profile,
    token: token
  };
}

function handleSignup(data) {
  const nome = String(data.nome || '').trim();
  const email = String(data.email || '').trim().toLowerCase();
  const senha = String(data.senha || '').trim();
  const admin = Boolean(data.admin);
  if (!nome || !email || !senha) throw new Error('Preencha nome, e-mail e senha.');
  if (getUserByEmail(email)) throw new Error('E-mail já cadastrado.');

  const id = Utilities.getUuid();
  const hashed = hashPassword(senha);
  writeRow(SHEETS.USERS, {
    ID: id,
    Nome: nome,
    Email: email,
    SenhaHash: hashed,
    Admin: admin,
    XP: 0,
    CriadoEm: new Date(),
    AtualizadoEm: new Date()
  });

  const token = issueToken(id);
  logAudit('signup', id, { email: email });
  return {
    token: token,
    user: buildUserProfile(id),
    ranking: buildRanking()
  };
}

function handleLogin(data) {
  const email = String(data.email || '').trim().toLowerCase();
  const senha = String(data.senha || '').trim();
  if (!email || !senha) throw new Error('Informe e-mail e senha.');
  const userRow = getUserByEmail(email);
  if (!userRow) throw new Error('Usuário não encontrado.');
  if (userRow.SenhaHash !== hashPassword(senha)) throw new Error('Senha incorreta.');
  const token = issueToken(userRow.ID);
  logAudit('login', userRow.ID, {});
  return {
    token: token,
    user: buildUserProfile(userRow.ID),
    ranking: buildRanking(),
    adminUsers: String(userRow.Admin).toLowerCase() === 'true' ? buildAdminList() : []
  };
}

function handleLogout(tokenRow) {
  if (tokenRow?.Token) {
    removeToken(tokenRow.Token);
  }
  return { message: 'Sessão finalizada.' };
}

function handleCheckin(user) {
  const cfg = loadConfig();
  const xpGain = Number(cfg.xpCheckin || 5);
  const today = Utilities.formatDate(new Date(), 'Etc/UTC', 'yyyy-MM-dd');
  const checkins = readSheet(SHEETS.CHECKINS);
  const already = checkins.find(c => String(c.UserID) === String(user.id) && String(c.Date) === today);
  if (already) {
    return {
      message: 'Presença de hoje já registrada.',
      user: buildUserProfile(user.id)
    };
  }
  writeRow(SHEETS.CHECKINS, {
    UserID: user.id,
    Date: today,
    XP: xpGain,
    RegistradoEm: new Date()
  });
  const current = getUserById(user.id);
  setUserXp(user.id, Number(current.XP || 0) + xpGain);
  logAudit('checkin', user.id, { date: today, xp: xpGain });
  return {
    message: `Presença registrada. +${xpGain} XP!`,
    user: buildUserProfile(user.id),
    ranking: buildRanking(),
    adminUsers: current.Admin ? buildAdminList() : []
  };
}

function handleSubmitProgress(user, data) {
  const moduleId = Number(data.moduleId);
  if (!moduleId) throw new Error('Módulo inválido.');
  const modules = loadModules();
  const module = modules.find(m => Number(m.id) === moduleId);
  if (!module) throw new Error('Módulo não localizado.');
  const answers = Array.isArray(data.answers) ? data.answers : [];
  const progressSheet = ensureSheet(SHEETS.PROGRESS);
  const rows = readSheet(SHEETS.PROGRESS);
  const existing = rows.find(r => String(r.UserID) === String(user.id) && Number(r.ModuleID) === moduleId);

  let score = 0;
  let totalPeso = 0;
  const answerMap = {};
  module.perguntas.forEach(q => {
    const peso = Number(q.peso || 1);
    totalPeso += peso;
    const ans = answers.find(a => String(a.questionId) === String(q.id));
    let isCorrect = false;
    if (q.tipo === 'mc') {
      const selected = ans ? Number(ans.answer) : null;
      isCorrect = selected !== null && selected === Number(q.correta);
    } else {
      const resp = ans ? String(ans.answer || '').trim() : '';
      isCorrect = resp.length >= Number(q.minCaracteres || 10);
    }
    if (isCorrect) score += peso;
    answerMap[q.id] = ans ? ans.answer : null;
  });
  const pct = totalPeso ? Math.round((score / totalPeso) * 100) : 0;
  const earned = Math.round(Number(module.xp || 0) * (pct / 100));

  const record = {
    UserID: user.id,
    ModuleID: moduleId,
    Score: pct,
    Done: pct >= 70,
    XP: earned,
    AnswersJSON: JSON.stringify(answerMap),
    AtualizadoEm: new Date()
  };

  const previousXp = existing ? Number(existing.XP || 0) : 0;
  if (existing) {
    writeRow(SHEETS.PROGRESS, Object.assign(existing, record), existing._rowNumber);
  } else {
    writeRow(SHEETS.PROGRESS, record);
  }
  const delta = Math.max(earned - previousXp, 0);
  if (delta > 0) {
    const current = getUserById(user.id);
    setUserXp(user.id, Number(current.XP || 0) + delta);
  }
  logAudit('submitProgress', user.id, { moduleId: moduleId, score: pct, earned: earned });
  const refreshedUser = buildUserProfile(user.id);
  return {
    message: `Você concluiu ${pct}% da avaliação.`,
    user: refreshedUser,
    ranking: buildRanking(),
    adminUsers: refreshedUser.admin ? buildAdminList() : []
  };
}

function handleUpdateEmbed(user, data) {
  const type = String(data.type || '').toLowerCase();
  const url = String(data.url || '').trim();
  if (!type || !url) throw new Error('Informe o tipo e o link.');
  const embeds = readSheet(SHEETS.EMBEDS);
  const existing = embeds.find(e => String(e.Tipo || '').toLowerCase() === type);
  const payload = {
    Tipo: type,
    URL: url,
    AtualizadoPor: user.id,
    AtualizadoEm: new Date()
  };
  if (existing) {
    writeRow(SHEETS.EMBEDS, Object.assign(existing, payload), existing._rowNumber);
  } else {
    writeRow(SHEETS.EMBEDS, payload);
  }
  logAudit('updateEmbed', user.id, { type: type });
  return {
    embeds: loadEmbeds()
  };
}

function handleAwardXp(adminUser, data) {
  const userId = String(data.userId || '').trim();
  const amount = Number(data.amount || 0);
  if (!userId || !amount) throw new Error('Informe o usuário e o XP.');
  const target = getUserById(userId);
  if (!target) throw new Error('Usuário não encontrado.');
  setUserXp(userId, Number(target.XP || 0) + amount);
  logAudit('awardXp', adminUser.id, { target: userId, amount: amount });
  return {
    ranking: buildRanking(),
    adminUsers: buildAdminList()
  };
}

function handleExportData() {
  const exportObj = {
    users: readSheet(SHEETS.USERS),
    modules: readSheet(SHEETS.MODULES),
    questions: readSheet(SHEETS.QUESTIONS),
    progress: readSheet(SHEETS.PROGRESS),
    checkins: readSheet(SHEETS.CHECKINS),
    config: readSheet(SHEETS.CONFIG),
    embeds: readSheet(SHEETS.EMBEDS)
  };
  return { export: exportObj };
}

function handleImportData(data) {
  if (!data || typeof data !== 'object') throw new Error('JSON inválido.');
  if (data.users) overwriteSheet(SHEETS.USERS, data.users);
  if (data.modules) overwriteSheet(SHEETS.MODULES, data.modules);
  if (data.questions) overwriteSheet(SHEETS.QUESTIONS, data.questions);
  if (data.progress) overwriteSheet(SHEETS.PROGRESS, data.progress);
  if (data.checkins) overwriteSheet(SHEETS.CHECKINS, data.checkins);
  if (data.config) overwriteSheet(SHEETS.CONFIG, data.config);
  if (data.embeds) overwriteSheet(SHEETS.EMBEDS, data.embeds);
  logAudit('importData', null, {});
  return { message: 'Importação concluída.' };
}

function overwriteSheet(name, rows) {
  const sheet = ensureSheet(name);
  sheet.clearContents();
  const headers = TABLE_HEADERS[name];
  if (headers) sheet.appendRow(headers);
  if (!Array.isArray(rows)) return;
  const data = rows.map(r => headers.map(h => r[h] !== undefined ? r[h] : ''));
  if (data.length) sheet.getRange(sheet.getLastRow() + 1, 1, data.length, headers.length).setValues(data);
}
