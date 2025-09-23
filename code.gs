/** =================== CONFIG GERAL =================== **/
const SPREADSHEET_ID = '1DCzIOIcRBaJ3WJVQCOWg6KXyghjRGMJUHekV1fGygJ4';
const ADMIN_SECURITY_CODE = 'xbY4nu'; // código exigido p/ criar admins

// Abas usadas
const SHEET_USERS            = 'Users';            // id, name, email, passHash, isAdmin, xp, createdAt
const SHEET_PROGRESS         = 'Progress';         // userId, moduleId, scorePct, earnedXP, completedAt
const SHEET_CHECKIN          = 'Checkins';         // userId, dateISO, xp, streak
const SHEET_CONFIG           = 'Config';           // key, value
const SHEET_EMBEDS           = 'Embeds';           // key, url
const SHEET_SESSIONS         = 'Sessions';         // userId, tokenHash, expiresAt, createdAt
const SHEET_USER_ACHIEVEMENT = 'UserAchievements'; // userId, achievementId, unlockedAt, rewardXP
const SHEET_WALL             = 'Wall';             // id, userId, authorName, message, createdAt, visibility, targetUserIds, attachmentIds, removedAt, removedBy
const SHEET_CONFIRMATIONS    = 'UserConfirmations';// userId, email, codeHash, createdAt, expiresAt, confirmedAt, lastSentAt
const SHEET_WALL_ATTACHMENTS = 'WallAttachments';  // attachmentId, postId, uploaderId, fileId, fileName, mimeType, fileUrl, sizeBytes, uploadedAt, linkedAt

const CONFIG_RECOMMENDED_RESOURCES = 'recommended_resources';
const CONFIG_UPCOMING_EVENTS       = 'upcoming_events';

const COMMUNITY_WALL_CHAR_LIMIT = 240;
const CHECKIN_STREAK_XP_CAP = 24;
const COMMUNITY_ATTACHMENT_MAX_BYTES = 10 * 1024 * 1024; // 10 MB por arquivo
const COMMUNITY_FILES_FOLDER_ID = '1LBfNBUjTMjrEVL1CE-YJtA3ecW_gQMIR';

const ACHIEVEMENTS = [
  {
    id: 'checkin_first',
    title: 'Primeiro Passo',
    description: 'Registre seu primeiro check-in diário.',
    category: 'Check-ins',
    icon: '✅',
    rewardXP: 10,
    criteria: { type: 'checkins_total', target: 1 }
  },
  {
    id: 'checkin_7_streak',
    title: 'Rotina Consolidada',
    description: 'Mantenha um streak de check-ins por 7 dias seguidos.',
    category: 'Check-ins',
    icon: '🔥',
    rewardXP: 30,
    criteria: { type: 'checkin_streak', target: 7 }
  },
  {
    id: 'checkin_30_total',
    title: 'Assiduidade Máxima',
    description: 'Complete 30 check-ins no total.',
    category: 'Check-ins',
    icon: '📅',
    rewardXP: 50,
    criteria: { type: 'checkins_total', target: 30 }
  },
  {
    id: 'modules_5_highscore',
    title: 'Especialista em Excel',
    description: 'Conclua 5 módulos com nota igual ou superior a 90%.',
    category: 'Módulos',
    icon: '📈',
    rewardXP: 60,
    criteria: { type: 'modules_high_score', target: 5, minScore: 90 }
  },
  {
    id: 'modules_10_completed',
    title: 'Maratona de Estudos',
    description: 'Finalize 10 módulos com aproveitamento mínimo.',
    category: 'Módulos',
    icon: '🎓',
    rewardXP: 80,
    criteria: { type: 'modules_completed', target: 10, minScore: 70 }
  },
  {
    id: 'xp_500_total',
    title: 'Acumulador de XP',
    description: 'Alcance um total de 500 XP.',
    category: 'Progressão',
    icon: '🏅',
    rewardXP: 100,
    criteria: { type: 'xp_total', target: 500 }
  },
  {
    id: 'xp_1000_total',
    title: 'Lenda do Excel',
    description: 'Some 1000 XP na plataforma.',
    category: 'Progressão',
    icon: '👑',
    rewardXP: 150,
    criteria: { type: 'xp_total', target: 1000 }
  }
];

const MODULES = Object.freeze([
  Object.freeze({ id: 1, xpMax: 30 }),
  Object.freeze({ id: 2, xpMax: 35 }),
  Object.freeze({ id: 3, xpMax: 35 }),
  Object.freeze({ id: 4, xpMax: 40 }),
  Object.freeze({ id: 5, xpMax: 35 }),
  Object.freeze({ id: 6, xpMax: 35 }),
  Object.freeze({ id: 7, xpMax: 40 }),
  Object.freeze({ id: 8, xpMax: 35 }),
  Object.freeze({ id: 9, xpMax: 45 }),
  Object.freeze({ id: 10, xpMax: 40 }),
  Object.freeze({ id: 11, xpMax: 35 }),
  Object.freeze({ id: 12, xpMax: 35 }),
  Object.freeze({ id: 13, xpMax: 40 }),
  Object.freeze({ id: 14, xpMax: 45 }),
  Object.freeze({ id: 15, xpMax: 40 }),
  Object.freeze({ id: 16, xpMax: 45 }),
  Object.freeze({ id: 17, xpMax: 50 }),
  Object.freeze({ id: 18, xpMax: 45 }),
  Object.freeze({ id: 19, xpMax: 50 }),
  Object.freeze({ id: 20, xpMax: 35 }),
  Object.freeze({ id: 21, xpMax: 50 }),
  Object.freeze({ id: 22, xpMax: 55 }),
  Object.freeze({ id: 23, xpMax: 60 }),
  Object.freeze({ id: 24, xpMax: 80 })
]);

const MODULES_BY_ID = MODULES.reduce((acc, module) => {
  acc[String(module.id)] = module;
  return acc;
}, {});

const SESSION_DURATION_HOURS = 24 * 7; // validade de 7 dias
const SESSION_INVALID_MESSAGE = 'Sessão inválida ou expirada. Faça login novamente.';

/** =================== BOOTSTRAP =================== **/
function onOpen() {
  setup_();
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Plataforma Excel')
    .addItem('Recriar estrutura', 'setup_')
    .addToUi();
}

function doGet(e) {
  setup_();
  let route = '';
  if (e && typeof e.pathInfo === 'string') {
    const parts = e.pathInfo
      .split('/')
      .map(str => str.trim())
      .filter(Boolean);
    if (parts.length) {
      route = parts[0];
    }
  }
  const template = HtmlService.createTemplateFromFile('index');
  const initialRoute = (route || '')
    .toString()
    .toLowerCase()
    .replace(/[^a-z]/g, '');
  template.initialRoute = initialRoute;
  return template
    .evaluate()
    .setTitle('Plataforma Gamificada — Curso de Excel')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function renderStudentSubPage(payload) {
  const pageRaw = payload && payload.page ? payload.page.toString() : '';
  const page = pageRaw.replace(/[^a-z0-9\-]/gi, '').toLowerCase();
  let filename = '';
  if (page === 'next-lesson') {
    filename = 'next-lesson';
  } else if (page === 'achievements-catalog') {
    filename = 'achievements-catalog';
  } else {
    throw new Error('Página desconhecida.');
  }
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Cria abas e cabeçalhos se não existirem */
function setup_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ensure = (name, headers) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const range = sh.getRange(1,1,1,headers.length);
    range.setValues([headers]);
    sh.setFrozenRows(1);
  };

  ensure(SHEET_USERS,            ['id','name','email','passHash','isAdmin','xp','createdAt']);
  ensure(SHEET_PROGRESS,         ['userId','moduleId','scorePct','earnedXP','completedAt']);
  ensure(SHEET_CHECKIN,          ['userId','dateISO','xp','streak']);
  ensure(SHEET_CONFIG,           ['key','value']);
  ensure(SHEET_EMBEDS,           ['key','url']);
  ensure(SHEET_SESSIONS,         ['userId','tokenHash','expiresAt','createdAt']);
  ensure(SHEET_USER_ACHIEVEMENT, ['userId','achievementId','unlockedAt','rewardXP']);
  ensure(SHEET_WALL,             ['id','userId','authorName','message','createdAt','visibility','targetUserIds','attachmentIds','removedAt','removedBy']);
  ensure(SHEET_CONFIRMATIONS,    ['userId','email','codeHash','createdAt','expiresAt','confirmedAt','lastSentAt']);
  ensure(SHEET_WALL_ATTACHMENTS, ['attachmentId','postId','uploaderId','fileId','fileName','mimeType','fileUrl','sizeBytes','uploadedAt','linkedAt']);

  // Defaults de config
  const cfg = getConfig_();
  if (!cfg.xpCheckin) setConfig_('xpCheckin','5');
  if (!cfg.xpPerLevel) setConfig_('xpPerLevel','100');
  if (!cfg[CONFIG_RECOMMENDED_RESOURCES]) setConfig_(CONFIG_RECOMMENDED_RESOURCES, '[]');
  if (!cfg[CONFIG_UPCOMING_EVENTS]) setConfig_(CONFIG_UPCOMING_EVENTS, '[]');
}

/** =================== HELPERS =================== **/
function ss_()              { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function sh_(name)          { return ss_().getSheetByName(name); }
function nowISO_()          { return new Date().toISOString(); }
function todayISO_()        { return Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd'); }
function toB64_(bytes)      { return Utilities.base64Encode(bytes); }
function sha256_(str)       { return toB64_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8)); }
function parseDateValue_(value) {
  if (value instanceof Date) {
    const time = value.getTime();
    return isNaN(time) ? null : new Date(time);
  }
  if (typeof value === 'number') {
    const dateFromNumber = new Date(value);
    return isNaN(dateFromNumber.getTime()) ? null : dateFromNumber;
  }
  const str = (value || '').toString().trim();
  if (!str) return null;
  const normalized = str.length === 10 ? str + 'T00:00:00Z' : str;
  const parsed = new Date(normalized);
  return isNaN(parsed.getTime()) ? null : parsed;
}
function getAll_(sheet)     { const r=sheet.getDataRange().getValues(); return r.length>1 ? r.slice(1) : []; }
function findByEmail_(email){ const rows=getAll_(sh_(SHEET_USERS)); for (let i=0;i<rows.length;i++){ if ((rows[i][2]||'').toLowerCase()===email.toLowerCase()){ return { row:i+2, data:rows[i] }; } } return null; }
function getConfig_()       { const rows=getAll_(sh_(SHEET_CONFIG)); const m={}; rows.forEach(r=>m[(r[0]||'').toString()]=(r[1]||'').toString()); return m; }
function setConfig_(k,v)    { const s=sh_(SHEET_CONFIG); const rows=getAll_(s); for (let i=0;i<rows.length;i++){ if (rows[i][0]==k){ s.getRange(i+2,2).setValue(v); return; } } s.appendRow([k,v]); }

function generateConfirmationCode_() {
  const min = 100000;
  const max = 999999;
  const number = Math.floor(Math.random() * (max - min + 1)) + min;
  return String(number);
}

function hashConfirmationCode_(userId, code) {
  const safeUser = (userId || '').toString();
  const safeCode = (code || '').toString();
  return sha256_(safeUser + '|' + safeCode);
}

function getConfirmationRecordByUserId_(userId) {
  if (!userId) return null;
  const rows = getAll_(sh_(SHEET_CONFIRMATIONS));
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if ((row[0] || '') === userId) {
      return {
        row: i + 2,
        data: {
          userId: row[0] || '',
          email: row[1] || '',
          codeHash: row[2] || '',
          createdAt: row[3] || '',
          expiresAt: row[4] || '',
          confirmedAt: row[5] || '',
          lastSentAt: row[6] || ''
        }
      };
    }
  }
  return null;
}

function saveConfirmationCode_(userId, email, codeHash, expiresAtISO) {
  const sheet = sh_(SHEET_CONFIRMATIONS);
  const now = nowISO_();
  const payload = [userId, email, codeHash, now, expiresAtISO, '', now];
  const existing = getConfirmationRecordByUserId_(userId);
  if (existing) {
    sheet.getRange(existing.row, 1, 1, payload.length).setValues([payload]);
    return {
      row: existing.row,
      data: {
        userId,
        email,
        codeHash,
        createdAt: now,
        expiresAt: expiresAtISO,
        confirmedAt: '',
        lastSentAt: now
      }
    };
  }
  sheet.appendRow(payload);
  return {
    row: sheet.getLastRow(),
    data: {
      userId,
      email,
      codeHash,
      createdAt: now,
      expiresAt: expiresAtISO,
      confirmedAt: '',
      lastSentAt: now
    }
  };
}

function markConfirmationAsConfirmed_(userId) {
  const sheet = sh_(SHEET_CONFIRMATIONS);
  const existing = getConfirmationRecordByUserId_(userId);
  const now = nowISO_();
  if (existing) {
    sheet.getRange(existing.row, 3).setValue('');
    sheet.getRange(existing.row, 6).setValue(now);
    return { row: existing.row, confirmedAt: now };
  }
  sheet.appendRow([userId, '', '', now, now, now, now]);
  return { row: sheet.getLastRow(), confirmedAt: now };
}

function isUserConfirmed_(userId) {
  const record = getConfirmationRecordByUserId_(userId);
  if (!record) return true;
  return !!record.data.confirmedAt;
}

function buildConfirmationEmailBodies_(name, code) {
  const template = HtmlService.createTemplateFromFile('confirmation-email');
  template.name = name || 'Participante';
  template.code = code;
  const html = template.evaluate().getContent();
  const subject = 'Confirme seu cadastro na Plataforma Excel';
  const plain = [
    'Olá ' + (name || 'participante') + ',',
    '',
    'Use o código abaixo para confirmar seu cadastro na Plataforma Excel:',
    '',
    code,
    '',
    'Se você não solicitou este acesso, ignore esta mensagem.'
  ].join('\n');
  return { subject, plain, html };
}

function sendConfirmationEmail_(email, name, code) {
  if (!email || !code) return;
  const bodies = buildConfirmationEmailBodies_(name, code);
  MailApp.sendEmail({
    to: email,
    subject: bodies.subject,
    htmlBody: bodies.html,
    body: bodies.plain,
    name: 'Plataforma Excel',
    noReply: true
  });
}

function getModuleById_(moduleId) {
  const numericId = Number(moduleId);
  if (!Number.isFinite(numericId) || numericId <= 0 || !Number.isInteger(numericId)) return null;
  const module = MODULES_BY_ID[String(numericId)];
  if (!module) return null;
  return { id: module.id, xpMax: module.xpMax };
}

function normalizeStringArray_(value) {
  if (!Array.isArray(value)) return [];
  const result = [];
  for (let i = 0; i < value.length; i++) {
    const item = value[i];
    if (item === null || item === undefined) continue;
    const textItem = item.toString().trim();
    if (!textItem) continue;
    if (result.indexOf(textItem) === -1) result.push(textItem);
  }
  return result;
}

function deserializeIdList_(value) {
  if (!value && value !== 0) return [];
  if (Array.isArray(value)) return normalizeStringArray_(value);
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return [];
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) return normalizeStringArray_(parsed);
    } catch (err) {
      // fallback para separação simples
    }
    const parts = trimmed.split(/[,;\s]+/);
    return normalizeStringArray_(parts);
  }
  return [];
}

function serializeIdList_(value) {
  if (typeof value === 'string') {
    const parts = value.split(/[,;\s]+/);
    return JSON.stringify(normalizeStringArray_(parts));
  }
  return JSON.stringify(normalizeStringArray_(Array.isArray(value) ? value : []));
}

function parseConfigListValue_(raw) {
  if (!raw) return [];
  if (Array.isArray(raw)) return normalizeStringArray_(raw);
  if (typeof raw === 'string') {
    const trimmed = raw.trim();
    if (!trimmed) return [];
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) return normalizeStringArray_(parsed);
    } catch (err) {
      // fallback para tratamento simples
    }
    return normalizeStringArray_(trimmed.split('\n'));
  }
  return [];
}

function saveConfigList_(key, list) {
  const normalized = normalizeStringArray_(Array.isArray(list) ? list : []);
  const value = normalized.length ? JSON.stringify(normalized) : '[]';
  setConfig_(key, value);
  return normalized;
}

function getConfigList_(key) {
  const cfg = getConfig_();
  const raw = cfg[key];
  return parseConfigListValue_(raw);
}

function getUserMap_() {
  const rows = getAll_(sh_(SHEET_USERS));
  const map = {};
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const id = row[0] || '';
    if (!id) continue;
    map[id] = {
      name: row[1] || '',
      email: row[2] || '',
      isAdmin: !!row[4],
      xp: Number(row[5] || 0)
    };
  }
  return map;
}

function computeLevelInfo_(xpRaw, cfg) {
  const xpPerLevelRaw = Number((cfg && cfg.xpPerLevel) || 100);
  const xpPerLevel = (Number.isFinite(xpPerLevelRaw) && xpPerLevelRaw > 0) ? xpPerLevelRaw : 100;
  const safeXP = Math.max(0, Number(xpRaw || 0));
  const level = 1 + Math.floor(safeXP / xpPerLevel);
  const xpIntoLevel = safeXP - xpPerLevel * Math.max(0, level - 1);
  const xpToNextLevel = Math.max(0, xpPerLevel - xpIntoLevel);
  const nextLevel = level + 1;
  return { xpPerLevel, level, xpIntoLevel, xpToNextLevel, nextLevel };
}

/** =================== ACHIEVEMENTS =================== **/
function getUserAchievementRecords_(userId) {
  const rows = getAll_(sh_(SHEET_USER_ACHIEVEMENT));
  const list = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if ((row[0] || '') === userId) {
      list.push({
        row: i + 2,
        achievementId: row[1],
        unlockedAt: row[2] || '',
        rewardXP: Number(row[3] || 0)
      });
    }
  }
  return list;
}

function calculateStreakStats_(days) {
  if (!Array.isArray(days) || days.length === 0) {
    return { total: 0, current: 0, best: 0 };
  }
  const msPerDay = 24 * 60 * 60 * 1000;
  const uniqueDays = Array.from(new Set(days.filter(Boolean)));
  if (!uniqueDays.length) {
    return { total: 0, current: 0, best: 0 };
  }
  const timestamps = [];
  for (let i = 0; i < uniqueDays.length; i++) {
    const parsed = parseDateValue_(uniqueDays[i]);
    if (parsed) {
      const utc = Date.UTC(parsed.getUTCFullYear(), parsed.getUTCMonth(), parsed.getUTCDate());
      if (!Number.isNaN(utc)) timestamps.push(utc);
    }
  }
  if (!timestamps.length) {
    return { total: 0, current: 0, best: 0 };
  }
  timestamps.sort((a, b) => a - b);
  let best = 0;
  let streak = 0;
  let previous = null;
  for (let i = 0; i < timestamps.length; i++) {
    const time = timestamps[i];
    if (previous === null) {
      streak = 1;
    } else {
      const diffDays = Math.round((time - previous) / msPerDay);
      streak = diffDays === 1 ? streak + 1 : 1;
    }
    if (streak > best) best = streak;
    previous = time;
  }

  let current = 0;
  const today = new Date();
  const todayUTC = Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), today.getUTCDate());
  const last = timestamps[timestamps.length - 1];
  const diffToToday = Math.round((todayUTC - last) / msPerDay);
  if (diffToToday <= 1) {
    current = 1;
    let pointer = last;
    for (let i = timestamps.length - 2; i >= 0; i--) {
      const diff = Math.round((pointer - timestamps[i]) / msPerDay);
      if (diff === 1) {
        current += 1;
        pointer = timestamps[i];
      } else {
        break;
      }
    }
  }

  return {
    total: timestamps.length,
    current,
    best
  };
}

function computeUserAchievementMetrics_(userId) {
  const user = getUserById_(userId) || {};
  const xpTotal = Number(user.xp || 0);

  const checkinRows = getAll_(sh_(SHEET_CHECKIN));
  const checkinDays = [];
  for (let i = 0; i < checkinRows.length; i++) {
    const row = checkinRows[i];
    if ((row[0] || '') !== userId) continue;
    const parsed = parseDateValue_(row[1]);
    if (parsed) {
      const day = Utilities.formatDate(parsed, 'UTC', 'yyyy-MM-dd');
      checkinDays.push(day);
    } else if (row[1]) {
      const raw = String(row[1]);
      const day = raw.length >= 10 ? raw.slice(0, 10) : raw;
      if (day) checkinDays.push(day);
    }
  }
  const streakStats = calculateStreakStats_(checkinDays);

  const progressRows = getAll_(sh_(SHEET_PROGRESS));
  const bestScores = {};
  for (let i = 0; i < progressRows.length; i++) {
    const row = progressRows[i];
    if ((row[0] || '') !== userId) continue;
    const moduleId = String(row[1] || '');
    const score = Math.max(0, Math.round(Number(row[2] || 0)));
    if (!bestScores[moduleId] || score > bestScores[moduleId]) {
      bestScores[moduleId] = score;
    }
  }

  let modulesCompleted = 0;
  let modulesHighScore = 0;
  const completedThreshold = 70;
  const highScoreThreshold = 90;
  Object.keys(bestScores).forEach(key => {
    const score = bestScores[key];
    if (score >= completedThreshold) modulesCompleted += 1;
    if (score >= highScoreThreshold) modulesHighScore += 1;
  });

  return {
    xpTotal,
    totalCheckins: streakStats.total,
    checkinCurrentStreak: streakStats.current,
    checkinBestStreak: streakStats.best,
    modulesCompleted,
    modulesHighScore,
    moduleBestScores: bestScores
  };
}

function evaluateAchievementStatus_(achievement, metrics) {
  const criteria = achievement && achievement.criteria ? achievement.criteria : {};
  const type = criteria.type || '';
  const targetRaw = Number(criteria.target || 0);
  const target = Number.isFinite(targetRaw) && targetRaw > 0 ? targetRaw : 1;
  let currentValue = 0;
  let achieved = false;
  let progressLabel = '';
  const extra = {};

  if (type === 'checkins_total') {
    currentValue = Math.max(0, Number(metrics.totalCheckins || 0));
    achieved = currentValue >= target;
    progressLabel = `${Math.min(currentValue, target)} de ${target} check-ins`;
  } else if (type === 'checkin_streak') {
    const best = Math.max(0, Number(metrics.checkinBestStreak || 0));
    const current = Math.max(0, Number(metrics.checkinCurrentStreak || 0));
    currentValue = best;
    achieved = current >= target;
    extra.currentStreak = current;
    extra.bestStreak = best;
    progressLabel = `Melhor sequência: ${best} dia${best === 1 ? '' : 's'}`;
  } else if (type === 'modules_high_score' || type === 'modules_completed') {
    const minScoreRaw = Number(criteria.minScore || (type === 'modules_high_score' ? 90 : 70));
    const minScore = Number.isFinite(minScoreRaw) ? minScoreRaw : (type === 'modules_high_score' ? 90 : 70);
    const scores = metrics.moduleBestScores || {};
    let count = 0;
    Object.keys(scores).forEach(key => {
      if (Number(scores[key] || 0) >= minScore) count += 1;
    });
    currentValue = count;
    achieved = count >= target;
    progressLabel = `${Math.min(count, target)} de ${target} módulos (${minScore}%+)`;
    extra.minScore = minScore;
  } else if (type === 'xp_total') {
    currentValue = Math.max(0, Number(metrics.xpTotal || 0));
    achieved = currentValue >= target;
    progressLabel = `${Math.min(currentValue, target)} / ${target} XP`;
    extra.currentXP = currentValue;
  } else {
    currentValue = 0;
    achieved = false;
  }

  const capped = target > 0 ? Math.min(currentValue, target) : currentValue;
  const progressPct = target > 0 ? Math.min(100, Math.round((capped / target) * 100)) : (achieved ? 100 : 0);

  return {
    achieved,
    currentValue,
    target,
    progressPct,
    progressLabel,
    extra
  };
}

function buildAchievementsOverview_(userId, existingRecords) {
  if (!userId) {
    return { achievements: [], summary: { total: 0, unlocked: 0 }, metrics: {} };
  }
  const metrics = computeUserAchievementMetrics_(userId);
  const records = Array.isArray(existingRecords) ? existingRecords : getUserAchievementRecords_(userId);
  const recordMap = {};
  for (let i = 0; i < records.length; i++) {
    const record = records[i];
    if (record && record.achievementId) {
      recordMap[record.achievementId] = record;
    }
  }

  const achievements = ACHIEVEMENTS.map(achievement => {
    const evaluation = evaluateAchievementStatus_(achievement, metrics);
    const record = recordMap[achievement.id];
    const unlocked = !!record;
    const unlockedAt = record ? record.unlockedAt : null;
    const rewardXP = Number(achievement.rewardXP || 0);
    const progressValue = evaluation.currentValue;
    const target = evaluation.target;
    return {
      id: achievement.id,
      title: achievement.title,
      description: achievement.description,
      category: achievement.category || '',
      icon: achievement.icon || '',
      rewardXP,
      unlocked,
      achieved: !!evaluation.achieved,
      unlockedAt,
      target,
      progressValue,
      progress: target > 0 ? Math.min(progressValue, target) : progressValue,
      progressPct: evaluation.progressPct,
      progressLabel: evaluation.progressLabel,
      readyToUnlock: !unlocked && evaluation.achieved,
      extra: evaluation.extra || {}
    };
  });

  const summary = {
    total: achievements.length,
    unlocked: achievements.filter(item => item.unlocked).length
  };

  return { achievements, summary, metrics };
}

function processAchievementsOnProgress_(userId) {
  if (!userId) {
    return { achievements: [], summary: { total: 0, unlocked: 0 }, metrics: {}, newlyUnlocked: [], bonusXP: 0 };
  }

  const achievementSheet = sh_(SHEET_USER_ACHIEVEMENT);
  const records = getUserAchievementRecords_(userId);
  const recordMap = {};
  for (let i = 0; i < records.length; i++) {
    const record = records[i];
    if (record && record.achievementId) {
      recordMap[record.achievementId] = record;
    }
  }

  const newlyUnlocked = [];
  let bonusXP = 0;
  const guardLimit = ACHIEVEMENTS.length + 3;
  let guard = 0;

  while (guard < guardLimit) {
    guard += 1;
    const metrics = computeUserAchievementMetrics_(userId);
    let unlockedThisLoop = false;

    for (let i = 0; i < ACHIEVEMENTS.length; i++) {
      const achievement = ACHIEVEMENTS[i];
      if (!achievement || recordMap[achievement.id]) continue;
      const evaluation = evaluateAchievementStatus_(achievement, metrics);
      if (!evaluation.achieved) continue;

      const rewardXP = Number(achievement.rewardXP || 0);
      const unlockedAt = nowISO_();
      achievementSheet.appendRow([userId, achievement.id, unlockedAt, rewardXP]);
      const record = { achievementId: achievement.id, unlockedAt, rewardXP };
      records.push(record);
      recordMap[achievement.id] = record;
      newlyUnlocked.push({
        id: achievement.id,
        title: achievement.title,
        description: achievement.description,
        category: achievement.category || '',
        icon: achievement.icon || '',
        rewardXP,
        unlockedAt
      });
      if (rewardXP > 0) {
        addUserXP_(userId, rewardXP);
        bonusXP += rewardXP;
      }
      unlockedThisLoop = true;
    }

    if (!unlockedThisLoop) {
      break;
    }
  }

  const overview = buildAchievementsOverview_(userId, records);
  return Object.assign({ newlyUnlocked, bonusXP }, overview);
}

function getUserAchievementsOverview(payload) {
  const token = payload && payload.token;
  const userId = payload && payload.userId;
  const session = requireSessionUser_(token, userId);
  const effectiveUserId = session.user.id;
  return buildAchievementsOverview_(effectiveUserId);
}

/** Atualiza XP do usuário (soma delta) */
function addUserXP_(userId, delta) {
  const s = sh_(SHEET_USERS);
  const rows = getAll_(s);
  for (let i=0;i<rows.length;i++){
    if (rows[i][0] === userId){
      const xp = Number(rows[i][5]||0) + Number(delta||0);
      s.getRange(i+2,6).setValue(xp);
      return xp;
    }
  }
  return null;
}

/** Obtém usuário por id */
function getUserById_(userId) {
  const s = sh_(SHEET_USERS);
  const rows = getAll_(s);
  for (let i=0;i<rows.length;i++){
    if (rows[i][0]===userId){
      const r = rows[i];
      return { id:r[0], name:r[1], email:r[2], isAdmin:!!r[4], xp:Number(r[5]||0) };
    }
  }
  return null;
}

function createSession_(userId) {
  if (!userId) throw new Error('userId inválido.');
  const token = Utilities.getUuid().replace(/-/g, '');
  const tokenHash = sha256_(token);
  const now = new Date();
  const expiresAt = new Date(now.getTime() + SESSION_DURATION_HOURS * 60 * 60 * 1000);

  const s = sh_(SHEET_SESSIONS);
  const rows = getAll_(s);
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];
    const rowUserId = row[0];
    const rowExpires = row[2];
    let shouldDelete = rowUserId === userId;
    if (!shouldDelete && rowExpires) {
      const expDate = new Date(rowExpires);
      if (!isNaN(expDate.getTime()) && expDate.getTime() <= now.getTime()) {
        shouldDelete = true;
      }
    }
    if (shouldDelete) {
      s.deleteRow(i + 2);
    }
  }

  s.appendRow([userId, tokenHash, expiresAt.toISOString(), nowISO_()]);
  return { token, expiresAt: expiresAt.toISOString() };
}

function findSessionByHash_(tokenHash) {
  const s = sh_(SHEET_SESSIONS);
  const rows = getAll_(s);
  for (let i = 0; i < rows.length; i++) {
    if ((rows[i][1] || '') === tokenHash) {
      return { row: i + 2, data: rows[i] };
    }
  }
  return null;
}

function getSessionUserByToken_(token) {
  const tokenStr = (token || '').toString().trim();
  if (!tokenStr) return null;
  const tokenHash = sha256_(tokenStr);
  const hit = findSessionByHash_(tokenHash);
  if (!hit) return null;

  const expiresAtRaw = hit.data[2];
  if (expiresAtRaw) {
    const expiresAtDate = parseDateValue_(expiresAtRaw);
    if (expiresAtDate && expiresAtDate.getTime() < Date.now()) {
      sh_(SHEET_SESSIONS).deleteRow(hit.row);
      return null;
    }
  }

  const userId = hit.data[0];
  if (!userId) return null;
  const user = getUserById_(userId);
  if (!user) {
    sh_(SHEET_SESSIONS).deleteRow(hit.row);
    return null;
  }

  const normalizedId = (user.id || '').toString().trim();
  const normalizedUser = Object.assign({}, user, { id: normalizedId });

  return { user: normalizedUser, sessionRow: hit.row };
}

function requireSessionUser_(token, expectedUserId) {
  const throwSessionError = () => {
    const err = new Error(SESSION_INVALID_MESSAGE);
    err.name = 'SessionError';
    throw err;
  };

  const tokenStr = (token || '').toString().trim();
  if (!tokenStr) {
    throwSessionError();
  }

  const session = getSessionUserByToken_(tokenStr);
  if (!session || !session.user) {
    throwSessionError();
  }

  const sessionUserId = (session.user.id || '').toString().trim();
  if (!sessionUserId) {
    throwSessionError();
  }

  if (expectedUserId !== undefined && expectedUserId !== null) {
    const expected = expectedUserId.toString().trim();
    if (expected && expected !== sessionUserId) {
      throwSessionError();
    }
  }

  const normalizedUser = Object.assign({}, session.user, { id: sessionUserId });
  return { user: normalizedUser, sessionRow: session.sessionRow };
}

/** =================== API PÚBLICA (chamada pelo HTML) =================== **/
function registerUser(payload) {
  setup_();
  const nameRaw  = (payload.name || '');
  const emailRaw = (payload.email||'');
  const passRaw  = (payload.password||'');
  const adminCode = (payload.adminCode||'').trim();

  const name  = nameRaw.trim();
  const email = emailRaw.trim().toLowerCase();
  const pass  = passRaw.trim();

  const errors = {};
  if (!name) {
    errors.name = 'Informe o nome completo.';
  } else if (name.length < 3) {
    errors.name = 'O nome deve ter pelo menos 3 caracteres.';
  }

  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!email) {
    errors.email = 'Informe o e-mail.';
  } else if (!emailPattern.test(email)) {
    errors.email = 'Formato de e-mail inválido.';
  } else if (findByEmail_(email)) {
    errors.email = 'E-mail já cadastrado.';
  }

  const missing = [];
  if (pass.length < 8) missing.push('no mínimo 8 caracteres');
  if (!/[A-Z]/.test(pass)) missing.push('uma letra maiúscula');
  if (!/[a-z]/.test(pass)) missing.push('uma letra minúscula');
  if (!/\d/.test(pass)) missing.push('um número');
  if (!/[^A-Za-z0-9]/.test(pass)) missing.push('um símbolo');
  if (missing.length) {
    const last = missing.pop();
    const msg = missing.length ? `A senha deve conter ${missing.join(', ')} e ${last}.` : `A senha deve conter ${last}.`;
    errors.password = msg;
  }

  if (Object.keys(errors).length) {
    return { ok:false, errors, message:'Revise os campos destacados.' };
  }

  const passHash = sha256_(pass);
  const isAdmin  = adminCode === ADMIN_SECURITY_CODE;

  const id = Utilities.getUuid();
  sh_(SHEET_USERS).appendRow([id, name, email, passHash, isAdmin, 0, nowISO_()]);

  const confirmationCode = generateConfirmationCode_();
  const expiresAt = new Date(Date.now() + 30 * 60 * 1000);
  const expiresISO = expiresAt.toISOString();
  const codeHash = hashConfirmationCode_(id, confirmationCode);
  saveConfirmationCode_(id, email, codeHash, expiresISO);
  sendConfirmationEmail_(email, name, confirmationCode);

  return {
    ok: true,
    requiresConfirmation: true,
    user: { id, name, email, isAdmin, xp: 0, level: 1 },
    confirmation: {
      userId: id,
      email,
      expiresAt: expiresISO
    },
    message: 'Enviamos um código de confirmação para o seu e-mail. Utilize-o para concluir o cadastro.'
  };
}

function loginUser(payload) {
  setup_();
  const email = (payload.email||'').trim().toLowerCase();
  const pass  = (payload.password||'').trim();
  if (!email || !pass) throw new Error('Informe e-mail e senha.');

  const hit = findByEmail_(email);
  if (!hit) throw new Error('Usuário não encontrado.');

  const passHash = sha256_(pass);
  if (hit.data[3] !== passHash) throw new Error('Senha inválida.');

  const user = {
    id: hit.data[0],
    name: hit.data[1],
    email: hit.data[2],
    isAdmin: !!hit.data[4],
    xp: Number(hit.data[5]||0)
  };
  user.level = 1 + Math.floor(user.xp / Number(getConfig_().xpPerLevel||100));

  const confirmationRecord = getConfirmationRecordByUserId_(user.id);
  if (confirmationRecord && !confirmationRecord.data.confirmedAt) {
    return {
      ok: false,
      requiresConfirmation: true,
      confirmation: {
        userId: user.id,
        email: user.email,
        expiresAt: confirmationRecord.data.expiresAt || ''
      },
      message: 'Confirme seu cadastro para acessar a plataforma.'
    };
  }

  const session = createSession_(user.id);
  return {
    ok: true,
    user,
    token: session.token,
    expiresAt: session.expiresAt,
  };
}

function confirmUserRegistration(payload) {
  setup_();
  const email = (payload.email || '').trim().toLowerCase();
  const code = (payload.code || '').toString().trim();
  if (!email || !code) throw new Error('Informe e-mail e código.');

  const hit = findByEmail_(email);
  if (!hit) throw new Error('Usuário não encontrado.');

  const userId = hit.data[0];
  const record = getConfirmationRecordByUserId_(userId);
  if (!record) {
    // Usuário antigo sem confirmação pendente
    const session = createSession_(userId);
    const user = {
      id: hit.data[0],
      name: hit.data[1],
      email: hit.data[2],
      isAdmin: !!hit.data[4],
      xp: Number(hit.data[5] || 0)
    };
    user.level = 1 + Math.floor(user.xp / Number(getConfig_().xpPerLevel || 100));
    return { ok: true, user, token: session.token, expiresAt: session.expiresAt };
  }

  if (record.data.confirmedAt) {
    const session = createSession_(userId);
    const user = {
      id: hit.data[0],
      name: hit.data[1],
      email: hit.data[2],
      isAdmin: !!hit.data[4],
      xp: Number(hit.data[5] || 0)
    };
    user.level = 1 + Math.floor(user.xp / Number(getConfig_().xpPerLevel || 100));
    return { ok: true, alreadyConfirmed: true, user, token: session.token, expiresAt: session.expiresAt };
  }

  const expiresAt = record.data.expiresAt ? new Date(record.data.expiresAt) : null;
  if (expiresAt && !isNaN(expiresAt.getTime()) && expiresAt.getTime() < Date.now()) {
    return { ok: false, expired: true, message: 'O código informado expirou. Solicite um novo envio.' };
  }

  const expectedHash = record.data.codeHash || '';
  const providedHash = hashConfirmationCode_(userId, code);
  if (!expectedHash || expectedHash !== providedHash) {
    throw new Error('Código inválido. Verifique o e-mail e tente novamente.');
  }

  markConfirmationAsConfirmed_(userId);
  const session = createSession_(userId);
  const user = {
    id: hit.data[0],
    name: hit.data[1],
    email: hit.data[2],
    isAdmin: !!hit.data[4],
    xp: Number(hit.data[5] || 0)
  };
  user.level = 1 + Math.floor(user.xp / Number(getConfig_().xpPerLevel || 100));
  return { ok: true, user, token: session.token, expiresAt: session.expiresAt };
}

function resendConfirmationCode(payload) {
  setup_();
  const email = (payload.email || '').trim().toLowerCase();
  if (!email) throw new Error('Informe o e-mail.');

  const hit = findByEmail_(email);
  if (!hit) throw new Error('Usuário não encontrado.');

  const userId = hit.data[0];
  const record = getConfirmationRecordByUserId_(userId);
  if (record && record.data.confirmedAt) {
    return { ok: true, alreadyConfirmed: true };
  }

  const code = generateConfirmationCode_();
  const expiresAt = new Date(Date.now() + 30 * 60 * 1000);
  const expiresISO = expiresAt.toISOString();
  const hash = hashConfirmationCode_(userId, code);
  saveConfirmationCode_(userId, email, hash, expiresISO);
  sendConfirmationEmail_(email, hit.data[1] || '', code);

  return {
    ok: true,
    confirmation: {
      userId,
      email,
      expiresAt: expiresISO
    }
  };
}

function resumeSession(token) {
  setup_();
  const tokenStr = (token || '').toString().trim();
  if (!tokenStr) return null;

  const tokenHash = sha256_(tokenStr);
  const hit = findSessionByHash_(tokenHash);
  if (!hit) return null;

  const expiresAt = hit.data[2];
  if (expiresAt) {
    const expDate = new Date(expiresAt);
    if (!isNaN(expDate.getTime()) && expDate.getTime() < Date.now()) {
      sh_(SHEET_SESSIONS).deleteRow(hit.row);
      return null;
    }
  }

  const userId = hit.data[0];
  const user = getUserById_(userId);
  if (!user) {
    sh_(SHEET_SESSIONS).deleteRow(hit.row);
    return null;
  }

  user.level = 1 + Math.floor(user.xp / Number(getConfig_().xpPerLevel||100));
  return user;
}

function logout(token) {
  setup_();
  const tokenStr = (token || '').toString().trim();
  if (!tokenStr) return { ok: true };

  const tokenHash = sha256_(tokenStr);
  const hit = findSessionByHash_(tokenHash);
  if (hit) {
    sh_(SHEET_SESSIONS).deleteRow(hit.row);
  }
  return { ok: true };
}

function checkin(payload) {
  const token = payload && payload.token;
  const userId = payload && payload.userId;
  const session = requireSessionUser_(token, userId);
  const effectiveUserId = session.user.id;

  const today = todayISO_();
  const s = sh_(SHEET_CHECKIN);
  const rows = getAll_(s);
  const checkinDays = [];
  let already = false;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if ((row[0] || '') !== effectiveUserId) continue;

    const parsed = parseDateValue_(row[1]);
    let day = '';
    if (parsed) {
      day = Utilities.formatDate(parsed, 'UTC', 'yyyy-MM-dd');
    } else if (row[1]) {
      const raw = String(row[1]);
      day = raw.length >= 10 ? raw.slice(0, 10) : raw;
    }

    if (day) {
      checkinDays.push(day);
      if (day === today) {
        already = true;
      }
    } else if ((row[1] || '').toString() === today) {
      already = true;
    }
  }

  if (already) return { ok:false, msg:'Presença já registrada hoje.' };

  const streakStatsBefore = calculateStreakStats_(checkinDays);
  let previousStreak = Number(streakStatsBefore.current || 0);
  if (!Number.isFinite(previousStreak) || previousStreak < 0) previousStreak = 0;

  const daysWithToday = checkinDays.slice();
  daysWithToday.push(today);
  const streakStatsAfter = calculateStreakStats_(daysWithToday);
  let currentStreak = Number(streakStatsAfter.current || 0);
  if (!Number.isFinite(currentStreak) || currentStreak <= 0) currentStreak = 1;

  const streakCap = CHECKIN_STREAK_XP_CAP;
  const xp = Math.max(1, Math.min(currentStreak, streakCap));
  const streakReset = previousStreak > 0 && currentStreak === 1;

  s.appendRow([effectiveUserId, today, xp, currentStreak]);
  const newXP = addUserXP_(effectiveUserId, xp);

  const achievements = processAchievementsOnProgress_(effectiveUserId);
  const cfg = getConfig_();
  const totalXP = achievements && achievements.metrics && Number.isFinite(Number(achievements.metrics.xpTotal))
    ? Number(achievements.metrics.xpTotal)
    : newXP + Number(achievements && achievements.bonusXP || 0);
  const levelInfo = computeLevelInfo_(totalXP, cfg);

  return {
    ok:true,
    xpGanho: xp,
    totalXP,
    level: levelInfo.level,
    xpPerLevel: levelInfo.xpPerLevel,
    xpToNextLevel: levelInfo.xpToNextLevel,
    nextLevel: levelInfo.nextLevel,
    achievementBonusXP: achievements && achievements.bonusXP || 0,
    achievementsUnlocked: achievements && achievements.newlyUnlocked || [],
    achievementsOverview: achievements ? {
      achievements: achievements.achievements,
      summary: achievements.summary,
      metrics: achievements.metrics
    } : null,
    streak: {
      current: currentStreak,
      previous: previousStreak,
      reset: streakReset,
      xpAwarded: xp,
      cap: streakCap
    }
  };
}

/** Envia resultado de uma atividade
 * payload: { userId, moduleId, scorePct (0-100) }
 * EarnedXP = round(xpOfModule * scorePct/100). Upsert em Progress (mantém melhor score).
 */
function submitActivity(payload) {
  const token = payload && payload.token;
  const userId   = payload && payload.userId;
  const session = requireSessionUser_(token, userId);
  const effectiveUserId = session.user.id;
  const invalidDataMessage = 'Dados inválidos (userId/moduleId).';
  if (!effectiveUserId) throw new Error(invalidDataMessage);

  const moduleInfo = getModuleById_(payload && payload.moduleId);
  if (!moduleInfo) throw new Error(invalidDataMessage);

  const moduleId = Number(moduleInfo.id);
  const xpMax = Number(moduleInfo.xpMax);
  if (!Number.isFinite(moduleId) || moduleId <= 0) throw new Error(invalidDataMessage);
  if (!Number.isFinite(xpMax) || xpMax < 0) throw new Error(invalidDataMessage);
  const rawScore = Number(payload && payload.scorePct);
  const normalizedScore = Number.isFinite(rawScore) ? rawScore : 0;
  const scorePct = Math.max(0, Math.min(100, normalizedScore));

  const earned = Math.round(xpMax * (scorePct/100));

  const s = sh_(SHEET_PROGRESS);
  const rows = getAll_(s);
  let foundRow = null;
  let oldScore = -1;
  let storedEarned = 0;

  for (let i=0;i<rows.length;i++){
    const r = rows[i];
    if (r[0]===effectiveUserId && Number(r[1])===moduleId){
      foundRow = i+2;
      oldScore = Number(r[2]||0);
      if (!Number.isFinite(oldScore)) oldScore = 0;
      storedEarned = Number(r[3] || 0);
      if (!Number.isFinite(storedEarned)) storedEarned = 0;
      break;
    }
  }

  let deltaXP = 0;
  if (foundRow){
    // Se melhorou a nota, atualiza e concede diferença de XP
    if (scorePct > oldScore){
      const oldEarned = Math.round(xpMax * (oldScore/100));
      deltaXP = Math.max(0, earned - oldEarned);
      s.getRange(foundRow, 3, 1, 3).setValues([[scorePct, earned, nowISO_()]]); // scorePct, earnedXP, completedAt
    } else {
      const expectedEarned = Math.round(xpMax * (oldScore/100));
      if (storedEarned !== expectedEarned) {
        s.getRange(foundRow, 4).setValue(expectedEarned);
      }
    }
  } else {
    // novo registro
    s.appendRow([effectiveUserId, moduleId, scorePct, earned, nowISO_()]);
    deltaXP = earned;
  }

  const cfg = getConfig_();
  let finalXP = null;
  if (deltaXP > 0) {
    finalXP = addUserXP_(effectiveUserId, deltaXP);
  }
  if (finalXP === null || finalXP === undefined) {
    const userInfo = getUserById_(effectiveUserId);
    finalXP = Number(userInfo?.xp || 0);
  }
  const achievements = processAchievementsOnProgress_(effectiveUserId);
  const totalXP = achievements && achievements.metrics && Number.isFinite(Number(achievements.metrics.xpTotal))
    ? Number(achievements.metrics.xpTotal)
    : finalXP + Number(achievements && achievements.bonusXP || 0);
  const levelInfo = computeLevelInfo_(totalXP, cfg);

  return {
    ok:true,
    deltaXP,
    totalXP,
    level: levelInfo.level,
    xpPerLevel: levelInfo.xpPerLevel,
    xpToNextLevel: levelInfo.xpToNextLevel,
    nextLevel: levelInfo.nextLevel,
    achievementBonusXP: achievements && achievements.bonusXP || 0,
    achievementsUnlocked: achievements && achievements.newlyUnlocked || [],
    achievementsOverview: achievements ? {
      achievements: achievements.achievements,
      summary: achievements.summary,
      metrics: achievements.metrics
    } : null
  };
}

/** Ranking (somente alunos, admins fora) */
function getRanking() {
  const users = getAll_(sh_(SHEET_USERS))
    .map(r => ({ id:r[0], name:r[1], email:r[2], isAdmin:!!r[4], xp:Number(r[5]||0) }))
    .filter(u => !u.isAdmin)
    .sort((a,b)=> (b.xp||0)-(a.xp||0));

  const cfg = getConfig_();
  const xpPerLevel = Number(cfg.xpPerLevel||100);

  return users.map((u,i)=>({
    pos: i+1,
    id: u.id, name: u.name, email: u.email,
    xp: u.xp, level: 1 + Math.floor(u.xp/xpPerLevel)
  }));
}

/** Estado do usuário (XP, level, concluídos) */
function getUserState(payload) {
  const token = payload && payload.token;
  const userId = payload && payload.userId;
  const session = requireSessionUser_(token, userId);
  const effectiveUserId = session.user.id;
  const u = getUserById_(effectiveUserId);
  if (!u) throw new Error('Usuário não encontrado.');
  const prog = getAll_(sh_(SHEET_PROGRESS)).filter(r=>r[0]===effectiveUserId);
  const concluidos = prog.length;
  const cfg = getConfig_();
  const levelInfo = computeLevelInfo_(u.xp, cfg);
  return {
    id: u.id, name: u.name, email: u.email, isAdmin: u.isAdmin,
    xp: u.xp,
    level: levelInfo.level,
    xpPerLevel: levelInfo.xpPerLevel,
    xpToNextLevel: levelInfo.xpToNextLevel,
    nextLevel: levelInfo.nextLevel,
    concluidos
  };
}

function getCheckinHistory(payload) {
  const token = payload && typeof payload === 'object' ? payload.token : null;
  const userId = payload && typeof payload === 'object' ? payload.userId : null;
  const session = requireSessionUser_(token, userId);
  const id = (session.user.id || '').toString().trim();

  const rows = getAll_(sh_(SHEET_CHECKIN));
  const entries = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if ((row[0] || '').toString() !== id) continue;

    const dateObj = parseDateValue_(row[1]);
    const timestamp = dateObj ? dateObj.getTime() : null;
    const iso = dateObj ? dateObj.toISOString() : ((row[1] || '').toString());
    const day = dateObj
      ? Utilities.formatDate(dateObj, 'UTC', 'yyyy-MM-dd')
      : (iso.length >= 10 ? iso.slice(0, 10) : '');

    entries.push({
      date: iso,
      day,
      xp: Number(row[2] || 0),
      streak: Number(row[3] || 0),
      timestamp
    });
  }

  entries.sort((a, b) => {
    const aTime = typeof a.timestamp === 'number' ? a.timestamp : -Infinity;
    const bTime = typeof b.timestamp === 'number' ? b.timestamp : -Infinity;
    return bTime - aTime;
  });

  return entries;
}

function getActivityHistory(payload) {
  const token = payload && typeof payload === 'object' ? payload.token : null;
  const userId = payload && typeof payload === 'object' ? payload.userId : null;
  const session = requireSessionUser_(token, userId);
  const id = (session.user.id || '').toString().trim();

  const rows = getAll_(sh_(SHEET_PROGRESS));
  const entries = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if ((row[0] || '').toString() !== id) continue;

    const moduleId = Number(row[1] || 0);
    const scorePct = Math.round(Number(row[2] || 0));
    const earnedXP = Number(row[3] || 0);
    const dateObj = parseDateValue_(row[4]);
    const timestamp = dateObj ? dateObj.getTime() : null;
    const iso = dateObj ? dateObj.toISOString() : ((row[4] || '').toString());
    const day = dateObj
      ? Utilities.formatDate(dateObj, 'UTC', 'yyyy-MM-dd')
      : (iso.length >= 10 ? iso.slice(0, 10) : '');

    entries.push({
      moduleId,
      scorePct,
      earnedXP,
      date: iso,
      day,
      timestamp,
      progressPct: 0
    });
  }

  if (!entries.length) return entries;

  const totalModules = MODULES.length;
  const asc = entries.slice().sort((a, b) => {
    const aTime = typeof a.timestamp === 'number' ? a.timestamp : Number.MAX_SAFE_INTEGER;
    const bTime = typeof b.timestamp === 'number' ? b.timestamp : Number.MAX_SAFE_INTEGER;
    return aTime - bTime;
  });
  const seenModules = new Set();

  asc.forEach(entry => {
    if (entry.moduleId) {
      seenModules.add(String(entry.moduleId));
    }
    const completed = seenModules.size;
    const pct = totalModules > 0 ? Math.min(100, Math.round((completed / totalModules) * 100)) : 0;
    entry.progressPct = pct;
  });

  entries.sort((a, b) => {
    const aTime = typeof a.timestamp === 'number' ? a.timestamp : -Infinity;
    const bTime = typeof b.timestamp === 'number' ? b.timestamp : -Infinity;
    return bTime - aTime;
  });

  return entries;
}

/** Embeds (Excel/PPT) */
function saveEmbeds(payload) {
  setup_();
  const token = payload && payload.token;
  const session = requireSessionUser_(token);
  if (!session.user.isAdmin) {
    throw new Error('Apenas administradores podem atualizar os materiais.');
  }

  const excel = payload && payload.excel ? payload.excel.toString().trim() : '';
  const ppt = payload && payload.ppt ? payload.ppt.toString().trim() : '';
  setKeyUrl_(SHEET_EMBEDS, 'excel', excel);
  setKeyUrl_(SHEET_EMBEDS, 'ppt', ppt);

  const resources = saveConfigList_(CONFIG_RECOMMENDED_RESOURCES, payload && payload.recommendedResources);
  const events = saveConfigList_(CONFIG_UPCOMING_EVENTS, payload && payload.upcomingEvents);

  return { ok: true, recommendedResources: resources, upcomingEvents: events };
}
function getEmbeds() {
  setup_();
  const rows = getAll_(sh_(SHEET_EMBEDS));
  const map = {};
  rows.forEach(r=> map[(r[0]||'')] = (r[1]||''));
  const resources = getConfigList_(CONFIG_RECOMMENDED_RESOURCES);
  const events = getConfigList_(CONFIG_UPCOMING_EVENTS);
  return {
    excel: map['excel']||'',
    ppt: map['ppt']||'',
    recommendedResources: resources,
    upcomingEvents: events
  };
}

function getAllAttachmentsMap_() {
  const rows = getAll_(sh_(SHEET_WALL_ATTACHMENTS));
  const map = {};
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const id = row[0] || '';
    if (!id) continue;
    map[id] = {
      row: i + 2,
      attachmentId: id,
      postId: row[1] || '',
      uploaderId: row[2] || '',
      fileId: row[3] || '',
      fileName: row[4] || '',
      mimeType: row[5] || '',
      fileUrl: row[6] || '',
      sizeBytes: Number(row[7] || 0),
      uploadedAt: row[8] || '',
      linkedAt: row[9] || ''
    };
  }
  return map;
}

function formatAttachmentForClient_(record) {
  if (!record) return null;
  return {
    id: record.attachmentId,
    postId: record.postId || '',
    fileId: record.fileId,
    name: record.fileName,
    mimeType: record.mimeType,
    url: record.fileUrl,
    sizeBytes: record.sizeBytes,
    uploadedAt: record.uploadedAt,
    linkedAt: record.linkedAt
  };
}

function validateAttachmentOwnership_(attachmentIds, userId) {
  const ids = normalizeStringArray_(Array.isArray(attachmentIds) ? attachmentIds : []);
  if (!ids.length) return [];
  const map = getAllAttachmentsMap_();
  const owned = [];
  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    const record = map[id];
    if (!record) {
      throw new Error('Não foi possível localizar um dos anexos enviados.');
    }
    if (record.uploaderId !== userId) {
      throw new Error('Há anexos que pertencem a outro usuário.');
    }
    if (record.postId) {
      throw new Error('Um dos anexos já foi vinculado a outra publicação.');
    }
    owned.push(record);
  }
  return owned;
}

function markAttachmentsAsLinked_(attachments, postId) {
  if (!Array.isArray(attachments) || !attachments.length) return;
  const sheet = sh_(SHEET_WALL_ATTACHMENTS);
  const now = nowISO_();
  attachments.forEach(att => {
    sheet.getRange(att.row, 2, 1, 2).setValues([[postId, now]]);
    att.postId = postId;
    att.linkedAt = now;
  });
}

function listCommunityWallEntries(payload) {
  setup_();
  let viewerId = null;
  let viewerIsAdmin = false;
  if (payload && payload.token) {
    try {
      const session = requireSessionUser_(payload.token);
      viewerId = session.user.id;
      viewerIsAdmin = !!session.user.isAdmin;
    } catch (err) {
      viewerId = null;
      viewerIsAdmin = false;
    }
  }

  const rows = getAll_(sh_(SHEET_WALL));
  const attachmentsMap = getAllAttachmentsMap_();
  const userMap = getUserMap_();
  const entries = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;
    const removedAt = row[8];
    if (removedAt) continue;
    const createdAtRaw = row[4] || '';
    const createdAtDate = parseDateValue_(createdAtRaw);
    const createdAt = createdAtDate ? createdAtDate.toISOString() : createdAtRaw.toString();
    const timestamp = createdAtDate ? createdAtDate.getTime() : null;
    const id = row[0] || '';
    const userId = row[1] || '';
    const authorName = row[2] || '';
    const message = row[3] || '';
    const visibilityRaw = (row[5] || 'public').toString();
    const visibility = visibilityRaw === 'private' ? 'private' : 'public';
    const targetIds = deserializeIdList_(row[6]);
    const attachmentIds = deserializeIdList_(row[7]);

    if (visibility === 'private') {
      const allowedViewers = new Set([userId]);
      for (let t = 0; t < targetIds.length; t++) {
        allowedViewers.add(targetIds[t]);
      }
      if (!viewerIsAdmin && (!viewerId || !allowedViewers.has(viewerId))) {
        continue;
      }
    }

    const attachmentList = [];
    for (let k = 0; k < attachmentIds.length; k++) {
      const attId = attachmentIds[k];
      const record = attachmentsMap[attId];
      if (!record) continue;
      if (record.postId && record.postId !== id) continue;
      attachmentList.push(formatAttachmentForClient_(record));
    }

    const targetNames = [];
    for (let t = 0; t < targetIds.length; t++) {
      const info = userMap[targetIds[t]];
      if (info && info.name) targetNames.push(info.name);
    }

    entries.push({
      id,
      userId,
      authorName,
      message,
      createdAt,
      timestamp,
      visibility,
      targetUserIds: targetIds,
      targetUserNames: targetNames,
      attachments: attachmentList
    });
  }

  entries.sort((a, b) => {
    const aTime = typeof a.timestamp === 'number' ? a.timestamp : -Infinity;
    const bTime = typeof b.timestamp === 'number' ? b.timestamp : -Infinity;
    return bTime - aTime;
  });

  return { entries, limit: COMMUNITY_WALL_CHAR_LIMIT };
}

function addCommunityWallEntry(payload) {
  setup_();
  const token = payload && payload.token;
  const session = requireSessionUser_(token);

  const rawMessage = payload && payload.message ? payload.message.toString() : '';
  const message = rawMessage.trim();
  if (!message) throw new Error('Escreva uma mensagem para publicar.');
  if (message.length > COMMUNITY_WALL_CHAR_LIMIT) {
    throw new Error(`A mensagem deve ter até ${COMMUNITY_WALL_CHAR_LIMIT} caracteres.`);
  }

  const normalized = message.replace(/\s+\n/g, '\n').replace(/\n{3,}/g, '\n\n');
  const visibilityRaw = (payload && payload.visibility ? payload.visibility.toString() : 'public');
  const visibility = visibilityRaw === 'private' ? 'private' : 'public';

  const userMap = getUserMap_();
  const targetsRaw = payload && (payload.targetUserIds || payload.targetIds || payload.targets);
  const targetIdsNormalized = normalizeStringArray_(Array.isArray(targetsRaw) ? targetsRaw : deserializeIdList_(targetsRaw || []));
  const sanitizedTargets = [];
  for (let i = 0; i < targetIdsNormalized.length; i++) {
    const targetId = targetIdsNormalized[i];
    if (!targetId) continue;
    if (targetId === session.user.id) continue;
    if (!userMap[targetId]) {
      throw new Error('Não foi possível localizar um dos destinatários selecionados.');
    }
    if (sanitizedTargets.indexOf(targetId) === -1) sanitizedTargets.push(targetId);
  }
  if (visibility === 'private' && !sanitizedTargets.length) {
    throw new Error('Selecione ao menos um destinatário para a publicação privada.');
  }

  const attachmentsRaw = payload && (payload.attachments || payload.attachmentIds);
  const attachmentIdList = normalizeStringArray_(Array.isArray(attachmentsRaw) ? attachmentsRaw : deserializeIdList_(attachmentsRaw || []));
  const attachmentRecords = validateAttachmentOwnership_(attachmentIdList, session.user.id);

  const sheet = sh_(SHEET_WALL);
  const id = Utilities.getUuid();
  const createdAt = nowISO_();
  const serializedTargets = serializeIdList_(sanitizedTargets);
  const serializedAttachments = serializeIdList_(attachmentRecords.map(att => att.attachmentId));
  sheet.appendRow([id, session.user.id, session.user.name || '', normalized, createdAt, visibility, serializedTargets, serializedAttachments, '', '']);
  markAttachmentsAsLinked_(attachmentRecords, id);

  const attachmentList = attachmentRecords.map(formatAttachmentForClient_).filter(Boolean);
  const targetNames = sanitizedTargets.map(t => (userMap[t] && userMap[t].name) ? userMap[t].name : '').filter(Boolean);

  return {
    ok: true,
    entry: {
      id,
      userId: session.user.id,
      authorName: session.user.name || '',
      message: normalized,
      createdAt,
      visibility,
      targetUserIds: sanitizedTargets,
      targetUserNames: targetNames,
      attachments: attachmentList
    }
  };
}

function removeCommunityWallEntry(payload) {
  setup_();
  const token = payload && payload.token;
  const session = requireSessionUser_(token);
  if (!session.user.isAdmin) {
    throw new Error('Apenas administradores podem remover publicações.');
  }

  const postId = payload && payload.postId ? payload.postId.toString().trim() : '';
  if (!postId) throw new Error('Identificador do post inválido.');

  const sheet = sh_(SHEET_WALL);
  const rows = getAll_(sheet);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if ((row[0] || '') === postId) {
      if (row[8]) {
        return { ok: true, alreadyRemoved: true };
      }
      sheet.getRange(i + 2, 9, 1, 2).setValues([[nowISO_(), session.user.id]]);
      return { ok: true };
    }
  }

  throw new Error('Publicação não encontrada.');
}

function uploadCommunityAttachment(payload) {
  setup_();
  const token = payload && payload.token;
  const session = requireSessionUser_(token);

  const base64Data = payload && payload.data ? payload.data.toString() : '';
  const fileNameRaw = payload && payload.fileName ? payload.fileName.toString() : 'anexo';
  const mimeType = payload && payload.mimeType ? payload.mimeType.toString() : 'application/octet-stream';
  if (!base64Data) throw new Error('Arquivo inválido.');

  const cleanName = fileNameRaw.trim() || 'anexo';
  const parts = base64Data.split(',');
  const encoded = parts.length > 1 ? parts[1] : parts[0];
  let bytes;
  try {
    bytes = Utilities.base64Decode(encoded);
  } catch (err) {
    throw new Error('Não foi possível processar o arquivo enviado.');
  }
  if (!bytes || !bytes.length) throw new Error('Arquivo vazio.');
  if (bytes.length > COMMUNITY_ATTACHMENT_MAX_BYTES) {
    throw new Error('O arquivo excede o limite de 10 MB.');
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(COMMUNITY_FILES_FOLDER_ID);
  } catch (err) {
    throw new Error('Pasta de armazenamento não encontrada. Verifique a configuração.');
  }

  const blob = Utilities.newBlob(bytes, mimeType, cleanName);
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    // ignora caso a configuração de compartilhamento não seja permitida
  }

  const attachmentId = Utilities.getUuid();
  const fileId = file.getId();
  const url = file.getUrl();
  const uploadedAt = nowISO_();
  const sheet = sh_(SHEET_WALL_ATTACHMENTS);
  sheet.appendRow([attachmentId, '', session.user.id, fileId, cleanName, mimeType, url, bytes.length, uploadedAt, '']);

  const record = {
    attachmentId,
    postId: '',
    uploaderId: session.user.id,
    fileId,
    fileName: cleanName,
    mimeType,
    fileUrl: url,
    sizeBytes: bytes.length,
    uploadedAt,
    linkedAt: ''
  };

  return {
    ok: true,
    attachment: formatAttachmentForClient_(record)
  };
}

function discardCommunityAttachment(payload) {
  setup_();
  const token = payload && payload.token;
  const session = requireSessionUser_(token);
  const attachmentId = payload && payload.attachmentId ? payload.attachmentId.toString().trim() : '';
  if (!attachmentId) throw new Error('Identificador do anexo inválido.');

  const map = getAllAttachmentsMap_();
  const record = map[attachmentId];
  if (!record) throw new Error('Anexo não encontrado.');
  if (record.uploaderId !== session.user.id && !session.user.isAdmin) {
    throw new Error('Você não tem permissão para remover este anexo.');
  }
  if (record.postId) {
    throw new Error('O anexo já foi vinculado a uma publicação e não pode ser removido.');
  }

  try {
    const file = DriveApp.getFileById(record.fileId);
    file.setTrashed(true);
  } catch (err) {
    // ignore falhas ao mover para lixeira
  }

  sh_(SHEET_WALL_ATTACHMENTS).deleteRow(record.row);
  return { ok: true };
}

function listShareableUsers(payload) {
  setup_();
  const token = payload && payload.token;
  const session = requireSessionUser_(token);
  const currentId = session.user.id;
  const rows = getAll_(sh_(SHEET_USERS));
  const users = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const id = row[0] || '';
    if (!id || id === currentId) continue;
    users.push({
      id,
      name: row[1] || '',
      email: row[2] || '',
      isAdmin: !!row[4]
    });
  }
  users.sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));
  return { ok: true, users };
}
function setKeyUrl_(sheetName, key, url) {
  const s = sh_(sheetName);
  const rows = getAll_(s);
  for (let i=0;i<rows.length;i++){
    if (rows[i][0]===key){ s.getRange(i+2,2).setValue(url); return; }
  }
  s.appendRow([key, url]);
}
