/** =================== CONFIG GERAL =================== **/
const SPREADSHEET_ID = '1DCzIOIcRBaJ3WJVQCOWg6KXyghjRGMJUHekV1fGygJ4';
const ADMIN_SECURITY_CODE = 'xbY4nu'; // código exigido p/ criar admins

// Abas usadas
const SHEET_USERS            = 'Users';            // id, name, email, passHash, isAdmin, xp, createdAt
const SHEET_PROGRESS         = 'Progress';         // userId, moduleId, scorePct, earnedXP, completedAt
const SHEET_CHECKIN          = 'Checkins';         // userId, dateISO, xp
const SHEET_CONFIG           = 'Config';           // key, value
const SHEET_EMBEDS           = 'Embeds';           // key, url
const SHEET_SESSIONS         = 'Sessions';         // userId, tokenHash, expiresAt, createdAt
const SHEET_USER_ACHIEVEMENT = 'UserAchievements'; // userId, achievementId, unlockedAt, rewardXP
const SHEET_WALL             = 'Wall';             // id, userId, authorName, message, createdAt, removedAt, removedBy

const CONFIG_RECOMMENDED_RESOURCES = 'recommended_resources';
const CONFIG_UPCOMING_EVENTS       = 'upcoming_events';

const COMMUNITY_WALL_CHAR_LIMIT = 240;

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
  ensure(SHEET_CHECKIN,          ['userId','dateISO','xp']);
  ensure(SHEET_CONFIG,           ['key','value']);
  ensure(SHEET_EMBEDS,           ['key','url']);
  ensure(SHEET_SESSIONS,         ['userId','tokenHash','expiresAt','createdAt']);
  ensure(SHEET_USER_ACHIEVEMENT, ['userId','achievementId','unlockedAt','rewardXP']);
  ensure(SHEET_WALL,             ['id','userId','authorName','message','createdAt','removedAt','removedBy']);

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

  const session = createSession_(id);
  return {
    ok: true,
    user: {
      id, name, email, isAdmin, xp: 0,
      level: 1,
    },
    token: session.token,
    expiresAt: session.expiresAt,
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
  const session = createSession_(user.id);
  return {
    ok: true,
    user,
    token: session.token,
    expiresAt: session.expiresAt,
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
  const already = rows.some(r => r[0]===effectiveUserId && r[1]===today);
  if (already) return { ok:false, msg:'Presença já registrada hoje.' };

  const cfg = getConfig_();
  const xp = Number(cfg.xpCheckin||5);
  s.appendRow([effectiveUserId, today, xp]);
  const newXP = addUserXP_(effectiveUserId, xp);

  const achievements = processAchievementsOnProgress_(effectiveUserId);
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
    } : null
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

function listCommunityWallEntries() {
  setup_();
  const rows = getAll_(sh_(SHEET_WALL));
  const entries = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;
    const removedAt = row[5];
    if (removedAt) continue;
    const createdAtRaw = row[4] || '';
    const createdAtDate = parseDateValue_(createdAtRaw);
    const createdAt = createdAtDate ? createdAtDate.toISOString() : createdAtRaw.toString();
    const timestamp = createdAtDate ? createdAtDate.getTime() : null;
    entries.push({
      id: row[0] || '',
      userId: row[1] || '',
      authorName: row[2] || '',
      message: row[3] || '',
      createdAt,
      timestamp
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
  const sheet = sh_(SHEET_WALL);
  const id = Utilities.getUuid();
  const createdAt = nowISO_();
  sheet.appendRow([id, session.user.id, session.user.name || '', normalized, createdAt, '', '']);

  return {
    ok: true,
    entry: {
      id,
      userId: session.user.id,
      authorName: session.user.name || '',
      message: normalized,
      createdAt
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
      if (row[5]) {
        return { ok: true, alreadyRemoved: true };
      }
      sheet.getRange(i + 2, 6, 1, 2).setValues([[nowISO_(), session.user.id]]);
      return { ok: true };
    }
  }

  throw new Error('Publicação não encontrada.');
}
function setKeyUrl_(sheetName, key, url) {
  const s = sh_(sheetName);
  const rows = getAll_(s);
  for (let i=0;i<rows.length;i++){
    if (rows[i][0]===key){ s.getRange(i+2,2).setValue(url); return; }
  }
  s.appendRow([key, url]);
}
