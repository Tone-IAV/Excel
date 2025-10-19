const SPREADSHEET_ID = '1DCzIOIcRBaJ3WJVQCOWg6KXyghjRGMJUHekV1fGygJ4';
const ADMIN_SECURITY_CODE = 'xbY4nu';

const SHEET_USERS = 'Users';
const SHEET_SESSIONS = 'Sessions';
const SHEET_CONFIRMATIONS = 'UserConfirmations';
const SHEET_PASSWORD_RESETS = 'PasswordResets';
const SHEET_USER_PROFILES = 'UserProfiles';

const USER_FILES_ROOT_FOLDER_NAME = 'ExcelPlatformUserFiles';

const SESSION_DURATION_HOURS = 24 * 7;
const SESSION_INVALID_MESSAGE = 'Sessão inválida ou expirada. Faça login novamente.';
const PASSWORD_RESET_WINDOW_MINUTES = 30;

const APP_TITLE = 'Plataforma Excel — Login';
const APP_SHELL_FILE = 'index';

const APP_PAGE_REGISTRY = (function buildAppPageRegistry_() {
  const definitions = [
    { id: 'home', title: 'Visão geral', file: 'pages/home' },
    { id: 'conta', title: 'Minha conta', file: 'pages/conta' }
  ];
  return createAppPageRegistry_(definitions);
})();

function doGet() {
  setup_();
  return renderAppShell_();
}

function include(filename) {
  return renderPartial_(filename);
}

function renderAppShell_() {
  return HtmlService.createTemplateFromFile(APP_SHELL_FILE)
    .evaluate()
    .setTitle(APP_TITLE);
}

function renderPartial_(filename) {
  const safeName = (filename || '').toString().trim();
  if (!safeName) {
    return '';
  }
  const loadFile = name =>
    HtmlService.createTemplateFromFile(name).evaluate().getContent();
  try {
    return loadFile(safeName);
  } catch (err) {
    const baseName = safeName.includes('/') ? safeName.split('/').pop() : safeName;
    const fallbackName = !safeName.startsWith('pages/') && baseName ? `pages/${baseName}` : '';
    if (fallbackName && fallbackName !== safeName) {
      try {
        return loadFile(fallbackName);
      } catch (fallbackErr) {
        throw err;
      }
    }
    throw err;
  }
}

function createAppPageRegistry_(definitions) {
  const normalizedList = [];
  const normalizedMap = {};
  (definitions || []).forEach((definition, index) => {
    const normalized = normalizeAppPageDefinition_(definition, index);
    if (!normalized) {
      return;
    }
    normalizedList.push(normalized);
    normalizedMap[normalized.id] = normalized;
  });
  return Object.freeze({
    list: Object.freeze(normalizedList),
    map: Object.freeze(normalizedMap)
  });
}

function normalizeAppPageDefinition_(definition, order) {
  if (!definition) {
    return null;
  }
  const id = (definition.id || '').toString().trim();
  if (!id) {
    return null;
  }
  const rawTitle = definition.title !== undefined ? definition.title : 'Página';
  const title = rawTitle === null ? 'Página' : rawTitle.toString();
  const rawFile = definition.file !== undefined ? definition.file : '';
  const file = (rawFile || '').toString().trim() || 'pages/' + id;
  const pageOrder = Number.isFinite(order) ? Number(order) : 0;
  return Object.freeze({ id, title, file, order: pageOrder });
}

function listAppPages_() {
  return APP_PAGE_REGISTRY.list.map(page => ({ id: page.id, title: page.title }));
}

function resolveAppPage_(pageId) {
  const safeId = (pageId || '').toString().trim();
  if (!safeId) {
    return null;
  }
  return APP_PAGE_REGISTRY.map[safeId] || null;
}

function setup_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ensureSheet = (name, headers) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    if (sheet.getFrozenRows() !== 1) {
      sheet.setFrozenRows(1);
    }
  };

  ensureSheet(SHEET_USERS, ['id', 'name', 'email', 'passHash', 'isAdmin', 'createdAt']);
  ensureSheet(SHEET_SESSIONS, ['userId', 'tokenHash', 'expiresAt', 'createdAt']);
  ensureSheet(SHEET_CONFIRMATIONS, [
    'userId',
    'email',
    'codeHash',
    'createdAt',
    'expiresAt',
    'confirmedAt',
    'lastSentAt',
    'pendingName',
    'pendingPassHash',
    'pendingIsAdmin'
  ]);
  ensureSheet(SHEET_PASSWORD_RESETS, [
    'userId',
    'email',
    'codeHash',
    'createdAt',
    'expiresAt',
    'usedAt',
    'lastSentAt',
    'attempts'
  ]);
  ensureSheet(
    SHEET_USER_PROFILES,
    ['userId', 'phone', 'role', 'bio', 'photoFileId', 'photoUrl', 'folderId', 'updatedAt']
  );
}

function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function sh_(name) {
  return ss_().getSheetByName(name);
}

function nowISO_() {
  return new Date().toISOString();
}

function toB64_(bytes) {
  return Utilities.base64Encode(bytes);
}

function sha256_(str) {
  return toB64_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8));
}

function generateSessionToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

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

function getAll_(sheet) {
  if (!sheet) return [];
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (!values || values.length <= 1) return [];
  return values.slice(1);
}

function normalizeBoolean_(value) {
  if (value === true) return true;
  if (value === false) return false;
  if (typeof value === 'number') return value !== 0;
  const text = (value || '').toString().trim().toLowerCase();
  if (!text) return false;
  return text === 'true' || text === '1' || text === 'yes' || text === 'sim';
}

function validatePasswordStrength_(password) {
  const value = (password || '').toString();
  if (!value || value.length < 8) {
    return 'A senha deve ter pelo menos 8 caracteres.';
  }
  const missing = [];
  if (!/[A-Z]/.test(value)) missing.push('uma letra maiúscula');
  if (!/[a-z]/.test(value)) missing.push('uma letra minúscula');
  if (!/\d/.test(value)) missing.push('um número');
  if (!/[^A-Za-z0-9]/.test(value)) missing.push('um símbolo');
  if (!missing.length) {
    return '';
  }
  const last = missing.pop();
  return missing.length
    ? `A senha deve conter ${missing.join(', ')} e ${last}.`
    : `A senha deve conter ${last}.`;
}

function mapUserRow_(row) {
  if (!row) return null;
  return {
    id: row[0] || '',
    name: row[1] || '',
    email: row[2] || '',
    passHash: row[3] || '',
    isAdmin: normalizeBoolean_(row[4]),
    createdAt: row[5] || ''
  };
}

function mapUserProfileRow_(row) {
  if (!row) {
    return {
      userId: '',
      phone: '',
      role: '',
      bio: '',
      photoFileId: '',
      photoUrl: '',
      folderId: '',
      updatedAt: ''
    };
  }
  return {
    userId: row[0] || '',
    phone: row[1] || '',
    role: row[2] || '',
    bio: row[3] || '',
    photoFileId: row[4] || '',
    photoUrl: row[5] || '',
    folderId: row[6] || '',
    updatedAt: row[7] || ''
  };
}

function normalizeUserForClient_(user) {
  if (!user) return null;
  return {
    id: user.id || '',
    name: user.name || '',
    email: user.email || '',
    isAdmin: !!user.isAdmin
  };
}

function findByEmail_(email) {
  const safeEmail = (email || '').toString().trim().toLowerCase();
  if (!safeEmail) return null;
  const sheet = sh_(SHEET_USERS);
  const rows = getAll_(sheet);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowEmail = (row[2] || '').toString().trim().toLowerCase();
    if (rowEmail && rowEmail === safeEmail) {
      return {
        row: i + 2,
        data: mapUserRow_(row)
      };
    }
  }
  return null;
}

function getUserById_(userId) {
  if (!userId) return null;
  const sheet = sh_(SHEET_USERS);
  const rows = getAll_(sheet);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if ((row[0] || '') === userId) {
      return {
        row: i + 2,
        data: mapUserRow_(row)
      };
    }
  }
  return null;
}

function getUserProfile_(userId) {
  if (!userId) return null;
  const sheet = sh_(SHEET_USER_PROFILES);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return null;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if ((row[0] || '') === userId) {
      return {
        row: i + 1,
        data: mapUserProfileRow_(row)
      };
    }
  }
  return null;
}

function ensureUserProfile_(userId) {
  const profile = getUserProfile_(userId);
  if (profile) {
    return profile;
  }
  const defaults = mapUserProfileRow_(null);
  defaults.userId = userId;
  defaults.updatedAt = nowISO_();
  const sheet = sh_(SHEET_USER_PROFILES);
  sheet.appendRow([
    defaults.userId,
    defaults.phone,
    defaults.role,
    defaults.bio,
    defaults.photoFileId,
    defaults.photoUrl,
    defaults.folderId,
    defaults.updatedAt
  ]);
  return getUserProfile_(userId);
}

function saveUserProfile_(userId, updates) {
  if (!userId) {
    throw new Error('Usuário inválido para atualizar perfil.');
  }
  const sheet = sh_(SHEET_USER_PROFILES);
  if (!sheet) {
    throw new Error('Aba de perfis não encontrada.');
  }
  const record = ensureUserProfile_(userId);
  const current = record && record.data ? record.data : mapUserProfileRow_(null);
  const merged = Object.assign({}, current, updates || {}, { userId, updatedAt: nowISO_() });
  const values = [
    merged.userId,
    merged.phone || '',
    merged.role || '',
    merged.bio || '',
    merged.photoFileId || '',
    merged.photoUrl || '',
    merged.folderId || '',
    merged.updatedAt || ''
  ];
  if (record && record.row) {
    sheet.getRange(record.row, 1, 1, values.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }
  return merged;
}

function ensureUserFilesRootFolder_() {
  const existing = DriveApp.getFoldersByName(USER_FILES_ROOT_FOLDER_NAME);
  if (existing.hasNext()) {
    return existing.next();
  }
  return DriveApp.createFolder(USER_FILES_ROOT_FOLDER_NAME);
}

function getOrCreateUserFolder_(userId) {
  const root = ensureUserFilesRootFolder_();
  const folderName = 'user-' + userId;
  let folder = null;
  const iterator = root.getFoldersByName(folderName);
  if (iterator.hasNext()) {
    folder = iterator.next();
  } else {
    folder = root.createFolder(folderName);
  }
  return folder;
}

function buildDriveViewUrl_(fileId) {
  if (!fileId) return '';
  return 'https://drive.google.com/uc?export=view&id=' + fileId;
}

function buildDriveDownloadUrl_(fileId) {
  if (!fileId) return '';
  return 'https://drive.google.com/uc?export=download&id=' + fileId;
}

function requireValidSession_(token) {
  const safeToken = (token || '').toString().trim();
  if (!safeToken) {
    throw new Error(SESSION_INVALID_MESSAGE);
  }
  const sessionRecord = findSessionByToken_(safeToken);
  if (!sessionRecord || !sessionRecord.data) {
    throw new Error(SESSION_INVALID_MESSAGE);
  }
  const expiresAt = sessionRecord.data.expiresAt ? new Date(sessionRecord.data.expiresAt) : null;
  if (!expiresAt || isNaN(expiresAt.getTime()) || expiresAt.getTime() <= Date.now()) {
    removeSessionByToken_(safeToken);
    throw new Error(SESSION_INVALID_MESSAGE);
  }
  const userHit = getUserById_(sessionRecord.data.userId);
  if (!userHit || !userHit.data) {
    removeSessionByToken_(safeToken);
    throw new Error(SESSION_INVALID_MESSAGE);
  }
  return {
    token: safeToken,
    session: sessionRecord,
    userHit
  };
}

function updateUserPasswordHash_(userId, passHash) {
  const hit = getUserById_(userId);
  if (!hit || !hit.row) {
    throw new Error('Usuário não encontrado para atualização de senha.');
  }
  const sheet = sh_(SHEET_USERS);
  sheet.getRange(hit.row, 4).setValue(passHash);
}

function cleanupSessions_(sheet, options) {
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return;
  const now = options && options.now instanceof Date ? options.now : new Date();
  const userId = options && options.userId ? options.userId : '';
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const rowUserId = row[0] || '';
    const expiresRaw = row[2] || '';
    const expiresDate = expiresRaw ? new Date(expiresRaw) : null;
    const isExpired = !expiresDate || isNaN(expiresDate.getTime()) || expiresDate.getTime() <= now.getTime();
    const shouldRemove = isExpired || (userId && rowUserId === userId);
    if (shouldRemove) {
      sheet.deleteRow(i + 1);
    }
  }
}

function createSession_(userId) {
  if (!userId) {
    throw new Error('ID de usuário inválido para criar sessão.');
  }
  const sheet = sh_(SHEET_SESSIONS);
  if (!sheet) {
    throw new Error('Aba de sessões não encontrada.');
  }
  const now = new Date();
  cleanupSessions_(sheet, { userId, now });
  const token = generateSessionToken_();
  const tokenHash = sha256_(token);
  const expiresAt = new Date(now.getTime() + SESSION_DURATION_HOURS * 60 * 60 * 1000);
  sheet.appendRow([userId, tokenHash, expiresAt.toISOString(), now.toISOString()]);
  return { token, expiresAt: expiresAt.toISOString() };
}

function removeSessionByToken_(token) {
  if (!token) return false;
  const hash = sha256_(token);
  const sheet = sh_(SHEET_SESSIONS);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return false;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowHash = (row[1] || '').toString();
    if (rowHash === hash) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function findSessionByToken_(token) {
  if (!token) return null;
  const hash = sha256_(token);
  const sheet = sh_(SHEET_SESSIONS);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return null;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowHash = (row[1] || '').toString();
    if (rowHash === hash) {
      return {
        row: i + 1,
        data: {
          userId: row[0] || '',
          tokenHash: rowHash,
          expiresAt: row[2] || '',
          createdAt: row[3] || ''
        }
      };
    }
  }
  return null;
}

function revokeOtherSessions_(userId, tokenToKeep) {
  if (!userId) return;
  const sheet = sh_(SHEET_SESSIONS);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) return;
  const keepHash = tokenToKeep ? sha256_(tokenToKeep) : '';
  const now = new Date();
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const rowUserId = row[0] || '';
    if (rowUserId !== userId) {
      continue;
    }
    const rowHash = (row[1] || '').toString();
    const expiresRaw = row[2] || '';
    const expiresDate = expiresRaw ? new Date(expiresRaw) : null;
    const isExpired = !expiresDate || isNaN(expiresDate.getTime()) || expiresDate.getTime() <= now.getTime();
    const shouldRemove = isExpired || (keepHash && rowHash !== keepHash);
    if (shouldRemove) {
      sheet.deleteRow(i + 1);
    }
  }
}

function mapConfirmationRow_(row) {
  if (!row) {
    return {
      userId: '',
      email: '',
      codeHash: '',
      createdAt: '',
      expiresAt: '',
      confirmedAt: '',
      lastSentAt: '',
      pendingName: '',
      pendingPassHash: '',
      pendingIsAdmin: false
    };
  }
  return {
    userId: row[0] || '',
    email: (row[1] || '').toString().trim().toLowerCase(),
    codeHash: row[2] || '',
    createdAt: row[3] || '',
    expiresAt: row[4] || '',
    confirmedAt: row[5] || '',
    lastSentAt: row[6] || '',
    pendingName: row[7] || '',
    pendingPassHash: row[8] || '',
    pendingIsAdmin: normalizeBoolean_(row[9])
  };
}

function getConfirmationRecordByUserId_(userId) {
  if (!userId) return null;
  const sheet = sh_(SHEET_CONFIRMATIONS);
  const rows = getAll_(sheet);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if ((row[0] || '') === userId) {
      return {
        row: i + 2,
        data: mapConfirmationRow_(row)
      };
    }
  }
  return null;
}

function mapPasswordResetRow_(row) {
  if (!row) {
    return {
      userId: '',
      email: '',
      codeHash: '',
      createdAt: '',
      expiresAt: '',
      usedAt: '',
      lastSentAt: '',
      attempts: 0
    };
  }
  return {
    userId: row[0] || '',
    email: (row[1] || '').toString().trim().toLowerCase(),
    codeHash: row[2] || '',
    createdAt: row[3] || '',
    expiresAt: row[4] || '',
    usedAt: row[5] || '',
    lastSentAt: row[6] || '',
    attempts: Number(row[7]) || 0
  };
}

function getPasswordResetRecordByUserId_(userId) {
  if (!userId) return null;
  const sheet = sh_(SHEET_PASSWORD_RESETS);
  const rows = getAll_(sheet);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if ((row[0] || '') === userId) {
      return {
        row: i + 2,
        data: mapPasswordResetRow_(row)
      };
    }
  }
  return null;
}

function getPasswordResetRecordByEmail_(email) {
  const safeEmail = (email || '').toString().trim().toLowerCase();
  if (!safeEmail) return null;
  const sheet = sh_(SHEET_PASSWORD_RESETS);
  const rows = getAll_(sheet);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowEmail = (row[1] || '').toString().trim().toLowerCase();
    if (rowEmail && rowEmail === safeEmail) {
      return {
        row: i + 2,
        data: mapPasswordResetRow_(row)
      };
    }
  }
  return null;
}

function savePasswordResetRecord_(userId, email, options) {
  const safeUserId = (userId || '').toString().trim();
  const normalizedEmail = (email || '').toString().trim().toLowerCase();
  if (!safeUserId || !normalizedEmail) {
    throw new Error('Dados inválidos para recuperação de senha.');
  }
  const sheet = sh_(SHEET_PASSWORD_RESETS);
  const now = nowISO_();
  const opts = options || {};
  let existing = getPasswordResetRecordByUserId_(safeUserId);
  if (!existing || (existing.data.email || '') !== normalizedEmail) {
    const byEmail = getPasswordResetRecordByEmail_(normalizedEmail);
    if (byEmail) {
      existing = byEmail;
    }
  }
  const existingData = existing ? existing.data : mapPasswordResetRow_(null);
  const payload = [
    safeUserId,
    normalizedEmail,
    Object.prototype.hasOwnProperty.call(opts, 'codeHash') ? opts.codeHash : existingData.codeHash,
    Object.prototype.hasOwnProperty.call(opts, 'createdAt')
      ? opts.createdAt
      : existingData.createdAt || now,
    Object.prototype.hasOwnProperty.call(opts, 'expiresAt') ? opts.expiresAt : existingData.expiresAt,
    Object.prototype.hasOwnProperty.call(opts, 'usedAt') ? opts.usedAt : existingData.usedAt,
    Object.prototype.hasOwnProperty.call(opts, 'lastSentAt') ? opts.lastSentAt : existingData.lastSentAt || now,
    Object.prototype.hasOwnProperty.call(opts, 'attempts') ? opts.attempts : existingData.attempts || 0
  ];
  if (existing && existing.row) {
    sheet.getRange(existing.row, 1, 1, payload.length).setValues([payload]);
  } else {
    sheet.appendRow(payload);
  }
}

function getConfirmationRecordByEmail_(email) {
  const safeEmail = (email || '').toString().trim().toLowerCase();
  if (!safeEmail) return null;
  const sheet = sh_(SHEET_CONFIRMATIONS);
  const rows = getAll_(sheet);
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowEmail = (row[1] || '').toString().trim().toLowerCase();
    if (rowEmail && rowEmail === safeEmail) {
      return {
        row: i + 2,
        data: mapConfirmationRow_(row)
      };
    }
  }
  return null;
}

function saveConfirmationCode_(userId, email, codeHash, expiresAtISO, options) {
  const safeUserId = (userId || '').toString().trim();
  if (!safeUserId) {
    throw new Error('ID de usuário inválido para confirmação.');
  }
  const sheet = sh_(SHEET_CONFIRMATIONS);
  const now = nowISO_();
  const opts = options || {};
  const normalizedEmail = (email || '').toString().trim().toLowerCase();
  let existing = getConfirmationRecordByUserId_(safeUserId);
  if (!existing || (existing.data.email || '') !== normalizedEmail) {
    const byEmail = getConfirmationRecordByEmail_(normalizedEmail);
    if (byEmail) {
      existing = byEmail;
    }
  }
  const existingData = existing ? existing.data : null;
  const payload = [
    safeUserId,
    normalizedEmail,
    codeHash || '',
    opts.createdAt || (existingData ? existingData.createdAt : now),
    expiresAtISO || (existingData ? existingData.expiresAt : ''),
    opts.confirmedAt !== undefined
      ? opts.confirmedAt
      : (existingData ? existingData.confirmedAt : ''),
    now,
    Object.prototype.hasOwnProperty.call(opts, 'pendingName')
      ? opts.pendingName
      : (existingData ? existingData.pendingName : ''),
    Object.prototype.hasOwnProperty.call(opts, 'pendingPassHash')
      ? opts.pendingPassHash
      : (existingData ? existingData.pendingPassHash : ''),
    Object.prototype.hasOwnProperty.call(opts, 'pendingIsAdmin')
      ? !!opts.pendingIsAdmin
      : (existingData ? existingData.pendingIsAdmin : false)
  ];
  if (existing && existing.row) {
    sheet.getRange(existing.row, 1, 1, payload.length).setValues([payload]);
  } else {
    sheet.appendRow(payload);
  }
}

function markConfirmationAsConfirmed_(userId) {
  const record = getConfirmationRecordByUserId_(userId);
  if (!record) return;
  const now = nowISO_();
  saveConfirmationCode_(userId, record.data.email, record.data.codeHash, record.data.expiresAt, {
    confirmedAt: now,
    pendingName: record.data.pendingName,
    pendingPassHash: '',
    pendingIsAdmin: record.data.pendingIsAdmin,
    createdAt: record.data.createdAt
  });
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

function buildPasswordResetEmailBodies_(name, code) {
  const template = HtmlService.createTemplateFromFile('password-reset-email');
  template.name = name || 'Participante';
  template.code = code;
  const html = template.evaluate().getContent();
  const subject = 'Redefina sua senha na Plataforma Excel';
  const plain = [
    'Olá ' + (name || 'participante') + ',',
    '',
    'Recebemos uma solicitação para redefinir sua senha na Plataforma Excel.',
    'Utilize o código abaixo dentro de 30 minutos para concluir o processo:',
    '',
    code,
    '',
    'Se você não solicitou esta redefinição, ignore este e-mail.'
  ].join('\n');
  return { subject, plain, html };
}

function sendPasswordResetEmail_(email, name, code) {
  if (!email || !code) return;
  const bodies = buildPasswordResetEmailBodies_(name, code);
  MailApp.sendEmail({
    to: email,
    subject: bodies.subject,
    htmlBody: bodies.html,
    body: bodies.plain,
    name: 'Plataforma Excel',
    noReply: true
  });
}

function registerUser(payload) {
  setup_();
  const name = (payload && payload.name ? payload.name : '').toString().trim();
  const email = (payload && payload.email ? payload.email : '').toString().trim().toLowerCase();
  const password = (payload && payload.password ? payload.password : '').toString().trim();
  const adminCode = (payload && payload.adminCode ? payload.adminCode : '').toString().trim();

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
  }

  const passwordError = validatePasswordStrength_(password);
  if (passwordError) {
    errors.password = passwordError;
  }

  if (Object.keys(errors).length) {
    return { ok: false, errors, message: 'Revise os campos destacados.' };
  }

  const existingUser = findByEmail_(email);
  if (existingUser && existingUser.data) {
    return { ok: false, message: 'E-mail já cadastrado. Faça login para continuar.' };
  }

  const confirmationRecord = getConfirmationRecordByEmail_(email);
  if (confirmationRecord && confirmationRecord.data && confirmationRecord.data.confirmedAt) {
    return { ok: false, message: 'Este e-mail já foi confirmado. Faça login para continuar.' };
  }

  const userId = confirmationRecord && confirmationRecord.data.userId
    ? confirmationRecord.data.userId
    : Utilities.getUuid();

  const passHash = sha256_(password);
  const isAdmin = adminCode && adminCode === ADMIN_SECURITY_CODE;
  const code = generateConfirmationCode_();
  const expiresAt = new Date(Date.now() + 30 * 60 * 1000);
  const expiresISO = expiresAt.toISOString();

  saveConfirmationCode_(userId, email, hashConfirmationCode_(userId, code), expiresISO, {
    pendingName: name,
    pendingPassHash: passHash,
    pendingIsAdmin: isAdmin,
    confirmedAt: '',
    createdAt: nowISO_()
  });

  sendConfirmationEmail_(email, name, code);

  return {
    ok: true,
    requiresConfirmation: true,
    confirmation: {
      userId,
      email,
      expiresAt: expiresISO
    },
    message: 'Enviamos um código de confirmação para o seu e-mail. Utilize-o para ativar o acesso.'
  };
}

function loginUser(payload) {
  setup_();
  const email = (payload && payload.email ? payload.email : '').toString().trim().toLowerCase();
  const password = (payload && payload.password ? payload.password : '').toString().trim();
  if (!email || !password) {
    throw new Error('Informe e-mail e senha.');
  }

  const hit = findByEmail_(email);
  if (hit && hit.data) {
    const storedHash = hit.data.passHash || '';
    const providedHash = sha256_(password);
    if (!storedHash || storedHash !== providedHash) {
      throw new Error('Senha incorreta.');
    }
    const session = createSession_(hit.data.id);
    const user = normalizeUserForClient_(hit.data);
    return {
      ok: true,
      user,
      token: session.token,
      expiresAt: session.expiresAt
    };
  }

  const record = getConfirmationRecordByEmail_(email);
  if (record && record.data && !record.data.confirmedAt) {
    const expiresAt = record.data.expiresAt || '';
    const expiresDate = expiresAt ? new Date(expiresAt) : null;
    const expired = !expiresDate || isNaN(expiresDate.getTime()) || expiresDate.getTime() < Date.now();
    return {
      ok: false,
      requiresConfirmation: true,
      confirmation: {
        userId: record.data.userId,
        email: record.data.email,
        expiresAt: record.data.expiresAt
      },
      message: expired
        ? 'O código informado expirou. Solicite um novo envio.'
        : 'Confirme o código enviado para o seu e-mail para concluir o acesso.'
    };
  }

  throw new Error('Usuário não encontrado.');
}

function confirmSignup(payload) {
  setup_();
  const email = (payload && payload.email ? payload.email : '').toString().trim().toLowerCase();
  const code = (payload && payload.code ? payload.code : '').toString().trim();
  if (!email || !code) {
    throw new Error('Informe e-mail e código de confirmação.');
  }

  const record = getConfirmationRecordByEmail_(email);
  if (!record) {
    throw new Error('Usuário não encontrado.');
  }

  const userId = record.data.userId;
  if (!userId) {
    throw new Error('Usuário não encontrado.');
  }

  const expiresAt = record.data.expiresAt ? new Date(record.data.expiresAt) : null;
  if (!expiresAt || isNaN(expiresAt.getTime()) || expiresAt.getTime() < Date.now()) {
    return {
      ok: false,
      requiresConfirmation: true,
      confirmation: {
        userId,
        email: record.data.email,
        expiresAt: record.data.expiresAt
      },
      message: 'O código informado expirou. Solicite um novo envio.'
    };
  }

  const providedHash = hashConfirmationCode_(userId, code);
  if (!record.data.codeHash || providedHash !== record.data.codeHash) {
    throw new Error('Código inválido. Verifique o e-mail e tente novamente.');
  }

  let userHit = getUserById_(userId);
  if (!userHit || !userHit.data) {
    const pendingName = record.data.pendingName || '';
    const pendingPassHash = record.data.pendingPassHash || '';
    if (!pendingName || !pendingPassHash) {
      throw new Error('Não foi possível concluir o cadastro. Solicite um novo registro.');
    }
    const sheet = sh_(SHEET_USERS);
    sheet.appendRow([
      userId,
      pendingName,
      record.data.email,
      pendingPassHash,
      record.data.pendingIsAdmin ? true : false,
      nowISO_()
    ]);
    userHit = getUserById_(userId);
  }

  markConfirmationAsConfirmed_(userId);
  const session = createSession_(userId);
  const user = normalizeUserForClient_(userHit.data);
  return {
    ok: true,
    user,
    token: session.token,
    expiresAt: session.expiresAt
  };
}

function resendConfirmationCode(payload) {
  setup_();
  const email = (payload && payload.email ? payload.email : '').toString().trim().toLowerCase();
  if (!email) {
    throw new Error('Informe o e-mail.');
  }

  const userHit = findByEmail_(email);
  if (userHit && userHit.data) {
    return {
      ok: true,
      message: 'Esta conta já foi confirmada. Faça login para continuar.'
    };
  }

  const record = getConfirmationRecordByEmail_(email);
  if (!record) {
    throw new Error('Usuário não encontrado.');
  }
  if (record.data.confirmedAt) {
    return {
      ok: true,
      message: 'Esta conta já foi confirmada. Faça login para continuar.'
    };
  }

  const code = generateConfirmationCode_();
  const expiresAt = new Date(Date.now() + 30 * 60 * 1000);
  const expiresISO = expiresAt.toISOString();
  saveConfirmationCode_(record.data.userId, email, hashConfirmationCode_(record.data.userId, code), expiresISO, {
    pendingName: record.data.pendingName,
    pendingPassHash: record.data.pendingPassHash,
    pendingIsAdmin: record.data.pendingIsAdmin,
    confirmedAt: record.data.confirmedAt,
    createdAt: record.data.createdAt
  });
  sendConfirmationEmail_(email, record.data.pendingName || '', code);
  return {
    ok: true,
    confirmation: {
      userId: record.data.userId,
      email,
      expiresAt: expiresISO
    },
    message: 'Enviamos um novo código para o seu e-mail.'
  };
}

function requestPasswordReset(payload) {
  setup_();
  const email = (payload && payload.email ? payload.email : '').toString().trim().toLowerCase();
  if (!email) {
    throw new Error('Informe o e-mail cadastrado.');
  }

  const userHit = findByEmail_(email);
  if (!userHit || !userHit.data) {
    return {
      ok: true,
      message: 'Se o e-mail estiver cadastrado, você receberá um código de redefinição em instantes.'
    };
  }

  const now = new Date();
  const nowIso = now.toISOString();
  const record =
    getPasswordResetRecordByUserId_(userHit.data.id) || getPasswordResetRecordByEmail_(email);
  if (record && record.data && record.data.lastSentAt) {
    const lastSentAt = new Date(record.data.lastSentAt);
    if (!isNaN(lastSentAt.getTime()) && now.getTime() - lastSentAt.getTime() < 60 * 1000) {
      return {
        ok: true,
        message:
          'Aguarde alguns instantes e verifique sua caixa de entrada. Um código recente pode estar disponível.'
      };
    }
  }

  const code = generateConfirmationCode_();
  const expiresAt = new Date(now.getTime() + PASSWORD_RESET_WINDOW_MINUTES * 60 * 1000);
  const attempts = record && record.data ? (Number(record.data.attempts) || 0) + 1 : 1;
  const createdAt = record && record.data && record.data.createdAt ? record.data.createdAt : nowIso;
  savePasswordResetRecord_(userHit.data.id, email, {
    codeHash: hashConfirmationCode_(userHit.data.id, code),
    createdAt,
    expiresAt: expiresAt.toISOString(),
    usedAt: '',
    lastSentAt: nowIso,
    attempts
  });
  sendPasswordResetEmail_(email, userHit.data.name || '', code);
  return {
    ok: true,
    message: 'Enviamos um código para o seu e-mail. Verifique a caixa de entrada e o spam.'
  };
}

function completePasswordReset(payload) {
  setup_();
  const email = (payload && payload.email ? payload.email : '').toString().trim().toLowerCase();
  const code = (payload && payload.code ? payload.code : '').toString().trim();
  const password = (payload && payload.password ? payload.password : '').toString().trim();

  const errors = {};
  if (!email) {
    errors.email = 'Informe o e-mail cadastrado.';
  }
  if (!code) {
    errors.code = 'Informe o código recebido.';
  }
  const passwordError = validatePasswordStrength_(password);
  if (passwordError) {
    errors.password = passwordError;
  }
  if (Object.keys(errors).length) {
    return { ok: false, errors, message: 'Revise os campos destacados.' };
  }

  const userHit = findByEmail_(email);
  if (!userHit || !userHit.data) {
    return { ok: false, message: 'Não encontramos uma conta ativa para este e-mail.' };
  }

  const record =
    getPasswordResetRecordByUserId_(userHit.data.id) || getPasswordResetRecordByEmail_(email);
  if (!record || !record.data || !record.data.codeHash) {
    return { ok: false, message: 'Solicite um novo código de redefinição.' };
  }

  if (record.data.usedAt) {
    return {
      ok: false,
      errors: { code: 'O código informado já foi utilizado. Solicite um novo envio.' },
      message: 'O código informado já foi utilizado. Solicite um novo envio.'
    };
  }

  const expiresAt = record.data.expiresAt ? new Date(record.data.expiresAt) : null;
  if (!expiresAt || isNaN(expiresAt.getTime()) || expiresAt.getTime() < Date.now()) {
    return {
      ok: false,
      errors: { code: 'O código informado expirou. Solicite um novo envio.' },
      message: 'O código informado expirou. Solicite um novo envio.'
    };
  }

  const expectedHash = record.data.codeHash;
  const providedHash = hashConfirmationCode_(userHit.data.id, code);
  if (!expectedHash || providedHash !== expectedHash) {
    return {
      ok: false,
      errors: { code: 'Código inválido. Confira o e-mail e tente novamente.' },
      message: 'Código inválido. Confira o e-mail e tente novamente.'
    };
  }

  const passHash = sha256_(password);
  updateUserPasswordHash_(userHit.data.id, passHash);

  const now = new Date();
  const nowIso = now.toISOString();
  savePasswordResetRecord_(userHit.data.id, email, {
    codeHash: '',
    createdAt: record.data.createdAt || nowIso,
    expiresAt: '',
    usedAt: nowIso,
    lastSentAt: record.data.lastSentAt || nowIso,
    attempts: record.data.attempts || 0
  });
  cleanupSessions_(sh_(SHEET_SESSIONS), { userId: userHit.data.id, now });

  return {
    ok: true,
    message: 'Senha atualizada com sucesso. Faça login novamente com a nova senha.'
  };
}

function restoreSession(payload) {
  setup_();
  const token = (payload && payload.token ? payload.token : '').toString().trim();
  if (!token) {
    return { ok: false, message: SESSION_INVALID_MESSAGE };
  }
  const sessionRecord = findSessionByToken_(token);
  if (!sessionRecord) {
    return { ok: false, message: SESSION_INVALID_MESSAGE };
  }
  const expiresAt = sessionRecord.data.expiresAt ? new Date(sessionRecord.data.expiresAt) : null;
  if (!expiresAt || isNaN(expiresAt.getTime()) || expiresAt.getTime() <= Date.now()) {
    removeSessionByToken_(token);
    return { ok: false, message: SESSION_INVALID_MESSAGE };
  }
  const userHit = getUserById_(sessionRecord.data.userId);
  if (!userHit || !userHit.data) {
    removeSessionByToken_(token);
    return { ok: false, message: SESSION_INVALID_MESSAGE };
  }
  const user = normalizeUserForClient_(userHit.data);
  return {
    ok: true,
    user,
    token,
    expiresAt: sessionRecord.data.expiresAt
  };
}

function logoutSession(payload) {
  setup_();
  const token = (payload && payload.token ? payload.token : '').toString().trim();
  if (!token) {
    return { ok: true };
  }
  removeSessionByToken_(token);
  return { ok: true };
}

function getAppPages() {
  setup_();
  return listAppPages_();
}

function getAppPageContent(pageId) {
  setup_();
  const page = resolveAppPage_(pageId);
  if (!page) {
    throw new Error('Página solicitada não está disponível.');
  }
  return renderPartial_(page.file);
}

function getAccountDetails(payload) {
  setup_();
  const token = payload && payload.token ? payload.token : '';
  const context = requireValidSession_(token);
  const userId = context.userHit.data.id;
  const profileRecord = ensureUserProfile_(userId);
  let profileData = profileRecord && profileRecord.data ? profileRecord.data : mapUserProfileRow_(null);

  let folder = null;
  let folderUrl = '';
  let folderId = profileData.folderId || '';
  if (folderId) {
    try {
      folder = DriveApp.getFolderById(folderId);
      folderUrl = folder.getUrl();
    } catch (err) {
      folder = null;
      folderId = '';
    }
  }
  if (!folder) {
    folder = getOrCreateUserFolder_(userId);
    folderId = folder.getId();
    folderUrl = folder.getUrl();
    profileData = saveUserProfile_(userId, { folderId });
  }

  let photoUrl = profileData.photoUrl || '';
  if (!photoUrl && profileData.photoFileId) {
    photoUrl = buildDriveViewUrl_(profileData.photoFileId);
    profileData = saveUserProfile_(userId, { photoUrl });
  }

  const responseProfile = {
    phone: profileData.phone || '',
    role: profileData.role || '',
    bio: profileData.bio || '',
    photoFileId: profileData.photoFileId || '',
    photoUrl: photoUrl || '',
    photoDownloadUrl: profileData.photoFileId ? buildDriveDownloadUrl_(profileData.photoFileId) : '',
    folderId: profileData.folderId || folderId || '',
    folderUrl: folderUrl || '',
    updatedAt: profileData.updatedAt || ''
  };

  return {
    ok: true,
    user: normalizeUserForClient_(context.userHit.data),
    profile: responseProfile
  };
}

function updateAccountDetails(payload) {
  setup_();
  const token = payload && payload.token ? payload.token : '';
  const context = requireValidSession_(token);
  const userHit = context.userHit;
  const userId = userHit.data.id;

  const name = (payload && payload.name ? payload.name : '').toString().trim();
  const emailRaw = (payload && payload.email ? payload.email : '').toString().trim();
  const email = emailRaw.toLowerCase();
  const phone = (payload && payload.phone ? payload.phone : '').toString().trim();
  const role = (payload && payload.role ? payload.role : '').toString().trim();
  const bio = (payload && payload.bio ? payload.bio : '').toString().trim();
  const photoPayload = payload && payload.photo ? payload.photo : null;
  const removePhoto = payload && payload.removePhoto ? true : false;

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
  } else {
    const hit = findByEmail_(email);
    if (hit && hit.data && hit.data.id !== userId) {
      errors.email = 'Este e-mail já está em uso por outra conta.';
    }
  }

  if (phone && phone.length > 40) {
    errors.phone = 'O telefone pode ter no máximo 40 caracteres.';
  }
  if (role && role.length > 120) {
    errors.role = 'O cargo ou função pode ter no máximo 120 caracteres.';
  }
  if (bio && bio.length > 800) {
    errors.bio = 'A biografia pode ter no máximo 800 caracteres.';
  }

  let photoData = null;
  if (photoPayload && photoPayload.data) {
    const base64 = (photoPayload.data || '').toString();
    if (!base64) {
      errors.photo = 'Não foi possível ler a imagem selecionada.';
    } else {
      let cleanBase64 = base64;
      const commaIndex = cleanBase64.indexOf(',');
      if (commaIndex >= 0) {
        cleanBase64 = cleanBase64.substring(commaIndex + 1);
      }
      try {
        const bytes = Utilities.base64Decode(cleanBase64);
        const mimeType = photoPayload.type || 'image/png';
        const fileName = photoPayload.name || 'foto-perfil.png';
        photoData = Utilities.newBlob(bytes, mimeType, fileName);
      } catch (err) {
        errors.photo = 'Não foi possível processar a imagem enviada.';
      }
    }
  }

  if (Object.keys(errors).length) {
    return { ok: false, message: 'Revise os campos destacados.', errors };
  }

  const usersSheet = sh_(SHEET_USERS);
  usersSheet.getRange(userHit.row, 2).setValue(name);
  usersSheet.getRange(userHit.row, 3).setValue(email);
  userHit.data.name = name;
  userHit.data.email = email;

  const profileUpdates = {
    phone,
    role,
    bio
  };

  let profile = ensureUserProfile_(userId).data;
  let folder = null;
  let folderId = profile.folderId || '';
  if (folderId) {
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (err) {
      folderId = '';
    }
  }
  if (!folder) {
    folder = getOrCreateUserFolder_(userId);
    folderId = folder.getId();
    profileUpdates.folderId = folderId;
  }

  let photoUrl = profile.photoUrl || '';
  let photoDownloadUrl = profile.photoFileId ? buildDriveDownloadUrl_(profile.photoFileId) : '';
  if (photoData) {
    const file = folder.createFile(photoData);
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
    const fileId = file.getId();
    photoUrl = buildDriveViewUrl_(fileId);
    photoDownloadUrl = buildDriveDownloadUrl_(fileId);
    profileUpdates.photoFileId = fileId;
    profileUpdates.photoUrl = photoUrl;
    if (profile.photoFileId) {
      try {
        DriveApp.getFileById(profile.photoFileId).setTrashed(true);
      } catch (err) {
        // Ignored: file might have been removed manually.
      }
    }
  } else if (removePhoto && profile.photoFileId) {
    try {
      DriveApp.getFileById(profile.photoFileId).setTrashed(true);
    } catch (err) {
      // Ignorado: arquivo já removido.
    }
    profileUpdates.photoFileId = '';
    profileUpdates.photoUrl = '';
    photoUrl = '';
    photoDownloadUrl = '';
  }

  profile = saveUserProfile_(userId, profileUpdates);
  if (!photoUrl && profile.photoFileId) {
    photoUrl = buildDriveViewUrl_(profile.photoFileId);
    photoDownloadUrl = buildDriveDownloadUrl_(profile.photoFileId);
    profile = saveUserProfile_(userId, { photoUrl });
  }

  let folderUrl = '';
  const resolvedFolderId = profile.folderId || folderId;
  if (resolvedFolderId) {
    try {
      folderUrl = DriveApp.getFolderById(resolvedFolderId).getUrl();
    } catch (err) {
      folderUrl = '';
    }
  }

  return {
    ok: true,
    message: 'Dados atualizados com sucesso.',
    user: normalizeUserForClient_(userHit.data),
    profile: {
      phone: profile.phone || '',
      role: profile.role || '',
      bio: profile.bio || '',
      photoFileId: profile.photoFileId || '',
      photoUrl: photoUrl || '',
      photoDownloadUrl: photoDownloadUrl || '',
      folderId: profile.folderId || folderId || '',
      folderUrl,
      updatedAt: profile.updatedAt || ''
    }
  };
}

function updateAccountPassword(payload) {
  setup_();
  const token = payload && payload.token ? payload.token : '';
  const context = requireValidSession_(token);
  const userHit = context.userHit;

  const currentPassword = (payload && payload.currentPassword ? payload.currentPassword : '')
    .toString()
    .trim();
  const newPassword = (payload && payload.newPassword ? payload.newPassword : '').toString().trim();
  const confirmPassword = (payload && payload.confirmPassword ? payload.confirmPassword : '')
    .toString()
    .trim();

  const errors = {};
  if (!currentPassword) {
    errors.currentPassword = 'Informe a senha atual.';
  }
  if (!newPassword) {
    errors.newPassword = 'Informe a nova senha.';
  }
  if (!confirmPassword) {
    errors.confirmPassword = 'Confirme a nova senha.';
  }

  if (newPassword && confirmPassword && newPassword !== confirmPassword) {
    errors.confirmPassword = 'A confirmação deve ser igual à nova senha.';
  }

  if (Object.keys(errors).length) {
    return { ok: false, message: 'Revise os campos destacados.', errors };
  }

  const currentHash = sha256_(currentPassword);
  if (!userHit.data.passHash || currentHash !== userHit.data.passHash) {
    return {
      ok: false,
      message: 'A senha atual informada não confere.',
      errors: { currentPassword: 'A senha atual informada não confere.' }
    };
  }

  if (newPassword === currentPassword) {
    return {
      ok: false,
      message: 'A nova senha deve ser diferente da senha atual.',
      errors: { newPassword: 'Escolha uma senha diferente da atual.' }
    };
  }

  const passwordError = validatePasswordStrength_(newPassword);
  if (passwordError) {
    return { ok: false, message: passwordError, errors: { newPassword: passwordError } };
  }

  const newHash = sha256_(newPassword);
  updateUserPasswordHash_(userHit.data.id, newHash);
  userHit.data.passHash = newHash;

  revokeOtherSessions_(userHit.data.id, context.token);

  return {
    ok: true,
    message: 'Senha atualizada com sucesso. Utilize a nova senha no próximo acesso.'
  };
}
