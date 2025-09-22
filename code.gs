/** =================== CONFIG GERAL =================== **/
const SPREADSHEET_ID = '1DCzIOIcRBaJ3WJVQCOWg6KXyghjRGMJUHekV1fGygJ4';
const ADMIN_SECURITY_CODE = 'xbY4nu'; // código exigido p/ criar admins

// Abas usadas
const SHEET_USERS    = 'Users';     // id, name, email, passHash, isAdmin, xp, createdAt
const SHEET_PROGRESS = 'Progress';  // userId, moduleId, scorePct, earnedXP, completedAt
const SHEET_CHECKIN  = 'Checkins';  // userId, dateISO, xp
const SHEET_CONFIG   = 'Config';    // key, value
const SHEET_EMBEDS   = 'Embeds';    // key, url

/** =================== BOOTSTRAP =================== **/
function onOpen() {
  setup_();
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Plataforma Excel')
    .addItem('Recriar estrutura', 'setup_')
    .addToUi();
}

function doGet() {
  setup_();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Plataforma Gamificada — Curso de Excel')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
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

  ensure(SHEET_USERS,    ['id','name','email','passHash','isAdmin','xp','createdAt']);
  ensure(SHEET_PROGRESS, ['userId','moduleId','scorePct','earnedXP','completedAt']);
  ensure(SHEET_CHECKIN,  ['userId','dateISO','xp']);
  ensure(SHEET_CONFIG,   ['key','value']);
  ensure(SHEET_EMBEDS,   ['key','url']);

  // Defaults de config
  const cfg = getConfig_();
  if (!cfg.xpCheckin) setConfig_('xpCheckin','5');
  if (!cfg.xpPerLevel) setConfig_('xpPerLevel','100');
}

/** =================== HELPERS =================== **/
function ss_()              { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function sh_(name)          { return ss_().getSheetByName(name); }
function nowISO_()          { return new Date().toISOString(); }
function todayISO_()        { return Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd'); }
function toB64_(bytes)      { return Utilities.base64Encode(bytes); }
function sha256_(str)       { return toB64_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8)); }
function getAll_(sheet)     { const r=sheet.getDataRange().getValues(); return r.length>1 ? r.slice(1) : []; }
function findByEmail_(email){ const rows=getAll_(sh_(SHEET_USERS)); for (let i=0;i<rows.length;i++){ if ((rows[i][2]||'').toLowerCase()===email.toLowerCase()){ return { row:i+2, data:rows[i] }; } } return null; }
function getConfig_()       { const rows=getAll_(sh_(SHEET_CONFIG)); const m={}; rows.forEach(r=>m[(r[0]||'').toString()]=(r[1]||'').toString()); return m; }
function setConfig_(k,v)    { const s=sh_(SHEET_CONFIG); const rows=getAll_(s); for (let i=0;i<rows.length;i++){ if (rows[i][0]==k){ s.getRange(i+2,2).setValue(v); return; } } s.appendRow([k,v]); }

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

/** =================== API PÚBLICA (chamada pelo HTML) =================== **/
function registerUser(payload) {
  setup_();
  const name  = (payload.name || '').trim();
  const email = (payload.email||'').trim().toLowerCase();
  const pass  = (payload.password||'').trim();
  const adminCode = (payload.adminCode||'').trim();

  if (!name || !email || !pass) throw new Error('Preencha nome, e-mail e senha.');

  if (findByEmail_(email)) throw new Error('E-mail já cadastrado.');

  const passHash = sha256_(pass);
  const isAdmin  = adminCode === ADMIN_SECURITY_CODE;

  const id = Utilities.getUuid();
  sh_(SHEET_USERS).appendRow([id, name, email, passHash, isAdmin, 0, nowISO_()]);

  return {
    id, name, email, isAdmin, xp: 0,
    level: 1,
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
  return user;
}

function checkin(payload) {
  const userId = payload.userId;
  if (!userId) throw new Error('userId inválido.');

  const today = todayISO_();
  const s = sh_(SHEET_CHECKIN);
  const rows = getAll_(s);
  const already = rows.some(r => r[0]===userId && r[1]===today);
  if (already) return { ok:false, msg:'Presença já registrada hoje.' };

  const xp = Number(getConfig_().xpCheckin||5);
  s.appendRow([userId, today, xp]);
  const newXP = addUserXP_(userId, xp);

  return { ok:true, xpGanho: xp, totalXP: newXP };
}

/** Envia resultado de uma atividade
 * payload: { userId, moduleId, scorePct (0-100), maxXP }
 * EarnedXP = round(maxXP * scorePct/100). Upsert em Progress (mantém melhor score).
 */
function submitActivity(payload) {
  const userId   = payload.userId;
  const moduleId = Number(payload.moduleId);
  const scorePct = Math.max(0, Math.min(100, Number(payload.scorePct||0)));
  const maxXP    = Math.max(0, Number(payload.maxXP||0));

  if (!userId || !moduleId) throw new Error('Dados inválidos (userId/moduleId).');

  const earned = Math.round(maxXP * (scorePct/100));

  const s = sh_(SHEET_PROGRESS);
  const rows = getAll_(s);
  let foundRow = null;
  let oldScore = -1;

  for (let i=0;i<rows.length;i++){
    const r = rows[i];
    if (r[0]===userId && Number(r[1])===moduleId){
      foundRow = i+2;
      oldScore = Number(r[2]||0);
      break;
    }
  }

  let deltaXP = 0;
  if (foundRow){
    // Se melhorou a nota, atualiza e concede diferença de XP
    if (scorePct > oldScore){
      const oldEarned = Math.round(maxXP * (oldScore/100));
      deltaXP = (earned - oldEarned);
      s.getRange(foundRow, 3, 1, 3).setValues([[scorePct, earned, nowISO_()]]); // scorePct, earnedXP, completedAt
    }
  } else {
    // novo registro
    s.appendRow([userId, moduleId, scorePct, earned, nowISO_()]);
    deltaXP = earned;
  }

  let totalXP = null;
  if (deltaXP>0) totalXP = addUserXP_(userId, deltaXP);
  const cfg = getConfig_();
  const level = 1 + Math.floor((totalXP!==null? totalXP : getUserById_(userId)?.xp || 0) / Number(cfg.xpPerLevel||100));

  return {
    ok:true,
    deltaXP,
    totalXP: totalXP ?? getUserById_(userId)?.xp,
    level
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
  const userId = payload.userId;
  const u = getUserById_(userId);
  if (!u) throw new Error('Usuário não encontrado.');
  const prog = getAll_(sh_(SHEET_PROGRESS)).filter(r=>r[0]===userId);
  const concluidos = prog.length;
  const cfg = getConfig_();
  return {
    id: u.id, name: u.name, email: u.email, isAdmin: u.isAdmin,
    xp: u.xp, level: 1 + Math.floor(u.xp / Number(cfg.xpPerLevel||100)),
    concluidos
  };
}

/** Embeds (Excel/PPT) */
function saveEmbeds(payload) {
  const { excel, ppt } = payload;
  setKeyUrl_(SHEET_EMBEDS, 'excel', excel||'');
  setKeyUrl_(SHEET_EMBEDS, 'ppt', ppt||'');
  return { ok:true };
}
function getEmbeds() {
  const rows = getAll_(sh_(SHEET_EMBEDS));
  const map = {};
  rows.forEach(r=> map[(r[0]||'')] = (r[1]||''));
  return { excel: map['excel']||'', ppt: map['ppt']||'' };
}
function setKeyUrl_(sheetName, key, url) {
  const s = sh_(sheetName);
  const rows = getAll_(s);
  for (let i=0;i<rows.length;i++){
    if (rows[i][0]===key){ s.getRange(i+2,2).setValue(url); return; }
  }
  s.appendRow([key, url]);
}
