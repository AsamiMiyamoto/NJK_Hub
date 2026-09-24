/**
 * NJK社内システム 共通基盤 バックエンド処理 (Code.gs)
 * 正本スプレッドシートID: 1oxCXcQj-428e7O7MNO64uLOjAeG6hu7Ioa9v3_5rbGg
 */

const DATA_SS_ID = '1oxCXcQj-428e7O7MNO64uLOjAeG6hu7Ioa9v3_5rbGg';

function getCommonSpreadsheet() {
  const ss = SpreadsheetApp.openById(DATA_SS_ID);
  if (!ss) throw new Error("スプレッドシートが開けません。IDを確認してください。");
  return ss;
}

// 管理画面へのアクセスを許可するGoogleアカウント一覧
const ADMIN_ALLOWED_EMAILS_ = ['admin@j-shelter.com'];

// 社員マスタの列番号（1始まり）
const EMP_COL_PASSWORD_ = 8;     // H列：パスワード
const EMP_COL_MUST_CHANGE_ = 9;  // I列：PW変更要
const EMP_COL_PW_CHANGED_AT_ = 10; // J列：PW変更日時

// CacheServiceのキー接頭辞（SSOトークンとログインセッションの取り違え防止）
const SSO_TOKEN_PREFIX_ = 'SSO_';
const SESSION_PREFIX_ = 'SESSION_';
const SSO_TOKEN_TTL_SEC_ = 300;
const SESSION_TTL_SEC_ = 21600; // 6時間（CacheServiceの上限）

function isAdmin_() {
  return ADMIN_ALLOWED_EMAILS_.indexOf(Session.getActiveUser().getEmail()) !== -1;
}

function assertAdmin_() {
  if (!isAdmin_()) throw new Error("管理者権限がありません。");
}

function doGet(e) {
  const isAdminRequest = e && e.parameter && e.parameter.admin === 'true';

  if (isAdminRequest) {
    if (!isAdmin_()) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family: sans-serif; padding: 40px; text-align:center; color:#555;">' +
        '<h2>アクセス権がありません</h2>' +
        '<p>この画面は管理者用アカウントでログインしている場合のみ表示されます。</p>' +
        '</div>'
      ).setTitle('アクセス拒否');
    }
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('NJK社内システム 共通基盤（管理画面）')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('login')
    .evaluate()
    .setTitle('Next Journey Keynote ログイン')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ----------------------------------------------------
// ハッシュ処理
// ----------------------------------------------------

/**
 * SHA-256ハッシュ化（IdPアプリと同一方式）
 */
function computeHash_(text) {
  const signature = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
  return signature.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

/**
 * パスワードをソルト付きでハッシュ化する
 * 形式：v2$<ソルト>$<SHA-256(ソルト+PW)の16進>
 */
function hashPassword_(plain) {
  const salt = Utilities.getUuid().replace(/-/g, '');
  return 'v2$' + salt + '$' + computeHash_(salt + plain);
}

/**
 * 保存済みハッシュとの照合（新形式 v2$…、旧形式：ソルトなしSHA-256の64桁16進 の両方に対応）
 */
function verifyPassword_(plain, stored) {
  const s = String(stored || '');
  if (s.indexOf('v2$') === 0) {
    const parts = s.split('$');
    if (parts.length !== 3) return false;
    return computeHash_(parts[1] + plain) === parts[2];
  }
  if (/^[0-9a-f]{64}$/i.test(s)) return computeHash_(plain) === s.toLowerCase();
  return false;
}

/**
 * パスワードルールの検証（違反時は日本語メッセージで例外）
 */
function validatePasswordRule_(newPw, employeeId) {
  if (typeof newPw !== 'string' || newPw.length < 8) throw new Error("パスワードは8文字以上で入力してください。");
  if (newPw === employeeId) throw new Error("社員IDと同じパスワードは使用できません。");
}

// ----------------------------------------------------
// ログイン・認証 処理
// ----------------------------------------------------

function isActiveEmployee_(emp) {
  return !!emp && emp.isValid === '有効' && emp.status !== '退職';
}

/**
 * メールアドレス＋パスワードの照合。成功時は社員レコード（内部用）を返す
 */
function authenticate_(email, password) {
  const emp = getEmployeeRecords_().find(e => e.email === email);

  if (!emp) throw new Error("メールアドレスまたはパスワードが間違っています。");
  if (emp.isValid !== '有効') throw new Error("無効化されているアカウントです。");
  if (emp.status === '退職') throw new Error("退職済みのアカウントです。");
  if (!verifyPassword_(password, emp.password)) throw new Error("メールアドレスまたはパスワードが間違っています。");
  return emp;
}

/**
 * ログイン完了時の戻り値（ログインセッションを発行する）
 */
function buildLoginResult_(emp, sessionCreatedAt) {
  return {
    success: true,
    empId: emp.empId,
    name: emp.name,
    departmentName: emp.departmentName,
    sectionName: emp.sectionName,
    sessionId: createSession_(emp.empId, sessionCreatedAt)
  };
}

/**
 * ログイン認証処理
 * PW変更要（I列TRUE）の場合はセッションを発行せず mustChangePassword:true を返す
 */
function verifyLogin(email, password) {
  if (!email || !password) throw new Error("メールアドレスとパスワードを入力してください。");

  const emp = authenticate_(email, password);

  if (emp.mustChangePassword) {
    return {
      success: true,
      mustChangePassword: true,
      empId: emp.empId,
      name: emp.name,
      departmentName: emp.departmentName,
      sectionName: emp.sectionName
    };
  }
  return buildLoginResult_(emp);
}

/**
 * パスワード変更：現PW照合→ルール検証→新形式で保存→PW変更要をFALSE→ログイン完了
 */
function changePassword(email, currentPw, newPw) {
  if (!email || !currentPw || !newPw) throw new Error("メールアドレス・現在のパスワード・新しいパスワードを入力してください。");

  const emp = authenticate_(email, currentPw);
  validatePasswordRule_(newPw, emp.empId);

  const sheet = getCommonSpreadsheet().getSheetByName('社員マスタ');
  const targetRow = findEmployeeRow_(sheet, emp.empId);
  // PW変更日時と同じ時刻でセッションを作成し、変更を行ったセッション自体は継続利用できるようにする
  const now = new Date().getTime();
  sheet.getRange(targetRow, EMP_COL_PASSWORD_, 1, 3).setValues([[hashPassword_(newPw), false, toPwChangedAt_(now)]]);

  return buildLoginResult_(emp, now);
}

/**
 * 管理者によるPWリセット：初期PW（社員ID）に戻し、PW変更要をTRUEにする
 */
function adminResetPassword(employeeId) {
  assertAdmin_();
  if (!employeeId) throw new Error("社員IDが指定されていません。");

  const sheet = getCommonSpreadsheet().getSheetByName('社員マスタ');
  if (!sheet) throw new Error("「社員マスタ」シートが見つかりません。");
  const targetRow = findEmployeeRow_(sheet, employeeId);
  sheet.getRange(targetRow, EMP_COL_PASSWORD_, 1, 3).setValues([[hashPassword_(employeeId), true, toPwChangedAt_(new Date().getTime())]]);

  return { success: true, employeeId: employeeId };
}

// ----------------------------------------------------
// ログインセッション（共通基盤内の再遷移用）
// ----------------------------------------------------

/**
 * PW変更日時としてJ列に記録する値
 * 秒単位に切り捨てる（シート保存時のミリ秒丸めで、記録が実際の変更時刻より後にずれるのを防ぐ）
 */
function toPwChangedAt_(timeMs) {
  return new Date(Math.floor(timeMs / 1000) * 1000);
}

function createSession_(employeeId, createdAt) {
  const sessionId = SESSION_PREFIX_ + Utilities.getUuid();
  const sessionData = { employeeId: employeeId, createdAt: createdAt || new Date().getTime() };
  CacheService.getScriptCache().put(sessionId, JSON.stringify(sessionData), SESSION_TTL_SEC_);
  return sessionId;
}

/**
 * ログインセッションから遷移用SSOトークンを都度発行する
 * 社員の状態（有効・退職・PW変更要・セッション作成後のPW変更）を毎回確認し、NGならセッションも破棄する
 */
function issueSsoTokenForSession(sessionId) {
  const expiredMsg = "セッションの有効期限が切れました。再度ログインしてください。";
  if (!sessionId || String(sessionId).indexOf(SESSION_PREFIX_) !== 0) throw new Error(expiredMsg);

  const cache = CacheService.getScriptCache();
  const sessionStr = cache.get(sessionId);
  if (!sessionStr) throw new Error(expiredMsg);

  const session = JSON.parse(sessionStr);
  const emp = getEmployeeRecords_().find(e => e.empId === session.employeeId);
  if (!isActiveEmployee_(emp) || emp.mustChangePassword || session.createdAt < emp.pwChangedAt) {
    cache.remove(sessionId);
    throw new Error("アカウントの状態が変更されました。再度ログインしてください。");
  }
  return generateSsoToken_(emp);
}

function logoutSession(sessionId) {
  if (sessionId && String(sessionId).indexOf(SESSION_PREFIX_) === 0) {
    CacheService.getScriptCache().remove(sessionId);
  }
  return { success: true };
}

// ----------------------------------------------------
// SSOトークン
// ----------------------------------------------------

/**
 * 共通認証SSOトークンの発行（内部用。呼び出し側で社員の状態確認を済ませること）
 */
function generateSsoToken_(emp) {
  const token = SSO_TOKEN_PREFIX_ + Utilities.getUuid();
  const cache = CacheService.getScriptCache();

  const tokenData = {
    employeeId: emp.empId,
    email: emp.email,
    name: emp.name,
    createdAt: new Date().getTime()
  };

  cache.put(token, JSON.stringify(tokenData), SSO_TOKEN_TTL_SEC_);

  return { success: true, token: token, employeeId: emp.empId };
}

/**
 * 管理画面からのSSOログイン（管理者の代理ログインのため、PW変更要は無視する）
 */
function adminGenerateSsoToken(employeeId) {
  assertAdmin_();
  if (!employeeId) throw new Error("社員IDが指定されていません。");

  const emp = getEmployeeRecords_().find(e => e.empId === employeeId);
  if (!isActiveEmployee_(emp)) throw new Error("アクセス権が無効化されています。");
  return generateSsoToken_(emp);
}

function verifySsoToken(token) {
  if (!token) return { isValid: false, error: "トークンが提示されていません。" };
  if (String(token).indexOf(SSO_TOKEN_PREFIX_) !== 0) return { isValid: false, error: "トークンの期限が切れているか、無効です。" };
  const cache = CacheService.getScriptCache();
  const cachedDataStr = cache.get(token);
  if (!cachedDataStr) return { isValid: false, error: "トークンの期限が切れているか、無効です。" };

  const tokenData = JSON.parse(cachedDataStr);
  const emp = getEmployeeRecords_().find(e => e.empId === tokenData.employeeId);

  if (!isActiveEmployee_(emp)) {
    return { isValid: false, error: "共通基盤上でアクセス権が無効化されています。" };
  }
  if (tokenData.createdAt < emp.pwChangedAt) {
    return { isValid: false, error: "パスワードが変更されたため、トークンは無効です。" };
  }
  cache.remove(token);
  return {
    isValid: true, employeeId: emp.empId, email: emp.email, name: emp.name,
    departmentName: emp.departmentName, sectionName: emp.sectionName
  };
}

// ----------------------------------------------------
// 社員マスタ 処理
// ----------------------------------------------------
function generateNewEmployeeId() {
  const sheet = getCommonSpreadsheet().getSheetByName('社員マスタ');
  if (!sheet) throw new Error("「社員マスタ」シートが見つかりません。");
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const empId = data[i][0];
    if (empId && typeof empId === 'string' && empId.startsWith('E')) {
      const num = parseInt(empId.substring(1), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }
  return 'E' + ('0000' + (maxNum + 1)).slice(-4);
}

function registerEmployeeFromWeb(name, email, status, startDateStr, departmentId, sectionId) {
  assertAdmin_();
  const ss = getCommonSpreadsheet();
  const sheet = ss.getSheetByName('社員マスタ');
  if (!sheet) throw new Error("「社員マスタ」シートが見つかりません。");

  const newId = generateNewEmployeeId();
  let formattedDate = startDateStr ? Utilities.formatDate(new Date(startDateStr), Session.getScriptTimeZone(), 'yyyy/MM/dd') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');

  // 初期パスワードは社員IDとする（H列）。保存時はハッシュ化し、I列（PW変更要）をTRUEにする
  const initialPassword = newId;

  sheet.appendRow([newId, name, email, status || '在籍', formattedDate, '', true, hashPassword_(initialPassword), true]);

  if (departmentId) {
    registerAssignment({
      employeeId: newId, departmentId: departmentId, sectionId: sectionId || '',
      type: '主所属', startDate: formattedDate, endDate: '9999/12/31'
    });
  }
  return newId;
}

/**
 * 社員マスタの行番号（1始まり）を返す。見つからなければ例外
 */
function findEmployeeRow_(sheet, employeeId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === employeeId) return i + 1;
  }
  throw new Error("指定された社員IDが見つかりません: " + employeeId);
}

/**
 * 管理画面用の社員一覧（管理者のみ）
 */
function getEmployeeDataForWeb() {
  assertAdmin_();
  return buildEmployeeListForDisplay_();
}

/**
 * 表示・連携用の社員一覧（パスワード関連の項目は含めない）
 */
function buildEmployeeListForDisplay_() {
  return getEmployeeRecords_().map(emp => {
    const { password, mustChangePassword, pwChangedAt, ...publicFields } = emp;
    return publicFields;
  });
}

/**
 * 認証用の社員一覧（内部用。パスワードハッシュ・PW変更要を含む）
 */
function getEmployeeRecords_() {
  const ss = getCommonSpreadsheet();
  const empSheet = ss.getSheetByName('社員マスタ');
  if (!empSheet) return [];
  const empData = empSheet.getDataRange().getValues();
  if (empData.length <= 1) return [];

  const depts = getDepartments_();
  const secs = getSections_();
  const assignSheet = ss.getSheetByName('所属履歴');
  const assignData = assignSheet ? assignSheet.getDataRange().getValues() : [];
  const deptMap = {}; depts.forEach(d => deptMap[d.id] = d.name);
  const secMap = {}; secs.forEach(s => secMap[s.id] = s.name);

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const currentAssignMap = {};
  const concurrentAssignMap = {}; // 兼務（複数可）

  for (let i = 1; i < assignData.length; i++) {
    const row = assignData[i];
    const type = row[4];
    const start = new Date(row[5]);
    const end = row[6] ? new Date(row[6]) : new Date('9999/12/31');
    start.setHours(0, 0, 0, 0); end.setHours(23, 59, 59, 999);
    if (!(today >= start && today <= end)) continue;

    if (type === '主所属') {
      currentAssignMap[row[1]] = {
        deptId: row[2] || '', secId: row[3] || '',
        deptName: row[2] ? (deptMap[row[2]] || row[2]) : '-',
        secName: row[3] ? (secMap[row[3]] || row[3]) : '(部署なし)'
      };
    } else if (type === '兼務') {
      if (!concurrentAssignMap[row[1]]) concurrentAssignMap[row[1]] = [];
      concurrentAssignMap[row[1]].push({
        departmentId: row[2] || '',
        sectionId: row[3] || '',
        departmentName: deptMap[row[2]] || row[2],
        sectionName: secMap[row[3]] || row[3]
      });
    }
  }

  return empData.slice(1).map(row => {
    const empId = row[0] || '';
    const currentAssign = currentAssignMap[empId] || { deptId: '', secId: '', deptName: '-', secName: '-' };
    return {
      empId: empId, name: row[1] || '', email: row[2] || '', status: row[3] || '',
      startDate: formatDate_(row[4]), endDate: formatDate_(row[5]),
      isValid: row[6] === false ? '無効' : '有効',
      password: row[7] || '', // H列（8列目）ハッシュ値を取得
      mustChangePassword: row[8] === true || String(row[8]).toUpperCase() === 'TRUE', // I列（空欄はFALSE扱い）
      pwChangedAt: row[9] instanceof Date ? row[9].getTime() : 0, // J列（空欄＝記録なしは0）
      departmentId: currentAssign.deptId, sectionId: currentAssign.secId,
      departmentName: currentAssign.deptName, sectionName: currentAssign.secName,
      concurrentAssignments: concurrentAssignMap[empId] || []
    };
  });
}

// ----------------------------------------------------
// 事業部・部署・所属マスタ 処理
// ----------------------------------------------------
/**
 * 管理画面用（管理者のみ）
 */
function getDepartments() {
  assertAdmin_();
  return getDepartments_();
}

function getDepartments_() {
  const sheet = getCommonSpreadsheet().getSheetByName('事業部マスタ');
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).map(row => ({
    id: row[0] || '', name: row[1] || '', sortOrder: row[2] || 0, active: row[3] !== false
  }));
}

function generateNewDepartmentId() {
  const sheet = getCommonSpreadsheet().getSheetByName('事業部マスタ');
  if (!sheet) throw new Error("「事業部マスタ」シートが見つかりません。");
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const id = data[i][0];
    if (id && typeof id === 'string' && id.startsWith('ORG')) {
      const num = parseInt(id.substring(3), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }
  return 'ORG' + ('00' + (maxNum + 1)).slice(-2);
}

function registerDepartment(name, sortOrder) {
  assertAdmin_();
  const sheet = getCommonSpreadsheet().getSheetByName('事業部マスタ');
  const newId = generateNewDepartmentId();
  sheet.appendRow([newId, name, Number(sortOrder) || 10, true]);
  return newId;
}

/**
 * 事業部マスタの編集（ID以外の項目：名称／表示順／有効フラグ）
 */
function updateDepartment(id, name, sortOrder, active) {
  assertAdmin_();
  if (!id) throw new Error("事業部IDが指定されていません。");
  const sheet = getCommonSpreadsheet().getSheetByName('事業部マスタ');
  if (!sheet) throw new Error("「事業部マスタ」シートが見つかりません。");

  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { targetRow = i + 1; break; }
  }
  if (targetRow < 0) throw new Error("指定された事業部IDが見つかりません: " + id);

  sheet.getRange(targetRow, 2).setValue(name);
  sheet.getRange(targetRow, 3).setValue(Number(sortOrder) || 10);
  sheet.getRange(targetRow, 4).setValue(active === true || active === 'true');

  return { success: true, id: id };
}

/**
 * 管理画面用（管理者のみ）
 */
function getSections() {
  assertAdmin_();
  return getSections_();
}

function getSections_() {
  const sheet = getCommonSpreadsheet().getSheetByName('部署マスタ');
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).map(row => ({
    id: row[0] || '', name: row[1] || '', departmentId: row[2] || '', sortOrder: row[3] || 0,
    startDate: formatDate_(row[4]), endDate: formatDate_(row[5]), active: row[6] !== false
  }));
}

function generateNewSectionId() {
  const sheet = getCommonSpreadsheet().getSheetByName('部署マスタ');
  if (!sheet) throw new Error("「部署マスタ」シートが見つかりません。");
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const id = data[i][0];
    if (id && typeof id === 'string' && id.startsWith('SEC')) {
      const num = parseInt(id.substring(3), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }
  return 'SEC' + ('00' + (maxNum + 1)).slice(-2);
}

function registerSection(name, departmentId, sortOrder, startDateStr) {
  assertAdmin_();
  const sheet = getCommonSpreadsheet().getSheetByName('部署マスタ');
  const newId = generateNewSectionId();
  let formattedDate = startDateStr ? Utilities.formatDate(new Date(startDateStr), Session.getScriptTimeZone(), 'yyyy/MM/dd') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  sheet.appendRow([newId, name, departmentId, Number(sortOrder) || 10, formattedDate, '', true]);
  return newId;
}

/**
 * 部署マスタの編集（ID以外の項目：名称／所属事業部／表示順／利用開始日／有効フラグ）
 */
function updateSection(id, name, departmentId, sortOrder, startDateStr, active) {
  assertAdmin_();
  if (!id) throw new Error("部署IDが指定されていません。");
  const sheet = getCommonSpreadsheet().getSheetByName('部署マスタ');
  if (!sheet) throw new Error("「部署マスタ」シートが見つかりません。");

  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { targetRow = i + 1; break; }
  }
  if (targetRow < 0) throw new Error("指定された部署IDが見つかりません: " + id);

  const formattedDate = startDateStr ? Utilities.formatDate(new Date(startDateStr), Session.getScriptTimeZone(), 'yyyy/MM/dd') : sheet.getRange(targetRow, 5).getValue();

  sheet.getRange(targetRow, 2).setValue(name);
  sheet.getRange(targetRow, 3).setValue(departmentId);
  sheet.getRange(targetRow, 4).setValue(Number(sortOrder) || 10);
  sheet.getRange(targetRow, 5).setValue(formattedDate);
  sheet.getRange(targetRow, 7).setValue(active === true || active === 'true');

  return { success: true, id: id };
}

/**
 * 管理画面用（管理者のみ）
 */
function getAssignments() {
  assertAdmin_();
  return getAssignments_();
}

function getAssignments_() {
  const sheet = getCommonSpreadsheet().getSheetByName('所属履歴');
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).map(row => ({
    historyId: row[0] || '', employeeId: row[1] || '', departmentId: row[2] || '',
    sectionId: row[3] || '', type: row[4] || '', startDate: formatDate_(row[5]), endDate: formatDate_(row[6]) || '9999/12/31'
  }));
}

function registerAssignment(param) {
  assertAdmin_();
  const sheet = getCommonSpreadsheet().getSheetByName('所属履歴');
  const newStart = new Date(param.startDate);
  const newEnd = param.endDate ? new Date(param.endDate) : new Date('9999/12/31');

  const data = sheet.getDataRange().getValues();

  // 主所属は期間重複不可
  if (param.type === '主所属') {
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] !== param.employeeId || row[4] !== '主所属') continue;
      const existStart = new Date(row[5]);
      const existEnd = row[6] ? new Date(row[6]) : new Date('9999/12/31');
      if (newStart <= existEnd && newEnd >= existStart) {
        throw new Error("主所属の期間が既存の登録（" + formatDate_(row[5]) + " 〜 " + formatDate_(row[6]) + "）と重複しています。");
      }
    }
  }

  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const idStr = String(data[i][0] || '');
    if (idStr.startsWith('H')) {
      const num = parseInt(idStr.replace('H', ''), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }
  const newHistoryId = 'H' + ('0000' + (maxNum + 1)).slice(-4);
  sheet.appendRow([
    newHistoryId, param.employeeId, param.departmentId, param.sectionId, param.type,
    Utilities.formatDate(newStart, Session.getScriptTimeZone(), 'yyyy/MM/dd'),
    param.endDate ? Utilities.formatDate(newEnd, Session.getScriptTimeZone(), 'yyyy/MM/dd') : '9999/12/31'
  ]);
  return { success: true, historyId: newHistoryId };
}

function formatDate_(dateStr) {
  if (!dateStr) return '-';
  try { return Utilities.formatDate(new Date(dateStr), Session.getScriptTimeZone(), 'yyyy/MM/dd'); } catch(e) { return dateStr; }
}

/**
 * 外部システム（戦略AP等）連携用API
 * 共通基盤の最新マスタデータを一括返却する
 * ※ライブラリ経由で戦略APから呼ばれるため管理者チェックではなくAPIキーで保護する（パスワード関連の項目は含めない）
 * @param {string} apiKey 連携用APIキー（スクリプトプロパティ API_KEY_STRATEGY_AP と照合）
 * @return {Object} 全マスタデータを含むオブジェクト
 */
function exportCommonMasterData(apiKey) {
  assertStrategyApiKey_(apiKey);
  return {
    employees: buildEmployeeListForDisplay_(),
    departments: getDepartments_(),
    sections: getSections_(),
    assignments: getAssignments_()
  };
}

/**
 * 戦略AP連携用APIキーの照合。未指定・不一致・プロパティ未設定はいずれも拒否する
 */
function assertStrategyApiKey_(apiKey) {
  const expected = PropertiesService.getScriptProperties().getProperty('API_KEY_STRATEGY_AP');
  if (!expected || !apiKey || apiKey !== expected) throw new Error('Unauthorized');
}

/**
 * 退職処理：在籍状況を'退職'に、利用終了日を設定する。
 * 有効フラグはここでは変更しない（退職と無効化は別概念のため）。
 */
function retireEmployee(employeeId, endDateStr) {
  assertAdmin_();
  if (!employeeId) throw new Error("社員IDが指定されていません。");
  if (!endDateStr) throw new Error("利用終了日を指定してください。");

  const sheet = getCommonSpreadsheet().getSheetByName('社員マスタ');
  if (!sheet) throw new Error("「社員マスタ」シートが見つかりません。");

  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === employeeId) { targetRow = i + 1; break; }
  }
  if (targetRow < 0) throw new Error("指定された社員IDが見つかりません: " + employeeId);

  const formattedEndDate = Utilities.formatDate(new Date(endDateStr), Session.getScriptTimeZone(), 'yyyy/MM/dd');

  sheet.getRange(targetRow, 4).setValue('退職');       // D列：在籍状況
  sheet.getRange(targetRow, 6).setValue(formattedEndDate); // F列：利用終了日

  return { success: true, employeeId: employeeId, endDate: formattedEndDate };
}

/**
 * 社員情報の編集（氏名・メールアドレス・在籍状況）
 * 在籍状況が「退職」の場合のみ利用終了日を必須とし、そうでない場合は利用終了日をクリアする。
 * 主所属の付け替えはここでは扱わない（reassignPrimaryAssignmentを参照）。
 */
function updateEmployeeFromWeb(employeeId, name, email, status, endDateStr) {
  assertAdmin_();
  if (!employeeId) throw new Error("社員IDが指定されていません。");
  if (!name) throw new Error("氏名を入力してください。");
  if (!email) throw new Error("メールアドレスを入力してください。");

  const sheet = getCommonSpreadsheet().getSheetByName('社員マスタ');
  if (!sheet) throw new Error("「社員マスタ」シートが見つかりません。");

  const data = sheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === employeeId) { targetRow = i + 1; break; }
  }
  if (targetRow < 0) throw new Error("指定された社員IDが見つかりません: " + employeeId);

  sheet.getRange(targetRow, 2).setValue(name);
  sheet.getRange(targetRow, 3).setValue(email);

  if (status === '退職') {
    if (!endDateStr) throw new Error("退職の場合は利用終了日を指定してください。");
    const formattedEndDate = Utilities.formatDate(new Date(endDateStr), Session.getScriptTimeZone(), 'yyyy/MM/dd');
    sheet.getRange(targetRow, 4).setValue('退職');
    sheet.getRange(targetRow, 6).setValue(formattedEndDate);
  } else {
    sheet.getRange(targetRow, 4).setValue(status);
    sheet.getRange(targetRow, 6).setValue('');
  }

  return { success: true, employeeId: employeeId };
}

/**
 * 主所属の付け替え：現在の主所属の終了日を異動日の前日に設定し、
 * 新しい主所属を所属履歴に追加登録する（履歴は消さない）。
 * 事業部が未指定の場合は「所属なし」への異動として、現在の主所属を終了するのみで新規登録は行わない。
 * 部署は事業部配下に部署が存在しない場合を考慮し、未指定（事業部のみ）を許容する。
 */
function reassignPrimaryAssignment(employeeId, newDepartmentId, newSectionId, effectiveDateStr) {
  assertAdmin_();
  if (!employeeId) throw new Error("社員IDが指定されていません。");
  if (!effectiveDateStr) throw new Error("異動日を指定してください。");
  if (!newDepartmentId && newSectionId) throw new Error("部署を指定する場合は事業部も指定してください。");

  const unassigning = !newDepartmentId;

  const sheet = getCommonSpreadsheet().getSheetByName('所属履歴');
  if (!sheet) throw new Error("「所属履歴」シートが見つかりません。");

  const effectiveDate = new Date(effectiveDateStr);
  const data = sheet.getDataRange().getValues();

  // 現在の主所属（終了日が9999/12/31＝未設定）を探して、異動日前日を終了日として設定する
  let currentRow = -1;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] !== employeeId || row[4] !== '主所属') continue;
    const end = row[6] ? new Date(row[6]) : new Date('9999/12/31');
    if (end.getFullYear() === 9999) { currentRow = i + 1; break; }
  }

  if (currentRow > 0) {
    const prevEnd = new Date(effectiveDate);
    prevEnd.setDate(prevEnd.getDate() - 1);
    sheet.getRange(currentRow, 7).setValue(Utilities.formatDate(prevEnd, Session.getScriptTimeZone(), 'yyyy/MM/dd'));
  }

  if (unassigning) {
    return { success: true, employeeId: employeeId, unassigned: true };
  }

  return registerAssignment({
    employeeId: employeeId,
    departmentId: newDepartmentId,
    sectionId: newSectionId || '',
    type: '主所属',
    startDate: effectiveDateStr,
    endDate: ''
  });
}