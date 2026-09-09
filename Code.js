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

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('NJK社内システム 共通基盤')
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

// ----------------------------------------------------
// ログイン・認証 処理
// ----------------------------------------------------

/**
 * ログイン認証処理
 */
function verifyLogin(email, password) {
  if (!email || !password) throw new Error("メールアドレスとパスワードを入力してください。");

  const empList = getEmployeeDataForWeb();
  const emp = empList.find(e => e.email === email);

  if (!emp) throw new Error("メールアドレスまたはパスワードが間違っています。");
  if (emp.isValid !== '有効') throw new Error("無効化されているアカウントです。");
  if (emp.status === '退職') throw new Error("退職済みのアカウントです。");
  if (emp.password !== computeHash_(password)) throw new Error("メールアドレスまたはパスワードが間違っています。");

  // 認証成功時、ポータル遷移用のSSOトークンも同時に発行する
  const ssoResult = generateSsoToken(emp.empId);

  return {
    success: true,
    empId: emp.empId,
    name: emp.name,
    departmentName: emp.departmentName,
    sectionName: emp.sectionName,
    ssoToken: ssoResult.token
  };
}

/**
 * 共通認証SSOトークンの発行
 */
function generateSsoToken(employeeId) {
  if (!employeeId) throw new Error("社員IDが指定されていません。");

  const empList = getEmployeeDataForWeb();
  const emp = empList.find(e => e.empId === employeeId);
  if (!emp || emp.isValid !== '有効' || emp.status === '退職') {
    throw new Error("アクセス権が無効化されています。");
  }

  const token = 'SSO_' + Utilities.getUuid();
  const cache = CacheService.getScriptCache();

  const tokenData = {
    employeeId: emp.empId,
    email: emp.email,
    name: emp.name,
    createdAt: new Date().getTime()
  };

  cache.put(token, JSON.stringify(tokenData), 300);

  return { success: true, token: token, employeeId: emp.empId };
}

function verifySsoToken(token) {
  if (!token) return { isValid: false, error: "トークンが提示されていません。" };
  const cache = CacheService.getScriptCache();
  const cachedDataStr = cache.get(token);
  if (!cachedDataStr) return { isValid: false, error: "トークンの期限が切れているか、無効です。" };

  const tokenData = JSON.parse(cachedDataStr);
  const empList = getEmployeeDataForWeb();
  const emp = empList.find(e => e.empId === tokenData.employeeId);

  if (!emp || emp.isValid !== '有効' || emp.status === '退職') {
    return { isValid: false, error: "共通基盤上でアクセス権が無効化されています。" };
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
  const ss = getCommonSpreadsheet();
  const sheet = ss.getSheetByName('社員マスタ');
  if (!sheet) throw new Error("「社員マスタ」シートが見つかりません。");

  const newId = generateNewEmployeeId();
  let formattedDate = startDateStr ? Utilities.formatDate(new Date(startDateStr), Session.getScriptTimeZone(), 'yyyy/MM/dd') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');

  // 初期パスワードは社員IDとする（8列目に追加）。保存時はハッシュ化する
  const initialPassword = newId;

  sheet.appendRow([newId, name, email, status || '在籍', formattedDate, '', true, computeHash_(initialPassword)]);

  if (departmentId) {
    registerAssignment({
      employeeId: newId, departmentId: departmentId, sectionId: sectionId || '',
      type: '主所属', startDate: formattedDate, endDate: '9999/12/31'
    });
  }
  return newId;
}

function getEmployeeDataForWeb() {
  const ss = getCommonSpreadsheet();
  const empSheet = ss.getSheetByName('社員マスタ');
  if (!empSheet) return [];
  const empData = empSheet.getDataRange().getValues();
  if (empData.length <= 1) return [];

  const depts = getDepartments();
  const secs = getSections();
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
      departmentId: currentAssign.deptId, sectionId: currentAssign.secId,
      departmentName: currentAssign.deptName, sectionName: currentAssign.secName,
      concurrentAssignments: concurrentAssignMap[empId] || []
    };
  });
}

// ----------------------------------------------------
// 事業部・部署・所属マスタ 処理
// ----------------------------------------------------
function getDepartments() {
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
  const sheet = getCommonSpreadsheet().getSheetByName('事業部マスタ');
  const newId = generateNewDepartmentId();
  sheet.appendRow([newId, name, Number(sortOrder) || 10, true]);
  return newId;
}

/**
 * 事業部マスタの編集（ID以外の項目：名称／表示順／有効フラグ）
 */
function updateDepartment(id, name, sortOrder, active) {
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

function getSections() {
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

function getAssignments() {
  const sheet = getCommonSpreadsheet().getSheetByName('所属履歴');
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).map(row => ({
    historyId: row[0] || '', employeeId: row[1] || '', departmentId: row[2] || '',
    sectionId: row[3] || '', type: row[4] || '', startDate: formatDate_(row[5]), endDate: formatDate_(row[6]) || '9999/12/31'
  }));
}

function registerAssignment(param) {
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
 * @return {Object} 全マスタデータを含むオブジェクト
 */
function exportCommonMasterData() {
  return {
    employees: getEmployeeDataForWeb(),
    departments: getDepartments(),
    sections: getSections(),
    assignments: getAssignments()
  };
}

function testLogin() {
  const result = verifyLogin('asa.miyamoto.3@gmail.com', 'E0001');
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * 退職処理：在籍状況を'退職'に、利用終了日を設定する。
 * 有効フラグはここでは変更しない（退職と無効化は別概念のため）。
 */
function retireEmployee(employeeId, endDateStr) {
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