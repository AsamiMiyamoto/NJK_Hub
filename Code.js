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

// 管理画面へのアクセスを許可する固定管理者（コード上で固定し、画面からは削除できない）。
// それ以外の管理者はデータSSの「管理者設定」タブで管理する
const FIXED_ADMIN_EMAILS_ = ['admin@j-shelter.com'];
const ADMIN_EMAIL_DOMAIN_ = '@j-shelter.com';
const ADMIN_SHEET_NAME_ = '管理者設定';
// D列「社員ID」：問い合わせ通知の宛先（社員マスタのメールアドレス）を引くための紐付け。既存の行は空のまま残す
const ADMIN_SHEET_HEADERS_ = ['メールアドレス', '追加日時', '追加者', '社員ID'];
const ADMIN_HISTORY_SHEET_NAME_ = '管理者変更履歴';
const ADMIN_HISTORY_HEADERS_ = ['日時', '操作', '対象', '実行者'];
// メニューの「ポータル」（遷移先は人事評価）を準備中として管理者のみに制限する。公開時は false にする
const PORTAL_ADMIN_ONLY_ = true;
const PORTAL_SYSTEM_KEY_ = 'jinji';

// 共通基盤メニューからの遷移先（キーは goToSystem に渡すシステム名）。ログイン画面と管理画面のヘルプで使う
const SYSTEM_URLS_ = {
  strategy: 'https://script.google.com/macros/s/AKfycbwC2en3g6XedkK4rQcwFPv5WTe4NVxK1yNnY3S0cQODdwPqcE9INiXI_rSCxrU86s6R/exec',
  jinji: 'https://script.google.com/macros/s/AKfycbyqGY8dPwXWB3vVOmaN2Eu9ZHWRqzwlK4VlZ01aVlURj4N1uPCnV7Iin14hPCd9SMnG/exec'
};

// 「システム管理者に連絡」（問い合わせ）
const INQUIRY_SHEET_NAME_ = '問い合わせ';
const INQUIRY_HEADERS_ = ['受付ID', '受付日時', '社員ID', '氏名', 'メールアドレス', '用件', '内容', '状態', '対応者', '対応日時', '対応メモ'];
const INQUIRY_CATEGORIES_ = ['パスワードを忘れた', 'ロックされた', '社員情報の登録・変更', 'その他'];
const INQUIRY_CONTENT_REQUIRED_ = ['社員情報の登録・変更', 'その他'];
const INQUIRY_CONTENT_MAX_ = 2000;
const INQUIRY_STATUS_OPEN_ = '未対応';
const INQUIRY_STATUS_DONE_ = '対応済み';
// 送信回数の制限（同じメールアドレスで1時間に3回まで。キーは 接頭辞＋SHA-256(小文字化したメールアドレス)）
const INQUIRY_RATE_PREFIX_ = 'inquiryrate_';
const INQUIRY_RATE_LIMIT_ = 3;
const INQUIRY_RATE_WINDOW_SEC_ = 3600;

// 先頭は英数字に限る（=・+・- で始まる値はシートで数式として扱われるため）
const ADMIN_EMAIL_PATTERN_ = /^[a-z0-9][a-z0-9._%+-]*@[a-z0-9.-]+\.[a-z]{2,}$/;

// 社員マスタの列番号（1始まり）
const EMP_COL_PASSWORD_ = 8;     // H列：パスワード
const EMP_COL_MUST_CHANGE_ = 9;  // I列：PW変更要
const EMP_COL_PW_CHANGED_AT_ = 10; // J列：PW変更日時

// 在籍状況（D列）に入れられる値
const EMP_STATUSES_ = ['在籍', '休職', '退職'];

// CacheServiceのキー接頭辞（SSOトークンとログインセッションの取り違え防止）
const SSO_TOKEN_PREFIX_ = 'SSO_';
const SESSION_PREFIX_ = 'SESSION_';
const SSO_TOKEN_TTL_SEC_ = 300;
const SESSION_TTL_SEC_ = 21600; // 6時間（CacheServiceの上限）

// 試行回数制限（キーは 接頭辞＋SHA-256(入力メールアドレス)。ロックは同じキーの末尾に _lock を付ける）
const LOGIN_FAIL_PREFIX_ = 'loginfail_';
const PW_CHANGE_FAIL_PREFIX_ = 'pwchgfail_';
const AUTH_FAIL_LIMIT_ = 5;
const AUTH_FAIL_TTL_SEC_ = 900;

// 管理者リセット時の仮PW（紛らわしい 0 O 1 l I を除く）
const TEMP_PW_LENGTH_ = 12;
const TEMP_PW_CHARS_ = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';

// ----------------------------------------------------
// 管理者判定・管理者設定
// ----------------------------------------------------

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function getActiveUserEmail_() {
  return normalizeEmail_(Session.getActiveUser().getEmail());
}

function isFixedAdmin_(email) {
  return FIXED_ADMIN_EMAILS_.indexOf(normalizeEmail_(email)) !== -1;
}

/**
 * スクリプトロックを取得して fn を実行する（ID採番〜書き込みを他の実行と重ねないため）。
 * 解放前に flush し、次にロックを取った実行が書き込み後の値を読めるようにする
 */
function withScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

/**
 * タブを取得する。無ければ見出し付きで作成する
 */
function getOrCreateSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 管理者設定タブの登録行（見出しを除く、メールアドレスは小文字化済み）。空行は除く
 */
function readAdminRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, ADMIN_SHEET_HEADERS_.length).getValues()
    .map((r, i) => ({ rowNumber: i + 2, email: normalizeEmail_(r[0]), addedAt: r[1], addedBy: r[2], employeeId: String(r[3] || '').trim() }))
    .filter(r => r.email);
}

/**
 * 管理者か判定する（固定管理者 ∪ 管理者設定タブ）。
 * 管理者設定タブはシートを直接編集されることもあるため、@j-shelter.com 以外の行は無視する。
 * 判定の都度タブを読むため、追加・削除はすぐに反映される。
 */
function isAdmin_(email) {
  const target = normalizeEmail_(email);
  if (!target) return false;
  if (isFixedAdmin_(target)) return true;
  if (!target.endsWith(ADMIN_EMAIL_DOMAIN_)) return false;
  const sheet = getCommonSpreadsheet().getSheetByName(ADMIN_SHEET_NAME_);
  return readAdminRows_(sheet).some(r => r.email === target);
}

/**
 * 実行中のGoogleアカウントが管理者でなければ拒否する。管理者ならそのメールアドレス（小文字化済み）を返す
 */
function assertAdmin_() {
  const email = getActiveUserEmail_();
  if (!isAdmin_(email)) throw new Error("管理者権限がありません。");
  return email;
}

/**
 * 共通基盤メニューから遷移先システムを利用できるか（ポータルは準備中の間、管理者のみ）
 */
function canUseSystem_(systemKey) {
  if (PORTAL_ADMIN_ONLY_ && systemKey === PORTAL_SYSTEM_KEY_) return isAdmin_(getActiveUserEmail_());
  return true;
}

/**
 * 管理者設定タブを取得する（無ければ見出し付きで作成）。
 * 「社員ID」列の追加前に作られたタブには、D列の見出しだけを追加する（既存の行は空のまま）
 */
function getAdminSheet_(ss) {
  const sheet = getOrCreateSheet_(ss, ADMIN_SHEET_NAME_, ADMIN_SHEET_HEADERS_);
  const col = ADMIN_SHEET_HEADERS_.length;
  if (!sheet.getRange(1, col).getValue()) sheet.getRange(1, col).setValue(ADMIN_SHEET_HEADERS_[col - 1]);
  return sheet;
}

/**
 * 管理者に紐付けられる社員か確認して返す（在籍・休職のみ。退職・未登録は例外）
 */
function findLinkableEmployee_(employeeId) {
  const id = String(employeeId || '').trim();
  if (!id) throw new Error("紐付ける社員を選択してください。");
  const emp = getEmployeeRecords_().find(e => e.empId === id);
  if (!emp) throw new Error("指定された社員が見つかりません: " + id);
  if (emp.status === '退職') throw new Error("退職済みの社員は紐付けできません: " + id);
  return emp;
}

/**
 * 紐付けた社員の状態。ok（通知を受け取れる）／unlinked（未設定）／missing（社員が見つからない）／retired（退職済み）
 */
function adminLinkStatus_(employeeId, empMap) {
  if (!employeeId) return 'unlinked';
  const emp = empMap[employeeId];
  if (!emp) return 'missing';
  return emp.status === '退職' ? 'retired' : 'ok';
}

/**
 * 問い合わせ通知の宛先：管理者設定タブで社員が紐付いた管理者の、社員マスタのメールアドレス（小文字化・重複は1件）。
 * 固定管理者・社員が紐付いていない管理者・退職の社員は除く。isAdmin_ と同じく @j-shelter.com 以外の行は管理者として扱わない
 * @param {Array} employees getEmployeeRecords_() の結果（省略時は読み込む）
 */
function getAdminNotifyEmails_(employees) {
  const empMap = {};
  (employees || getEmployeeRecords_()).forEach(e => { empMap[e.empId] = e; });
  const emails = readAdminRows_(getCommonSpreadsheet().getSheetByName(ADMIN_SHEET_NAME_))
    .filter(r => r.email.endsWith(ADMIN_EMAIL_DOMAIN_) && !isFixedAdmin_(r.email))
    .filter(r => adminLinkStatus_(r.employeeId, empMap) === 'ok')
    .map(r => normalizeEmail_(empMap[r.employeeId].email))
    .filter(Boolean);
  return emails.filter((email, i, list) => list.indexOf(email) === i);
}

/**
 * 管理者の追加・削除・社員の紐付けを管理者変更履歴タブに記録する
 */
function recordAdminChange_(ss, operation, targetEmail, actorEmail) {
  getOrCreateSheet_(ss, ADMIN_HISTORY_SHEET_NAME_, ADMIN_HISTORY_HEADERS_)
    .appendRow([new Date(), operation, targetEmail, actorEmail]);
}

/**
 * 管理者一覧（固定管理者を先頭に、削除不可の印 fixed を付ける。self は実行中のアカウント）
 * linkStatus は紐付けた社員の状態（fixed／ok／unlinked／missing／retired。ok 以外は通知を受け取れない）
 * @return {Array<{email, addedAt, addedBy, employeeId, employeeName, linkStatus, fixed, self}>}
 */
function getAdmins() {
  const me = assertAdmin_();
  const sheet = getAdminSheet_(getCommonSpreadsheet());
  const tz = Session.getScriptTimeZone();
  const empMap = {};
  getEmployeeRecords_().forEach(e => { empMap[e.empId] = e; });
  const fixed = FIXED_ADMIN_EMAILS_.map(email => ({
    email: email, addedAt: '', addedBy: '', employeeId: '', employeeName: '', linkStatus: 'fixed', fixed: true, self: email === me
  }));
  const registered = readAdminRows_(sheet).map(r => ({
    email: r.email,
    addedAt: r.addedAt instanceof Date ? Utilities.formatDate(r.addedAt, tz, 'yyyy/MM/dd HH:mm') : String(r.addedAt || ''),
    addedBy: String(r.addedBy || ''),
    employeeId: r.employeeId,
    employeeName: empMap[r.employeeId] ? empMap[r.employeeId].name : '',
    linkStatus: adminLinkStatus_(r.employeeId, empMap),
    fixed: false,
    self: r.email === me
  }));
  return fixed.concat(registered);
}

/**
 * 管理者を追加する（@j-shelter.com のみ。固定管理者・登録済みとの重複は拒否）。
 * 通知の宛先にするため、在籍・休職の社員の紐付けを必須とする
 */
function addAdmin(email, employeeId) {
  const me = assertAdmin_();
  const target = normalizeEmail_(email);
  if (!target) throw new Error("メールアドレスを入力してください。");
  if (!ADMIN_EMAIL_PATTERN_.test(target)) throw new Error("メールアドレスの形式が正しくありません。");
  if (!target.endsWith(ADMIN_EMAIL_DOMAIN_)) throw new Error("管理者に登録できるのは " + ADMIN_EMAIL_DOMAIN_ + " のアカウントのみです。");
  const emp = findLinkableEmployee_(employeeId);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getCommonSpreadsheet();
    const sheet = getAdminSheet_(ss);
    if (isFixedAdmin_(target) || readAdminRows_(sheet).some(r => r.email === target)) {
      throw new Error("すでに管理者として登録されています: " + target);
    }
    sheet.appendRow([target, new Date(), me, emp.empId]);
    recordAdminChange_(ss, '追加', target + '（社員 ' + emp.empId + ' ' + emp.name + '）', me);
  } finally {
    lock.releaseLock();
  }
  return { success: true, email: target };
}

/**
 * 既存の管理者に社員を紐付ける（設定・変更）。固定管理者は対象外。変更は管理者変更履歴に記録する
 */
function linkAdminEmployee(email, employeeId) {
  const me = assertAdmin_();
  const target = normalizeEmail_(email);
  if (isFixedAdmin_(target)) throw new Error("固定管理者には社員を紐付けできません。");
  const emp = findLinkableEmployee_(employeeId);

  return withScriptLock_(() => {
    const ss = getCommonSpreadsheet();
    const sheet = getAdminSheet_(ss);
    const rows = readAdminRows_(sheet).filter(r => r.email === target);
    if (rows.length === 0) throw new Error("指定された管理者が見つかりません: " + target);
    const before = rows[0].employeeId;
    if (before === emp.empId && rows.every(r => r.employeeId === emp.empId)) return { success: true, email: target, employeeId: emp.empId };
    rows.forEach(r => sheet.getRange(r.rowNumber, ADMIN_SHEET_HEADERS_.length).setValue(emp.empId));
    recordAdminChange_(ss, '社員の紐付け', target + '：' + (before || '未設定') + ' → ' + emp.empId + '（' + emp.name + '）', me);
    return { success: true, email: target, employeeId: emp.empId };
  });
}

/**
 * 管理者を削除する（固定管理者・自分自身は拒否）
 */
function removeAdmin(email) {
  const me = assertAdmin_();
  const target = normalizeEmail_(email);
  if (isFixedAdmin_(target)) throw new Error("固定管理者は削除できません。");
  if (target === me) throw new Error("自分自身は削除できません。");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = getCommonSpreadsheet();
    const sheet = getAdminSheet_(ss);
    const rows = readAdminRows_(sheet).filter(r => r.email === target);
    if (rows.length === 0) throw new Error("指定された管理者が見つかりません: " + target);
    // 下の行から削除して行番号のずれを防ぐ（重複行があればまとめて削除）
    rows.map(r => r.rowNumber).sort((a, b) => b - a).forEach(n => sheet.deleteRow(n));
    recordAdminChange_(ss, '削除', target, me);
  } finally {
    lock.releaseLock();
  }
  return { success: true, email: target };
}

/**
 * 管理画面ヘルプの「各システムのURL」の表に出すURL（共通基盤自身のURLと SYSTEM_URLS_ から作る）
 */
function getHelpUrls_() {
  const selfUrl = ScriptApp.getService().getUrl();
  return {
    login: selfUrl,
    admin: selfUrl + '?admin=true',
    strategy: SYSTEM_URLS_.strategy,
    jinji: SYSTEM_URLS_.jinji
  };
}

// ----------------------------------------------------
// 問い合わせ（システム管理者に連絡）
// ----------------------------------------------------

/**
 * 利用者が入力した文字列をシートに書き込む前の処理（=・+・-・@ で始まる値が数式として扱われないようにする）
 */
function toSheetText_(value) {
  const text = String(value === null || value === undefined ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

/**
 * 送信回数の制限。同じメールアドレスで INQUIRY_RATE_WINDOW_SEC_ 秒に INQUIRY_RATE_LIMIT_ 回まで。
 * 最初の送信から数えた固定の時間枠で数える。上限を超えた場合は false
 */
function consumeInquiryQuota_(email) {
  const cache = CacheService.getScriptCache();
  const key = INQUIRY_RATE_PREFIX_ + computeHash_(email);
  const now = Date.now();
  const saved = JSON.parse(cache.get(key) || 'null');
  const state = (saved && now - saved.firstAt < INQUIRY_RATE_WINDOW_SEC_ * 1000) ? saved : { count: 0, firstAt: now };
  if (state.count >= INQUIRY_RATE_LIMIT_) return false;
  state.count++;
  const remainingSec = Math.max(1, Math.ceil(INQUIRY_RATE_WINDOW_SEC_ - (now - state.firstAt) / 1000));
  cache.put(key, JSON.stringify(state), remainingSec);
  return true;
}

/**
 * ログイン画面・メニューの「システム管理者に連絡」からの送信を受け付ける（ログイン不要）。
 * メールアドレスを小文字化して社員マスタと照合し、一致して退職でない場合のみ保存・通知する（無効の社員は受け付ける）。
 * 登録の有無を推測させないため、未登録・送信回数超過・保存時のエラーでも結果は常に同じにする。
 * 入力の形式エラー（必須の未入力など）だけは例外で知らせる。
 * @return {{success: boolean}}
 */
function submitInquiry(email, category, content) {
  const normalizedEmail = normalizeEmail_(email);
  const text = String(content || '').trim();
  if (!normalizedEmail) throw new Error("メールアドレスを入力してください。");
  if (INQUIRY_CATEGORIES_.indexOf(category) === -1) throw new Error("用件を選択してください。");
  if (INQUIRY_CONTENT_REQUIRED_.indexOf(category) !== -1 && !text) throw new Error("内容を入力してください。");
  if (text.length > INQUIRY_CONTENT_MAX_) throw new Error("内容は" + INQUIRY_CONTENT_MAX_ + "文字以内で入力してください。");

  const result = { success: true };
  try {
    if (!consumeInquiryQuota_(normalizedEmail)) {
      console.warn('submitInquiry: 送信回数の上限に達したため受け付けませんでした');
      return result;
    }
    const emp = getEmployeeRecords_().find(e => normalizeEmail_(e.email) === normalizedEmail);
    if (!emp || emp.status === '退職') return result;

    const ss = getCommonSpreadsheet();
    const inquiryId = withScriptLock_(() => {
      const sheet = getOrCreateSheet_(ss, INQUIRY_SHEET_NAME_, INQUIRY_HEADERS_);
      const id = generateNewInquiryId_(sheet);
      sheet.appendRow([id, new Date(), emp.empId, emp.name, normalizedEmail, category, toSheetText_(text), INQUIRY_STATUS_OPEN_, '', '', '']);
      return id;
    });
    notifyAdminsOfInquiry_(inquiryId, emp, category, text);
  } catch (e) {
    console.error('submitInquiry: 受付処理でエラーが発生しました: ' + (e && e.message ? e.message : e));
  }
  return result;
}

/**
 * 受付IDの採番（Q＋4桁、既存の最大番号＋1）。呼び出し側でロックを取ること
 */
function generateNewInquiryId_(sheet) {
  let maxNum = 0;
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach(r => {
      const id = String(r[0] || '');
      if (id.startsWith('Q')) {
        const num = parseInt(id.substring(1), 10);
        if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    });
  }
  return 'Q' + ('0000' + (maxNum + 1)).slice(-4);
}

/**
 * 問い合わせを、社員が紐付いた管理者に、社員マスタのメールアドレスで知らせる（失敗しても受付は取り消さない）。
 * 宛先が1件もないときは送らずにログに記録する
 */
function notifyAdminsOfInquiry_(inquiryId, emp, category, text) {
  try {
    const body = [
      'システム管理者への問い合わせを受け付けました。',
      '',
      '受付ID：' + inquiryId,
      '氏名：' + emp.name,
      '社員ID：' + emp.empId,
      '用件：' + category,
      '内容：',
      text || '（なし）',
      '',
      '管理画面：' + getHelpUrls_().admin
    ].join('\n');
    const recipients = getAdminNotifyEmails_();
    if (recipients.length === 0) {
      console.warn('notifyAdminsOfInquiry_: 通知を受け取れる管理者がいないため、通知メールを送りませんでした（受付ID ' + inquiryId + '）');
      return;
    }
    MailApp.sendEmail({ to: recipients.join(','), subject: '[NJK] 問い合わせ：' + category, body: body });
  } catch (e) {
    console.error('notifyAdminsOfInquiry_: 通知メールを送れませんでした（受付ID ' + inquiryId + '）: ' + (e && e.message ? e.message : e));
  }
}

/**
 * 管理画面「要対応」パネルのデータ（管理者のみ）
 * - inquiries：問い合わせ（新しい順。対応済みを含む。表示の絞り込みは画面側で行う）
 * - tempPasswordEmployees：仮パスワードのまま（PW変更要がTRUE）の社員（退職・無効は除く）
 * - inactiveOrgEmployees：無効な事業部・部署に主所属・兼務が残っている社員（退職は除く）
 * - notifiableAdminCount：問い合わせの通知を受け取れる宛先の数（0 のときは画面で注意を出す）
 */
function getActionItems() {
  assertAdmin_();
  const tz = Session.getScriptTimeZone();
  const fmt = v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy/MM/dd HH:mm') : String(v || '');

  const inquirySheet = getCommonSpreadsheet().getSheetByName(INQUIRY_SHEET_NAME_);
  const inquiryRows = (inquirySheet && inquirySheet.getLastRow() >= 2)
    ? inquirySheet.getRange(2, 1, inquirySheet.getLastRow() - 1, INQUIRY_HEADERS_.length).getValues()
    : [];
  const inquiries = inquiryRows.filter(r => r[0]).map(r => ({
    id: String(r[0]), receivedAt: fmt(r[1]), employeeId: String(r[2] || ''), name: String(r[3] || ''),
    email: String(r[4] || ''), category: String(r[5] || ''), content: String(r[6] || ''),
    status: String(r[7] || INQUIRY_STATUS_OPEN_), handledBy: String(r[8] || ''), handledAt: fmt(r[9]), memo: String(r[10] || '')
  })).reverse();

  const employees = getEmployeeRecords_().filter(e => e.status !== '退職');
  const tempPasswordEmployees = employees
    .filter(e => e.mustChangePassword && e.isValid === '有効')
    .map(e => ({ empId: e.empId, name: e.name }));

  const deptActive = {};
  getDepartments_().forEach(d => { deptActive[d.id] = d.active; });
  const secActive = {};
  getSections_().forEach(sec => { secActive[sec.id] = sec.active; });
  const inactiveOrgEmployees = [];
  employees.forEach(e => {
    const reasons = [];
    const check = (label, deptId, secId, deptName, secName) => {
      if (deptId && deptActive[deptId] === false) reasons.push(label + '：事業部「' + deptName + '」が無効');
      if (secId && secActive[secId] === false) reasons.push(label + '：部署「' + secName + '」が無効');
    };
    check('主所属', e.departmentId, e.sectionId, e.departmentName, e.sectionName);
    e.concurrentAssignments.forEach(c => check('兼務', c.departmentId, c.sectionId, c.departmentName, c.sectionName));
    if (reasons.length) inactiveOrgEmployees.push({ empId: e.empId, name: e.name, reasons: reasons });
  });

  return {
    inquiries: inquiries,
    tempPasswordEmployees: tempPasswordEmployees,
    inactiveOrgEmployees: inactiveOrgEmployees,
    notifiableAdminCount: getAdminNotifyEmails_(employees).length
  };
}

/**
 * 問い合わせを対応済みにする（管理者のみ）。対応者・対応日時は自動で記録する
 */
function resolveInquiry(inquiryId, memo) {
  const me = assertAdmin_();
  const text = String(memo || '').trim();
  if (!inquiryId) throw new Error("受付IDが指定されていません。");
  if (!text) throw new Error("対応メモを入力してください。");
  if (text.length > INQUIRY_CONTENT_MAX_) throw new Error("対応メモは" + INQUIRY_CONTENT_MAX_ + "文字以内で入力してください。");

  return withScriptLock_(() => {
    const sheet = getCommonSpreadsheet().getSheetByName(INQUIRY_SHEET_NAME_);
    if (!sheet || sheet.getLastRow() < 2) throw new Error("指定された問い合わせが見つかりません: " + inquiryId);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    const index = rows.findIndex(r => String(r[0]) === String(inquiryId));
    if (index < 0) throw new Error("指定された問い合わせが見つかりません: " + inquiryId);
    if (rows[index][7] === INQUIRY_STATUS_DONE_) throw new Error("この問い合わせはすでに対応済みです。");
    // H〜K列：状態／対応者／対応日時／対応メモ
    sheet.getRange(index + 2, 8, 1, 4).setValues([[INQUIRY_STATUS_DONE_, me, new Date(), toSheetText_(text)]]);
    return { success: true, id: String(inquiryId) };
  });
}

function doGet(e) {
  const isAdminRequest = e && e.parameter && e.parameter.admin === 'true';

  if (isAdminRequest) {
    if (!isAdmin_(getActiveUserEmail_())) {
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
function validatePasswordRule_(newPw, employeeId, currentPw) {
  if (typeof newPw !== 'string' || newPw.length < 8) throw new Error("パスワードは8文字以上で入力してください。");
  if (newPw === employeeId) throw new Error("社員IDと同じパスワードは使用できません。");
  if (newPw === currentPw) throw new Error("現在のパスワードと同じパスワードは使用できません。");
}

/**
 * 管理者リセット用の仮PWを生成する（乱数源は Utilities.getUuid()）
 * 偏りを避けるため、文字種数の倍数未満のバイトだけを採用する
 */
function generateTempPassword_() {
  const n = TEMP_PW_CHARS_.length;
  const limit = Math.floor(256 / n) * n;
  let pw = '';
  while (pw.length < TEMP_PW_LENGTH_) {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid());
    for (let i = 0; i < bytes.length && pw.length < TEMP_PW_LENGTH_; i++) {
      const b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
      if (b < limit) pw += TEMP_PW_CHARS_.charAt(b % n);
    }
  }
  return pw;
}

// ----------------------------------------------------
// 試行回数制限
// ----------------------------------------------------

function authFailKey_(prefix, email) {
  return prefix + computeHash_(normalizeEmail_(email));
}

/**
 * 失敗回数を加算し、上限に達したらロックする（厳密な原子性は求めない）
 */
function recordAuthFailure_(key) {
  const cache = CacheService.getScriptCache();
  const count = Number(cache.get(key) || 0) + 1;
  if (count >= AUTH_FAIL_LIMIT_) {
    cache.put(key + '_lock', '1', AUTH_FAIL_TTL_SEC_);
    cache.remove(key);
  } else {
    cache.put(key, String(count), AUTH_FAIL_TTL_SEC_);
  }
}

/**
 * 対象メールアドレスのログイン・PW変更の失敗回数とロックをすべて解除する
 */
function clearAuthFailures_(email) {
  const keys = [];
  [LOGIN_FAIL_PREFIX_, PW_CHANGE_FAIL_PREFIX_].forEach(prefix => {
    const key = authFailKey_(prefix, email);
    keys.push(key, key + '_lock');
  });
  CacheService.getScriptCache().removeAll(keys);
}

// ----------------------------------------------------
// ログイン・認証 処理
// ----------------------------------------------------

function isActiveEmployee_(emp) {
  return !!emp && emp.isValid === '有効' && emp.status !== '退職';
}

/**
 * メールアドレス＋パスワードの照合。成功時は社員レコード（内部用）を返す
 * failPrefix ごとに試行回数を数え、ロック中はPWを照合せずに拒否する。
 * 存在しないメールアドレスも同じく数え、無効・退職はPWが一致した場合のみ知らせる（登録有無を推測させないため）
 * メールアドレスは大文字・小文字を区別しない（試行回数も同じキーで数える）
 */
function authenticate_(email, password, failPrefix) {
  const cache = CacheService.getScriptCache();
  const failKey = authFailKey_(failPrefix, email);
  if (cache.get(failKey + '_lock')) throw new Error("試行回数が上限に達しました。15分ほどしてから再度お試しください。");

  // メールアドレスは入力・シートとも前後の空白を除いて小文字化して照合する
  const normalizedEmail = normalizeEmail_(email);
  const emp = getEmployeeRecords_().find(e => normalizeEmail_(e.email) === normalizedEmail);
  if (!emp || !verifyPassword_(password, emp.password)) {
    recordAuthFailure_(failKey);
    throw new Error("メールアドレスまたはパスワードが間違っています。");
  }
  cache.remove(failKey);

  if (emp.isValid !== '有効') throw new Error("無効化されているアカウントです。");
  if (emp.status === '退職') throw new Error("退職済みのアカウントです。");
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
    email: emp.email,
    sessionId: createSession_(emp.empId, sessionCreatedAt),
    portalEnabled: canUseSystem_(PORTAL_SYSTEM_KEY_),
    portalPreparing: PORTAL_ADMIN_ONLY_
  };
}

/**
 * ログイン認証処理
 * PW変更要（I列TRUE）の場合はセッションを発行せず mustChangePassword:true を返す
 */
function verifyLogin(email, password) {
  if (!email || !password) throw new Error("メールアドレスとパスワードを入力してください。");

  const emp = authenticate_(email, password, LOGIN_FAIL_PREFIX_);

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

  const emp = authenticate_(email, currentPw, PW_CHANGE_FAIL_PREFIX_);
  validatePasswordRule_(newPw, emp.empId, currentPw);

  const sheet = getCommonSpreadsheet().getSheetByName('社員マスタ');
  const targetRow = findEmployeeRow_(sheet, emp.empId);
  // PW変更日時と同じ時刻でセッションを作成し、変更を行ったセッション自体は継続利用できるようにする
  const now = new Date().getTime();
  sheet.getRange(targetRow, EMP_COL_PASSWORD_, 1, 3).setValues([[hashPassword_(newPw), false, toPwChangedAt_(now)]]);

  return buildLoginResult_(emp, now);
}

/**
 * 管理者によるPWリセット：ランダムな仮PWを設定し、PW変更要をTRUEにする。試行回数のロックも解除する
 * 仮PWは呼び出し元の管理者への戻り値でのみ返す（ログ・シート・プロパティには残さない）
 */
function adminResetPassword(employeeId) {
  assertAdmin_();
  if (!employeeId) throw new Error("社員IDが指定されていません。");

  const sheet = getCommonSpreadsheet().getSheetByName('社員マスタ');
  if (!sheet) throw new Error("「社員マスタ」シートが見つかりません。");
  const targetRow = findEmployeeRow_(sheet, employeeId);
  const tempPassword = generateTempPassword_();
  sheet.getRange(targetRow, EMP_COL_PASSWORD_, 1, 3).setValues([[hashPassword_(tempPassword), true, toPwChangedAt_(new Date().getTime())]]);
  clearAuthFailures_(sheet.getRange(targetRow, 3).getValue()); // C列：メールアドレス

  return { success: true, employeeId: employeeId, tempPassword: tempPassword };
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
 * 遷移先が利用できない（準備中のポータルを管理者以外が開こうとした）場合は、セッションを残したまま
 * { success: false, error } を返す
 * @param {string} sessionId ログインセッションID
 * @param {string} systemKey 遷移先システム（SYSTEM_URLS_ のキー）
 */
function issueSsoTokenForSession(sessionId, systemKey) {
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
  if (!canUseSystem_(systemKey)) {
    return { success: false, error: "このメニューは準備中のため、現在は管理者のみ利用できます。" };
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

/**
 * 社員の新規登録。メールアドレスは前後の空白を除いて小文字化し、既存社員（退職者を含む）との重複は拒否する
 */
function registerEmployeeFromWeb(name, email, status, startDateStr, departmentId, sectionId) {
  assertAdmin_();
  status = status || '在籍';
  assertEmployeeStatus_(status);
  email = normalizeEmail_(email);
  if (!email) throw new Error("メールアドレスを入力してください。");
  const ss = getCommonSpreadsheet();
  const sheet = ss.getSheetByName('社員マスタ');
  if (!sheet) throw new Error("「社員マスタ」シートが見つかりません。");
  let formattedDate = startDateStr ? Utilities.formatDate(new Date(startDateStr), Session.getScriptTimeZone(), 'yyyy/MM/dd') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');

  // 初期パスワードは仮PW（H列にハッシュで保存）とし、I列（PW変更要）をTRUEにする
  // 仮PWは呼び出し元の管理者への戻り値でのみ返す（ログ・シート・プロパティには残さない）
  const tempPassword = generateTempPassword_();

  // メールアドレスの重複確認・社員IDの採番・書き込み（主所属の登録を含む）をまとめてロックする
  const newId = withScriptLock_(() => {
    assertEmployeeEmailAvailable_(sheet.getDataRange().getValues(), email, '');
    const id = generateNewEmployeeId();
    sheet.appendRow([id, name, email, status, formattedDate, '', true, hashPassword_(tempPassword), true]);
    if (departmentId) {
      registerAssignment_({
        employeeId: id, departmentId: departmentId, sectionId: sectionId || '',
        type: '主所属', startDate: formattedDate, endDate: '9999/12/31'
      });
    }
    return id;
  });
  return { success: true, employeeId: newId, tempPassword: tempPassword };
}

/**
 * 社員マスタのメールアドレス（C列）の重複チェック。前後の空白を除いて小文字化して比較し、退職者も対象とする
 * @param {Array<Array>} data 社員マスタの全行（見出し行を含む）
 * @param {string} email 正規化済みのメールアドレス
 * @param {string} excludeEmployeeId 比較から除く社員ID（編集時の本人）。新規登録は空文字
 */
function assertEmployeeEmailAvailable_(data, email, excludeEmployeeId) {
  for (let i = 1; i < data.length; i++) {
    if (excludeEmployeeId && data[i][0] === excludeEmployeeId) continue;
    if (normalizeEmail_(data[i][2]) === email) {
      throw new Error("このメールアドレスは既に登録されています（社員ID: " + data[i][0] + "）。");
    }
  }
}

/**
 * 在籍状況の値の検証（在籍／休職／退職 以外は例外）
 */
function assertEmployeeStatus_(status) {
  if (EMP_STATUSES_.indexOf(status) === -1) throw new Error("在籍状況の値が不正です: " + status);
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
  return withScriptLock_(() => {
    const newId = generateNewDepartmentId();
    sheet.appendRow([newId, name, Number(sortOrder) || 10, true]);
    return newId;
  });
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
  let formattedDate = startDateStr ? Utilities.formatDate(new Date(startDateStr), Session.getScriptTimeZone(), 'yyyy/MM/dd') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  return withScriptLock_(() => {
    const newId = generateNewSectionId();
    sheet.appendRow([newId, name, departmentId, Number(sortOrder) || 10, formattedDate, '', true]);
    return newId;
  });
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
  return withScriptLock_(() => registerAssignment_(param));
}

/**
 * 所属履歴の登録（内部用。呼び出し側で管理者確認とロックの取得を済ませること）
 * 事業部・開始日は必須、部署は任意（部署なし）。部署を指定したときだけ、その事業部に属する部署かを確認する。
 * 主所属は期間の重複を拒否し、履歴IDを採番して追記する
 */
function registerAssignment_(param) {
  if (!param || !param.employeeId) throw new Error("社員IDが指定されていません。");
  if (!param.departmentId) throw new Error("事業部を指定してください。");
  if (!param.startDate) throw new Error("開始日を指定してください。");
  if (param.sectionId) {
    const section = getSections_().find(s => s.id === param.sectionId);
    if (!section) throw new Error("指定された部署が見つかりません: " + param.sectionId);
    if (section.departmentId !== param.departmentId) {
      throw new Error("部署（" + param.sectionId + "）は指定された事業部（" + param.departmentId + "）に属していません。");
    }
  }

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
 * 外部システム（戦略AP・人事評価）連携用API
 * 共通基盤の最新マスタデータを一括返却する
 * ※ライブラリ経由で外部システムから呼ばれるため管理者チェックではなくAPIキーで保護する（パスワード関連の項目は含めない）
 * @param {string} apiKey 連携用APIキー（スクリプトプロパティ API_KEY_STRATEGY_AP / API_KEY_JINJI と照合）
 * @return {Object} 全マスタデータを含むオブジェクト
 */
function exportCommonMasterData(apiKey) {
  assertApiKey_(apiKey, 'exportCommonMasterData');
  return {
    employees: buildEmployeeListForDisplay_(),
    departments: getDepartments_(),
    sections: getSections_(),
    assignments: getAssignments_()
  };
}

/**
 * 連携用APIキーとシステム名の対応表（スクリプトプロパティ名 → システム名）
 */
const API_KEY_SYSTEMS_ = {
  API_KEY_STRATEGY_AP: '戦略AP',
  API_KEY_JINJI: '人事評価'
};

/**
 * 連携用APIキーを照合し、呼び出し元のシステム名を返す。
 * 未指定・いずれのキーとも不一致・プロパティ未設定はいずれも拒否する。
 * ログにはシステム名のみ記録し、キー値は記録しない。
 * @param {string} apiKey 連携用APIキー
 * @param {string} apiName 呼び出されたAPI名（ログ用）
 * @return {string} システム名
 */
function assertApiKey_(apiKey, apiName) {
  if (apiKey) {
    const props = PropertiesService.getScriptProperties();
    for (const propName in API_KEY_SYSTEMS_) {
      const expected = props.getProperty(propName);
      if (expected && apiKey === expected) {
        const systemName = API_KEY_SYSTEMS_[propName];
        console.log(apiName + ': 呼び出し元=' + systemName);
        return systemName;
      }
    }
  }
  console.warn(apiName + ': APIキー認証に失敗しました');
  throw new Error('Unauthorized');
}

/**
 * 外部システム（戦略AP・人事評価）連携用API：セッションの有効性確認
 * 社員が存在しない・無効・退職、またはセッション作成後にPWが変更された場合は無効（休職は有効）
 * PW変更日時の比較は issueSsoTokenForSession と同じ（J列は秒単位に切り捨てて記録、空欄は0で無効化しない）
 * ※理由や社員情報は返さない
 * @param {string} apiKey 連携用APIキー（スクリプトプロパティ API_KEY_STRATEGY_AP / API_KEY_JINJI と照合）
 * @param {string} empId 社員ID
 * @param {number} sessionCreatedAtMs 戦略AP側セッションの作成時刻（ミリ秒）
 * @return {{valid: boolean}}
 */
function checkEmployeeSession(apiKey, empId, sessionCreatedAtMs) {
  const startedAt = Date.now();
  let employeeReadMs = 0;
  let employeeReadExecuted = false;
  try {
    assertApiKey_(apiKey, 'checkEmployeeSession');
    const createdAt = Number(sessionCreatedAtMs);
    if (!empId || !isFinite(createdAt) || createdAt <= 0) return { valid: false };

    const readStartedAt = Date.now();
    let emp;
    employeeReadExecuted = true;
    try {
      emp = getEmployeeSessionState_(empId);
    } finally {
      employeeReadMs = Date.now() - readStartedAt;
    }
    if (!isActiveEmployee_(emp) || createdAt < emp.pwChangedAt) return { valid: false };
    return { valid: true };
  } finally {
    try {
      console.log('[SESSION_CHECK_TRACE] ' + JSON.stringify({
        employeeReadMs: employeeReadMs,
        employeeReadExecuted: employeeReadExecuted,
        totalMs: Date.now() - startedAt
      }));
    } catch (ignored) {
      // 計測ログの失敗で認証結果を変えない。
    }
  }
}

/** セッション確認専用。社員マスタのみ読み、既存の値変換・先頭一致を維持する。 */
function getEmployeeSessionState_(empId) {
  const sheet = getCommonSpreadsheet().getSheetByName('社員マスタ');
  if (!sheet) return null;
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if ((row[0] || '') !== empId) continue;
    return {
      empId: row[0] || '',
      status: row[3] || '',
      isValid: row[6] === false ? '無効' : '有効',
      pwChangedAt: row[9] instanceof Date ? row[9].getTime() : 0
    };
  }
  return null;
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
 * メールアドレスは前後の空白を除いて小文字化し、本人以外の社員（退職者を含む）との重複は拒否する。
 * 在籍状況が「退職」の場合のみ利用終了日を必須とし、そうでない場合は利用終了日をクリアする。
 * 退職済みの社員は在籍状況を変更できない（退職のまま、他の項目の編集は可）。
 * 退職済みの社員は利用終了日を必須とせず、指定された場合のみ上書きする。
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
  email = normalizeEmail_(email);
  if (!email) throw new Error("メールアドレスを入力してください。");
  assertEmployeeEmailAvailable_(data, email, employeeId);
  assertEmployeeStatus_(status);
  const alreadyRetired = data[targetRow - 1][3] === '退職';
  if (alreadyRetired && status !== '退職') throw new Error("退職済みの社員の在籍状況は変更できません。");
  if (status === '退職' && !alreadyRetired && !endDateStr) throw new Error("退職の場合は利用終了日を指定してください。");

  sheet.getRange(targetRow, 2).setValue(name);
  sheet.getRange(targetRow, 3).setValue(email);

  if (status === '退職') {
    sheet.getRange(targetRow, 4).setValue('退職');
    if (endDateStr) {
      const formattedEndDate = Utilities.formatDate(new Date(endDateStr), Session.getScriptTimeZone(), 'yyyy/MM/dd');
      sheet.getRange(targetRow, 6).setValue(formattedEndDate);
    }
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

  // 現在の主所属の終了〜新しい主所属の採番・追記までをまとめてロックする
  return withScriptLock_(() => {
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

    return registerAssignment_({
      employeeId: employeeId,
      departmentId: newDepartmentId,
      sectionId: newSectionId || '',
      type: '主所属',
      startDate: effectiveDateStr,
      endDate: ''
    });
  });
}