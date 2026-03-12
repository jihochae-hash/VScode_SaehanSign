// ============================================================
// 품질 전자결재 시스템 - Google Apps Script 백엔드
// (주)새한화장품 품질관리팀
// ============================================================

// ========== 설정 ==========
const PROPS = PropertiesService.getScriptProperties();

function getSpreadsheet() {
  const id = PROPS.getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('품질전자결재_DB');
  PROPS.setProperty('SS_ID', ss.getId());
  initializeSheets(ss);
  return ss;
}

function getDriveRootFolder() {
  let folderId = PROPS.getProperty('DRIVE_ROOT_ID');
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch(e) {}
  }
  const folder = DriveApp.createFolder('품질전자결재');
  PROPS.setProperty('DRIVE_ROOT_ID', folder.getId());
  folder.createFolder('템플릿');
  folder.createFolder('서명');
  folder.createFolder('원본문서');
  folder.createFolder('승인문서');
  return folder;
}

// ========== 초기화 ==========
function initializeSheets(ss) {
  let sh = ss.getSheetByName('Sheet1');
  if (sh) sh.setName('사용자');
  else sh = ss.insertSheet('사용자');
  sh.getRange(1, 1, 1, 11).setValues([['id','username','password','name','department','position','role','signature_file_id','created_at','status','email']]);

  sh = ss.insertSheet('템플릿');
  sh.getRange(1, 1, 1, 10).setValues([['id','name','category','description','file_id','signature_config','approval_steps','created_at','created_by','status']]);

  sh = ss.insertSheet('문서');
  sh.getRange(1, 1, 1, 20).setValues([['id','doc_number','title','template_id','category','file_id','creator_id','creator_name','status','current_step','created_at','updated_at','pdf_file_id','pdf_url','save_path','approval_config','rejection_comment','rejection_by','file_type','product_code']]);

  sh = ss.insertSheet('결재이력');
  sh.getRange(1, 1, 1, 13).setValues([['id','doc_id','step_order','step_name','approver_id','approver_name','approver_dept','approver_position','status','signed_at','comment','signature_applied','signature_file_id']]);

  sh = ss.insertSheet('알림');
  sh.getRange(1, 1, 1, 8).setValues([['id','doc_id','user_id','type','message','read','created_at','doc_title']]);

  sh = ss.insertSheet('설정');
  sh.getRange(1, 1, 1, 2).setValues([['key','value']]);
  sh.getRange(2, 1, 7, 2).setValues([
    ['webhook_url', ''],
    ['save_path_template', '승인문서/{year}/{month}'],
    ['company_name', '(주)새한화장품'],
    ['default_approval_chain', '[]'],
    ['doc_categories', JSON.stringify([
      {value:'성적서', label:'완제품 시험 성적서'},
      {value:'검사기록', label:'검사 기록서'},
      {value:'시험의뢰', label:'시험 의뢰서'},
      {value:'기타', label:'기타'}
    ])],
    ['doc_type_paths', '{}'],
    ['author_sign_first', 'false']
  ]);

  sh = ss.insertSheet('BOM_QA');
  sh.getRange(1, 1, 1, 3).setValues([['product_code','product_name','updated_at']]);

  sh = ss.insertSheet('로그인이력');
  sh.getRange(1, 1, 1, 6).setValues([['id','user_id','username','name','login_at','ip']]);

  const userSheet = ss.getSheetByName('사용자');
  userSheet.getRange(2, 1, 1, 11).setValues([[
    generateId(), 'admin', hashPassword('admin1234'), '관리자', '품질관리팀', '팀장', 'admin', '', new Date().toISOString(), 'active', ''
  ]]);
}

// ========== 웹앱 ==========
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('품질 전자결재 시스템')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  const params = e.parameter;
  const action = params.action || (e.postData ? JSON.parse(e.postData.contents).action : '');
  let data = {};
  try {
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
  } catch(ex) {}
  Object.keys(params).forEach(k => { if (k !== 'action') data[k] = params[k]; });

  try {
    let result;
    switch(action) {
      case 'login': result = handleLogin(data); break;
      case 'logout': result = handleLogout(data); break;
      case 'validate_session': result = validateSession(data.token); break;
      case 'change_password': result = handleChangePassword(data); break;

      case 'user_list': result = getUsers(data); break;
      case 'user_add': result = addUser(data); break;
      case 'user_update': result = updateUser(data); break;
      case 'user_delete': result = deleteUser(data); break;
      case 'user_signature_upload': result = uploadSignature(data); break;

      case 'template_list': result = getTemplates(data); break;
      case 'template_upload': result = uploadTemplate(data); break;
      case 'template_update': result = updateTemplate(data); break;
      case 'template_delete': result = deleteTemplate(data); break;
      case 'template_file': result = getTemplateFile(data); break;

      case 'doc_list': result = getDocuments(data); break;
      case 'doc_create': result = createDocument(data); break;
      case 'doc_detail': result = getDocumentDetail(data); break;
      case 'doc_delete': result = deleteDocument(data); break;
      case 'doc_file': result = getDocumentFile(data); break;
      case 'doc_search': result = searchDocuments(data); break;

      case 'approval_submit': result = submitForApproval(data); break;
      case 'approval_approve': result = approveDocument(data); break;
      case 'approval_reject': result = rejectDocument(data); break;
      case 'approval_history': result = getApprovalHistory(data); break;

      case 'save_pdf': result = savePdfToDrive(data); break;
      case 'regen_pdf': result = regenPdf(data); break;

      case 'notifications': result = getNotifications(data); break;
      case 'notification_read': result = markNotificationRead(data); break;

      case 'settings_get': result = getSettings(); break;
      case 'settings_update': result = updateSettings(data); break;

      case 'dashboard': result = getDashboard(data); break;

      case 'get_signature_image': result = getSignatureImage(data); break;
      case 'bom_qa_upload': result = bomQaUpload(data); break;
      case 'bom_qa_search': result = bomQaSearch(data); break;
      case 'login_history': result = getLoginHistory(data); break;

      default: result = { success: false, error: '알 수 없는 요청: ' + action };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(ex) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false, error: ex.message, stack: ex.stack
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ========== 유틸리티 ==========
function generateId() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 12);
}

function hashPassword(pw) {
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw);
  return hash.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function generateToken() {
  return Utilities.getUuid() + '-' + Date.now();
}

function generateDocNumber() {
  const now = new Date();
  const prefix = 'QA' + now.getFullYear().toString().slice(-2) +
    ('0' + (now.getMonth()+1)).slice(-2);
  const sh = getSheet('문서');
  const data = sh.getDataRange().getValues();
  let maxSeq = 0;
  data.forEach(row => {
    if (row[1] && row[1].toString().startsWith(prefix)) {
      const seq = parseInt(row[1].toString().slice(-4));
      if (seq > maxSeq) maxSeq = seq;
    }
  });
  return prefix + ('0000' + (maxSeq + 1)).slice(-4);
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    initializeSheets(ss);
    sh = ss.getSheetByName(name);
  }
  return sh;
}

// ========== 캐시 유틸리티 ==========
// GAS CacheService 래퍼 (콜드 스타트 및 반복 Sheets 읽기 최소화)
const _CACHE = CacheService.getScriptCache();
const _CACHE_TTL = 300; // 기본 5분

function cacheGet(key) {
  try { const v = _CACHE.get('qa_' + key); return v ? JSON.parse(v) : null; } catch(e) { return null; }
}
function cachePut(key, data, ttl) {
  try { _CACHE.put('qa_' + key, JSON.stringify(data), ttl || _CACHE_TTL); } catch(e) {}
}
function cacheDel(key) {
  try { _CACHE.remove('qa_' + key); } catch(e) {}
}

function findRowIndex(sheet, colIndex, value) {
  const data = sheet.getDataRange().getValues();
  const strVal = String(value);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]) === strVal) return i + 1;
  }
  return -1;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ========== 인증 ==========
function handleLogin(data) {
  const sh = getSheet('사용자');
  const users = sheetToObjects(sh);
  const user = users.find(u => u.username === data.username && u.status === 'active');
  if (!user) return { success: false, error: '아이디 또는 비밀번호가 일치하지 않습니다.' };
  if (user.password !== hashPassword(data.password)) {
    return { success: false, error: '아이디 또는 비밀번호가 일치하지 않습니다.' };
  }
  const token = generateToken();
  const cache = CacheService.getScriptCache();
  cache.put('session_' + token, JSON.stringify({
    id: String(user.id), username: user.username, name: user.name,
    department: user.department, position: user.position, role: user.role,
    signature_file_id: user.signature_file_id || '', email: user.email || ''
  }), 21600);
  // 로그인 이력 기록
  try { recordLoginHistory(String(user.id), user.username, user.name, data.ip || ''); } catch(e) {}
  return {
    success: true,
    token: token,
    user: {
      id: String(user.id), username: user.username, name: user.name,
      department: user.department, position: user.position, role: user.role,
      signature_file_id: user.signature_file_id || '', email: user.email || ''
    }
  };
}

function validateSession(token) {
  if (!token) return { success: false, error: '세션이 만료되었습니다.' };
  const cache = CacheService.getScriptCache();
  const sessionStr = cache.get('session_' + token);
  if (!sessionStr) return { success: false, error: '세션이 만료되었습니다.' };
  return { success: true, user: JSON.parse(sessionStr) };
}

function getSessionUser(data) {
  const result = validateSession(data.token);
  if (!result.success) throw new Error('인증이 필요합니다.');
  return result.user;
}

function handleLogout(data) {
  if (data.token) CacheService.getScriptCache().remove('session_' + data.token);
  return { success: true };
}

function handleChangePassword(data) {
  const user = getSessionUser(data);
  const sh = getSheet('사용자');
  const rowIdx = findRowIndex(sh, 0, user.id);
  if (rowIdx < 0) return { success: false, error: '사용자를 찾을 수 없습니다.' };
  const currentPw = sh.getRange(rowIdx, 3).getValue();
  if (currentPw !== hashPassword(data.current_password)) {
    return { success: false, error: '현재 비밀번호가 일치하지 않습니다.' };
  }
  sh.getRange(rowIdx, 3).setValue(hashPassword(data.new_password));
  SpreadsheetApp.flush();
  return { success: true };
}

// ========== 사용자 관리 ==========
function getUsers(data) {
  getSessionUser(data);
  const cached = cacheGet('users');
  if (cached) return cached;
  const sh = getSheet('사용자');
  const users = sheetToObjects(sh).filter(u => u.status === 'active');
  const result = {
    success: true,
    users: users.map(u => ({
      id: String(u.id), username: u.username, name: u.name,
      department: u.department, position: u.position, role: u.role,
      signature_file_id: u.signature_file_id || '', email: u.email || '',
      created_at: u.created_at
    }))
  };
  cachePut('users', result, 300);
  return result;
}

function addUser(data) {
  const admin = getSessionUser(data);
  if (admin.role !== 'admin') return { success: false, error: '관리자 권한이 필요합니다.' };
  const sh = getSheet('사용자');
  const users = sheetToObjects(sh);
  if (users.find(u => u.username === data.username && u.status === 'active')) {
    return { success: false, error: '이미 존재하는 아이디입니다.' };
  }
  const id = generateId();
  sh.appendRow([
    id, data.username, hashPassword(data.password || '1234'),
    data.name, data.department, data.position, data.role || 'user',
    '', new Date().toISOString(), 'active', data.email || ''
  ]);
  SpreadsheetApp.flush();
  cacheDel('users'); // 사용자 캐시 무효화
  return { success: true, id: id };
}

function updateUser(data) {
  const admin = getSessionUser(data);
  if (admin.role !== 'admin' && admin.id !== data.id) {
    return { success: false, error: '권한이 없습니다.' };
  }
  const sh = getSheet('사용자');
  const rowIdx = findRowIndex(sh, 0, data.id);
  if (rowIdx < 0) return { success: false, error: '사용자를 찾을 수 없습니다.' };
  if (data.name) sh.getRange(rowIdx, 4).setValue(data.name);
  if (data.department) sh.getRange(rowIdx, 5).setValue(data.department);
  if (data.position) sh.getRange(rowIdx, 6).setValue(data.position);
  if (data.role && admin.role === 'admin') sh.getRange(rowIdx, 7).setValue(data.role);
  if (data.password) sh.getRange(rowIdx, 3).setValue(hashPassword(data.password));
  if (data.email !== undefined) sh.getRange(rowIdx, 11).setValue(data.email);
  SpreadsheetApp.flush();
  cacheDel('users'); // 사용자 캐시 무효화
  return { success: true };
}

function deleteUser(data) {
  const admin = getSessionUser(data);
  if (admin.role !== 'admin') return { success: false, error: '관리자 권한이 필요합니다.' };
  const sh = getSheet('사용자');
  const rowIdx = findRowIndex(sh, 0, data.id);
  if (rowIdx < 0) return { success: false, error: '사용자를 찾을 수 없습니다.' };
  sh.getRange(rowIdx, 10).setValue('inactive');
  SpreadsheetApp.flush();
  cacheDel('users'); // 사용자 캐시 무효화
  return { success: true };
}

function uploadSignature(data) {
  const user = getSessionUser(data);
  const targetId = data.target_id || user.id;
  if (String(targetId) !== String(user.id) && user.role !== 'admin') {
    return { success: false, error: '권한이 없습니다.' };
  }
  const root = getDriveRootFolder();
  const sigFolder = getSubFolder(root, '서명');
  const blob = Utilities.newBlob(Utilities.base64Decode(data.image_data), data.mime_type || 'image/png', 'sig_' + targetId + '.png');
  const sh = getSheet('사용자');
  const rowIdx = findRowIndex(sh, 0, targetId);
  const oldFileId = sh.getRange(rowIdx, 8).getValue();
  if (oldFileId) {
    try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch(e) {}
  }
  const file = sigFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  sh.getRange(rowIdx, 8).setValue(file.getId());
  SpreadsheetApp.flush();
  cacheDel('users'); // 사용자 캐시 무효화
  return { success: true, file_id: file.getId() };
}

// ========== 템플릿 관리 ==========
function getTemplates(data) {
  getSessionUser(data);
  const sh = getSheet('템플릿');
  const templates = sheetToObjects(sh).filter(t => t.status !== 'deleted');
  return { success: true, templates: templates };
}

function uploadTemplate(data) {
  const user = getSessionUser(data);
  const root = getDriveRootFolder();
  const tplFolder = getSubFolder(root, '템플릿');
  const mimeType = detectMimeType(data.file_name);
  const blob = Utilities.newBlob(Utilities.base64Decode(data.file_data), mimeType, data.file_name);
  const file = tplFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = generateId();
  const sh = getSheet('템플릿');
  sh.appendRow([
    id, data.name, data.category || '성적서', data.description || '',
    file.getId(), data.signature_config || '{}', data.approval_steps || '[]',
    new Date().toISOString(), user.id, 'active'
  ]);
  SpreadsheetApp.flush();
  return { success: true, id: id };
}

function updateTemplate(data) {
  const user = getSessionUser(data);
  if (user.role !== 'admin') return { success: false, error: '관리자 권한이 필요합니다.' };
  const sh = getSheet('템플릿');
  const rowIdx = findRowIndex(sh, 0, data.id);
  if (rowIdx < 0) return { success: false, error: '템플릿을 찾을 수 없습니다.' };
  if (data.name) sh.getRange(rowIdx, 2).setValue(data.name);
  if (data.category) sh.getRange(rowIdx, 3).setValue(data.category);
  if (data.description !== undefined) sh.getRange(rowIdx, 4).setValue(data.description);
  if (data.signature_config) sh.getRange(rowIdx, 6).setValue(data.signature_config);
  if (data.approval_steps) sh.getRange(rowIdx, 7).setValue(data.approval_steps);
  SpreadsheetApp.flush();
  return { success: true };
}

function deleteTemplate(data) {
  const user = getSessionUser(data);
  if (user.role !== 'admin') return { success: false, error: '관리자 권한이 필요합니다.' };
  const sh = getSheet('템플릿');
  const rowIdx = findRowIndex(sh, 0, data.id);
  if (rowIdx < 0) return { success: false, error: '템플릿을 찾을 수 없습니다.' };
  sh.getRange(rowIdx, 10).setValue('deleted');
  SpreadsheetApp.flush();
  return { success: true };
}

function getTemplateFile(data) {
  getSessionUser(data);
  const sh = getSheet('템플릿');
  const rowIdx = findRowIndex(sh, 0, data.id);
  if (rowIdx < 0) return { success: false, error: '템플릿을 찾을 수 없습니다.' };
  const fileId = sh.getRange(rowIdx, 5).getValue();
  const file = DriveApp.getFileById(fileId);
  const bytes = file.getBlob().getBytes();
  return { success: true, data: Utilities.base64Encode(bytes), name: file.getName() };
}

// ========== 문서 관리 ==========
function getDocuments(data) {
  const user = getSessionUser(data);
  const sh = getSheet('문서');
  let docs = sheetToObjects(sh).filter(d => d.status !== 'deleted');

  // ID를 모두 String으로 통일
  docs.forEach(d => { d.id = String(d.id); d.creator_id = String(d.creator_id); });

  if (data.filter === 'my_created') {
    docs = docs.filter(d => d.creator_id === String(user.id));
  } else if (data.filter === 'pending_my_approval') {
    const approvalSh = getSheet('결재이력');
    const approvals = sheetToObjects(approvalSh);
    const pendingDocIds = approvals
      .filter(a => String(a.approver_id) === String(user.id) && a.status === 'pending')
      .map(a => String(a.doc_id));
    docs = docs.filter(d => pendingDocIds.includes(d.id));
  } else if (data.filter === 'approved') {
    docs = docs.filter(d => d.status === 'approved');
  } else if (data.filter === 'rejected') {
    docs = docs.filter(d => d.status === 'rejected');
  }

  if (data.category) docs = docs.filter(d => d.category === data.category);
  docs.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
  return { success: true, documents: docs };
}

function createDocument(data) {
  const user = getSessionUser(data);
  const root = getDriveRootFolder();
  const docFolder = getSubFolder(root, '원본문서');

  // 파일 타입 감지 (엑셀 또는 PDF)
  const mimeType = detectMimeType(data.file_name);
  const fileType = mimeType === 'application/pdf' ? 'pdf' : 'excel';

  const blob = Utilities.newBlob(
    Utilities.base64Decode(data.file_data),
    mimeType,
    data.file_name
  );
  const file = docFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const id = generateId();
  const docNumber = generateDocNumber();
  const sh = getSheet('문서');
  const approvalConfig = data.approval_steps || '[]';

  sh.appendRow([
    id, docNumber, data.title, data.template_id || '', data.category || '성적서',
    file.getId(), user.id, user.name, 'draft', 0,
    new Date().toISOString(), new Date().toISOString(),
    '', '', data.save_path || '', approvalConfig, '', '', fileType,
    data.product_code || '' // col 20: 제품코드
  ]);

  // 결재이력 초기 생성
  const steps = JSON.parse(approvalConfig);
  const approvalSh = getSheet('결재이력');
  let stepOffset = 0;

  // 작성자 서명 자동 포함 옵션
  if (data.include_author_signature === true || data.include_author_signature === 'true') {
    if (steps.length > 0) {
      const userShA = getSheet('사용자');
      const creatorRow = findRowIndex(userShA, 0, user.id);
      const creatorSigId = creatorRow > 0 ? userShA.getRange(creatorRow, 8).getValue() || '' : '';
      approvalSh.appendRow([
        generateId(), id, 1, '작성', user.id, user.name,
        user.department, user.position, 'approved',
        new Date().toISOString(), '작성자 서명', 'Y', creatorSigId
      ]);
      stepOffset = 1;
    }
  }

  steps.forEach((step, idx) => {
    const actualStepNum = idx + 1 + stepOffset;
    // 결재자의 최신 서명 파일 ID를 사용자 시트에서 가져옴
    let sigFileId = step.signature_file_id || '';
    if (!sigFileId && step.approver_id) {
      const userSh = getSheet('사용자');
      const userRow = findRowIndex(userSh, 0, step.approver_id);
      if (userRow > 0) sigFileId = userSh.getRange(userRow, 8).getValue() || '';
    }
    approvalSh.appendRow([
      generateId(), id, actualStepNum, step.name, step.approver_id,
      step.approver_name, step.approver_dept || '', step.approver_position || '',
      idx === 0 ? 'pending' : 'waiting', '', '', '', sigFileId
    ]);
  });

  SpreadsheetApp.flush();
  cacheDel('dashboard_' + user.id); // 대시보드 캐시 무효화
  return { success: true, id: id, doc_number: docNumber };
}

function getDocumentDetail(data) {
  const user = getSessionUser(data);
  const sh = getSheet('문서');
  const docs = sheetToObjects(sh);
  const doc = docs.find(d => String(d.id) === String(data.id));
  if (!doc) return { success: false, error: '문서를 찾을 수 없습니다.' };
  doc.id = String(doc.id);
  doc.creator_id = String(doc.creator_id);

  const approvalSh = getSheet('결재이력');
  const approvals = sheetToObjects(approvalSh)
    .filter(a => String(a.doc_id) === String(data.id))
    .sort((a, b) => Number(a.step_order) - Number(b.step_order));

  // 서명 이미지를 base64로 가져오기
  approvals.forEach(a => {
    a.id = String(a.id);
    a.doc_id = String(a.doc_id);
    a.approver_id = String(a.approver_id);
    a.step_order = Number(a.step_order);
    if (a.signature_file_id && a.status === 'approved') {
      try {
        const file = DriveApp.getFileById(a.signature_file_id);
        const bytes = file.getBlob().getBytes();
        a.signature_base64 = Utilities.base64Encode(bytes);
        a.signature_mime = file.getMimeType();
      } catch(e) {
        a.signature_base64 = '';
        a.signature_mime = '';
      }
    }
  });

  // file_id로 Google Drive 미리보기 URL 제공
  doc.preview_url = doc.file_id ? 'https://drive.google.com/file/d/' + doc.file_id + '/preview' : '';

  return { success: true, document: doc, approvals: approvals };
}

function deleteDocument(data) {
  const user = getSessionUser(data);
  const sh = getSheet('문서');
  const rowIdx = findRowIndex(sh, 0, data.id);
  if (rowIdx < 0) return { success: false, error: '문서를 찾을 수 없습니다.' };
  const creatorId = String(sh.getRange(rowIdx, 7).getValue());
  const status = sh.getRange(rowIdx, 9).getValue();
  if (creatorId !== String(user.id) && user.role !== 'admin') {
    return { success: false, error: '삭제 권한이 없습니다.' };
  }
  if (status === 'approved') {
    return { success: false, error: '승인 완료된 문서는 삭제할 수 없습니다.' };
  }
  sh.getRange(rowIdx, 9).setValue('deleted');
  SpreadsheetApp.flush();
  return { success: true };
}

function getDocumentFile(data) {
  getSessionUser(data);
  const sh = getSheet('문서');
  const rowIdx = findRowIndex(sh, 0, data.id);
  if (rowIdx < 0) return { success: false, error: '문서를 찾을 수 없습니다.' };
  const fileId = sh.getRange(rowIdx, 6).getValue();
  if (!fileId) return { success: false, error: '파일이 없습니다.' };
  const file = DriveApp.getFileById(fileId);
  return {
    success: true,
    file_id: fileId,
    preview_url: 'https://drive.google.com/file/d/' + fileId + '/preview',
    name: file.getName(),
    mime_type: file.getMimeType()
  };
}

function searchDocuments(data) {
  const user = getSessionUser(data);
  const sh = getSheet('문서');
  let docs = sheetToObjects(sh).filter(d => d.status !== 'deleted');
  docs.forEach(d => { d.id = String(d.id); d.creator_id = String(d.creator_id); });
  const q = (data.query || '').toLowerCase();
  if (q) {
    docs = docs.filter(d =>
      (d.title || '').toLowerCase().includes(q) ||
      (d.doc_number || '').toLowerCase().includes(q) ||
      (d.creator_name || '').toLowerCase().includes(q) ||
      (d.category || '').toLowerCase().includes(q)
    );
  }
  if (data.status) docs = docs.filter(d => d.status === data.status);
  if (data.date_from) docs = docs.filter(d => d.created_at >= data.date_from);
  if (data.date_to) docs = docs.filter(d => d.created_at <= data.date_to + 'T23:59:59');
  docs.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
  return { success: true, documents: docs };
}

// ========== 결재 워크플로우 ==========
function submitForApproval(data) {
  const user = getSessionUser(data);
  const sh = getSheet('문서');
  const rowIdx = findRowIndex(sh, 0, data.doc_id);
  if (rowIdx < 0) return { success: false, error: '문서를 찾을 수 없습니다.' };

  const creatorId = String(sh.getRange(rowIdx, 7).getValue());
  if (creatorId !== String(user.id)) return { success: false, error: '작성자만 결재 요청할 수 있습니다.' };

  sh.getRange(rowIdx, 9).setValue('pending');
  sh.getRange(rowIdx, 10).setValue(1);
  sh.getRange(rowIdx, 12).setValue(new Date().toISOString());
  SpreadsheetApp.flush();

  const approvalSh = getSheet('결재이력');
  const approvals = sheetToObjects(approvalSh)
    .filter(a => String(a.doc_id) === String(data.doc_id))
    .sort((a, b) => Number(a.step_order) - Number(b.step_order));

  if (approvals.length > 0) {
    const title = sh.getRange(rowIdx, 3).getValue();
    // 작성자 서명이 이미 포함된 경우 첫 번째 pending 단계로 알림
    const firstPending = approvals.find(a => a.status === 'pending');
    const notifyTarget = firstPending || approvals[0];
    if (firstPending) {
      sh.getRange(rowIdx, 10).setValue(Number(firstPending.step_order));
      SpreadsheetApp.flush();
    }
    createNotification(String(notifyTarget.approver_id), data.doc_id, 'approval_request',
      user.name + '님이 "' + title + '" 문서의 결재를 요청했습니다.', title);
  }

  return { success: true };
}

function approveDocument(data) {
  const user = getSessionUser(data);
  const approvalSh = getSheet('결재이력');
  const allApprovals = sheetToObjects(approvalSh);

  // 모든 ID를 String으로 통일하고 step_order를 Number로 변환
  allApprovals.forEach(a => {
    a.id = String(a.id);
    a.doc_id = String(a.doc_id);
    a.approver_id = String(a.approver_id);
    a.step_order = Number(a.step_order);
  });

  const docApprovals = allApprovals
    .filter(a => a.doc_id === String(data.doc_id))
    .sort((a, b) => a.step_order - b.step_order);

  // 현재 사용자의 결재 단계 찾기
  const myStep = docApprovals.find(a => a.approver_id === String(user.id) && a.status === 'pending');
  if (!myStep) return { success: false, error: '결재 권한이 없거나 이미 처리되었습니다.' };

  // 결재자의 최신 서명 파일 가져오기
  const userSh = getSheet('사용자');
  const userRow = findRowIndex(userSh, 0, user.id);
  let currentSigFileId = '';
  if (userRow > 0) {
    currentSigFileId = userSh.getRange(userRow, 8).getValue() || '';
  }

  // 결재이력 업데이트
  const rowIdx = findRowIndex(approvalSh, 0, myStep.id);
  if (rowIdx < 0) return { success: false, error: '결재 이력을 찾을 수 없습니다.' };

  approvalSh.getRange(rowIdx, 9).setValue('approved');
  approvalSh.getRange(rowIdx, 10).setValue(new Date().toISOString());
  approvalSh.getRange(rowIdx, 11).setValue(data.comment || '');
  approvalSh.getRange(rowIdx, 12).setValue('Y');
  // 서명 파일 ID 업데이트 (최신 서명 사용)
  if (currentSigFileId) {
    approvalSh.getRange(rowIdx, 13).setValue(currentSigFileId);
  }
  SpreadsheetApp.flush();

  // 다음 단계 활성화 또는 최종 승인
  const nextStep = docApprovals.find(a => a.step_order === myStep.step_order + 1);
  const docSh = getSheet('문서');
  const docRowIdx = findRowIndex(docSh, 0, data.doc_id);
  if (docRowIdx < 0) return { success: false, error: '문서를 찾을 수 없습니다.' };
  const docTitle = docSh.getRange(docRowIdx, 3).getValue();
  const creatorId = String(docSh.getRange(docRowIdx, 7).getValue());

  if (nextStep) {
    // 다음 결재자 pending으로 변경
    const nextRowIdx = findRowIndex(approvalSh, 0, nextStep.id);
    if (nextRowIdx > 0) {
      approvalSh.getRange(nextRowIdx, 9).setValue('pending');
    }
    docSh.getRange(docRowIdx, 10).setValue(myStep.step_order + 1);
    SpreadsheetApp.flush();

    createNotification(nextStep.approver_id, data.doc_id, 'approval_request',
      user.name + '님이 "' + docTitle + '" 문서를 승인했습니다. 결재를 진행해주세요.', docTitle);
    createNotification(creatorId, data.doc_id, 'progress',
      user.name + '님이 "' + docTitle + '" 문서를 승인했습니다.', docTitle);
  } else {
    // 최종 승인
    docSh.getRange(docRowIdx, 9).setValue('approved');
    docSh.getRange(docRowIdx, 12).setValue(new Date().toISOString());
    SpreadsheetApp.flush();

    // 원본 파일을 PDF로 변환하여 저장
    try {
      autoSavePdf(data.doc_id, docRowIdx, docSh);
    } catch(pdfErr) {
      Logger.log('PDF 자동저장 실패: ' + pdfErr.message);
    }

    createNotification(creatorId, data.doc_id, 'approved',
      '"' + docTitle + '" 문서가 최종 승인되었습니다.', docTitle);
  }

  // 결재 처리 후 관련 사용자 대시보드 캐시 무효화
  cacheDel('dashboard_' + String(user.id));
  cacheDel('dashboard_' + creatorId);
  return { success: true, is_final: !nextStep };
}

function rejectDocument(data) {
  const user = getSessionUser(data);
  const approvalSh = getSheet('결재이력');
  const allApprovals = sheetToObjects(approvalSh);
  allApprovals.forEach(a => {
    a.id = String(a.id);
    a.doc_id = String(a.doc_id);
    a.approver_id = String(a.approver_id);
  });

  const myStep = allApprovals.find(a =>
    a.doc_id === String(data.doc_id) && a.approver_id === String(user.id) && a.status === 'pending'
  );
  if (!myStep) return { success: false, error: '결재 권한이 없습니다.' };

  const rowIdx = findRowIndex(approvalSh, 0, myStep.id);
  approvalSh.getRange(rowIdx, 9).setValue('rejected');
  approvalSh.getRange(rowIdx, 10).setValue(new Date().toISOString());
  approvalSh.getRange(rowIdx, 11).setValue(data.comment || '반려');

  const docSh = getSheet('문서');
  const docRowIdx = findRowIndex(docSh, 0, data.doc_id);
  docSh.getRange(docRowIdx, 9).setValue('rejected');
  docSh.getRange(docRowIdx, 12).setValue(new Date().toISOString());
  docSh.getRange(docRowIdx, 17).setValue(data.comment || '반려');
  docSh.getRange(docRowIdx, 18).setValue(user.name);
  SpreadsheetApp.flush();

  const creatorId = String(docSh.getRange(docRowIdx, 7).getValue());
  const docTitle = docSh.getRange(docRowIdx, 3).getValue();
  createNotification(creatorId, data.doc_id, 'rejected',
    user.name + '님이 "' + docTitle + '" 문서를 반려했습니다. 사유: ' + (data.comment || ''), docTitle);

  return { success: true };
}

function getApprovalHistory(data) {
  getSessionUser(data);
  const sh = getSheet('결재이력');
  const approvals = sheetToObjects(sh)
    .filter(a => String(a.doc_id) === String(data.doc_id))
    .sort((a, b) => Number(a.step_order) - Number(b.step_order));

  approvals.forEach(a => {
    a.id = String(a.id);
    a.approver_id = String(a.approver_id);
    a.step_order = Number(a.step_order);
    // 승인된 단계의 서명 이미지를 base64로
    if (a.signature_file_id && a.status === 'approved') {
      try {
        const file = DriveApp.getFileById(a.signature_file_id);
        const bytes = file.getBlob().getBytes();
        a.signature_base64 = Utilities.base64Encode(bytes);
        a.signature_mime = file.getMimeType();
      } catch(e) {
        a.signature_base64 = '';
      }
    }
  });
  return { success: true, approvals: approvals };
}

// ========== PDF 자동 저장 (최종 승인 시) ==========
function autoSavePdf(docId, docRowIdx, docSh) {
  const fileId = docSh.getRange(docRowIdx, 6).getValue();
  if (!fileId) return;

  const root = getDriveRootFolder();
  // 캐시된 설정 + 문서 종류별 경로
  const settingsResult = getSettings();
  const docCategory = docSh.getRange(docRowIdx, 5).getValue() || '';
  let docTypePaths = {};
  try { docTypePaths = JSON.parse(settingsResult.settings.doc_type_paths || '{}'); } catch(e) {}
  const basePathTemplate = settingsResult.settings.save_path_template || '승인문서/{year}/{month}';
  const pathTemplate = docTypePaths[docCategory] || (basePathTemplate + '/' + docCategory);
  const customPath = docSh.getRange(docRowIdx, 15).getValue();
  const now = new Date();
  let folderPath = pathTemplate
    .replace('{year}', String(now.getFullYear()))
    .replace('{month}', ('0' + (now.getMonth()+1)).slice(-2))
    .replace('{doc_type}', docCategory);
  if (customPath) folderPath = customPath;

  const targetFolder = createFolderPath(root, folderPath);
  const docNumber = docSh.getRange(docRowIdx, 2).getValue();
  const docTitle = docSh.getRange(docRowIdx, 3).getValue();
  const fileName = docNumber + '_' + docTitle + '.pdf';

  const originalFile = DriveApp.getFileById(fileId);
  let savedFile;
  let savedUrl;
  const origMime = originalFile.getMimeType();

  if (origMime === 'application/pdf') {
    // PDF 그대로 복사
    const blob = originalFile.getBlob().setName(fileName);
    savedFile = targetFolder.createFile(blob);
  } else {
    // Excel → Google Sheets 변환 후 PDF 추출
    try {
      let pdfBlob;
      try {
        pdfBlob = originalFile.getAs('application/pdf').setName(fileName);
      } catch(convErr) {
        // getAs 실패 시 xlsx 원본 파일명(.xlsx)으로 저장
        const origExt = originalFile.getName().split('.').pop();
        const origFileName = docNumber + '_' + docTitle + '.' + origExt;
        const origBlob = originalFile.getBlob().setName(origFileName);
        savedFile = targetFolder.createFile(origBlob);
        savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        docSh.getRange(docRowIdx, 13).setValue(savedFile.getId());
        docSh.getRange(docRowIdx, 14).setValue(savedFile.getUrl());
        docSh.getRange(docRowIdx, 15).setValue(folderPath);
        SpreadsheetApp.flush();
        Logger.log('PDF 변환 불가, 원본 저장: ' + convErr.message);
        return;
      }
      savedFile = targetFolder.createFile(pdfBlob);
    } catch(e) {
      Logger.log('PDF 저장 오류: ' + e.message);
      throw e;
    }
  }
  savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  docSh.getRange(docRowIdx, 13).setValue(savedFile.getId());
  docSh.getRange(docRowIdx, 14).setValue(savedFile.getUrl());
  docSh.getRange(docRowIdx, 15).setValue(folderPath);
  SpreadsheetApp.flush();
}

// ========== 승인 문서 PDF 재생성/저장 ==========
function regenPdf(data) {
  const user = getSessionUser(data);
  const docSh = getSheet('문서');
  const docRowIdx = findRowIndex(docSh, 0, data.doc_id);
  if (docRowIdx < 0) return { success: false, error: '문서를 찾을 수 없습니다.' };
  const status = docSh.getRange(docRowIdx, 9).getValue();
  if (status !== 'approved') return { success: false, error: '승인된 문서만 PDF 저장이 가능합니다.' };
  try {
    autoSavePdf(data.doc_id, docRowIdx, docSh);
    const pdfFileId = docSh.getRange(docRowIdx, 13).getValue();
    const pdfUrl   = docSh.getRange(docRowIdx, 14).getValue();
    return { success: true, pdf_file_id: pdfFileId, pdf_url: pdfUrl };
  } catch(e) {
    return { success: false, error: 'PDF 저장 실패: ' + e.message };
  }
}

// ========== 클라이언트에서 PDF 전송 시 저장 ==========
function savePdfToDrive(data) {
  const user = getSessionUser(data);
  const docSh = getSheet('문서');
  const docRowIdx = findRowIndex(docSh, 0, data.doc_id);
  if (docRowIdx < 0) return { success: false, error: '문서를 찾을 수 없습니다.' };

  const root = getDriveRootFolder();
  // 캐시된 설정 사용
  const settingsResult = getSettings();
  const pathTemplate = settingsResult.settings.save_path_template || '승인문서/{year}/{month}';

  const savePath = data.save_path || docSh.getRange(docRowIdx, 15).getValue() || '';
  const now = new Date();
  let folderPath = pathTemplate
    .replace('{year}', now.getFullYear())
    .replace('{month}', ('0' + (now.getMonth()+1)).slice(-2));
  if (savePath) folderPath = savePath;

  const targetFolder = createFolderPath(root, folderPath);
  const docNumber = docSh.getRange(docRowIdx, 2).getValue();
  const docTitle = docSh.getRange(docRowIdx, 3).getValue();
  const fileName = docNumber + '_' + docTitle + '.pdf';
  const blob = Utilities.newBlob(Utilities.base64Decode(data.pdf_data), 'application/pdf', fileName);
  const file = targetFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  docSh.getRange(docRowIdx, 13).setValue(file.getId());
  docSh.getRange(docRowIdx, 14).setValue(file.getUrl());
  docSh.getRange(docRowIdx, 15).setValue(folderPath);
  SpreadsheetApp.flush();

  return { success: true, file_id: file.getId(), url: file.getUrl(), path: folderPath };
}

// ========== 알림 ==========
function createNotification(userId, docId, type, message, docTitle) {
  const sh = getSheet('알림');
  sh.appendRow([
    generateId(), docId, userId, type, message, 'N', new Date().toISOString(), docTitle || ''
  ]);
  SpreadsheetApp.flush();
  sendWebhook(message);
}

function getNotifications(data) {
  const user = getSessionUser(data);
  const sh = getSheet('알림');
  const notifs = sheetToObjects(sh)
    .filter(n => String(n.user_id) === String(user.id))
    .sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
  const unread = notifs.filter(n => n.read === 'N').length;
  return { success: true, notifications: notifs.slice(0, 50), unread_count: unread };
}

function markNotificationRead(data) {
  const user = getSessionUser(data);
  const sh = getSheet('알림');
  if (data.id === 'all') {
    const allData = sh.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][2]) === String(user.id) && allData[i][5] === 'N') {
        sh.getRange(i + 1, 6).setValue('Y');
      }
    }
  } else {
    const rowIdx = findRowIndex(sh, 0, data.id);
    if (rowIdx > 0) sh.getRange(rowIdx, 6).setValue('Y');
  }
  SpreadsheetApp.flush();
  return { success: true };
}

function sendWebhook(message) {
  // 캐시된 설정 사용 (매 알림마다 Sheets 읽기 방지)
  const settingsResult = getSettings();
  const url = String(settingsResult.settings.webhook_url || '').trim();
  if (!url || url.length < 10) return;

  const fullMessage = '[품질전자결재] ' + message;

  try {
    let payload;
    let contentType = 'application/json';
    if (url.includes('chat.googleapis.com')) {
      // Google Chat 웹훅
      payload = JSON.stringify({ text: fullMessage });
    } else if (url.includes('hooks.slack.com')) {
      // Slack 웹훅
      payload = JSON.stringify({ text: fullMessage });
    } else if (url.includes('discord.com')) {
      // Discord 웹훅
      payload = JSON.stringify({ content: fullMessage });
    } else if (url.includes('jandi.com') || url.includes('wh.jandi')) {
      // 잔디(JANDI) 웹훅
      payload = JSON.stringify({
        body: fullMessage,
        connectColor: '#FAC11B',
        connectInfo: [{ title: '품질전자결재', description: fullMessage }]
      });
    } else {
      // 범용 (text 필드)
      payload = JSON.stringify({ text: fullMessage, message: fullMessage, content: fullMessage });
    }

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: contentType,
      payload: payload,
      muteHttpExceptions: true,
      followRedirects: true
    });

    const code = response.getResponseCode();
    if (code < 200 || code >= 300) {
      Logger.log('Webhook failed: HTTP ' + code + ' / URL: ' + url.substring(0, 50) + '... / Response: ' + response.getContentText().substring(0, 200));
    }
  } catch(e) {
    Logger.log('Webhook error: ' + e.message + ' / URL: ' + url.substring(0, 50));
  }
}

// ========== 설정 ==========
function getSettings() {
  const cached = cacheGet('settings');
  if (cached) return cached;
  const sh = getSheet('설정');
  const data = sheetToObjects(sh);
  const settings = {};
  data.forEach(row => settings[row.key] = row.value);
  const result = { success: true, settings: settings };
  cachePut('settings', result, 300);
  return result;
}

function updateSettings(data) {
  const user = getSessionUser(data);
  if (user.role !== 'admin') return { success: false, error: '관리자 권한이 필요합니다.' };
  const sh = getSheet('설정');
  const allData = sh.getDataRange().getValues();
  const updates = data.settings || {};
  Object.keys(updates).forEach(key => {
    let found = false;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === key) {
        sh.getRange(i + 1, 2).setValue(updates[key]);
        found = true;
        break;
      }
    }
    if (!found) sh.appendRow([key, updates[key]]);
  });
  SpreadsheetApp.flush();
  cacheDel('settings'); // 설정 캐시 무효화
  return { success: true };
}

// ========== 대시보드 ==========
function getDashboard(data) {
  const user = getSessionUser(data);
  const uid = String(user.id);
  // 사용자별 대시보드 60초 캐시
  const cacheKey = 'dashboard_' + uid;
  const cached = cacheGet(cacheKey);
  if (cached) return cached;
  const docSh = getSheet('문서');
  const docs = sheetToObjects(docSh).filter(d => d.status !== 'deleted');
  docs.forEach(d => { d.id = String(d.id); d.creator_id = String(d.creator_id); });
  const approvalSh = getSheet('결재이력');
  const approvals = sheetToObjects(approvalSh);
  approvals.forEach(a => { a.approver_id = String(a.approver_id); a.doc_id = String(a.doc_id); });

  const myCreated = docs.filter(d => d.creator_id === uid);
  const pendingApproval = approvals.filter(a => a.approver_id === uid && a.status === 'pending');
  const pendingDocIds = pendingApproval.map(a => a.doc_id);
  const pendingDocs = docs.filter(d => pendingDocIds.includes(d.id));

  const result = {
    success: true,
    stats: {
      total_docs: docs.length,
      my_created: myCreated.length,
      my_pending: myCreated.filter(d => d.status === 'pending').length,
      my_approved: myCreated.filter(d => d.status === 'approved').length,
      my_rejected: myCreated.filter(d => d.status === 'rejected').length,
      pending_my_approval: pendingDocs.length
    },
    recent_docs: docs.slice(0, 10),
    pending_docs: pendingDocs
  };
  cachePut(cacheKey, result, 60); // 60초 캐시
  return result;
}

// ========== 드라이브 유틸리티 ==========
function getSubFolder(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function createFolderPath(root, path) {
  const parts = path.split('/').filter(p => p);
  let current = root;
  parts.forEach(part => { current = getSubFolder(current, part); });
  return current;
}

function detectMimeType(fileName) {
  const ext = (fileName || '').split('.').pop().toLowerCase();
  const mimeTypes = {
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'xls': 'application/vnd.ms-excel',
    'pdf': 'application/pdf'
  };
  return mimeTypes[ext] || 'application/octet-stream';
}

// 서명 이미지 Base64 가져오기
function getSignatureImage(data) {
  getSessionUser(data);
  const fileId = data.file_id;
  if (!fileId) return { success: false, error: '서명 파일이 없습니다.' };
  try {
    const file = DriveApp.getFileById(fileId);
    const bytes = file.getBlob().getBytes();
    const mimeType = file.getMimeType();
    return { success: true, data: Utilities.base64Encode(bytes), mime_type: mimeType };
  } catch(e) {
    return { success: false, error: '서명 파일을 불러올 수 없습니다.' };
  }
}


// ========== 로그인 이력 ==========
function recordLoginHistory(userId, username, name, ip) {
  try {
    const sh = getSheet('로그인이력');
    sh.appendRow([generateId(), userId, username, name, new Date().toISOString(), ip || '']);
  } catch(e) { Logger.log('로그인이력 오류: ' + e.message); }
}

function getLoginHistory(data) {
  const user = getSessionUser(data);
  if (user.role !== 'admin') return { success: false, error: '관리자 권한이 필요합니다.' };
  const sh = getSheet('로그인이력');
  const history = sheetToObjects(sh).map(r => ({
    id: String(r.id), user_id: String(r.user_id), username: String(r.username),
    name: String(r.name), login_at: String(r.login_at), ip: String(r.ip || '')
  })).sort((a, b) => new Date(b.login_at) - new Date(a.login_at));
  return { success: true, history: history.slice(0, 200) };
}

// ========== BOM 제품 데이터 ==========
function bomQaUpload(data) {
  const user = getSessionUser(data);
  if (user.role !== 'admin') return { success: false, error: '관리자 권한이 필요합니다.' };
  if (!data.products || !Array.isArray(data.products)) return { success: false, error: '제품 데이터가 없습니다.' };
  const sh = getSheet('BOM_QA');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 3).clearContent();
  if (data.products.length > 0) {
    const now = new Date().toISOString();
    sh.getRange(2, 1, data.products.length, 3).setValues(
      data.products.map(p => [String(p.product_code || ''), String(p.product_name || ''), now])
    );
  }
  cacheDel('bom_qa');
  SpreadsheetApp.flush();
  return { success: true, count: data.products.length };
}

function bomQaSearch(data) {
  getSessionUser(data);
  const cached = cacheGet('bom_qa');
  let rows = cached;
  if (!rows) {
    const sh = getSheet('BOM_QA');
    rows = sheetToObjects(sh).filter(p => p.product_code || p.product_name);
    cachePut('bom_qa', rows, 600); // 10분
  }
  const q = (data.query || '').toLowerCase().trim();
  const filtered = q
    ? rows.filter(p => String(p.product_name||'').toLowerCase().includes(q) || String(p.product_code||'').toLowerCase().includes(q))
    : rows;
  return { success: true, products: filtered.slice(0, 30) };
}

// ========== Keep-Warm (콜드 스타트 방지) ==========
// 시간 기반 트리거로 5~10분마다 실행 설정 권장
// Apps Script 편집기 → 트리거 → 새 트리거 → keepWarm → 시간 기반 → 5분마다
function keepWarm() {
  // 인스턴스를 활성 상태로 유지 (GAS 콜드 스타트 방지)
  Logger.log('keepWarm: ' + new Date().toISOString());
}

// ========== 초기 설정 실행 ==========
function setup() {
  const ss = getSpreadsheet();
  const root = getDriveRootFolder();
  Logger.log('스프레드시트 ID: ' + ss.getId());
  Logger.log('드라이브 폴더 ID: ' + root.getId());
  Logger.log('설정 완료! 웹앱을 배포하세요.');
}
