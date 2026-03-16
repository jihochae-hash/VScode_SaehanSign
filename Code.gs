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
  sh.getRange(1, 1, 1, 12).setValues([['id','username','password','name','department','position','role','signature_file_id','created_at','status','email','webhook_url']]);

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
function getDeployedUrl() {
  return ScriptApp.getService().getUrl();
}

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.deployedUrl = getDeployedUrl();
  return template.evaluate()
    .setTitle('SaehanSign - 새한 문서 전자결재 시스템')
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
      case 'doc_batch_delete': result = batchDeleteDocuments(data); break;
      case 'doc_file': result = getDocumentFile(data); break;
      case 'doc_search': result = searchDocuments(data); break;

      case 'approval_submit': result = submitForApproval(data); break;
      case 'approval_approve': result = approveDocument(data); break;
      case 'approval_reject': result = rejectDocument(data); break;
      case 'approval_history': result = getApprovalHistory(data); break;

      case 'save_pdf': result = savePdfToDrive(data); break;
      case 'regen_pdf': result = regenPdf(data); break;
      case 'generate_pdf': result = generatePdfWithStamp(data); break;

      case 'notifications': result = getNotifications(data); break;
      case 'notification_read': result = markNotificationRead(data); break;
      case 'notification_delete': result = deleteNotification(data); break;

      case 'settings_get': result = getSettings(); break;
      case 'settings_update': result = updateSettings(data); break;

      case 'dashboard': result = getDashboard(data); break;

      case 'get_signature_image': result = getSignatureImage(data); break;
      case 'bom_qa_upload': result = bomQaUpload(data); break;
      case 'bom_qa_search': result = bomQaSearch(data); break;
      case 'login_history': result = getLoginHistory(data); break;
      case 'login_history_delete': result = deleteLoginHistory(data); break;

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
  if (!data.new_password || String(data.new_password).length < 6) {
    return { success: false, error: '새 비밀번호는 최소 6자 이상이어야 합니다.' };
  }
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
  if (!data.password || String(data.password).length < 6) {
    return { success: false, error: '비밀번호는 최소 6자 이상이어야 합니다.' };
  }
  if (!data.username || String(data.username).trim().length < 2) {
    return { success: false, error: '아이디는 최소 2자 이상이어야 합니다.' };
  }
  const sh = getSheet('사용자');
  const users = sheetToObjects(sh);
  if (users.find(u => u.username === data.username && u.status === 'active')) {
    return { success: false, error: '이미 존재하는 아이디입니다.' };
  }
  const id = generateId();
  sh.appendRow([
    id, data.username, hashPassword(data.password),
    data.name, data.department, data.position, data.role || 'user',
    '', new Date().toISOString(), 'active', data.email || '', data.webhook_url || ''
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
  if (data.webhook_url !== undefined) sh.getRange(rowIdx, 12).setValue(data.webhook_url);
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

  // 30초 캐시 (filter/category 조합별)
  const cacheKey = 'docs_' + (data.filter || 'all') + '_' + (data.category || '') + '_' + String(user.id);
  const cached = cacheGet(cacheKey);
  if (cached) return cached;

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
  const result = { success: true, documents: docs };
  cachePut(cacheKey, result, 30); // 30초 캐시
  return result;
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

  // 사용자 시트 한 번만 로드 (루프 밖)
  const userShForSig = getSheet('사용자');
  steps.forEach((step, idx) => {
    const actualStepNum = idx + 1 + stepOffset;
    // 결재자의 최신 서명 파일 ID를 사용자 시트에서 가져옴
    let sigFileId = step.signature_file_id || '';
    if (!sigFileId && step.approver_id) {
      const userRow = findRowIndex(userShForSig, 0, step.approver_id);
      if (userRow > 0) sigFileId = userShForSig.getRange(userRow, 8).getValue() || '';
    }
    approvalSh.appendRow([
      generateId(), id, actualStepNum, step.name, step.approver_id,
      step.approver_name, step.approver_dept || '', step.approver_position || '',
      idx === 0 ? 'pending' : 'waiting', '', '', '', sigFileId
    ]);
  });

  SpreadsheetApp.flush();
  cacheDel('dashboard_' + user.id); // 대시보드 캐시 무효화
  // 문서 목록 캐시 무효화 (카테고리 없는 기본 키)
  ['all', 'my_created', 'pending_my_approval', 'approved', 'rejected'].forEach(f => {
    cacheDel('docs_' + f + '__' + String(user.id));
  });
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

  // 작성자 칸이 결재이력에 없으면 앞에 추가 (서명 base64 포함)
  const hasAuthorStep = approvals.length > 0 && approvals[0].step_name === '작성';
  if (!hasAuthorStep) {
    const creatorInfo = getCreatorInfo(doc.creator_id, doc.creator_name);
    creatorInfo.signed_at = doc.created_at;
    creatorInfo.status = (doc.status !== 'draft') ? 'approved' : 'pending';
    // 작성자 서명 base64
    if (creatorInfo.signature_file_id) {
      try {
        const sigFile = DriveApp.getFileById(creatorInfo.signature_file_id);
        creatorInfo.signature_base64 = Utilities.base64Encode(sigFile.getBlob().getBytes());
        creatorInfo.signature_mime = sigFile.getMimeType();
      } catch(e) { creatorInfo.signature_base64 = ''; }
    }
    approvals.unshift(creatorInfo);
  }

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
  // 관리자는 승인 완료 문서도 삭제 가능, 일반 사용자는 불가
  if (status === 'approved' && user.role !== 'admin') {
    return { success: false, error: '승인 완료된 문서는 삭제할 수 없습니다.' };
  }
  sh.getRange(rowIdx, 9).setValue('deleted');
  SpreadsheetApp.flush();
  return { success: true };
}

function batchDeleteDocuments(data) {
  const user = getSessionUser(data);
  if (user.role !== 'admin') return { success: false, error: '관리자만 일괄 삭제할 수 있습니다.' };
  const ids = Array.isArray(data.ids) ? data.ids : String(data.ids).split(',');
  if (!ids.length) return { success: false, error: '삭제할 문서가 없습니다.' };
  const sh = getSheet('문서');
  let deleted = 0;
  ids.forEach(function(id) {
    const rowIdx = findRowIndex(sh, 0, String(id).trim());
    if (rowIdx >= 0) {
      sh.getRange(rowIdx, 9).setValue('deleted');
      deleted++;
    }
  });
  SpreadsheetApp.flush();
  return { success: true, deleted_count: deleted };
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

  // 문서 상태 확인
  const docSh = getSheet('문서');
  const docRowIdx = findRowIndex(docSh, 0, data.doc_id);
  if (docRowIdx < 0) return { success: false, error: '문서를 찾을 수 없습니다.' };
  const docStatus = String(docSh.getRange(docRowIdx, 9).getValue());
  if (docStatus === 'approved') return { success: false, error: '이미 최종 승인된 문서입니다.' };
  if (docStatus !== 'pending') return { success: false, error: '결재 진행중인 문서가 아닙니다.' };

  // 현재 사용자의 결재 단계 찾기 (pending 또는 waiting 중 순서가 된 것)
  let myStep = docApprovals.find(a => a.approver_id === String(user.id) && a.status === 'pending');

  // pending이 아니면, 이전 단계가 모두 approved이고 내 차례인 waiting 단계 찾기
  if (!myStep) {
    const myWaiting = docApprovals.find(a => a.approver_id === String(user.id) && a.status === 'waiting');
    if (myWaiting) {
      const allBefore = docApprovals.filter(a => a.step_order < myWaiting.step_order);
      const allBeforeApproved = allBefore.every(a => a.status === 'approved');
      if (allBeforeApproved) {
        // waiting → pending 전환 후 진행
        const waitRowIdx = findRowIndex(approvalSh, 0, myWaiting.id);
        if (waitRowIdx > 0) {
          approvalSh.getRange(waitRowIdx, 9).setValue('pending');
          SpreadsheetApp.flush();
        }
        myStep = myWaiting;
        myStep.status = 'pending';
      }
    }
  }

  if (!myStep) return { success: false, error: '결재 권한이 없거나 이미 처리되었습니다.' };

  // 결재자의 최신 서명 파일 가져오기
  const userSh = getSheet('사용자');
  const userRow = findRowIndex(userSh, 0, user.id);
  let currentSigFileId = '';
  if (userRow > 0) {
    currentSigFileId = userSh.getRange(userRow, 8).getValue() || '';
  }

  // 결재이력 업데이트 (배치 쓰기로 속도 개선)
  const rowIdx = findRowIndex(approvalSh, 0, myStep.id);
  if (rowIdx < 0) return { success: false, error: '결재 이력을 찾을 수 없습니다.' };

  const now = new Date().toISOString();
  approvalSh.getRange(rowIdx, 9, 1, 5).setValues([[
    'approved', now, data.comment || '', 'Y', currentSigFileId || myStep.signature_file_id || ''
  ]]);

  // 다음 단계 활성화 또는 최종 승인
  const nextStep = docApprovals.find(a => a.step_order === myStep.step_order + 1);
  const docRow = docSh.getRange(docRowIdx, 1, 1, 20).getValues()[0];
  const docTitle = String(docRow[2] || '');
  const creatorId = String(docRow[6] || '');

  if (nextStep) {
    // 다음 결재자 pending으로 변경
    const nextRowIdx = findRowIndex(approvalSh, 0, nextStep.id);
    if (nextRowIdx > 0) approvalSh.getRange(nextRowIdx, 9).setValue('pending');
    docSh.getRange(docRowIdx, 10).setValue(myStep.step_order + 1);
    SpreadsheetApp.flush(); // 단 1회만 flush

    createNotification(nextStep.approver_id, data.doc_id, 'approval_request',
      user.name + '님이 "' + docTitle + '" 문서를 승인했습니다. 결재를 진행해주세요.', docTitle);
    createNotification(creatorId, data.doc_id, 'progress',
      user.name + '님이 "' + docTitle + '" 문서를 승인했습니다.', docTitle);
  } else {
    // 최종 승인
    docSh.getRange(docRowIdx, 9).setValue('approved');
    docSh.getRange(docRowIdx, 12).setValue(now);
    SpreadsheetApp.flush(); // 단 1회만 flush

    // 원본 파일을 PDF로 변환하여 저장
    try {
      autoSavePdf(data.doc_id, docRowIdx, docSh);
    } catch(pdfErr) {
      Logger.log('PDF 자동저장 실패: ' + pdfErr.message);
    }

    createNotification(creatorId, data.doc_id, 'approved',
      '"' + docTitle + '" 문서가 최종 승인되었습니다.', docTitle);
  }

  // 승인 의견이 있으면 문서 시트에도 저장 (리스트 표시용)
  if (data.comment && String(data.comment).trim()) {
    const docHeaders = docSh.getRange(1, 1, 1, docSh.getLastColumn()).getValues()[0];
    let acCol = docHeaders.indexOf('approval_comment') + 1;
    let abCol = docHeaders.indexOf('approval_by') + 1;
    if (acCol === 0) {
      acCol = docSh.getLastColumn() + 1;
      abCol = acCol + 1;
      docSh.getRange(1, acCol).setValue('approval_comment');
      docSh.getRange(1, abCol).setValue('approval_by');
    }
    docSh.getRange(docRowIdx, acCol).setValue(String(data.comment).trim());
    docSh.getRange(docRowIdx, abCol).setValue(user.name);
    SpreadsheetApp.flush();
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

  const docApprovals = allApprovals
    .filter(a => a.doc_id === String(data.doc_id))
    .sort((a, b) => Number(a.step_order) - Number(b.step_order));

  let myStep = docApprovals.find(a => a.approver_id === String(user.id) && a.status === 'pending');
  // waiting 상태이지만 이전 단계 모두 승인된 경우 처리
  if (!myStep) {
    const myWaiting = docApprovals.find(a => a.approver_id === String(user.id) && a.status === 'waiting');
    if (myWaiting) {
      const allBefore = docApprovals.filter(a => a.step_order < myWaiting.step_order);
      if (allBefore.every(a => a.status === 'approved')) {
        myStep = myWaiting;
      }
    }
  }
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
  // 설정 + 문서 종류별 경로
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
    .replace('{month_kr}', (now.getMonth()+1) + '월')
    .replace('{doc_type}', docCategory);
  if (customPath) folderPath = customPath;

  const targetFolder = createFolderPath(root, folderPath);
  const docNumber = docSh.getRange(docRowIdx, 2).getValue();
  const docTitle = docSh.getRange(docRowIdx, 3).getValue();
  const fileName = docNumber + '_' + docTitle + '.pdf';

  const originalFile = DriveApp.getFileById(fileId);
  const origMime = originalFile.getMimeType();
  let savedFile;

  // 결재이력 가져오기
  const approvalSh = getSheet('결재이력');
  const approvals = sheetToObjects(approvalSh)
    .filter(a => String(a.doc_id) === String(docId))
    .sort((a, b) => Number(a.step_order) - Number(b.step_order));

  // 작성자 정보 가져오기
  const creatorId = String(docSh.getRange(docRowIdx, 7).getValue());
  const creatorName = docSh.getRange(docRowIdx, 8).getValue() || '';
  const creatorInfo = getCreatorInfo(creatorId, creatorName);

  // 작성자 칸이 결재이력에 없으면 추가
  const hasAuthorStep = approvals.length > 0 && approvals[0].step_name === '작성';
  const stampApprovals = hasAuthorStep ? approvals : [creatorInfo].concat(approvals);

  // Excel/PDF 모두 동일한 이미지화+결재란 합성 방식
  let sourceFileId = fileId;
  let tempConvertId = null;
  try {
    if (origMime !== 'application/pdf') {
      // Excel → Sheets → 스탬프 없는 깨끗한 PDF 생성 → Drive 임시 저장
      tempConvertId = convertFileToSheets(fileId, 'temp_clean_' + docId, root.getId());
      Utilities.sleep(2000);
      const ss = SpreadsheetApp.openById(tempConvertId);
      SpreadsheetApp.flush();
      const cleanPdf = exportSheetAsCleanPdf(ss, '_clean.pdf');
      const tempPdfFile = root.createFile(cleanPdf);
      sourceFileId = tempPdfFile.getId();
      // Sheets 임시 파일 삭제 (PDF는 썸네일 생성 후 삭제)
      try { DriveApp.getFileById(tempConvertId).setTrashed(true); } catch(e) {}
      tempConvertId = null;
      Utilities.sleep(3000); // 썸네일 생성 대기
    }
    // PDF 이미지화 + 결재란 합성 (공통 로직)
    const pdfResult = buildPdfBlobWithDocStamp(sourceFileId, stampApprovals, fileName, root);
    savedFile = targetFolder.createFile(pdfResult.blob);
  } finally {
    if (tempConvertId) { try { DriveApp.getFileById(tempConvertId).setTrashed(true); } catch(e) {} }
    // Excel에서 생성한 임시 PDF 삭제
    if (sourceFileId !== fileId) { try { DriveApp.getFileById(sourceFileId).setTrashed(true); } catch(e) {} }
  }

  savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  docSh.getRange(docRowIdx, 13).setValue(savedFile.getId());
  docSh.getRange(docRowIdx, 14).setValue(savedFile.getUrl());
  docSh.getRange(docRowIdx, 15).setValue(folderPath);
  SpreadsheetApp.flush();
}

// ========== 작성자 정보 조회 ==========
function getCreatorInfo(creatorId, creatorName) {
  const userSh = getSheet('사용자');
  const userRow = findRowIndex(userSh, 0, creatorId);
  let dept = '', position = '', sigFileId = '';
  if (userRow > 0) {
    dept = userSh.getRange(userRow, 5).getValue() || '';
    position = userSh.getRange(userRow, 6).getValue() || '';
    sigFileId = userSh.getRange(userRow, 8).getValue() || '';
  }
  return {
    step_name: '작성',
    approver_id: creatorId,
    approver_name: creatorName,
    approver_dept: dept,
    approver_position: position,
    status: 'approved',
    signed_at: new Date().toISOString(),
    signature_file_id: sigFileId,
    signature_applied: 'Y'
  };
}

// ========== 결재란 표시 항목 설정 가져오기 ==========
function getStampDisplay() {
  try {
    var s = getSettings().settings;
    if (s.stamp_display) return JSON.parse(s.stamp_display);
  } catch(e) {}
  return { dept: true, name: true, position: true, date: true };
}

// ========== 결재란 삽입 (상단 5행 삽입 → 결재란 우측 상단 배치) ==========
// 반환값: { r1, c1, r2, c2 } (0-indexed PDF 인쇄 범위)
function embedApprovalStamp(sheet, approvals) {
  if (!approvals || approvals.length === 0) return null;

  const stepCount = approvals.length;
  const STAMP_ROWS = 5;
  const sd = getStampDisplay();

  // ── 삽입 전 원본 컨텐츠 범위 파악 ──
  const contentLastCol = Math.max(sheet.getLastColumn(), 1);
  const contentLastRow = Math.max(sheet.getLastRow(), 1);

  // ── 상단에 5행 삽입 (원본 내용이 아래로 밀림) ──
  sheet.insertRowsBefore(1, STAMP_ROWS);

  // ── 결재란 위치: 기존 열 우측에 새 열 추가 (기존 열 너비 절대 변경 안함) ──
  const stampGap = 1;                          // 문서와 결재란 사이 빈 열 수
  const startCol = contentLastCol + stampGap + 1;
  const endCol   = startCol + stepCount - 1;

  // 새 열 확보
  while (sheet.getMaxColumns() < endCol) {
    sheet.insertColumnAfter(sheet.getMaxColumns());
  }

  // 결재란 열 너비 (새 열에만 적용)
  for (let i = 0; i < stepCount; i++) {
    sheet.setColumnWidth(startCol + i, 87);
  }

  // 결재란 행 높이 (20% 증가)
  sheet.setRowHeight(1, 22);   // 단계명
  sheet.setRowHeight(2, 31);   // 서명 상단
  sheet.setRowHeight(3, 31);   // 서명 하단
  sheet.setRowHeight(4, 31);   // 이름/부서/직책
  sheet.setRowHeight(5, 17);   // 결재일

  // ── 웹 화면과 동일한 폰트/서식 ──
  var stampFont = 'Noto Sans KR';

  // ── Row 1: 단계명 헤더 ──
  for (let i = 0; i < stepCount; i++) {
    const col = startCol + i;
    const a = approvals[i];
    sheet.getRange(1, col).setValue(a.step_name || '결재')
      .setFontFamily(stampFont).setFontSize(9).setFontWeight('bold')
      .setFontColor('#004d4d')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground('#dff0f0');
  }

  // ── Row 2-3: 서명 이미지 (병합, 가운데 배치) ──
  for (let i = 0; i < stepCount; i++) {
    const col = startCol + i;
    const a = approvals[i];
    sheet.getRange(2, col, 2, 1).merge()
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground('#ffffff');
    if (a.status === 'approved') {
      let sigInserted = false;
      if (a.signature_file_id) {
        try {
          const sigFile = DriveApp.getFileById(a.signature_file_id);
          const sigBlob = sigFile.getBlob().setContentType(sigFile.getMimeType() || 'image/png');
          var sigImg = sheet.insertImage(sigBlob, col, 2);
          sigImg.setWidth(66).setHeight(50);
          sigImg.setAnchorCellXOffset(10);
          sigImg.setAnchorCellYOffset(6);
          sigInserted = true;
        } catch(e) { Logger.log('서명 이미지 삽입 실패(col=' + col + '): ' + e.message); }
      }
      if (!sigInserted) {
        sheet.getRange(2, col).setValue(a.approver_name || '')
          .setFontFamily(stampFont).setFontSize(10).setFontWeight('bold')
          .setFontColor('#006666');
      }
    }
  }

  // ── Row 4: 부서 + 이름 직책 (설정에 따라) ──
  for (let i = 0; i < stepCount; i++) {
    const col = startCol + i;
    const a = approvals[i];
    const parts = [];
    if (sd.dept !== false && a.approver_dept) parts.push(a.approver_dept);
    const nameParts = [];
    if (sd.name !== false && a.approver_name) nameParts.push(a.approver_name);
    if (sd.position !== false && a.approver_position) nameParts.push(a.approver_position);
    if (nameParts.length) parts.push(nameParts.join(' '));
    sheet.getRange(4, col)
      .setValue(parts.join('\n'))
      .setFontFamily(stampFont).setFontSize(8)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true);
  }

  // ── Row 5: 결재일 (설정에 따라) ──
  for (let i = 0; i < stepCount; i++) {
    const col = startCol + i;
    const a = approvals[i];
    if (sd.date === false) { sheet.getRange(5, col).setValue(''); continue; }
    const dateStr = (a.signed_at || a.approved_at) ? formatShortDate(a.signed_at || a.approved_at) : '';
    sheet.getRange(5, col).setValue(dateStr)
      .setFontFamily(stampFont).setFontSize(8).setFontColor('#666666')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  }

  // ── 테두리 (웹과 동일한 teal) ──
  sheet.getRange(1, startCol, STAMP_ROWS, stepCount)
    .setBorder(true, true, true, true, true, true, '#8ababa', SpreadsheetApp.BorderStyle.SOLID);

  // ── 메모(Threaded comment 포함) 전체 삭제 → 2페이지 원인 차단 ──
  try { sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearNote(); } catch(e) {}

  // ── 시트 트리밍: 인쇄 범위 밖의 행/열 삭제 → 빈 영역이 없어 scale=4 정확히 동작 ──
  const totalRows = STAMP_ROWS + contentLastRow;
  const maxRows = sheet.getMaxRows();
  if (maxRows > totalRows) {
    try { sheet.deleteRows(totalRows + 1, maxRows - totalRows); } catch(e) {}
  }
  const maxCols = sheet.getMaxColumns();
  if (maxCols > endCol) {
    try { sheet.deleteColumns(endCol + 1, maxCols - endCol); } catch(e) {}
  }

  // 시트 자체를 트리밍했으므로 명시적 인쇄 범위 불필요 → null 반환
  return null;
}

// ========== A4 한 페이지 PDF 내보내기 ==========
// printRange: { r1, c1, r2, c2 } (0-indexed) — 지정 시 해당 범위만 출력
function exportSheetAsPdf(ss, fileName, printRange) {
  const sheetId = ss.getSheets()[0].getSheetId();
  // scale=4 (Fit to Page): 시트 전체를 A4 한 장에 맞춤
  // fitw/fith는 scale=4와 충돌하므로 제거
  // 시트를 미리 트리밍했으므로 r1/c1/r2/c2 불필요
  let url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' +
    'format=pdf' +
    '&size=A4' +
    '&portrait=true' +
    '&scale=4' +           // Fit to Page (가로+세로 모두 A4에 맞춤)
    '&sheetnames=false' +
    '&gridlines=false' +
    '&printtitle=false' +
    '&pagenumbers=false' +
    '&notes=false' +
    '&top_margin=0.5' +
    '&bottom_margin=0.5' +
    '&left_margin=0.5' +
    '&right_margin=0.5' +
    '&gid=' + sheetId;

  // 시트 트리밍 후에도 명시적 범위가 전달된 경우 사용 (하위 호환)
  if (printRange) {
    url += '&r1=' + printRange.r1 + '&c1=' + printRange.c1 +
           '&r2=' + printRange.r2 + '&c2=' + printRange.c2;
  }

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDF export failed: ' + response.getResponseCode());
  }
  return response.getBlob().setName(fileName);
}

// ========== Excel→이미지용 깨끗한 PDF 내보내기 (최소 여백, 가운데 정렬) ==========
function exportSheetAsCleanPdf(ss, fileName) {
  const sheetId = ss.getSheets()[0].getSheetId();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' +
    'format=pdf&size=A4&portrait=true&scale=4' +
    '&sheetnames=false&gridlines=false&printtitle=false&pagenumbers=false&notes=false' +
    '&top_margin=0.20&bottom_margin=0.20&left_margin=0.20&right_margin=0.20' +
    '&horizontal_alignment=CENTER' +
    '&gid=' + sheetId;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) throw new Error('Clean PDF export failed');
  return response.getBlob().setName(fileName);
}

// ========== 파일을 Google Sheets 형식으로 변환 (multipart upload) ==========
// Drive Advanced Service 없이 Excel/CSV → Sheets 변환
function convertFileToSheets(fileId, tempName, parentId) {
  var token = ScriptApp.getOAuthToken();
  var originalFile = DriveApp.getFileById(fileId);
  var originalBlob = originalFile.getBlob();
  var originalBytes = originalBlob.getBytes();
  var originalMime = originalBlob.getContentType() || 'application/octet-stream';

  var boundary = 'boundary' + Utilities.getUuid().replace(/-/g, '');
  var metadata = JSON.stringify({
    name: tempName,
    mimeType: 'application/vnd.google-apps.spreadsheet',
    parents: [parentId]
  });

  // multipart body (binary safe)
  var headerBytes = Utilities.newBlob(
    '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
    metadata + '\r\n--' + boundary + '\r\nContent-Type: ' + originalMime + '\r\n\r\n'
  ).getBytes();
  var footerBytes = Utilities.newBlob('\r\n--' + boundary + '--').getBytes();

  var bodyBytes = [].concat(headerBytes).concat(originalBytes).concat(footerBytes);

  var resp = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
    {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token },
      contentType: 'multipart/related; boundary=' + boundary,
      payload: bodyBytes,
      muteHttpExceptions: true
    }
  );

  if (resp.getResponseCode() !== 200) {
    throw new Error('파일→Sheets 변환 실패(' + resp.getResponseCode() + '): ' + resp.getContentText().substring(0, 200));
  }
  return JSON.parse(resp.getContentText()).id;
}

// ========== PDF 원본 결재란 삽입 (Google Sheets 방식) ==========
// Slides API 불필요 — 이미 동작 중인 SpreadsheetApp만 사용
// 1) Drive thumbnail로 PDF 1페이지 이미지를 가져옴
// 2) 빈 Google Sheets에 결재란(우측 상단) + PDF 이미지(아래) 배치
// 3) A4 세로 PDF로 내보내기
function buildPdfBlobWithDocStamp(fileId, stampApprovals, fileName, root) {
  var token = ScriptApp.getOAuthToken();
  var tempSsId = null;
  try {
    // ── 1. PDF 페이지 이미지(썸네일) 가져오기 ──
    var pdfPageBlob = null;
    // Sheets insertImage 제한: 2MB / 100만 픽셀 → A4 비율 최대 ~800px
    var thumbSizes = ['w800', 'w700', 'w600'];
    for (var ti = 0; ti < thumbSizes.length && !pdfPageBlob; ti++) {
      try {
        var tr = UrlFetchApp.fetch(
          'https://drive.google.com/thumbnail?id=' + fileId + '&sz=' + thumbSizes[ti],
          { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
        );
        if (tr.getResponseCode() === 200 && tr.getBlob().getBytes().length > 500) {
          pdfPageBlob = tr.getBlob();
        }
      } catch(e) {}
    }
    // thumbnailLink fallback
    if (!pdfPageBlob) {
      try {
        var mr = UrlFetchApp.fetch(
          'https://www.googleapis.com/drive/v3/files/' + fileId + '?fields=thumbnailLink',
          { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
        );
        var thumbLink = JSON.parse(mr.getContentText()).thumbnailLink;
        if (thumbLink) {
          var hiRes = thumbLink.replace(/=s\d+/, '=s800');
          var tr2 = UrlFetchApp.fetch(hiRes, { muteHttpExceptions: true });
          if (tr2.getResponseCode() === 200) pdfPageBlob = tr2.getBlob();
        }
      } catch(e) {}
    }
    if (!pdfPageBlob) throw new Error('PDF 썸네일을 가져올 수 없습니다.');

    // ── 2. 빈 Google Sheets 생성 ──
    var ss = SpreadsheetApp.create('_qa_tmp_pdf_' + Utilities.getUuid().substring(0, 6));
    tempSsId = ss.getId();
    var sheet = ss.getSheets()[0];

    var stepCount = stampApprovals.length;
    var stampColW = 87;  // 결재 열 너비(px)
    var stampTotalW = stampColW * stepCount;

    // ── 3. 열/행 구성: A4 인쇄 영역 맞춤 + 결재란 여백 ──
    // 1 Sheet px ≈ 0.75pt
    // 결재란 여백: 위 2cm(≈76px), 오른쪽 1.5cm(≈57px)
    var topMarginPx = 38;   // 1cm (PDF 원본 상단 여백과 합쳐 실질 2cm)
    var rightMarginPx = 57; // 1.5cm

    var totalWidth = 775;
    var mainColW = totalWidth - stampTotalW - rightMarginPx;
    if (mainColW < 300) mainColW = 300;

    // 열: [메인콘텐츠] [결재열1] [결재열2] ... [오른쪽여백열]
    sheet.setColumnWidth(1, mainColW);
    for (var i = 0; i < stepCount; i++) {
      sheet.setColumnWidth(2 + i, stampColW);
    }
    var rightMarginCol = 2 + stepCount;
    // 여백열 확보
    while (sheet.getMaxColumns() < rightMarginCol) {
      sheet.insertColumnAfter(sheet.getMaxColumns());
    }
    sheet.setColumnWidth(rightMarginCol, rightMarginPx);

    // 행: [위여백] [결재란 5행] [PDF이미지]
    // 엑셀 결재란과 동일한 크기
    sheet.setRowHeight(1, topMarginPx); // 위 여백 2cm
    sheet.setRowHeight(2, 22);   // 단계명
    sheet.setRowHeight(3, 31);   // 서명 상단
    sheet.setRowHeight(4, 31);   // 서명 하단
    sheet.setRowHeight(5, 31);   // 이름/직책
    sheet.setRowHeight(6, 17);   // 날짜

    // ── 4. 결재란 (행 2~6, 열 2~) — 웹 화면과 동일한 폰트/서식 ──
    var stampFont = 'Noto Sans KR'; // 웹과 유사한 한국어 폰트

    for (var i = 0; i < stepCount; i++) {
      var col = 2 + i;
      var a = stampApprovals[i];

      // 행2: 단계명 (teal 배경)
      sheet.getRange(2, col).setValue(a.step_name || '결재')
        .setFontFamily(stampFont).setFontSize(9).setFontWeight('bold')
        .setFontColor('#004d4d')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBackground('#dff0f0');

      // 행3-4: 서명 (병합)
      sheet.getRange(3, col, 2, 1).merge()
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBackground('#ffffff');

      if (a.status === 'approved') {
        var sigOk = false;
        if (a.signature_file_id) {
          try {
            var sf = DriveApp.getFileById(a.signature_file_id);
            var sigImg2 = sheet.insertImage(sf.getBlob().setContentType(sf.getMimeType() || 'image/png'), col, 3);
            sigImg2.setWidth(66).setHeight(50);
            sigImg2.setAnchorCellXOffset(10); // (87-66)/2 = 가운데
            sigImg2.setAnchorCellYOffset(6);  // (62-50)/2 = 가운데
            sigOk = true;
          } catch(e) { Logger.log('PDF 서명 삽입 실패: ' + e.message); }
        }
        if (!sigOk) {
          sheet.getRange(3, col).setValue(a.approver_name || '')
            .setFontFamily(stampFont).setFontSize(10).setFontWeight('bold')
            .setFontColor('#006666');
        }
      }

      // 행5: 부서 + 이름 직책 (설정에 따라)
      var sd = getStampDisplay();
      var infoParts = [];
      if (sd.dept !== false && a.approver_dept) infoParts.push(a.approver_dept);
      var nameParts2 = [];
      if (sd.name !== false && a.approver_name) nameParts2.push(a.approver_name);
      if (sd.position !== false && a.approver_position) nameParts2.push(a.approver_position);
      if (nameParts2.length) infoParts.push(nameParts2.join(' '));
      sheet.getRange(5, col).setValue(infoParts.join('\n'))
        .setFontFamily(stampFont).setFontSize(8)
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setWrap(true);

      // 행6: 날짜 (설정에 따라)
      if (sd.date === false) {
        sheet.getRange(6, col).setValue('');
      } else {
        var dateStr = (a.signed_at || a.approved_at) ? formatShortDate(a.signed_at || a.approved_at) : '';
        sheet.getRange(6, col).setValue(dateStr)
          .setFontFamily(stampFont).setFontSize(8).setFontColor('#666666')
          .setHorizontalAlignment('center').setVerticalAlignment('middle');
      }
    }

    // 결재란 테두리 (웹과 동일한 teal 색상)
    sheet.getRange(2, 2, 5, stepCount)
      .setBorder(true, true, true, true, true, true, '#8ababa', SpreadsheetApp.BorderStyle.SOLID);

    // ── 5. PDF 이미지 삽입 (행7 앵커, 원본 크기 유지) ──
    var stampH = topMarginPx + 22 + 31 + 31 + 31 + 17; // =170px (여백38+결재란132)
    var imgW = totalWidth;
    var fullPageH = Math.round(totalWidth * 1.414); // A4 전체 높이 (≈1096px)
    var visibleH = fullPageH - stampH; // 1페이지 내 보이는 영역 (≈916px)

    // 이미지는 원본 A4 크기 그대로 (축소 없음)
    // 행7 높이 = 보이는 영역만 → scale=4에서 정확히 1페이지
    // 이미지 하단은 행7 경계에서 자동 클립됨
    sheet.setRowHeight(7, visibleH);

    // 행8 이후 모두 삭제
    if (sheet.getMaxRows() > 7) {
      try { sheet.deleteRows(8, sheet.getMaxRows() - 7); } catch(e) {}
    }

    // 열 트리밍 (오른쪽 여백열까지 유지)
    var endCol = rightMarginCol;
    if (sheet.getMaxColumns() > endCol) {
      try { sheet.deleteColumns(endCol + 1, sheet.getMaxColumns() - endCol); } catch(e) {}
    }

    // 이미지: 원본 크기로 삽입 (결재란 바로 아래, 행7)
    var img = sheet.insertImage(pdfPageBlob, 1, 7);
    img.setWidth(imgW);
    img.setHeight(fullPageH);

    try { sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearNote(); } catch(e) {}
    SpreadsheetApp.flush();
    Utilities.sleep(1500);

    // ── 6. A4 세로 PDF 내보내기 (최소 여백, 1페이지 맞춤) ──
    var sheetId = sheet.getSheetId();
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' +
      'format=pdf&size=A4&portrait=true&scale=4' +
      '&sheetnames=false&gridlines=false&printtitle=false&pagenumbers=false&notes=false' +
      '&top_margin=0.1&bottom_margin=0.1&left_margin=0.1&right_margin=0.1' +
      '&gid=' + sheetId;

    var pdfResp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (pdfResp.getResponseCode() !== 200) {
      throw new Error('PDF export 실패: ' + pdfResp.getResponseCode());
    }

    // 하단 클립 비율 계산
    var clipPct = Math.round(stampH / fullPageH * 100);

    return {
      blob: pdfResp.getBlob().setName(fileName),
      clip_percent: clipPct
    };

  } finally {
    if (tempSsId) { try { DriveApp.getFileById(tempSsId).setTrashed(true); } catch(e) {} }
  }
}


function formatShortDate(s) {
  if (!s) return '';
  try {
    const d = new Date(s);
    if (isNaN(d)) return '';
    return d.getFullYear() + '.' + (d.getMonth()+1) + '.' + d.getDate();
  } catch(e) { return ''; }
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

// ========== PDF 다운로드 (결재 스탬프 포함) ==========
function generatePdfWithStamp(data) {
  const user = getSessionUser(data);
  const docSh = getSheet('문서');
  const docRowIdx = findRowIndex(docSh, 0, data.doc_id);
  if (docRowIdx < 0) return { success: false, error: '문서를 찾을 수 없습니다.' };

  const fileId = docSh.getRange(docRowIdx, 6).getValue();
  if (!fileId) return { success: false, error: '원본 파일이 없습니다.' };

  const originalFile = DriveApp.getFileById(fileId);
  const origMime = originalFile.getMimeType();

  // 결재이력 가져오기
  const approvalSh = getSheet('결재이력');
  const approvals = sheetToObjects(approvalSh)
    .filter(a => String(a.doc_id) === String(data.doc_id))
    .sort((a, b) => Number(a.step_order) - Number(b.step_order));

  // 작성자 정보
  const creatorId = String(docSh.getRange(docRowIdx, 7).getValue());
  const creatorName = docSh.getRange(docRowIdx, 8).getValue() || '';
  const creatorInfo = getCreatorInfo(creatorId, creatorName);
  const hasAuthorStep = approvals.length > 0 && approvals[0].step_name === '작성';
  const stampApprovals = hasAuthorStep ? approvals : [creatorInfo].concat(approvals);

  // Excel/PDF 모두 동일한 이미지화+결재란 합성 방식
  const docNumber = docSh.getRange(docRowIdx, 2).getValue();
  const docTitle = docSh.getRange(docRowIdx, 3).getValue();
  const pdfFileName = docNumber + '_' + docTitle + '.pdf';
  const root = getDriveRootFolder();
  let sourceFileId = fileId;
  let tempConvertId = null;
  let tempPdfFileId = null;
  try {
    if (origMime !== 'application/pdf') {
      // Excel → Sheets → 깨끗한 PDF 생성 → Drive 임시 저장
      tempConvertId = convertFileToSheets(fileId, 'temp_dl_' + data.doc_id, root.getId());
      Utilities.sleep(2000);
      const ss = SpreadsheetApp.openById(tempConvertId);
      SpreadsheetApp.flush();
      const cleanPdf = exportSheetAsCleanPdf(ss, '_clean.pdf');
      const tempPdfFile = root.createFile(cleanPdf);
      tempPdfFileId = tempPdfFile.getId();
      sourceFileId = tempPdfFileId;
      try { DriveApp.getFileById(tempConvertId).setTrashed(true); } catch(e) {}
      tempConvertId = null;
      Utilities.sleep(3000); // 썸네일 생성 대기
    }
    const pdfResult = buildPdfBlobWithDocStamp(sourceFileId, stampApprovals, pdfFileName, root);
    return {
      success: true,
      pdf_base64: Utilities.base64Encode(pdfResult.blob.getBytes()),
      file_name: pdfFileName,
      clip_percent: pdfResult.clip_percent
    };
  } catch(e) {
    Logger.log('PDF 생성 실패: ' + e.message);
    return { success: false, error: 'PDF 생성 실패: ' + e.message };
  } finally {
    if (tempConvertId) { try { DriveApp.getFileById(tempConvertId).setTrashed(true); } catch(e) {} }
    if (tempPdfFileId) { try { DriveApp.getFileById(tempPdfFileId).setTrashed(true); } catch(e) {} }
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
    .replace('{month}', ('0' + (now.getMonth()+1)).slice(-2))
    .replace('{month_kr}', (now.getMonth()+1) + '월');
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
  sendWebhook(message, userId);
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

function deleteNotification(data) {
  const user = getSessionUser(data);
  const sh = getSheet('알림');
  if (data.id === 'all') {
    // 내 알림 전체 삭제 (아래→위로 삭제해야 행번호 안 밀림)
    const allData = sh.getDataRange().getValues();
    for (var i = allData.length - 1; i >= 1; i--) {
      if (String(allData[i][2]) === String(user.id)) {
        sh.deleteRow(i + 1);
      }
    }
  } else {
    const rowIdx = findRowIndex(sh, 0, data.id);
    if (rowIdx > 0) {
      // 본인 알림인지 확인
      if (String(sh.getRange(rowIdx, 3).getValue()) === String(user.id)) {
        sh.deleteRow(rowIdx);
      }
    }
  }
  SpreadsheetApp.flush();
  return { success: true };
}

function sendWebhook(message, targetUserId) {
  const fullMessage = '[SaehanSign] ' + message;

  // 전송할 URL 목록 수집 (중복 제거)
  const urls = new Set();

  // 1) 글로벌 웹훅 URL
  const settingsResult = getSettings();
  const globalUrl = String(settingsResult.settings.webhook_url || '').trim();
  if (globalUrl && globalUrl.length >= 10) urls.add(globalUrl);

  // 2) 대상 사용자 개인 웹훅 URL
  if (targetUserId) {
    try {
      const userSh = getSheet('사용자');
      const userData = userSh.getDataRange().getValues();
      for (let i = 1; i < userData.length; i++) {
        if (String(userData[i][0]) === String(targetUserId)) {
          const userWebhook = String(userData[i][11] || '').trim();
          if (userWebhook && userWebhook.length >= 10) urls.add(userWebhook);
          break;
        }
      }
    } catch(e) {}
  }

  if (urls.size === 0) return;

  urls.forEach(function(url) {
    try {
      let payload;
      let contentType = 'application/json';
      if (url.includes('chat.googleapis.com')) {
        payload = JSON.stringify({ text: fullMessage });
      } else if (url.includes('hooks.slack.com')) {
        payload = JSON.stringify({ text: fullMessage });
      } else if (url.includes('discord.com')) {
        payload = JSON.stringify({ content: fullMessage });
      } else if (url.includes('jandi.com') || url.includes('wh.jandi')) {
        payload = JSON.stringify({
          body: fullMessage,
          connectColor: '#FAC11B',
          connectInfo: [{ title: 'SaehanSign', description: fullMessage }]
        });
      } else {
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
        Logger.log('Webhook failed: HTTP ' + code + ' / URL: ' + url.substring(0, 50) + '...');
      }
    } catch(e) {
      Logger.log('Webhook error: ' + e.message + ' / URL: ' + url.substring(0, 50));
    }
  });
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

function deleteLoginHistory(data) {
  const user = getSessionUser(data);
  if (user.role !== 'admin') return { success: false, error: '관리자만 삭제할 수 있습니다.' };
  const sh = getSheet('로그인이력');
  if (data.id === 'all') {
    // 헤더 행만 남기고 전체 삭제
    const lastRow = sh.getLastRow();
    if (lastRow > 1) sh.deleteRows(2, lastRow - 1);
  } else {
    const rowIdx = findRowIndex(sh, 0, data.id);
    if (rowIdx > 0) sh.deleteRow(rowIdx);
  }
  SpreadsheetApp.flush();
  return { success: true };
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
  // 빈 쿼리는 전체 반환 (프론트엔드 캐시용), 검색 시 50개 제한
  return { success: true, products: q ? filtered.slice(0, 50) : filtered };
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
