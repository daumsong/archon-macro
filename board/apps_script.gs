/**
 * Archon Macro 1:1 문의 게시판 - Google Apps Script 백엔드
 *
 * 셋업:
 * 1) 구글 드라이브에서 새 시트 생성 (예: "ArchonInquiries")
 * 2) 첫 행에 헤더 입력:
 *    A1: id  B1: created_at  C1: user_email  D1: title
 *    E1: content  F1: status  G1: admin_reply  H1: replied_at
 * 3) 시트의 [확장 프로그램] → [Apps Script] 클릭
 * 4) 이 파일 전체 내용을 Code.gs에 붙여넣기
 * 5) 아래 ADMIN_TOKEN을 본인만 아는 임의 문자열로 교체
 * 6) 상단 우측 [배포] → [새 배포] → 유형: 웹 앱
 *    - 다음 사용자로 실행: 나
 *    - 액세스 권한: 모든 사용자
 *    배포 → 웹 앱 URL 복사 (https://script.google.com/macros/s/.../exec)
 * 7) 그 URL을 board/index.html의 SCRIPT_URL 상수에 붙여넣기
 *
 * 코드 수정 후엔 [배포] → [배포 관리] → 새 버전으로 다시 배포해야 반영됨.
 */

const ADMIN_TOKEN = 'CHANGE_ME_TO_RANDOM_STRING';  // ← 본인만 아는 비밀 문자열로 교체
const SHEET_NAME = 'Sheet1';  // 시트 탭 이름 (기본 Sheet1)

function doGet(e) {
  return handleRequest(e.parameter);
}

function doPost(e) {
  let params = {};
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    return _json({ ok: false, error: 'invalid json' });
  }
  return handleRequest(params);
}

function handleRequest(p) {
  const action = (p.action || '').toLowerCase();
  try {
    if (action === 'list')   return _json(listInquiries(p));
    if (action === 'get')    return _json(getInquiry(p));
    if (action === 'create') return _json(createInquiry(p));
    if (action === 'reply')  return _json(replyInquiry(p));
    return _json({ ok: false, error: 'unknown action' });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function _sheet() {
  return SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _rows() {
  const sh = _sheet();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(2, 1, last - 1, 8).getValues();
  return data.map((r, idx) => ({
    _row: idx + 2,
    id: r[0],
    created_at: r[1],
    user_email: r[2],
    title: r[3],
    content: r[4],
    status: r[5] || '대기',
    admin_reply: r[6] || '',
    replied_at: r[7] || '',
  }));
}

function listInquiries(p) {
  const isAdmin = p.token === ADMIN_TOKEN;
  let rows = _rows();
  if (!isAdmin) {
    const email = (p.email || '').trim().toLowerCase();
    if (!email) return { ok: true, items: [] };
    rows = rows.filter(r => String(r.user_email).trim().toLowerCase() === email);
  }
  rows.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
  // 목록에서는 content는 제외 (상세에서 가져감)
  const items = rows.map(r => ({
    id: r.id,
    created_at: r.created_at,
    user_email: r.user_email,
    title: r.title,
    status: r.status,
    has_reply: !!r.admin_reply,
  }));
  return { ok: true, items: items, is_admin: isAdmin };
}

function getInquiry(p) {
  const id = String(p.id || '');
  const isAdmin = p.token === ADMIN_TOKEN;
  const email = (p.email || '').trim().toLowerCase();
  const row = _rows().find(r => String(r.id) === id);
  if (!row) return { ok: false, error: 'not found' };
  if (!isAdmin && String(row.user_email).trim().toLowerCase() !== email) {
    return { ok: false, error: 'forbidden' };
  }
  return { ok: true, item: row };
}

function createInquiry(p) {
  const email = (p.email || '').trim();
  const title = (p.title || '').trim();
  const content = (p.content || '').trim();
  if (!email || !title || !content) {
    return { ok: false, error: 'missing fields' };
  }
  if (title.length > 200 || content.length > 10000) {
    return { ok: false, error: 'too long' };
  }
  const sh = _sheet();
  const id = Utilities.getUuid().slice(0, 8);
  const now = new Date();
  sh.appendRow([id, now, email, title, content, '대기', '', '']);
  return { ok: true, id: id };
}

function replyInquiry(p) {
  if (p.token !== ADMIN_TOKEN) return { ok: false, error: 'forbidden' };
  const id = String(p.id || '');
  const reply = (p.reply || '').trim();
  if (!reply) return { ok: false, error: 'empty reply' };
  const row = _rows().find(r => String(r.id) === id);
  if (!row) return { ok: false, error: 'not found' };
  const sh = _sheet();
  sh.getRange(row._row, 6).setValue('답변완료');
  sh.getRange(row._row, 7).setValue(reply);
  sh.getRange(row._row, 8).setValue(new Date());
  return { ok: true };
}
