/**
 * Archon Macro 클라우드 설정 저장소 - Google Apps Script 백엔드
 *
 * 여러 사용자가 ID + 비밀번호로 자기 macro_config.json을 클라우드에 저장/불러옴.
 *
 * 셋업:
 * 1) 구글 드라이브에서 새 시트 생성 (예: "ArchonCloudConfig")
 * 2) 첫 행에 헤더 입력 (5개 열):
 *    A1: id  B1: password  C1: config_json  D1: created_at  E1: updated_at
 * 3) 시트 [확장 프로그램] → [Apps Script] → 이 코드 전체 붙여넣기
 * 4) 우측 상단 [배포] → [새 배포] → 유형: 웹 앱
 *    - 다음 사용자로 실행: 나
 *    - 액세스 권한: 모든 사용자
 *    배포 → 웹 앱 URL 복사
 * 5) 그 URL을 앱(Archon)의 클라우드 설정 다이얼로그 코드에 박음
 *
 * 코드 수정 후엔 [배포] → [배포 관리] → 새 버전으로 다시 배포해야 반영됨.
 *
 * ※ 비밀번호는 시트에 평문 저장됩니다. 시트 공유 권한 관리에 주의.
 */

const SHEET_NAME = 'Sheet1';
const MAX_CONFIG_SIZE = 200000;  // 200 KB 제한

function doGet(e)  { return handleRequest(e.parameter); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { return _json({ ok:false, error:'invalid json' }); }
  return handleRequest(p);
}

function handleRequest(p) {
  const a = (p.action || '').toLowerCase();
  try {
    if (a === 'register') return _json(register(p));
    if (a === 'save')     return _json(saveConfig(p));
    if (a === 'load')     return _json(loadConfig(p));
    if (a === 'change_password') return _json(changePassword(p));
    return _json({ ok:false, error:'unknown action' });
  } catch (err) {
    return _json({ ok:false, error:String(err) });
  }
}

function _sheet() { return SpreadsheetApp.getActive().getSheetByName(SHEET_NAME); }
function _json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function _findUser(id) {
  const sh = _sheet();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, 5).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === id) {
      return {
        _row: i + 2,
        id: data[i][0],
        password: String(data[i][1]),
        config_json: data[i][2],
        created_at: data[i][3],
        updated_at: data[i][4],
      };
    }
  }
  return null;
}

function _validId(id) {
  return typeof id === 'string' && /^[A-Za-z0-9_]{3,30}$/.test(id);
}

function _validPw(pw) {
  return typeof pw === 'string' && pw.length >= 4 && pw.length <= 128;
}

function register(p) {
  const id = (p.id || '').trim();
  const pw = p.password || '';
  if (!_validId(id)) return { ok:false, error:'아이디는 영문/숫자/_ 3~30자' };
  if (!_validPw(pw)) return { ok:false, error:'비밀번호는 4~128자' };
  if (_findUser(id)) return { ok:false, error:'이미 사용 중인 아이디' };
  const sh = _sheet();
  const now = new Date();
  sh.appendRow([id, pw, '', now, now]);
  return { ok:true };
}

function saveConfig(p) {
  const id = (p.id || '').trim();
  const pw = p.password || '';
  const cfg = p.config_json || '';
  if (!_validId(id) || !_validPw(pw)) return { ok:false, error:'아이디/비밀번호 형식 오류' };
  if (typeof cfg !== 'string' || cfg.length === 0) return { ok:false, error:'설정값 비어있음' };
  if (cfg.length > MAX_CONFIG_SIZE) return { ok:false, error:'설정값 크기 초과 (200KB 제한)' };
  try { JSON.parse(cfg); } catch (e) { return { ok:false, error:'유효한 JSON 아님' }; }
  const u = _findUser(id);
  if (!u) return { ok:false, error:'존재하지 않는 아이디' };
  if (u.password !== pw) return { ok:false, error:'비밀번호 불일치' };
  const sh = _sheet();
  sh.getRange(u._row, 3).setValue(cfg);
  sh.getRange(u._row, 5).setValue(new Date());
  return { ok:true, size: cfg.length };
}

function loadConfig(p) {
  const id = (p.id || '').trim();
  const pw = p.password || '';
  if (!_validId(id) || !_validPw(pw)) return { ok:false, error:'아이디/비밀번호 형식 오류' };
  const u = _findUser(id);
  if (!u) return { ok:false, error:'존재하지 않는 아이디' };
  if (u.password !== pw) return { ok:false, error:'비밀번호 불일치' };
  if (!u.config_json) return { ok:false, error:'저장된 설정 없음' };
  return { ok:true, config_json: String(u.config_json), updated_at: u.updated_at };
}

function changePassword(p) {
  const id = (p.id || '').trim();
  const oldPw = p.old_password || '';
  const newPw = p.new_password || '';
  if (!_validId(id) || !_validPw(oldPw) || !_validPw(newPw)) return { ok:false, error:'형식 오류' };
  const u = _findUser(id);
  if (!u) return { ok:false, error:'존재하지 않는 아이디' };
  if (u.password !== oldPw) return { ok:false, error:'기존 비밀번호 불일치' };
  _sheet().getRange(u._row, 2).setValue(newPw);
  return { ok:true };
}
