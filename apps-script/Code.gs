/**
 * 9층 자리 배치표 — 구글 시트 저장용 Apps Script
 * 시트 메뉴 [확장 프로그램 > Apps Script]에 이 코드를 붙여넣고
 * [배포 > 새 배포 > 웹 앱] (실행: 나, 액세스: 모든 사용자)로 배포하세요.
 */
const PASSWORD = 'kjh0928';      // 현재 버전 저장 비밀번호 (여기서 바꾸면 됩니다)
const SHEET = 'versions';
const CHUNK = 40000;             // 셀 하나 최대 50,000자 → 나눠서 저장

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['id', 'kind', 'name', 'updated', 'data']);
    sh.setFrozenRows(1);
    sh.getRange('A:D').setNumberFormat('@');
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function readAll_() {
  const v = sheet_().getDataRange().getValues();
  const out = { ok: true, current: null, versions: [] };
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!r[0]) continue;
    const raw = r.slice(4).join('');
    let data = null;
    try { data = raw ? JSON.parse(raw) : null; } catch (e) {}
    const o = { id: String(r[0]), name: String(r[2]), updated: String(r[3]), data: data };
    if (r[1] === 'current') out.current = o; else out.versions.push(o);
  }
  return out;
}

function rowOf_(sh, id) {
  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return -1;
}

function write_(id, kind, name, data) {
  const sh = sheet_();
  const s = JSON.stringify(data);
  const chunks = [];
  for (let i = 0; i < s.length; i += CHUNK) chunks.push(s.slice(i, i + CHUNK));
  const now = new Date().toISOString();
  const row = [String(id), kind, String(name), now].concat(chunks);
  let r = rowOf_(sh, id);
  if (r < 0) r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, sh.getMaxColumns()).clearContent();
  if (sh.getMaxColumns() < row.length) sh.insertColumnsAfter(sh.getMaxColumns(), row.length - sh.getMaxColumns());
  const rg = sh.getRange(r, 1, 1, row.length);
  rg.setNumberFormat('@');
  rg.setValues([row]);
  return now;
}

function doGet() {
  return json_(readAll_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const b = JSON.parse(e.postData.contents);
    switch (b.action) {
      case 'checkPassword':
        return json_(b.password === PASSWORD ? { ok: true } : { ok: false, error: 'password' });
      case 'saveCurrent':
        if (b.password !== PASSWORD) return json_({ ok: false, error: 'password' });
        return json_({ ok: true, updated: write_('current', 'current', b.name, b.data) });
      case 'saveVersion':
        if (!b.id || b.id === 'current') return json_({ ok: false, error: 'bad id' });
        return json_({ ok: true, updated: write_(b.id, 'draft', b.name, b.data) });
      case 'deleteVersion': {
        const sh = sheet_(), r = rowOf_(sh, b.id);
        if (r > 1 && b.id !== 'current') sh.deleteRow(r);
        return json_({ ok: true });
      }
      default:
        return json_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
