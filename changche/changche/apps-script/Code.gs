/**
 * 호남고 2학년 창체활동 도우미 — 저장 서버 (Google Apps Script)
 *
 * 하는 일
 *  - 학생이 올린 파일을 구글 드라이브에 저장
 *      [ROOT 폴더] / 2322_홍길동 / 자율활동 / 03_독서활동 / 2322_홍길동_독서활동_원래파일명.hwp
 *  - 학생의 선택 내용을 구글 시트 '제출현황' 탭에 학번·영역별 한 줄로 기록 (다시 제출하면 그 줄을 갱신)
 *  - 같은 항목을 다시 올리면 그 항목의 기존 파일은 휴지통으로 이동 (덮어쓰기)
 *  - 다시 제출할 때 선택에서 빠진 항목의 파일도 휴지통으로 이동
 *  - (선택) '명단' 탭에 학번·이름이 있으면 일치하는 학생만 제출 가능
 *
 * 설정: 아래 두 값만 바꾸세요. (사용 설명서.md 참고)
 */
const ROOT_FOLDER_ID = '여기에_드라이브_폴더_ID';
const SHEET_ID = '여기에_구글시트_ID';

const STATUS_SHEET = '제출현황';
const ROSTER_SHEET = '명단';
const HEADERS = ['학번', '이름', '영역', '방법', '선택 항목', '후속탐구', '탐구동기 메모', '제출 파일', '최종 제출', '상태(JSON)'];
const ALLOWED_EXT = ['hwp', 'hwpx', 'doc', 'docx'];
const MAX_BYTES = 20 * 1024 * 1024;

/* ---------- 진입점 ---------- */

function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'status') return json_(getStatus_(p.sid, p.name));
    return json_({ ok: true, msg: '창체활동 도우미 서버가 작동 중입니다.' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'upload') return json_(upload_(d));
    if (d.action === 'finalize') return json_(finalize_(d));
    return json_({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* ---------- 기능 ---------- */

function getStatus_(sid, name) {
  checkStudent_(sid, name);
  const sh = statusSheet_();
  const rows = sh.getDataRange().getValues();
  const sections = {};
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === sid) {
      let st = {};
      try { st = JSON.parse(rows[i][9] || '{}'); } catch (x) {}
      st.updated = rows[i][8] ? Utilities.formatDate(new Date(rows[i][8]), 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : '';
      st.files = listFiles_(sid, name, rows[i][2]);
      sections[rows[i][2]] = st;
    }
  }
  return { ok: true, sections: sections };
}

function upload_(d) {
  checkStudent_(d.sid, d.name);
  const ext = String(d.fileName).split('.').pop().toLowerCase();
  if (ALLOWED_EXT.indexOf(ext) < 0) throw new Error('한글(hwp, hwpx) 또는 워드(doc, docx) 파일만 올릴 수 있습니다.');
  const bytes = Utilities.base64Decode(d.data);
  if (bytes.length > MAX_BYTES) throw new Error('파일이 너무 큽니다. (20MB 이하)');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const folder = itemFolder_(d.sid, d.name, d.section, d.itemOrder, d.itemLabel, true);
    if (d.replace) trashAll_(folder); // 이 항목의 첫 파일이면 기존 파일 정리(덮어쓰기)
    const safe = String(d.fileName).replace(/[\\/:*?"<>|]/g, '_');
    const fname = d.sid + '_' + d.name + '_' + d.itemLabel + '_' + safe;
    folder.createFile(Utilities.newBlob(bytes, d.mimeType || 'application/octet-stream', fname));
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function finalize_(d) {
  checkStudent_(d.sid, d.name);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // 선택에서 빠진 항목 폴더 정리
    const sec = sectionFolder_(d.sid, d.name, d.section, false);
    if (sec) {
      const keep = (d.items || []).map(function (it) { return folderName_(it.order, it.label); });
      const it = sec.getFolders();
      while (it.hasNext()) {
        const f = it.next();
        if (keep.indexOf(f.getName()) < 0) { trashAll_(f); f.setTrashed(true); }
      }
    }
    const files = listFiles_(d.sid, d.name, d.section);
    const fileText = Object.keys(files).map(function (k) { return k + ': ' + files[k].join(', '); }).join('\n');
    const notes = (d.items || []).filter(function (x) { return x.note; })
      .map(function (x) { return x.label + ': ' + x.note; }).join('\n');

    const state = { method: d.method, methodLabel: d.methodLabel, items: d.items, followup: !!d.followup };
    const row = [
      d.sid, d.name, d.section, d.methodLabel,
      (d.items || []).map(function (x) { return x.label; }).join(', '),
      d.followup ? '포함' : '-', notes, fileText, new Date(), JSON.stringify(state)
    ];
    const sh = statusSheet_();
    const rows = sh.getDataRange().getValues();
    let at = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === d.sid && rows[i][2] === d.section) { at = i + 1; break; }
    }
    if (at > 0) sh.getRange(at, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/* ---------- 도우미 ---------- */

function checkStudent_(sid, name) {
  if (!/^2[1-9]\d{2}$/.test(String(sid || ''))) throw new Error('학번 형식이 올바르지 않습니다.');
  if (!name || String(name).trim().length < 2) throw new Error('이름을 확인해 주세요.');
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const r = ss.getSheetByName(ROSTER_SHEET);
  if (!r || r.getLastRow() < 2) return; // 명단 탭이 없으면 검사 생략
  const vals = r.getRange(2, 1, r.getLastRow() - 1, 2).getValues();
  const ok = vals.some(function (v) { return String(v[0]).trim() === sid && String(v[1]).trim() === String(name).trim(); });
  if (!ok) throw new Error('학번과 이름이 명단과 맞지 않습니다. 담임 선생님께 확인하세요.');
}

function statusSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(STATUS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STATUS_SHEET);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.hideColumns(10);
  }
  return sh;
}

function folderName_(order, label) {
  return ('0' + order).slice(-2) + '_' + label;
}

function studentFolder_(sid, name, create) {
  const root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  const it = root.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf(sid + '_') === 0) return f;
  }
  return create ? root.createFolder(sid + '_' + name) : null;
}

function child_(parent, name, create) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return create ? parent.createFolder(name) : null;
}

function sectionFolder_(sid, name, section, create) {
  const s = studentFolder_(sid, name, create);
  return s ? child_(s, section, create) : null;
}

function itemFolder_(sid, name, section, order, label, create) {
  const sec = sectionFolder_(sid, name, section, create);
  return sec ? child_(sec, folderName_(order, label), create) : null;
}

function trashAll_(folder) {
  const it = folder.getFiles();
  while (it.hasNext()) it.next().setTrashed(true);
}

function listFiles_(sid, name, section) {
  const out = {};
  const sec = sectionFolder_(sid, name, section, false);
  if (!sec) return out;
  const it = sec.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    const label = f.getName().replace(/^\d+_/, '');
    const names = [];
    const fi = f.getFiles();
    while (fi.hasNext()) names.push(fi.next().getName().split('_' + label + '_').pop());
    if (names.length) out[label] = names;
  }
  return out;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 처음 한 번 실행: 권한 승인 + 시트 준비 확인용 */
function setup() {
  DriveApp.getFolderById(ROOT_FOLDER_ID).getName();
  statusSheet_();
  Logger.log('준비 완료');
}
