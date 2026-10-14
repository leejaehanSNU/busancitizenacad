/**
 * Google Apps Script backend for the GitHub Pages survey.
 * Bind this script to a Google Sheet, run setup(), then deploy as a Web App.
 */
const RESPONSE_SHEET = 'responses';
const HEADERS = ['timestamp', 'participant_code', 'stage', 'vote', 'comment'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('이 스크립트를 Google Sheets에서 확장 프로그램 > Apps Script로 열어 실행하세요.');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  let sh = ss.getSheetByName(RESPONSE_SHEET);
  if (!sh) sh = ss.insertSheet(RESPONSE_SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  let pin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!pin) {
    pin = Math.random().toString(36).slice(2, 8).toUpperCase();
    PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', pin);
  }
  Logger.log('설정 완료. 호스트 결과 URL: ?view=results');
  Logger.log('Google Sheet: ' + ss.getUrl());
  Logger.log('관리 확인 코드(필요 시): ' + pin);
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'health';
  let result;
  try {
    if (action === 'results') result = getResults_();
    else result = { ok: true, message: 'survey backend ready' };
  } catch (err) {
    result = { ok: false, error: String(err.message || err) };
  }
  return respond_(result, e && e.parameter && e.parameter.callback);
}

function doPost(e) {
  let result;
  try {
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw) throw new Error('요청 본문이 비어 있습니다.');
    const data = JSON.parse(raw);
    if (data.action !== 'submit') throw new Error('지원하지 않는 요청입니다.');
    saveResponse_(data);
    result = { ok: true };
  } catch (err) {
    result = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function saveResponse_(data) {
  const code = String(data.code || '').trim().toUpperCase();
  const stage = String(data.stage || '').trim();
  const vote = String(data.vote || '').trim();
  const comment = String(data.comment || '').trim().slice(0, 240);
  const allowed = ['가능하다', '일부 가능하다', '어렵다'];
  if (!/^[A-Z0-9]{4,8}$/.test(code)) throw new Error('참여 코드는 영문 대문자와 숫자 4~8자리로 입력해 주세요.');
  if (!['pre', 'post'].includes(stage)) throw new Error('설문 단계가 올바르지 않습니다.');
  if (!allowed.includes(vote)) throw new Error('응답 선택지를 확인해 주세요.');
  const sh = getSheet_();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const rows = sh.getDataRange().getValues();
    let match = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).toUpperCase() === code && String(rows[i][2]) === stage) match = i + 1;
    }
    const row = [new Date(), code, stage, vote, comment];
    if (match > 0) sh.getRange(match, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
}

function getResults_() {
  const rows = getSheet_().getDataRange().getValues();
  const byCode = {};
  for (let i = 1; i < rows.length; i++) {
    const [timestamp, codeRaw, stage, vote, comment] = rows[i];
    const code = String(codeRaw || '').trim().toUpperCase();
    if (!code || !['pre', 'post'].includes(String(stage))) continue;
    if (!byCode[code]) byCode[code] = {};
    byCode[code][String(stage)] = { vote: String(vote || ''), comment: String(comment || ''), timestamp: timestamp };
  }
  const all = Object.values(byCode);
  const paired = all.filter(x => x.pre && x.post);
  const summarize = (items, key) => {
    const votes = { '가능하다': 0, '일부 가능하다': 0, '어렵다': 0 };
    items.forEach(x => { if (votes[x[key].vote] !== undefined) votes[x[key].vote]++; });
    return { total: items.filter(x => x[key]).length, votes, comments: items.filter(x => x[key] && x[key].comment).map(x => ({ comment: x[key].comment })) };
  };
  return { ok: true, pre: summarize(all, 'pre'), post: summarize(all, 'post'), paired: paired.length, pairedPre: summarize(paired, 'pre'), pairedPost: summarize(paired, 'post') };
}

function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('setup()을 먼저 실행해 주세요.');
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(RESPONSE_SHEET);
  if (!sh) { sh = ss.insertSheet(RESPONSE_SHEET); sh.appendRow(HEADERS); }
  return sh;
}

function respond_(data, callback) {
  const json = JSON.stringify(data);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]{0,80}$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
