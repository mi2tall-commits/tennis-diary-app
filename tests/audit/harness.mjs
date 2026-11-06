// 시험용 가짜 Apps Script 환경. Code.gs를 node vm에 올려 실제 함수를 그대로 실행한다.
// 외부 API(텔레그램·Gemini·Drive)는 호출하지 않고 기록만 한다.
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
export const CODE_GS = fs.readFileSync(path.join(ROOT, 'Code.gs'), 'utf8');
export const INDEX_HTML = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');

const HEADER = ["일자", "구분", "운동제목", "시간(분)", "강도(RPE)", "경기결과", "스코어", "함께한사람", "코트종류", "장소", "집중기술", "컨디션/수기메모", "사진ID목록", "댓글", "AI코칭분석"];

export function makeSheet(name, rows) {
  const data = rows.map(r => r.slice());
  const sheet = {
    name, data,
    getName: () => name,
    setName(n) { name = n; return sheet; },
    getLastRow: () => data.length,
    getLastColumn: () => data.reduce((m, r) => Math.max(m, r.length), 0),
    appendRow(r) { data.push(r.slice()); return sheet; },
    deleteRow(i) {
      if (typeof i !== 'number' || i < 1 || i > data.length) throw new Error('bad row ' + i);
      data.splice(i - 1, 1); return sheet;
    },
    setFrozenRows() { return sheet; },
    getDataRange() { return { getValues: () => data.map(r => r.slice()) }; },
    getRange(r, c, nr = 1, nc = 1) {
      const rng = {
        getValue: () => (data[r - 1] || [])[c - 1] ?? '',
        setValue(v) { while (data.length < r) data.push([]); data[r - 1][c - 1] = v; return rng; },
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (data[r - 1 + i] || [])[c - 1 + j] ?? '')),
        setValues(vals) { vals.forEach((row, i) => row.forEach((v, j) => { while (data.length < r + i) data.push([]); data[r - 1 + i][c - 1 + j] = v; })); return rng; },
        setBackground() { return rng; }, setFontColor() { return rng; }, setFontWeight() { return rng; },
      };
      return rng;
    },
  };
  return sheet;
}

export function logRow(date, title, extra = {}) {
  const r = [date, extra.category || '레슨', title, 60, 5, '', '', '', '하드코트', '', '', extra.notes || '', extra.photoIds || '', extra.comments || '[]', ''];
  return r;
}

function formatDate(d, tz, fmt) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false })
    .formatToParts(d).map(p => [p.type, p.value]));
  const hh = parts.hour === '24' ? '00' : parts.hour;
  return fmt.replace('yyyy', parts.year).replace('MM', parts.month).replace('dd', parts.day).replace('HH', hh).replace('mm', parts.minute);
}

/**
 * opts.logs: 테니스일지 데이터 행(헤더 제외), opts.strings: 스트링교체 행, opts.props: Script Properties,
 * opts.now: 고정 현재 시각(ISO), opts.driveFails: Drive 파일 생성 실패 여부
 */
export function loadGas(opts = {}) {
  const props = Object.assign({}, opts.props || {});
  const tennis = makeSheet('테니스일지', [HEADER, ...(opts.logs || [])]);
  const strings = makeSheet('스트링교체', [["교체일자", "라켓모델", "스트링명", "스트링종류", "텐션(lbs)", "교체비용", "스트링샵", "메모", "등록일시"], ...(opts.strings || [])]);
  const fetches = [];
  const lockCalls = [];
  let uuid = 0;

  const RealDate = Date;
  const fixedNow = opts.now ? new RealDate(opts.now).getTime() : null;
  class FakeDate extends RealDate {
    constructor(...a) { if (a.length === 0 && fixedNow !== null) super(fixedNow); else super(...a); }
    static now() { return fixedNow !== null ? fixedNow : RealDate.now(); }
  }

  const ctx = {
    console, JSON, Math, String, Number, Array, Object, Error, isNaN, parseInt, parseFloat, encodeURIComponent,
    Date: FakeDate,
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = v; } }) },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({
      getSheetByName: n => (n === '테니스일지' ? tennis : n === '스트링교체' ? strings : null),
      getSheets: () => [tennis], insertSheet: () => strings,
    }) },
    Session: { getScriptTimeZone: () => 'Asia/Seoul' },
    Utilities: {
      formatDate,
      getUuid: () => 'uuid-' + (++uuid),
      base64Decode: s => Array.from(Buffer.from(s, 'base64')),
      newBlob: (bytes, type, name) => ({ bytes, type, name }),
    },
    DriveApp: {
      Access: { ANYONE_WITH_LINK: 'A' }, Permission: { VIEW: 'V' },
      getFoldersByName: () => ({ hasNext: () => false }),
      createFolder: () => ({
        setSharing() {},
        createFile(blob) {
          if (opts.driveFails) throw new Error('Drive 용량 초과(시험)');
          return { setSharing() {}, getId: () => 'FILEID_' + blob.name.length + '_abcdefghijklmnopqrstuvwxyz' };
        },
      }),
    },
    UrlFetchApp: { fetch: (url, o) => { fetches.push({ url, o }); return { getResponseCode: () => 200, getContentText: () => '{}' }; } },
    LockService: { getScriptLock: () => ({ waitLock: ms => lockCalls.push('wait:' + ms), tryLock: ms => { lockCalls.push('try:' + ms); return true; }, releaseLock: () => lockCalls.push('release') }) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ text: s, setMimeType() { return this; } }) },
    HtmlService: {},
    Logger: { log() {} },
  };
  vm.createContext(ctx);
  vm.runInContext(CODE_GS, ctx, { filename: 'Code.gs' });
  ctx.__tennis = tennis; ctx.__strings = strings; ctx.__fetches = fetches; ctx.__lockCalls = lockCalls; ctx.__props = props;
  return ctx;
}

export function post(ctx, action, payload) {
  const out = ctx.doPost({ postData: { contents: JSON.stringify({ action, payload }) } });
  return JSON.parse(out.text);
}

/** index.html 안의 이름 있는 함수 하나를 소스 그대로 잘라낸다(중괄호 짝 맞추기). */
export function extractFunction(src, name) {
  const start = src.indexOf('function ' + name + '(');
  if (start < 0) throw new Error('function not found: ' + name);
  let i = src.indexOf('{', start), depth = 0, q = null;
  for (; i < src.length; i++) {
    const ch = src[i];
    if (q) { if (ch === '\\') { i++; continue; } if (ch === q) q = null; continue; }
    if (ch === '"' || ch === "'" || ch === '`') { q = ch; continue; }
    if (ch === '{') depth++;
    else if (ch === '}') { depth--; if (depth === 0) return src.slice(start, i + 1); }
  }
  throw new Error('unbalanced: ' + name);
}
