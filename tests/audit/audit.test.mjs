// 2026-11-06 점검에서 확인한 버그 재현 시험 + 핵심 동작 회귀 시험.
// 실행: node --test tests/audit/
process.env.TZ = 'Asia/Seoul';
import test from 'node:test';
import assert from 'node:assert/strict';
import { loadGas, post, logRow, CODE_GS, INDEX_HTML, extractFunction } from './harness.mjs';

// ---------- #1 코드에 박힌 Gemini 키 ----------
test('#1 Script Properties가 비어 있으면 Gemini 기본 키도 비어 있어야 한다(코드에 키 없음)', () => {
  const ctx = loadGas({ props: {} });
  assert.equal(ctx.DEFAULT_GEMINI_API_KEY === '', true, '코드 안에 대체 키가 들어 있음');
  assert.ok(!/\[\s*"AQ"\s*,/.test(CODE_GS), '쪼갠 키 문자열이 남아 있음');
});

test('#1 회귀: Script Properties의 GEMINI_API_KEY는 그대로 쓴다', () => {
  const ctx = loadGas({ props: { GEMINI_API_KEY: 'dummy-test-key' } });
  assert.equal(ctx.DEFAULT_GEMINI_API_KEY, 'dummy-test-key');
});

// ---------- #3 수정하면 사진 연결이 지워짐 ----------
test('#3 수정 창은 서버가 보낸 photo_ids를 기존 사진으로 읽어야 한다', () => {
  const src = extractFunction(INDEX_HTML, 'openEditTennisModal');
  const line = src.split('\n').find(l => /editExistingPhotoIds\s*=/.test(l));
  assert.ok(line, '대입 줄을 찾지 못함');
  const rhs = line.split('=').slice(1).join('=').replace(/;\s*$/, '');
  const log = { photo_ids: ['AAAAAAAAAAAAAAAAAAAAAAAA', 'vid:BBBBBBBBBBBBBBBBBBBBBBBB'] };
  const val = new Function('log', 'return (' + rhs + ');')(log);
  assert.deepEqual(val, log.photo_ids);
});

test('#3 회귀: 서버 getTennisData는 사진 ID를 photo_ids로 돌려준다', () => {
  const id = 'AAAAAAAAAAAAAAAAAAAAAAAA';
  const ctx = loadGas({ logs: [logRow('2026-11-01', '레슨A', { photoIds: id })] });
  const d = ctx.getTennisData();
  assert.deepEqual(Array.from(d.logs[0].photo_ids), [id]);
});

test('#3 회귀: updateTennisLog는 existingPhotoIds를 13열에 그대로 남긴다', () => {
  const id = 'AAAAAAAAAAAAAAAAAAAAAAAA';
  const ctx = loadGas({ logs: [logRow('2026-11-01', '레슨A', { photoIds: id })] });
  const r = post(ctx, 'updateTennisLog', { rowIndex: 2, date: '2026-11-01', title: '레슨A', notes: '메모 수정', existingPhotoIds: [id], photos: [] });
  assert.equal(r.success, true);
  assert.equal(ctx.__tennis.data[1][12], id);
  assert.equal(ctx.__tennis.data[1][11], '메모 수정');
});

// ---------- #4 행 번호로 지움 ----------
test('#4 행 번호가 밀린 뒤 삭제 요청이 와도 다른 일지를 지우면 안 된다', () => {
  const ctx = loadGas({ logs: [
    logRow('2026-11-01', 'A'), logRow('2026-11-02', 'B'), logRow('2026-11-03', 'C'), logRow('2026-11-04', 'D'),
  ] });
  // 화면은 B=3행, C=4행으로 알고 있음. B를 지운 뒤(이제 C가 3행, D가 4행) 옛 번호 4로 C를 지우려 함.
  post(ctx, 'deleteTennisLog', { rowIndex: 3, expectDate: '2026-11-02', expectTitle: 'B' });
  post(ctx, 'deleteTennisLog', { rowIndex: 4, expectDate: '2026-11-03', expectTitle: 'C' });
  const titles = ctx.__tennis.data.slice(1).map(r => r[2]);
  assert.ok(titles.includes('D'), 'D가 대신 지워짐: ' + titles.join(','));
  assert.ok(!titles.includes('C'), 'C가 남아 있음: ' + titles.join(','));
});

test('#4 회귀: 예전 방식(행 번호만)의 삭제 요청도 동작한다', () => {
  const ctx = loadGas({ logs: [logRow('2026-11-01', 'A'), logRow('2026-11-02', 'B')] });
  const r = post(ctx, 'deleteTennisLog', 3);
  assert.equal(r.success, true);
  assert.deepEqual(ctx.__tennis.data.slice(1).map(r => r[2]), ['A']);
});

test('#4 같은 날짜·제목 일지가 여럿이라 확신할 수 없으면 지우지 않는다', () => {
  const ctx = loadGas({ logs: [logRow('2026-11-01', '클럽 게임'), logRow('2026-11-01', '클럽 게임'), logRow('2026-11-02', 'X')] });
  const r = post(ctx, 'deleteTennisLog', { rowIndex: 4, expectDate: '2026-11-01', expectTitle: '클럽 게임' });
  assert.equal(r.success, false);
  assert.equal(ctx.__tennis.data.length, 4);
});

test('#4 댓글도 행이 밀렸으면 맞는 일지에 달린다', () => {
  const ctx = loadGas({ logs: [logRow('2026-11-01', 'A'), logRow('2026-11-02', 'B')] });
  const r = post(ctx, 'addTennisComment', { rowIndex: 2, commentData: { author: '나', text: 'B에 댓글', expectDate: '2026-11-02', expectTitle: 'B' } });
  assert.equal(r.success, true);
  assert.equal(ctx.__tennis.data[1][13], '[]', 'A 일지에 잘못 달림');
  assert.match(String(ctx.__tennis.data[2][13]), /B에 댓글/);
});

test('#4 화면의 임시(저장 중) 일지 번호가 날짜순 첫 항목 번호+1이 아니다', () => {
  assert.ok(!/allLogs\[0\]\.rowIndex \+ 1/.test(INDEX_HTML), '날짜순 첫 항목 번호+1을 임시 번호로 씀');
});

// ---------- #5 댓글 수정 버튼 XSS ----------
test('#5 작은따옴표가 든 댓글이 수정 버튼에서 스크립트로 실행되면 안 된다', () => {
  const line = INDEX_HTML.split('\n').find(l => l.includes('onclick="editComment('));
  assert.ok(line);
  const expr = line.trim().replace(/\s*\+\s*$/, '');
  const c = { id: 'uuid-1', text: "'-alert(1)-'" };
  const log = { rowIndex: 5 };
  // 고치기 전 코드는 safeText(encodeURIComponent 결과), 고친 뒤 코드는 safeId를 쓴다. 둘 다 원래 코드와 같은 방식으로 만든다.
  const html = new Function('c', 'log', 'escapeHtml', 'safeText', 'safeId', 'return ' + expr + ';')(c, log, s => String(s), encodeURIComponent(c.text), c.id);
  const onclick = html.match(/onclick="([^"]*)"/)[1].replace(/&#039;/g, "'").replace(/&quot;/g, '"').replace(/&amp;/g, '&');
  let alerted = false; const calls = [];
  try { new Function('editComment', 'alert', onclick)((...a) => calls.push(a), () => { alerted = true; }); } catch (e) { calls.push(['SYNTAX', e.message]); }
  assert.equal(alerted, false, '댓글 내용이 코드로 실행됨');
  assert.equal(calls.length, 1);
  assert.notEqual(calls[0][0], 'SYNTAX', '수정 버튼이 문법 오류로 동작 안 함');
});

// ---------- #6 날짜 ----------
test('#6 KST 07:30에 앱을 열면 오늘(KST) 날짜가 기본값이어야 한다', () => {
  const init = extractFunction(INDEX_HTML, 'initApp');
  assert.ok(!/toISOString\(\)\.split\('T'\)\[0\]/.test(init), 'initApp가 UTC 날짜를 씀');
  const fn = new Function('return ' + extractFunction(INDEX_HTML, 'getLocalDateString'))();
  assert.equal(fn(new Date('2026-11-05T22:30:00Z')), '2026-11-06');
});

test('#6 서버 경과일: KST 2026-11-06 07:30 기준 11-01 교체는 5일 경과', () => {
  const ctx = loadGas({ now: '2026-11-05T22:30:00Z', strings: [['2026-11-01', '라켓', '스트링', '폴리', 50, 0, '', '', '']] });
  const d = ctx.getTennisStringData();
  assert.equal(d.statsSince.daysElapsed, 5);
});

test('#6 회귀: 서버 경과일 KST 오후에도 같은 값', () => {
  const ctx = loadGas({ now: '2026-11-06T06:00:00Z', strings: [['2026-11-01', '라켓', '스트링', '폴리', 50, 0, '', '', '']] });
  assert.equal(ctx.getTennisStringData().statsSince.daysElapsed, 5);
});

// ---------- #7 업로드 실패를 숨김 ----------
test('#7 사진 업로드가 실패하면 결과에 실패 수가 담겨야 한다', () => {
  const ctx = loadGas({ driveFails: true });
  const r = post(ctx, 'saveTennisLog', { date: '2026-11-06', title: 'T', photos: ['data:image/jpeg;base64,AAAA', 'data:image/jpeg;base64,BBBB'] });
  assert.equal(r.success, true);
  assert.equal(r.failedMedia, 2);
});

test('#7 회귀: 업로드가 성공하면 사진 ID가 13열에 저장되고 failedMedia는 0', () => {
  const ctx = loadGas({});
  const r = post(ctx, 'saveTennisLog', { date: '2026-11-06', title: 'T', photos: ['data:image/jpeg;base64,AAAA'] });
  assert.equal(r.success, true);
  assert.equal(r.failedMedia || 0, 0);
  assert.match(String(ctx.__tennis.data[1][12]), /^FILEID_/);
  assert.equal(ctx.__tennis.data[1].length, 15);
});

test('#7 화면은 failedMedia가 있으면 사용자에게 알린다', () => {
  const at = INDEX_HTML.indexOf("callGasApi('saveTennisLog'");
  assert.ok(at > 0);
  assert.match(INDEX_HTML.slice(at, at + 600), /failedMedia/);
});

// ---------- #8 잠금 ----------
test('#8 댓글 쓰기는 스크립트 잠금(LockService) 안에서 실행된다', () => {
  const ctx = loadGas({ logs: [logRow('2026-11-01', 'A')] });
  post(ctx, 'addTennisComment', { rowIndex: 2, commentData: { author: '나', text: 'hi' } });
  assert.ok(ctx.__lockCalls.some(c => c.startsWith('wait')), '잠금 없음');
  assert.ok(ctx.__lockCalls.includes('release'), '잠금 해제 없음');
});

test('#8 회귀: 댓글 추가·수정·삭제가 순서대로 동작한다', () => {
  const ctx = loadGas({ logs: [logRow('2026-11-01', 'A')] });
  const a = post(ctx, 'addTennisComment', { rowIndex: 2, commentData: { author: '나', text: '첫 댓글' } });
  assert.equal(a.success, true);
  const id = a.comments[0].id;
  const u = post(ctx, 'updateTennisComment', { rowIndex: 2, commentId: id, newText: '고친 댓글' });
  assert.equal(u.comments[0].text, '고친 댓글');
  const d = post(ctx, 'deleteTennisComment', { rowIndex: 2, commentId: id });
  assert.deepEqual(d.comments, []);
});

// ---------- #9 텔레그램 HTML ----------
test('#9 메모의 < & 는 텔레그램 HTML용으로 바꿔 보낸다', () => {
  const ctx = loadGas({ props: { TELEGRAM_BOT_TOKEN: 'fake', TELEGRAM_CHAT_ID: '1' } });
  post(ctx, 'saveTennisLog', { date: '2026-11-06', title: 'A&B 복식', notes: '포핸드 <3 성공', photos: [] });
  assert.equal(ctx.__fetches.length, 1);
  const text = JSON.parse(ctx.__fetches[0].o.payload).text;
  assert.match(text, /포핸드 &lt;3 성공/);
  assert.match(text, /A&amp;B 복식/);
  assert.match(text, /<b>/, '굵은 글씨 태그는 유지');
});

test('#9 댓글 알림도 이스케이프한다', () => {
  const ctx = loadGas({ props: { TELEGRAM_BOT_TOKEN: 'fake', TELEGRAM_CHAT_ID: '1' }, logs: [logRow('2026-11-01', 'A')] });
  post(ctx, 'addTennisComment', { rowIndex: 2, commentData: { author: '다람이', text: '<a href="http://x">클릭</a>' } });
  const text = JSON.parse(ctx.__fetches[0].o.payload).text;
  assert.ok(!text.includes('<a href'), text);
});

// ---------- #10 GitHub Pages 판 텔레그램 테스트 ----------
function runCallGasApi(action, fetchImpl) {
  const src = extractFunction(INDEX_HTML, 'callGasApi');
  const f = new Function('fetch', 'GAS_API_URL', 'alert', 'google', src + '; return callGasApi;')(fetchImpl, 'https://example.invalid/exec', () => {}, undefined);
  return new Promise(resolve => f(action, undefined, res => resolve({ ok: true, res }), err => resolve({ ok: false, err })));
}

test('#10 텔레그램 테스트는 서버가 처리하는 GET으로 보낸다', async () => {
  const seen = [];
  const fetchImpl = (url, opt) => {
    const method = (opt && opt.method) || 'GET';
    seen.push(method);
    // 실제 서버: doGet은 이 action을 처리, doPost는 분기가 없어 {success:false}
    return Promise.resolve({ json: () => Promise.resolve(method === 'GET' ? { success: true } : { success: false }) });
  };
  const r = await runCallGasApi('testTennisTelegramAlert', fetchImpl);
  assert.equal(seen[0], 'GET');
  assert.equal(r.ok, true);
});

test('#10 서버가 {success:false}만 돌려주면 성공으로 처리하지 않는다', async () => {
  const fetchImpl = () => Promise.resolve({ json: () => Promise.resolve({ success: false }) });
  const r = await runCallGasApi('testTennisTelegramAlert', fetchImpl);
  assert.equal(r.ok, false);
});

test('회귀: 일지 조회(GET getTennisData)는 {success:false} 없이도 성공으로 처리된다', async () => {
  const fetchImpl = () => Promise.resolve({ json: () => Promise.resolve({ logs: [], overview: {} }) });
  const r = await runCallGasApi('getTennisData', fetchImpl);
  assert.equal(r.ok, true);
});
