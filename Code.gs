/**
 * 동명 급식 도우미 · 「급식실에 한마디」 공유 서버 (구글 앱스크립트)
 *
 * 설치 방법 (알림장 공유 서버와 같은 방법)
 * 1. 구글 시트를 새로 만들고 제목을 「동명 급식 한마디」로 바꿔요.
 * 2. [확장 프로그램] → [Apps Script]를 열고, 원래 코드를 지운 뒤 이 파일 내용을 붙여 넣고 저장해요.
 * 3. [배포] → [새 배포] → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → [배포] → 권한 허용
 * 4. 나온 웹 앱 주소(https://script.google.com/macros/s/…/exec)를 config.js의 SHARE_URL에 붙여 넣어요.
 *
 * 메뉴를 평가하지 않아요. 모으는 것은 세 가지뿐이에요.
 *   - likes  : 좋아하는 메뉴 ❤️ (누르거나 안 누르거나)
 *   - checks : 오늘 급식 ✅ 다 먹었어요 / 🙂 조금 남겼어요 (내 실천) + 💌 급식실에 한마디
 * 학생 이름은 모으지 않아요(반만). 부적절한 한마디는 checks 탭 '숨김' 칸에 1을 적으면 앱에서 안 보여요.
 */

// 시트에 붙은 스크립트면 비워 두고, 따로 만든 스크립트면 시트 주소의 /d/ 뒤 ID를 넣어요.
const SHEET_ID = '';
const CLASSES = ['1-1', '1-2', '1-3', '2-1', '2-2', '2-3', '3-1', '3-2'];
const label_ = gc => gc.split('-')[0] + '학년 ' + gc.split('-')[1] + '반';
// 아주 기본적인 말 거르기: 걸리면 저장은 하되 숨김 처리
const BAD = ['시발', '씨발', 'ㅅㅂ', '병신', 'ㅂㅅ', '존나', 'ㅈㄴ', '개새', '꺼져', '죽어', '미친', '맛없', '토나', '쓰레기'];

function sheet_(name, head) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); sh.setFrozenRows(1); }
  return sh;
}
const likesSheet_ = () => sheet_('likes', ['날짜', '반', '메뉴', '기기', '시각']);
const checksSheet_ = () => sheet_('checks', ['날짜', '반', '먹은 정도', '한마디', '기기', '시각', '숨김']);
const clean_ = (v, max) => String(v == null ? '' : v).replace(/[\r\n\t]/g, ' ').trim().slice(0, max);
const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const ymd_ = d => Utilities.formatDate(d, 'Asia/Seoul', 'yyyyMMdd');
const now_ = () => Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');

const likes_ = () => likesSheet_().getDataRange().getValues().slice(1)
  .map(r => ({ date: String(r[0]), gc: String(r[1]), dish: String(r[2]), did: String(r[3]) }));
const checks_ = () => checksSheet_().getDataRange().getValues().slice(1)
  .map(r => ({ date: String(r[0]), gc: String(r[1]), ate: String(r[2]), text: String(r[3]), did: String(r[4]), hide: String(r[6]) === '1' }));

function summary_(from, to, did) {
  const likes = {}, mine = [];
  likes_().filter(l => l.date >= from && l.date <= to).forEach(l => {
    const k = l.date + '|' + l.dish; likes[k] = (likes[k] || 0) + 1;
    if (did && l.did === did) mine.push(k);
  });
  const days = {};
  checks_().filter(c => c.date >= from && c.date <= to).forEach(c => {
    const d = days[c.date] || (days[c.date] = { all: 0, some: 0, notes: [] });
    if (c.ate === '다 먹음') d.all++; else if (c.ate === '조금 남김') d.some++;
    if (c.text && !c.hide) d.notes.push(c.text);
  });
  Object.keys(days).forEach(k => { days[k].notes = days[k].notes.slice(-5); });
  return { ok: true, likes: likes, days: days, mine: mine };
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.story) {   // 「급식 이야기」 탭: 최근 30일
    const since = ymd_(new Date(Date.now() - 30 * 864e5));
    const cnt = {};
    likes_().filter(l => l.date >= since).forEach(l => { cnt[l.dish] = (cnt[l.dish] || 0) + 1; });
    const top = Object.keys(cnt).map(k => ({ dish: k, n: cnt[k] })).sort((a, b) => b.n - a.n).slice(0, 10);
    const cls = {};
    const recent = checks_().filter(c => c.date >= since);
    recent.forEach(c => {
      const s = cls[c.gc] || (cls[c.gc] = { gc: c.gc, all: 0, n: 0 });
      if (c.ate === '다 먹음' || c.ate === '조금 남김') { s.n++; if (c.ate === '다 먹음') s.all++; }
    });
    const notes = recent.filter(c => c.text && !c.hide).slice(-10).reverse().map(c => ({ date: c.date, text: c.text }));
    return json_({ ok: true, top: top, classes: Object.keys(cls).sort().map(k => cls[k]), notes: notes, noteCount: recent.filter(c => c.text && !c.hide).length });
  }
  const from = clean_(p.from, 8), to = clean_(p.to, 8);
  if (!/^\d{8}$/.test(from) || !/^\d{8}$/.test(to)) return json_({ ok: false, error: '날짜가 올바르지 않아요.' });
  return json_(summary_(from, to, clean_(p.did, 20)));
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '잘못된 요청이에요.' }); }
  const gc = String(b.gc || ''), date = clean_(b.date, 8), did = clean_(b.did, 20);
  if (CLASSES.indexOf(gc) < 0) return json_({ ok: false, error: '반 정보가 올바르지 않아요.' });
  if (!/^\d{8}$/.test(date) || date > ymd_(new Date())) return json_({ ok: false, error: '먹은 날 급식에만 남길 수 있어요.' });
  if (!did) return json_({ ok: false, error: '잘못된 요청이에요.' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (b.action === 'like') {
      const dish = clean_(b.dish, 40);
      if (!dish) return json_({ ok: false, error: '메뉴가 없어요.' });
      const sh = likesSheet_(), v = sh.getDataRange().getValues();
      for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === date && String(v[i][2]) === dish && String(v[i][3]) === did) sh.deleteRow(i + 1);
      if (b.on) sh.appendRow([date, label_(gc), dish, did, now_()]);
    } else if (b.action === 'check') {
      const ate = b.ate === 'all' ? '다 먹음' : b.ate === 'some' ? '조금 남김' : '';
      const text = clean_(b.text, 40);
      const hide = BAD.some(w => text.indexOf(w) >= 0) ? 1 : '';
      const row = [date, label_(gc), ate, text, did, now_(), hide];
      const sh = checksSheet_(), v = sh.getDataRange().getValues();
      let at = -1;
      for (let i = 1; i < v.length; i++) if (String(v[i][0]) === date && String(v[i][4]) === did) { at = i + 1; break; }
      if (at > 0) sh.getRange(at, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
    } else {
      return json_({ ok: false, error: '알 수 없는 요청이에요.' });
    }
  } finally { lock.releaseLock(); }
  return json_(summary_(date, date, did));
}
