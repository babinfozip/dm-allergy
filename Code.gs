/**
 * 동명 급식 도우미 · 메뉴 평가 공유 서버 (구글 앱스크립트)
 *
 * 설치 방법 (알림장 공유 서버와 같은 방법)
 * 1. 구글 시트를 새로 만들고 제목을 「동명 급식 평가」로 바꿔요.
 * 2. [확장 프로그램] → [Apps Script]를 열고, 원래 코드를 지운 뒤 이 파일 내용을 붙여 넣고 저장해요.
 * 3. [배포] → [새 배포] → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → [배포] → 권한 허용
 * 4. 나온 웹 앱 주소(https://script.google.com/macros/s/…/exec)를 config.js의 SHARE_URL에 붙여 넣어요.
 *
 * 시트 'ratings' 탭이 자동으로 생겨요. 한 기기는 같은 날 같은 메뉴에 한 번만 평가하고, 다시 하면 고쳐져요.
 * 부적절한 한 줄 평은 시트의 '숨김' 칸에 1을 적으면 앱에서 안 보여요. 학생 이름은 모으지 않아요.
 */

const RATE = 'ratings';
const HEAD = ['날짜', '반', '메뉴', '별점', '남길것같아요', '한줄평', '기기', '시각', '숨김'];
const CLASSES = ['1-1', '1-2', '1-3', '2-1', '2-2', '2-3', '3-1', '3-2'];
const label_ = gc => gc.split('-')[0] + '학년 ' + gc.split('-')[1] + '반';
// 아주 기본적인 말 거르기: 걸리면 한 줄 평은 저장하되 숨김 처리
const BAD = ['시발', '씨발', 'ㅅㅂ', '병신', 'ㅂㅅ', '존나', 'ㅈㄴ', '개새', '꺼져', '죽어', '미친'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(RATE);
  if (!sh) { sh = ss.insertSheet(RATE); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}
const clean_ = (v, max) => String(v == null ? '' : v).replace(/[\r\n\t]/g, ' ').trim().slice(0, max);
const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const today_ = () => Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
const rows_ = () => sheet_().getDataRange().getValues().slice(1).map(r => ({
  date: String(r[0]), gc: String(r[1]), dish: String(r[2]), stars: +r[3], left: +r[4], text: String(r[5]), hide: String(r[8]) === '1'
}));

function summarize_(list) {
  const by = {};
  list.forEach(r => {
    const k = r.date + '|' + r.dish;
    const s = by[k] || (by[k] = { date: r.date, dish: r.dish, sum: 0, n: 0, left: 0, reviews: [] });
    s.sum += r.stars; s.n++; s.left += r.left ? 1 : 0;
    if (r.text && !r.hide) s.reviews.push(r.text);
  });
  return Object.keys(by).map(k => {
    const s = by[k];
    return { date: s.date, dish: s.dish, avg: s.sum / s.n, n: s.n, left: s.left, reviews: s.reviews.slice(-3) };
  });
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const all = rows_();
  if (p.rank) {
    const since = Utilities.formatDate(new Date(Date.now() - 30 * 864e5), 'Asia/Seoul', 'yyyyMMdd');
    const recent = all.filter(r => r.date >= since);
    const byDish = {};
    recent.forEach(r => {
      const s = byDish[r.dish] || (byDish[r.dish] = { dish: r.dish, sum: 0, n: 0, left: 0 });
      s.sum += r.stars; s.n++; s.left += r.left ? 1 : 0;
    });
    const list = Object.keys(byDish).map(k => byDish[k]).filter(s => s.n >= 3).map(s => ({ dish: s.dish, avg: s.sum / s.n, n: s.n, left: s.left }));
    const top = list.slice().sort((a, b) => b.avg - a.avg || b.n - a.n).slice(0, 10);
    const left = list.filter(s => s.left > 0).sort((a, b) => b.left / b.n - a.left / a.n).slice(0, 5);
    const cnt = {};
    recent.forEach(r => { cnt[r.gc] = (cnt[r.gc] || 0) + 1; });
    const classes = Object.keys(cnt).sort().map(gc => ({ gc: gc, n: cnt[gc] }));
    return json_({ ok: true, top: top, left: left, classes: classes });
  }
  const from = clean_(p.from, 8), to = clean_(p.to, 8);
  if (!/^\d{8}$/.test(from) || !/^\d{8}$/.test(to)) return json_({ ok: false, error: '날짜가 올바르지 않아요.' });
  return json_({ ok: true, items: summarize_(all.filter(r => r.date >= from && r.date <= to)) });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '잘못된 요청이에요.' }); }
  if (b.action !== 'rate') return json_({ ok: false, error: '알 수 없는 요청이에요.' });
  const gc = String(b.gc || ''), date = clean_(b.date, 8), dish = clean_(b.dish, 40), did = clean_(b.did, 20);
  const stars = Math.round(+b.stars);
  if (CLASSES.indexOf(gc) < 0) return json_({ ok: false, error: '반 정보가 올바르지 않아요.' });
  if (!/^\d{8}$/.test(date) || date > today_()) return json_({ ok: false, error: '먹은 날 급식만 평가할 수 있어요.' });
  if (!dish || !did || !(stars >= 1 && stars <= 5)) return json_({ ok: false, error: '별점을 골라 주세요.' });
  const text = clean_(b.text, 30);
  const hide = BAD.some(w => text.indexOf(w) >= 0) ? 1 : '';
  const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  const row = [date, label_(gc), dish, stars, b.leftover ? 1 : 0, text, did, now, hide];

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const v = sh.getDataRange().getValues();
    let at = -1;
    for (let i = 1; i < v.length; i++) if (String(v[i][0]) === date && String(v[i][2]) === dish && String(v[i][6]) === did) { at = i + 1; break; }
    if (at > 0) sh.getRange(at, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  } finally { lock.releaseLock(); }

  const item = summarize_(rows_().filter(r => r.date === date && r.dish === dish))[0];
  return json_({ ok: true, item: item });
}
