/**
 * 방명록 저장소 (Google 스프레드시트 + Apps Script, 무료)
 *
 * 설정 방법
 * 1. Google 스프레드시트 새로 만들기 → 1행에 A1 "date", B1 "name", C1 "message" 입력
 * 2. 메뉴 [확장 프로그램] → [Apps Script] → 이 파일 내용을 전부 붙여넣고 저장
 * 3. [배포] → [새 배포] → 유형: 웹 앱
 *    - 실행 사용자: 나 / 액세스 권한: 모든 사용자 → 배포 → 권한 승인
 * 4. 나온 웹 앱 URL(https://script.google.com/macros/s/.../exec)을
 *    index.html 의 GUESTBOOK_URL 에 붙여넣기
 *
 * 메시지 삭제는 스프레드시트에서 해당 행을 지우면 됩니다.
 */
const sheet = () => SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
const json = obj => ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);

function doGet() {
  const rows = sheet().getDataRange().getValues().slice(1);
  return json(rows.reverse().map(([date, name, message]) => ({ date, name, message })));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const { name, message, hp } = JSON.parse(e.postData.contents);
    // 허니팟: 화면에 안 보이는 칸이라 사람은 비워두고 봇만 채움 → 성공인 척 무시
    if (hp) return json({ ok: true });
    const n = String(name || '').trim().slice(0, 15);
    const m = String(message || '').trim().slice(0, 200);
    if (!n || !m) return json({ ok: false });
    // 링크 스팸 차단 (방명록에 URL 쓸 하객은 없음)
    if (/https?:\/\/|www\.|[a-z0-9-]+\.(com|net|org|co|io|xyz|ru|top|shop|link)\b/i.test(m)) {
      return json({ ok: false, reason: 'link' });
    }
    if (!lock.tryLock(5000)) return json({ ok: false });
    const sh = sheet(), last = sh.getLastRow();
    if (last > 1) {
      const n20 = Math.min(20, last - 1);
      const recent = sh.getRange(last - n20 + 1, 1, n20, 3).getValues();
      // 같은 이름이 쓴 글만 검사. 전역으로 하면 여러 하객이 동시에 쓸 때 정상 글이 거부됨
      const mine = recent.filter(r => String(r[1]).trim() === n);
      if (mine.length) {
        // 연타 차단: 같은 사람이 10초 안에 또 쓰기
        if (new Date() - new Date(mine[mine.length - 1][0]) < 10000) return json({ ok: false, reason: 'slow' });
        // 도배 차단: 같은 사람이 같은 내용 또 쓰기
        if (mine.some(r => String(r[2]).trim() === m)) return json({ ok: false, reason: 'dup' });
      }
    }
    // 스프레드시트 수식 주입 방지
    const safe = s => (/^[=+\-@]/.test(s) ? "'" + s : s);
    sh.appendRow([new Date(), safe(n), safe(m)]);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false });
  } finally {
    lock.releaseLock();
  }
}

/** Apps Script 편집기에서 실행해 스팸 필터 확인 (시트에 쓰지 않는 경우만 검사) */
function test_doPost() {
  const call = o => JSON.parse(doPost({ postData: { contents: JSON.stringify(o) } }).getContent());
  const ok = (c, m) => { if (!c) throw new Error(m); };
  ok(call({ name: '봇', message: 'hi', hp: 'x' }).ok, '허니팟은 ok:true 로 조용히 무시');
  ok(call({ name: '봇', message: '대출 https://spam.ru' }).reason === 'link', '링크 차단');
  ok(call({ name: '봇', message: '여기 www.spam.shop 보세요' }).reason === 'link', 'www 차단');
  ok(call({ name: '', message: '' }).ok === false, '빈 입력 차단');
  ok(call({ name: '하객', message: '결혼 축하해! 2.5년 연애 끝' }).reason !== 'link', '정상 메시지 통과');
  Logger.log('모두 통과 ✅ (마지막 정상 메시지 1건은 시트에 기록됨 → 지워주세요)');
}
