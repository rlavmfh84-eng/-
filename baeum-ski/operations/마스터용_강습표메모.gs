/** @OnlyCurrentDoc */
// 배움스키아카데미 마스터 · 강습표 메모 자동 생성 (권한 요청 없음)
// 설치: 확장 프로그램 → Apps Script → 기존 내용 모두 지우고 이 내용 붙여넣기 → 저장(실행 버튼은 누르지 않기)
// 이후 관리자_MASTER 또는 강사DB를 수정할 때마다 강습표 칸 메모가 자동으로 새로 만들어집니다.

const MASTER_SHEET = '관리자_MASTER';
const SCHEDULE_PREFIX = '강습표_';
const FIRST_ROW = 5;
const BLOCK_WIDTH = 14;   // 요일/날짜 + 시간 + 강사 12명
const TEACHERS = 12;
// 9시 칸은 9시 이전 시작 강습 포함, 야간 = 19시 이후
const SLOTS = [[0, 600]]
  .concat([10, 11, 12, 13, 14, 15, 16, 17, 18].map(h => [h * 60, h * 60 + 60]))
  .concat([[1140, 1440]]);

// MASTER나 강사DB를 수정하면 자동으로 메모를 새로 만듭니다 (권한 요청 없음)
function onEdit(e) {
  const name = e && e.range ? e.range.getSheet().getName() : '';
  if (name === MASTER_SHEET || name === '강사DB') updateNotes();
}

function hhmm(min) {
  const h = Math.floor(min / 60), m = min % 60;
  return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m;
}

function updateNotes() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(MASTER_SHEET);
  const last = master.getLastRow();
  const rows = last >= FIRST_ROW ? master.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, 23).getValues() : [];

  ss.getSheets().filter(s => s.getName().indexOf(SCHEDULE_PREFIX) === 0).forEach(sheet => {
    const month = sheet.getRange('C2').getValue();
    if (!(month instanceof Date)) return;
    const first = new Date(month.getFullYear(), month.getMonth(), 1);
    const weekStart = new Date(first.getFullYear(), first.getMonth(), 1 - ((first.getDay() + 6) % 7));
    const nRows = 7 * SLOTS.length;
    const nCols = sheet.getLastColumn();
    const nWeeks = Math.floor(nCols / BLOCK_WIDTH);
    const names = sheet.getRange(4, 3, 1, TEACHERS).getValues()[0];
    const notes = Array.from({ length: nRows }, () => Array(nCols).fill(''));

    rows.forEach(r => {
      const status = r[1], booker = r[3], phone = r[4], students = r[5], type = r[11],
            teacher = r[12], memo = r[17], equip = r[19], date = r[20], start = r[21], end = r[22];
      if (!booker || (status !== '결제완료' && status !== '상담중')) return;
      if (!(date instanceof Date) || start === '' || end === '') return;
      if (date.getFullYear() !== month.getFullYear() || date.getMonth() !== month.getMonth()) return;
      const t = names.indexOf(teacher);
      if (t < 0) return;
      const days = Math.round((Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) -
                               Date.UTC(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate())) / 86400000);
      const b = Math.floor(days / 7), k = days % 7;
      if (b < 0 || b >= nWeeks) return;

      const text = [
        '강습생: ' + (students || booker) + (status === '상담중' ? '  (상담중 · 입금 전)' : ''),
        '예약자: ' + booker,
        '연락처: ' + (phone || '-'),
        '시간: ' + hhmm(start) + ' ~ ' + hhmm(end) + (type ? '  (' + type + ')' : ''),
        '장비대여: ' + (equip || '-'),
        '특이사항: ' + (memo || '-'),
      ].join('\n');
      SLOTS.forEach((slot, j) => {
        if (start < slot[1] && end > slot[0]) {
          const row = k * SLOTS.length + j, col = b * BLOCK_WIDTH + 2 + t;
          notes[row][col] = notes[row][col] ? notes[row][col] + '\n──────\n' + text : text;
        }
      });
    });
    sheet.getRange(FIRST_ROW, 1, nRows, nCols).setNotes(notes);
  });
}
