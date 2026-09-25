/**
 * 실무 서류 자동 생성 세트 — 설치_한번에 (편집기에 이 파일 하나만 붙여넣기)
 * ------------------------------------------------------------
 * Code.js + 양식.js + 설치.js 를 합친 파일입니다.
 * 붙여넣은 뒤 실행 순서:  설치_전체  →  설치_샘플데이터
 * (clasp 로 올릴 때는 이 파일을 올리지 않습니다 — 함수 이름이 겹칩니다.
 *  .claspignore 에서 제외해 두었습니다.)
 * ------------------------------------------------------------
 */

/* ===================== Code.js ===================== */

/**
 * 실무 서류 자동 생성 세트 — 본체
 * ------------------------------------------------------------
 * 실행 환경: Google Sheets + Apps Script (무료 계정)
 * 데이터: 전부 합성(가상) 데이터. 실제 회사 자료 사용 금지.
 *
 * 만드는 것
 *   1) 점검기록(20대 × 28일 × 9항목 = 5,040행) → 요약 2종(COUNTIFS)
 *      → 점검대장 .xlsx 내보내기 (엑셀에서 열리는 실무용 대장)
 *   2) 작업계획·설비 데이터 → 작업계획서 / 점검결과보고서
 *      (구글독스 + PDF + DOCX 각각 20부, 합본 PDF)
 *
 * 시트 탭: 안내 / 설비마스터 / 점검기록 / 작업계획 / 설정 /
 *          요약_설비별 / 요약_항목별 / 발행이력 / 측정로그 (/ 피벗)
 *
 * 대장 수식은 엑셀에서도 도는 것만 쓴다(COUNTIFS·IF·TEXT·XLOOKUP).
 * QUERY·FILTER·ARRAYFORMULA 는 엑셀에 없으므로 금지.
 */

const TZ = 'Asia/Seoul';

const SH = {
  안내: '안내',
  설비: '설비마스터',
  점검: '점검기록',
  계획: '작업계획',
  설정: '설정',
  요약설비: '요약_설비별',
  요약항목: '요약_항목별',
  발행: '발행이력',
  측정: '측정로그',
  피벗: '피벗',
};

const 항목목록 = ['토출압력', '오일레벨', '진동소음', '베어링온도', '탱크액위', '배관누설', '밸브잠금', '안전커버', '윤활급유'];

/* ============================ 공통 도구 ============================ */

const 로그모음 = [];

/** Logger.log + 배열 누적(설치 결과를 한 번에 돌려주기 위함) */
function _로그(...args) {
  const s = args.map(String).join(' ');
  Logger.log(s);
  로그모음.push(s);
  return s;
}

function 로그비우기() {
  로그모음.length = 0;
  _캐시비우기();
}

function 로그전체() {
  return 로그모음.join('\n');
}

/** 실패해도 전체를 멈추지 않는 실행 */
function _안전(이름, fn) {
  try {
    return fn();
  } catch (e) {
    _로그('[건너뜀] ' + 이름 + ': ' + e.message);
    return null;
  }
}

function _시트(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('탭이 없습니다: ' + name + ' — 편집기에서 설치_전체를 먼저 실행하세요.');
  return sh;
}

function _헤더행(sh, 행) {
  const 열수 = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(행, 1, 1, 열수).getValues()[0].map(String);
}

function _헤더(sh) {
  return _헤더행(sh, 1);
}

function _열(sh, name) {
  const i = _헤더(sh).indexOf(name);
  if (i < 0) throw new Error('헤더 없음: "' + name + '" in ' + sh.getName());
  return i + 1;
}

/** 행 배열 → {헤더: 값} */
function _행객체(h, r) {
  const o = {};
  h.forEach((k, j) => {
    o[k] = r[j] === null || r[j] === undefined ? '' : r[j];
  });
  return o;
}

/**
 * 설정·점검기록은 한 번 실행 안에서 여러 번 읽는다.
 * 그냥 읽을 때마다 시트를 다시 부르면 호출이 수만 번이 되어 6분 제한에 걸린다
 * (수식 15,120개를 만들면서 설정을 행마다 읽던 것이 실제로 그랬다).
 * 한 실행 안에서만 쓰는 메모리 캐시를 두고, 값을 쓰면 그 키만 갱신한다.
 */
const _설정캐시 = {};
let _점검값캐시 = null;

function _캐시비우기() {
  Object.keys(_설정캐시).forEach((k) => delete _설정캐시[k]);
  _점검값캐시 = null;
}

function _점검캐시비우기() {
  _점검값캐시 = null;
}

/** 점검기록 전체 값(헤더 포함). 부르는 쪽에서 slice() 해서 쓴다 */
function _점검값() {
  if (!_점검값캐시) _점검값캐시 = _시트(SH.점검).getDataRange().getValues();
  return _점검값캐시;
}

function 설정(key) {
  if (key in _설정캐시) return _설정캐시[key];
  const v = _시트(SH.설정).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === key) {
      _설정캐시[key] = v[i][1];
      return v[i][1];
    }
  }
  throw new Error('설정 탭에 없는 키: ' + key);
}

function _설정쓰기(key, 값) {
  const sh = _시트(SH.설정);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(값);
      _설정캐시[key] = 값;
      return;
    }
  }
  sh.appendRow([key, 값, '']);
  _설정캐시[key] = 값;
}

function _날짜(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

/**
 * 셀 값을 'yyyy-MM-dd' 문자열로 맞춘다.
 * 시트가 날짜로 바꿔 저장하면 Date 객체가 되고, 문자열이면 시각이 붙을 수 있어
 * 그대로 두면 날짜 비교가 어긋난다.
 */
function _날짜문자열(raw) {
  if (raw === null || raw === undefined) return '';
  if (Object.prototype.toString.call(raw) === '[object Date]') return _날짜(raw);
  const s = String(raw).trim();
  return s.length > 10 ? s.slice(0, 10) : s;
}

function _오늘0시_기준(d) {
  const s = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  return new Date(s + 'T00:00:00+09:00');
}

function _요일(d) {
  const w = ['일', '월', '화', '수', '목', '금', '토'];
  const 날 = _날짜문자열(d);
  if (!날) return '';
  const x = new Date(날 + 'T00:00:00+09:00');
  return isNaN(x.getTime()) ? '' : '(' + w[x.getDay()] + ')';
}

/** WP-2026-0007 */
function _문서번호(접두어, n) {
  const 연도 = Utilities.formatDate(new Date(), TZ, 'yyyy');
  return 접두어 + '-' + 연도 + '-' + ('0000' + Number(n)).slice(-4);
}

/** 폴더를 이름으로 찾고, 없으면 만든다(멱등) */
function _폴더보장(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** 1 → A, 27 → AA */
function _열문자(n) {
  let s = '';
  let x = Number(n);
  while (x > 0) {
    const r = (x - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

/* ============================ 대장 서식·수식 ============================ */

/**
 * 설비마스터의 실제 데이터 범위($A$2:$A$21 형태).
 * 전체 열($A:$A) 참조는 100만 행을 대상으로 잡혀 계산이 느려진다 —
 * 5,040행 수식에서는 체감 차이가 크므로 범위를 좁혀 쓴다.
 */
function _설비범위() {
  const 마지막 = Math.max(_시트(SH.설비).getLastRow(), 2);
  return { 태그: '$A$2:$A$' + 마지막, 열: (c) => '$' + c + '$2:$' + c + '$' + 마지막 };
}

/** 설정 조회함수 값에 따라 XLOOKUP 또는 INDEX/MATCH 수식 문자열 */
function _조회수식(찾을셀, 반환열문자, 범위) {
  const r = 범위 || _설비범위();
  const 방식 = String(설정('조회함수') || 'XLOOKUP').toUpperCase();
  if (방식 === 'INDEXMATCH') {
    return '=IFERROR(INDEX(' + SH.설비 + '!' + r.열(반환열문자) + ',MATCH(' + 찾을셀 + ',' + SH.설비 + '!' + r.태그 + ',0)),"미등록")';
  }
  return '=IFERROR(XLOOKUP(' + 찾을셀 + ',' + SH.설비 + '!' + r.태그 + ',' + SH.설비 + '!' + r.열(반환열문자) + '),"미등록")';
}

/**
 * 데이터 검증(드롭다운) 한 열.
 * requireValueInRange 는 같은 스프레드시트 안에서만 허용된다 —
 * 내보낸 점검대장은 임시 사본(다른 파일)이라 범위 참조를 쓰면 거기서 죽는다.
 * 값 목록(문자열 배열)으로 넣으면 어느 시트에서나 동작한다.
 */
function _검증목록(sh, 머리행, 열이름, 값들) {
  const h = _헤더행(sh, 머리행);
  const i = h.indexOf(열이름);
  if (i < 0) return false;
  const 목록 = 값들.map(String).filter((v) => v.trim() !== '');
  if (!목록.length) return false;
  const 검증 = SpreadsheetApp.newDataValidation().requireValueInList(목록, true)
    .setAllowInvalid(false).setHelpText(열이름 + ' 목록에서 고르세요').build();
  const 데이터수 = sh.getLastRow() - 머리행;
  if (데이터수 > 0) sh.getRange(머리행 + 1, i + 1, 데이터수, 1).setDataValidation(검증);
  return true;
}

/**
 * 점검기록(또는 내보낸 점검대장) 시트에 조회 수식·검증·조건부 서식·틀 고정·열 너비를 넣는다.
 * 시작행을 받는 이유: 내보낸 대장은 1~4행이 제목 블록이라 머리글이 5행이다.
 */
function 대장서식적용(sh, 시작행) {
  const 머리 = Number(시작행 || 1);
  const h = _헤더행(sh, 머리);
  const 마지막 = sh.getLastRow();
  const 데이터수 = 마지막 - 머리;
  if (!h.length || 데이터수 <= 0) {
    _로그('대장 서식: 데이터가 없어 건너뜁니다 (' + sh.getName() + ')');
    return;
  }

  const 태그열 = _열문자(h.indexOf('설비태그') + 1);
  const 조회범위 = _설비범위();

  // 1) 조회 수식 3열 — 엑셀에 ARRAYFORMULA 가 없으므로 행마다 기록한다
  const 매핑 = [['설비명', 'B'], ['구역', 'D'], ['담당자', 'F']];
  let 수식수 = 0;
  매핑.forEach(([이름, 반환열]) => {
    const c = h.indexOf(이름) + 1;
    if (!c) return;
    const 수식들 = [];
    for (let r = 머리 + 1; r <= 마지막; r++) 수식들.push([_조회수식('$' + 태그열 + r, 반환열, 조회범위)]);
    sh.getRange(머리 + 1, c, 수식들.length, 1).setFormulas(수식들);
    수식수 += 수식들.length;
  });

  // 2) 데이터 검증 4열
  const 설비태그들 = _시트(SH.설비).getDataRange().getValues().slice(1)
    .map((r) => String(r[0]).trim()).filter((v) => v);
  let 검증수 = 0;
  if (_검증목록(sh, 머리, '설비태그', 설비태그들)) 검증수++;
  if (_검증목록(sh, 머리, '판정', ['정상', '이상'])) 검증수++;
  if (_검증목록(sh, 머리, '심각도', ['상', '중', '하'])) 검증수++;
  if (_검증목록(sh, 머리, '조치상태', ['미조치', '진행중', '완료'])) 검증수++;

  // 3) 조건부 서식 4규칙 — 같은 시트만 참조해야 엑셀에서도 산다
  const 범위 = (이름) => sh.getRange(머리 + 1, h.indexOf(이름) + 1, 데이터수, 1);
  const L = _열문자(h.indexOf('판정') + 1);
  const N = _열문자(h.indexOf('조치상태') + 1);
  const 규칙 = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('이상')
      .setBackground('#f4cccc').setFontColor('#990000').setRanges([범위('판정')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('정상')
      .setBackground('#d9ead3').setRanges([범위('판정')]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($' + L + (머리 + 1) + '="이상",$' + N + (머리 + 1) + '="미조치")')
      .setBold(true).setFontColor('#c00000').setRanges([범위('조치상태')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('상')
      .setBackground('#fce5cd').setRanges([범위('심각도')]).build(),
  ];
  sh.setConditionalFormatRules(규칙);

  // 4) 틀 고정·필터·열 너비 (A4 가로 1페이지 폭에 들어가게)
  sh.setFrozenRows(머리);
  sh.setFrozenColumns(3);
  const 너비 = {
    점검ID: 70, 점검일: 85, 설비태그: 65, 설비명: 95, 구역: 70, 담당자: 65, 점검자: 65,
    항목명: 90, 측정값: 60, 기준: 60, 단위: 45, 판정: 50, 심각도: 50, 조치상태: 65, 조치내용: 120, 비고: 45,
  };
  Object.keys(너비).forEach((이름) => {
    const i = h.indexOf(이름);
    if (i >= 0) sh.setColumnWidth(i + 1, 너비[이름]);
  });
  _안전('필터', () => sh.getRange(머리, 1, 데이터수 + 1, h.length).createFilter());

  _로그('대장 서식: 수식 ' + 수식수 + '개 / 검증 ' + 검증수 + '열 / 조건부 서식 ' + 규칙.length + '규칙 / 틀 고정 ' + 머리 + '행 3열');
}

/* ============================ 요약 (COUNTIFS) ============================ */

/**
 * COUNTIFS 한 줄: 조건들 = [[열문자, 조건], ...]
 * 마지막행을 받아 점검기록의 실제 데이터 범위만 본다(전체 열 참조보다 훨씬 빠르다).
 */
function _카운트(조건들, 마지막행) {
  const m = Math.max(Number(마지막행) || 2, 2);
  return '=COUNTIFS(' + 조건들.map(([c, v]) => SH.점검 + '!$' + c + '$2:$' + c + '$' + m + ',' + v).join(',') + ')';
}

function 요약갱신() {
  const 점검 = _시트(SH.점검);
  if (점검.getLastRow() < 2) {
    _로그('요약: 점검기록이 비어 있어 건너뜁니다');
    return;
  }
  const h = _헤더(점검);
  const 마지막 = Math.max(점검.getLastRow(), 2);
  const 태그열 = _열문자(h.indexOf('설비태그') + 1);
  const 항목열 = _열문자(h.indexOf('항목명') + 1);
  const 판정열 = _열문자(h.indexOf('판정') + 1);
  const 심각열 = _열문자(h.indexOf('심각도') + 1);
  const 조치열 = _열문자(h.indexOf('조치상태') + 1);

  // 요약_설비별 — 설비 20행 × 항목 9열
  const 설비들 = _시트(SH.설비).getDataRange().getValues().slice(1).filter((r) => String(r[0]).trim());
  const 설비탭 = _시트(SH.요약설비);
  const 설비행들 = 설비들.map((r, i) => {
    const 행 = i + 2;
    const 값 = [String(r[0]).trim(), r[1], r[3]];
    const 항목수식 = 항목목록.map((_, k) => {
      const 열 = _열문자(4 + k);
      return _카운트([[태그열, '$A' + 행], [항목열, 열 + '$1'], [판정열, '"이상"']], 마지막);
    });
    값.push(...항목수식);
    값.push('=SUM(D' + 행 + ':L' + 행 + ')');
    값.push('=ROUND(COUNTIF(' + SH.점검 + '!$' + 태그열 + '$2:$' + 태그열 + '$' + 마지막 + ',$A' + 행 + ')/' + 항목목록.length + ',0)');
    값.push('=IFERROR(ROUND(M' + 행 + '/(N' + 행 + '*' + 항목목록.length + ')*100,1),0)');
    return 값;
  });
  if (설비행들.length) {
    설비탭.getRange(2, 1, 설비행들.length, 설비행들[0].length).setValues(설비행들);
    // 이상합계 색 눈금 (구글시트·엑셀 모두 지원)
    _안전('색 눈금', () => {
      설비탭.setConditionalFormatRules([
        SpreadsheetApp.newConditionalFormatRule().setGradientMinpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '0')
          .setGradientMaxpointWithValue('#F4CCCC', SpreadsheetApp.InterpolationType.NUMBER, '10')
          .setRanges([설비탭.getRange(2, 13, 설비행들.length, 1)]).build(),
      ]);
    });
  }

  // 요약_항목별 — 항목 9행
  const 항목탭 = _시트(SH.요약항목);
  const 항목행들 = 항목목록.map((항목, i) => {
    const 행 = i + 2;
    return [
      항목,
      _카운트([[항목열, '$A' + 행], [판정열, '"이상"']], 마지막),
      _카운트([[항목열, '$A' + 행], [판정열, '"이상"'], [심각열, '"상"']], 마지막),
      _카운트([[항목열, '$A' + 행], [판정열, '"이상"'], [심각열, '"중"']], 마지막),
      _카운트([[항목열, '$A' + 행], [판정열, '"이상"'], [심각열, '"하"']], 마지막),
      _카운트([[항목열, '$A' + 행], [판정열, '"이상"'], [조치열, '"미조치"']], 마지막),
      '=IFERROR(TEXT(MAXIFS(' + SH.점검 + '!$B$2:$B$' + 마지막 + ',' + SH.점검 + '!$' + 항목열 + '$2:$' + 항목열 + '$' + 마지막 + ',$A' + 행 +
        ',' + SH.점검 + '!$' + 판정열 + '$2:$' + 판정열 + '$' + 마지막 + ',"이상"),"yyyy-mm-dd"),"—")',
    ];
  });
  항목탭.getRange(2, 1, 항목행들.length, 항목행들[0].length).setValues(항목행들);

  _로그('요약: 설비 ' + 설비행들.length + ' × 항목 ' + 항목목록.length + ' 수식, 항목별 ' + 항목행들.length + '행 (범위 $2:$' + 마지막 + ')');
}

/* ============================ 피벗 (선택) ============================ */

/**
 * Sheets API 고급 서비스가 있을 때만 네이티브 피벗을 만든다.
 * .xlsx 로 내보내면 피벗 정의가 빠질 수 있어 요약 2탭이 항상 본체다.
 */
function 피벗만들기() {
  if (typeof Sheets === 'undefined') {
    _로그('[건너뜀] Sheets API 고급 서비스가 없습니다. 요약 탭으로 대체합니다.');
    return false;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = _시트(SH.점검);
  if (src.getLastRow() < 2) {
    _로그('피벗: 점검기록이 비어 있어 건너뜁니다');
    return false;
  }
  const 피벗탭 = ss.getSheetByName(SH.피벗) || ss.insertSheet(SH.피벗);
  피벗탭.clear();
  const 요청 = {
    updateCells: {
      start: { sheetId: 피벗탭.getSheetId(), rowIndex: 0, columnIndex: 0 },
      fields: 'pivotTable',
      rows: [{
        values: [{
          pivotTable: {
            source: {
              sheetId: src.getSheetId(), startRowIndex: 0, endRowIndex: src.getLastRow(),
              startColumnIndex: 0, endColumnIndex: src.getLastColumn(),
            },
            rows: [{ sourceColumnOffset: _열(src, '설비태그') - 1, showTotals: true, sortOrder: 'ASCENDING' }],
            columns: [{ sourceColumnOffset: _열(src, '항목명') - 1, showTotals: true, sortOrder: 'ASCENDING' }],
            values: [{ summarizeFunction: 'COUNTA', sourceColumnOffset: _열(src, '판정') - 1 }],
            criteria: { [_열(src, '판정') - 1]: { visibleValues: ['이상'] } },
          },
        }],
      }],
    },
  };
  Sheets.Spreadsheets.batchUpdate({ requests: [요청] }, ss.getId());
  _로그('피벗 탭 생성: 행 설비 × 열 항목 (판정=이상)');
  return true;
}

/* ============================ 점검대장 .xlsx 내보내기 ============================ */

function _표지시트(ss, 기간시작, 기간종료) {
  const sh = ss.getSheetByName('표지') || ss.insertSheet('표지', 0);
  sh.clear();
  sh.getRange(1, 1).setValue(String(설정('회사명'))).setFontSize(12).setFontColor('#666666');
  sh.getRange(3, 1).setValue('설비 점검 대장').setFontSize(20).setFontWeight('bold').setFontColor('#1F3864');
  sh.getRange(4, 1).setValue('집계기간   ' + 기간시작 + ' ~ ' + 기간종료).setFontSize(11);
  sh.getRange(5, 1).setValue('작성일     ' + _날짜(new Date())).setFontSize(11);
  sh.getRange(6, 1).setValue('부서       ' + String(설정('부서명'))).setFontSize(11);
  sh.getRange(9, 1, 1, 3).setValues([['담당', '검토', '승인']])
    .setBackground('#F1F3F4').setFontColor('#666666').setHorizontalAlignment('center');
  sh.getRange(10, 1, 1, 3).setValues([['', '', '']]);
  sh.setRowHeight(10, 52);
  sh.getRange(12, 1).setValue('※ ' + String(설정('합성데이터고지'))).setFontColor('#999999').setFontSize(9);
  sh.setColumnWidth(1, 120);
  [2, 3].forEach((c) => sh.setColumnWidth(c, 120));
  return sh;
}

/** 판정=이상 행만 값으로 복사(수식 없음 → 어떤 엑셀에서도 열림) */
function _이상목록시트(ss) {
  const v = _점검값().slice();
  const h = v.shift().map(String);
  const L = h.indexOf('판정');
  const 이상 = v.filter((r) => String(r[L]).trim() === '이상');
  const sh = ss.getSheetByName('이상목록') || ss.insertSheet('이상목록');
  sh.clear();
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#C00000').setFontColor('#FFFFFF');
  if (이상.length) {
    sh.getRange(2, 1, 이상.length, h.length).setValues(이상);
    const N = h.indexOf('조치상태');
    if (N >= 0) {
      sh.setConditionalFormatRules([
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('미조치')
          .setBold(true).setFontColor('#C00000').setRanges([sh.getRange(2, N + 1, 이상.length, 1)]).build(),
      ]);
    }
    sh.setFrozenRows(1);
  }
  _로그('이상목록: ' + 이상.length + '행 (값 복사)');
  return sh;
}

function 대장내보내기() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 시작시각 = Date.now();
  const 기간시작 = _날짜문자열(설정('대장기간_시작'));
  const 기간종료 = _날짜문자열(설정('대장기간_종료'));
  const 이름 = '점검대장_' + (기간종료 ? 기간종료.slice(0, 7) : _날짜(new Date()).slice(0, 7));
  const 출력 = DriveApp.getFolderById(String(설정('출력폴더ID')));

  let 임시ID = null;
  let 결과 = null;
  let 데이터수 = 0;
  try {
    임시ID = DriveApp.getFileById(ss.getId()).makeCopy('임시_' + 이름).getId();
    const 임시 = SpreadsheetApp.openById(임시ID);

    // 요약 탭을 값으로 굳힌다 — 아래에서 원본 점검기록 탭을 지우므로
    // 수식(COUNTIFS)이 남아 있으면 내보낸 파일에서 #REF! 가 된다
    [SH.요약설비, SH.요약항목].forEach((n) => {
      const s = 임시.getSheetByName(n);
      if (!s || s.getLastRow() < 2) return;
      const 범위 = s.getDataRange();
      범위.setValues(범위.getValues());
    });

    // 내보내기에 필요 없는 탭 제거
    [SH.안내, SH.설정, SH.발행, SH.측정, SH.계획, SH.피벗, SH.점검].forEach((n) => {
      const s = 임시.getSheetByName(n);
      if (s) 임시.deleteSheet(s);
    });

    // 점검대장: 1~3행 제목 블록, 5행 머리글, 6행부터 데이터
    const h = _헤더(_시트(SH.점검));
    const 값 = _점검값().slice();
    값.shift();
    데이터수 = 값.length;
    const 대장 = 임시.insertSheet('점검대장', 0);
    대장.getRange(1, 1).setValue(String(설정('회사명')) + ' ' + String(설정('부서명')) + ' 설비 점검 대장')
      .setFontSize(14).setFontWeight('bold').setFontColor('#1F3864');
    대장.getRange(2, 1).setValue('집계기간 ' + 기간시작 + ' ~ ' + 기간종료 + ' · 작성 ' + _날짜(new Date()));
    대장.getRange(3, 1).setValue('※ ' + String(설정('합성데이터고지'))).setFontColor('#999999').setFontSize(9);
    대장.getRange(1, 13, 1, 3).setValues([['담당', '검토', '승인']])
      .setBackground('#F1F3F4').setFontColor('#666666').setHorizontalAlignment('center');
    대장.getRange(2, 13, 1, 3).setValues([['', '', '']]);
    대장.setRowHeight(2, 40);
    대장.getRange(5, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
    if (값.length) 대장.getRange(6, 1, 값.length, h.length).setValues(값);
    대장서식적용(대장, 5);

    _표지시트(임시, 기간시작, 기간종료);
    _이상목록시트(임시);

    // 수식 계산을 기다린 뒤 변환
    SpreadsheetApp.flush();
    Utilities.sleep(1500);

    const 파일 = DriveApp.getFileById(임시ID);
    const xlsx = 출력.createFile(파일.getAs(MimeType.MICROSOFT_EXCEL).setName(이름 + '.xlsx'));
    _안전('대장 PDF', () => 출력.createFile(파일.getAs('application/pdf').setName(이름 + '.pdf')));
    결과 = xlsx.getUrl();
    _로그('xlsx 저장: ' + 결과);
  } finally {
    if (임시ID) _안전('임시본 정리', () => DriveApp.getFileById(임시ID).setTrashed(true));
  }

  _측정로그('자동', '점검대장 내보내기', 데이터수, (Date.now() - 시작시각) / 1000, '자동', 이름 + '.xlsx');
  return 결과;
}

/** 매월 1일 트리거용(영문명) */
function monthlyLedger() {
  return 대장내보내기();
}

/* ============================ 측정 ============================ */

function _측정로그(구분, 작업, 건수, 소요초, 방식, 비고) {
  const sh = _시트(SH.측정);
  const 초 = Number(소요초) || 0;
  const n = Number(건수) || 0;
  sh.appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    구분, 작업, n, 초.toFixed(1), n ? (초 / n).toFixed(2) : '', 방식 || '자동', 비고 || '',
  ]);
  return 초;
}

/* ============================ 발행이력 ============================ */

/**
 * 발행이력 한 줄 기록. 같은 (문서종류, 대상ID)가 있으면 그 줄을 갱신한다.
 * 넘기지 않은 칸은 이전 값을 그대로 둔다 — 상태만 '오류'로 바꿀 때
 * 문서 링크(구글독스ID·PDF링크)가 지워지면 어떤 문서가 문제인지 추적할 수 없다.
 */
function _발행이력기록(종류, 대상ID, o) {
  const sh = _시트(SH.발행);
  const h = _헤더(sh);
  // 짧은 이름으로 넘겨도 헤더 이름으로 맞춰 준다
  const 별칭 = { 독스ID: '구글독스ID', pdf: 'PDF링크', docx: 'DOCX링크' };
  const 값 = {
    발행일시: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    문서종류: 종류,
    대상ID: 대상ID,
  };
  Object.keys(o || {}).forEach((k) => {
    값[별칭[k] || k] = o[k];
  });

  const 기존 = sh.getDataRange().getValues();
  const iID = h.indexOf('대상ID');
  const i종 = h.indexOf('문서종류');
  let 행번호 = 0;
  for (let i = 1; i < 기존.length; i++) {
    if (String(기존[i][iID]).trim() === 대상ID && String(기존[i][i종]).trim() === 종류) {
      행번호 = i + 1;
      break;
    }
  }

  const 이전 = 행번호 ? 기존[행번호 - 1] : null;
  const 행 = h.map((k, j) => {
    if (k in 값) return 값[k];
    return 이전 ? 이전[j] : '';
  });

  if (행번호) sh.getRange(행번호, 1, 1, h.length).setValues([행]);
  else sh.appendRow(행);
}

function _발행이력완료행(종류) {
  const sh = _시트(SH.발행);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const h = v.shift().map(String);
  const i종 = h.indexOf('문서종류');
  const i상 = h.indexOf('상태');
  return v
    .filter((r) => String(r[i종]).trim() === 종류 && String(r[i상]).trim() === '완료')
    .map((r) => _행객체(h, r));
}

function _발행이력완료집합(종류) {
  const 집합 = {};
  _발행이력완료행(종류).forEach((r) => (집합[String(r.대상ID).trim()] = true));
  return 집합;
}


/* ===================== 양식.js ===================== */

/**
 * 실무 서류 자동 생성 세트 — 문서(양식) 부분
 * ------------------------------------------------------------
 * 작업계획서 / 점검결과보고서를 구글독스로 만들고 PDF·DOCX 로 내보낸다.
 * 템플릿 문서를 미리 만들지 않는다 — 서식을 코드로 그린다.
 *   · 본문 폭: A4(595pt) - 좌우 여백 2cm(56.7pt×2) ≈ 481pt
 *   · 표 열 너비 합계를 480 근처로 맞춰야 A4 1장에 들어간다
 *   · 서식은 실패해도 내용은 남도록 구간마다 _서식시도로 감싼다
 */

const 보고서양식 = {
  글꼴: 'Noto Sans KR',
  여백: 56.7,
  진남: '#1F3864',
  회색: '#666666',
  본문: '#202124',
  라벨배경: '#F1F3F4',
  빨강: '#C00000',
  음영: '#F8F9FA',
  선: '#DADCE0',
  정상배경: '#E2EFDA',
  정상글: '#274E13',
  바닥글: '#999999',
  심각도배경: { 상: '#F4CCCC', 중: '#FCE5CD', 하: '#EFEFEF' },
  최대상세행: 15,
};

/* ============================ 문서 서식 도구 ============================ */

/** 서식 한 구간 실행. 실패해도 문서 생성은 계속한다 */
function _서식시도(이름, 작업) {
  try {
    작업();
  } catch (e) {
    Logger.log('[서식 건너뜀] ' + 이름 + ': ' + e.message);
  }
}

/** 문단·셀 글자 서식. 글꼴·크기는 따로 감싸 하나가 실패해도 굵기·색은 들어가게 한다 */
function _글자(el, 크기, 굵게, 색) {
  _서식시도('글자', () => {
    const t = el.editAsText();
    try { t.setFontFamily(보고서양식.글꼴); } catch (e) {}
    try { t.setFontSize(크기); } catch (e) { t.setFontSize(Math.round(크기)); }
    t.setBold(!!굵게);
    t.setForegroundColor(색 || 보고서양식.본문);
  });
}

function _문단간격(p, 앞, 뒤) {
  _서식시도('문단 간격', () => p.setSpacingBefore(앞).setSpacingAfter(뒤).setLineSpacing(1.15));
}

/** 표와 표 사이를 원하는 높이로 띄우는 빈 문단(기본 줄 높이를 없애려고 글자 크기를 1로) */
function _간격(body, pt) {
  const p = body.appendParagraph('');
  _서식시도('간격', () => {
    p.setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 1 });
    p.setSpacingBefore(0).setSpacingAfter(pt).setLineSpacing(1);
  });
  return p;
}

function _셀문단(cell) {
  return cell.getChild(0).asParagraph();
}

/** 표를 붙이고 공통 서식(테두리·열 너비·안쪽 여백)을 준다 */
function _표(body, 행들, 열너비, 안쪽, 테두리색) {
  const t = body.appendTable(행들.map((r) => r.map((x) => String(x === null || x === undefined ? '' : x))));
  const 여백 = 안쪽 || 4;
  const 셀마다 = (fn) => {
    for (let r = 0; r < t.getNumRows(); r++) {
      const row = t.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) fn(row.getCell(c));
    }
  };
  _서식시도('표 테두리', () => { t.setBorderColor(테두리색 || 보고서양식.선); t.setBorderWidth(0.75); });
  _서식시도('표 열 너비', () => 열너비.forEach((w, i) => t.setColumnWidth(i, w)));
  _서식시도('셀 여백', () => 셀마다((cell) => cell.setPaddingTop(여백).setPaddingBottom(여백).setPaddingLeft(6).setPaddingRight(6)));
  _서식시도('셀 세로 정렬', () => 셀마다((cell) => cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER)));
  _서식시도('셀 문단 간격', () => 셀마다((cell) => _셀문단(cell).setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1.1)));
  return t;
}

/** 셀 하나 서식. o = { 크기, 굵게, 색, 배경, 가운데 } */
function _셀(cell, o) {
  _글자(cell, o.크기, o.굵게, o.색);
  if (o.배경) _서식시도('셀 배경', () => cell.setBackgroundColor(o.배경));
  if (o.가운데) _서식시도('셀 가로 정렬', () => _셀문단(cell).setAlignment(DocumentApp.HorizontalAlignment.CENTER));
}

/** 머리글 아래 굵은 줄. 문단 테두리는 못 쓰니 색 채운 1칸 표로 대신한다 */
function _구분선(body, 색) {
  const t = body.appendTable([[' ']]);
  const cell = t.getCell(0, 0);
  _서식시도('구분선 색', () => { t.setBorderColor(색); cell.setBackgroundColor(색); });
  _서식시도('구분선 두께', () => {
    cell.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
    _셀문단(cell).setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    cell.editAsText().setFontSize(2);
  });
  _서식시도('구분선 테두리', () => t.setBorderWidth(0));
  return t;
}

/** 새 문서의 빈 첫 문단을 거의 안 보이게 줄인다(합본에서는 건드리지 않는다) */
function _첫문단정리(body) {
  _서식시도('여백', () => body.setMarginTop(보고서양식.여백).setMarginBottom(보고서양식.여백)
    .setMarginLeft(보고서양식.여백).setMarginRight(보고서양식.여백));
  if (body.getNumChildren() !== 1) return null; // 합본: 이미 내용이 있음
  const p = body.getChild(0).asParagraph();
  if (!p.getText().trim()) {
    _서식시도('첫 문단', () => {
      p.setAttributes({ [DocumentApp.Attribute.FONT_SIZE]: 1 });
      p.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    });
  }
  return p;
}

function _소제목(body, 글) {
  const p = body.appendParagraph('▶ ' + 글);
  _글자(p, 13, true, 보고서양식.진남);
  _문단간격(p, 14, 4);
  return p;
}

function _문서출처() {
  let 이름 = '서류자동화_포트폴리오';
  try { 이름 = SpreadsheetApp.getActiveSpreadsheet().getName() || 이름; } catch (e) {}
  return 이름;
}

/** 문서 바닥글. 바닥글 영역이 안 되면 본문 끝에 붙인다 */
function _바닥글(doc, 문구) {
  let 바닥 = null;
  try {
    const f = doc.getFooter() || doc.addFooter();
    바닥 = f.getParagraphs()[0] || f.appendParagraph('');
    바닥.setText(문구);
  } catch (e) {
    Logger.log('[서식 건너뜀] 바닥글 영역: ' + e.message);
    바닥 = null;
  }
  if (!바닥) {
    바닥 = doc.getBody().appendParagraph(문구);
    _문단간격(바닥, 18, 0);
  }
  _글자(바닥, 8, false, 보고서양식.바닥글);
  _서식시도('바닥글 정렬', () => 바닥.setAlignment(DocumentApp.HorizontalAlignment.RIGHT));
  return 바닥;
}

/** 제목 + 결재 3칸(담당·검토·승인). 부제 줄에 문서번호·기간을 넣는다 */
function _제목결재표(body, 제목, 부제) {
  const S = 보고서양식;
  const t = _표(body, [[제목, '담당', '검토', '승인'], [부제, '', '', '']], [281, 66, 66, 66], 6);
  _셀(t.getCell(0, 0), { 크기: 20, 굵게: true, 색: S.진남 });
  for (let c = 1; c <= 3; c++) {
    _셀(t.getCell(0, c), { 크기: 9, 색: S.회색, 배경: S.라벨배경, 가운데: true });
    _셀(t.getCell(1, c), { 크기: 10, 가운데: true });
    _서식시도('서명칸 높이', () => t.getCell(1, c).setPaddingTop(18).setPaddingBottom(18));
  }
  _셀(t.getCell(1, 0), { 크기: 10, 색: S.회색 });
  return t;
}

/**
 * 라벨/값 격자. 행들 = [[라벨, 값, 라벨, 값], ...]
 * 라벨열 = 회색 배경을 줄 열 번호(기본: 짝수 열 전부)
 */
function _라벨값표(body, 행들, 너비, 라벨열) {
  const t = _표(body, 행들, 너비, 6);
  const 라벨 = 라벨열 && 라벨열.length ? 라벨열 : 너비.map((_, i) => i).filter((i) => i % 2 === 0);
  라벨.forEach((c) => {
    if (c + 1 >= 너비.length) return;
    for (let r = 0; r < t.getNumRows(); r++) {
      _셀(t.getCell(r, c), { 크기: 9, 색: 보고서양식.회색, 배경: 보고서양식.라벨배경 });
      _셀(t.getCell(r, c + 1), { 크기: 10, 색: 보고서양식.본문 });
    }
  });
  return t;
}

/** 서명란: 윗줄 라벨, 아랫줄 빈칸 */
function _서명표(body, 라벨들, 너비) {
  const S = 보고서양식;
  const t = _표(body, [라벨들, 라벨들.map(() => '')], 너비, 8);
  라벨들.forEach((_, c) => {
    _셀(t.getCell(0, c), { 크기: 9, 색: S.회색, 배경: S.라벨배경, 가운데: true });
    _셀(t.getCell(1, c), { 크기: 10, 가운데: true });
    _서식시도('서명칸 높이', () => t.getCell(1, c).setPaddingTop(16).setPaddingBottom(16));
  });
  return t;
}

/** 요약 카드 4칸 */
function _요약카드(body, 항목들) {
  const S = 보고서양식;
  const t = _표(body, [항목들.map((x) => x[0]), 항목들.map((x) => x[1])], [120, 120, 120, 120], 7);
  항목들.forEach((x, c) => {
    _셀(t.getCell(0, c), { 크기: 9, 색: S.회색, 배경: S.라벨배경, 가운데: true });
    _셀(t.getCell(1, c), { 크기: 14, 굵게: true, 색: x[2] ? S.빨강 : S.본문, 가운데: true });
  });
  return t;
}

/* ============================ 데이터 읽기 ============================ */

/** 기간 내 해당 설비의 이상 목록(심각도 → 최근일 순) */
function _이상상세(시작, 종료, 태그) {
  const v = _점검값().slice();
  const h = v.shift().map(String);
  const i = (k) => h.indexOf(k);
  const 값 = (r, k) => (i(k) < 0 || r[i(k)] === null ? '' : String(r[i(k)]).trim());
  const 순서 = { 상: 0, 중: 1, 하: 2 };

  const 목록 = [];
  v.forEach((r) => {
    if (값(r, '판정') !== '이상') return;
    if (태그 && 값(r, '설비태그') !== 태그) return;
    const d = _날짜문자열(r[i('점검일')]);
    if (!d || d < 시작 || d > 종료) return;
    목록.push({
      점검일: d,
      설비: 값(r, '설비태그'),
      항목: 값(r, '항목명'),
      측정값: 값(r, '측정값'),
      기준: 값(r, '기준'),
      심각도: 값(r, '심각도'),
      조치상태: 값(r, '조치상태') || '미조치',
    });
  });

  목록.sort((a, b) => {
    const 등급 = (s) => (s in 순서 ? 순서[s] : 9);
    return 등급(a.심각도) - 등급(b.심각도) || (a.점검일 < b.점검일 ? 1 : a.점검일 > b.점검일 ? -1 : 0);
  });

  const 분포 = { 상: 0, 중: 0, 하: 0 };
  목록.forEach((x) => {
    if (x.심각도 in 분포) 분포[x.심각도]++;
  });

  return { 행: 목록.slice(0, 보고서양식.최대상세행), 전체: 목록.length, 분포 };
}

/** 설비 1대의 기간 내 통계 */
function _설비통계(태그, 시작, 종료) {
  const v = _점검값().slice();
  const h = v.shift().map(String);
  const i = (k) => h.indexOf(k);
  const 값 = (r, k) => (i(k) < 0 || r[i(k)] === null ? '' : String(r[i(k)]).trim());

  const 행들 = v.filter((r) => 값(r, '설비태그') === 태그);
  const 기간내 = 행들.filter((r) => {
    const d = _날짜문자열(r[i('점검일')]);
    return d && d >= 시작 && d <= 종료;
  });
  // '-' 는 그 설비에 해당 없는 항목(예: 펌프의 탱크 액위) — 이상률 모집단에서 뺀다
  const 대상 = 기간내.filter((r) => 값(r, '측정값') !== '-');
  const 이상 = 기간내.filter((r) => 값(r, '판정') === '이상');

  const 날짜모음 = {};
  기간내.forEach((r) => {
    날짜모음[_날짜문자열(r[i('점검일')])] = true;
  });

  const 항목별 = {};
  const 분포 = { 상: 0, 중: 0, 하: 0 };
  const 조치 = { 완료: 0, 진행중: 0, 미조치: 0 };
  이상.forEach((r) => {
    const a = 값(r, '항목명');
    if (!항목별[a]) 항목별[a] = { 이상: 0, 상: 0, 중: 0, 하: 0, 미조치: 0, 최근: '' };
    const o = 항목별[a];
    o.이상++;
    const s = 값(r, '심각도');
    if (s in 분포) {
      분포[s]++;
      o[s]++;
    }
    const st = 값(r, '조치상태') || '미조치';
    if (st in 조치) 조치[st]++;
    if (st === '미조치') o.미조치++;
    const d = _날짜문자열(r[i('점검일')]);
    if (d > o.최근) o.최근 = d;
  });

  return {
    태그,
    점검횟수: Object.keys(날짜모음).length,
    판정건수: 대상.length,
    이상건수: 이상.length,
    이상률: 대상.length ? (이상.length / 대상.length) * 100 : 0,
    미조치: 조치.미조치,
    조치, 항목별, 분포,
  };
}

function _종합의견(t) {
  const r = t.이상률;
  let s = r < 5
    ? '전체 점검 대비 기준 초과가 낮아 상태는 양호합니다.'
    : r < 15
      ? '기준 초과가 반복되는 항목이 있어 점검 주기 단축을 검토하십시오.'
      : '기준 초과 비율이 높습니다. 원인 분석과 작업계획 수립을 권고합니다.';
  const 많은 = Object.keys(t.항목별).sort((a, b) => t.항목별[b].이상 - t.항목별[a].이상)[0];
  if (많은) s += ' 최다 이상 항목은 ' + 많은 + '(' + t.항목별[많은].이상 + '건)입니다.';
  if (t.미조치 > 0) s += ' 미조치 ' + t.미조치 + '건은 작업계획서 발행 대상입니다.';
  return s;
}

/* ============================ 대상 조회 ============================ */

function _대상목록(종류) {
  if (종류 === '작업계획서') {
    return _시트(SH.계획).getDataRange().getValues().slice(1)
      .filter((r) => String(r[0]).trim())
      .map((r) => ({ ID: String(r[0]).trim() }));
  }
  return _시트(SH.설비).getDataRange().getValues().slice(1)
    .filter((r) => String(r[0]).trim())
    .map((r) => ({ ID: String(r[0]).trim() }));
}

function _대상행(종류, ID) {
  if (종류 === '작업계획서') {
    const v = _시트(SH.계획).getDataRange().getValues();
    const h = v.shift().map(String);
    const r = v.filter((x) => String(x[0]).trim() === ID)[0];
    return r ? _행객체(h, r) : null;
  }
  const v = _시트(SH.설비).getDataRange().getValues();
  const h = v.shift().map(String);
  const i = v.findIndex((x) => String(x[0]).trim() === ID);
  if (i < 0) return null;
  const o = _행객체(h, v[i]);
  o._번호 = _문서번호(String(설정('문서번호접두어_결과보고') || 'IR'), i + 1);
  return o;
}

/* ============================ 문서 작성 ============================ */

/** 작업계획서 A4 1장 */
function _작업계획서작성(body, 행) {
  const S = 보고서양식;
  _첫문단정리(body);

  _제목결재표(body, '작 업 계 획 서', 행.계획ID + '  ·  작성일 ' + _날짜(new Date()));
  const 회사 = body.appendParagraph(String(설정('회사명')) + '  ' + String(설정('부서명')));
  _글자(회사, 10, false, S.회색);
  _문단간격(회사, 6, 4);
  _구분선(body, S.진남);
  _간격(body, 8);

  _소제목(body, '1. 기본 정보');
  _라벨값표(body, [
    ['작업명', 행.작업명, '작업유형', 행.작업유형],
    ['설비', 행.설비태그 + ' ' + 행.설비명, '구역', 행.구역],
    ['작업일', _날짜문자열(행.작업일) + ' ' + _요일(행.작업일), '예상 소요', 행['예상소요(h)'] + ' h'],
    ['담당자', 행.담당자, '작업조', 행.작업조인원 + ' 명'],
    ['결재자', 행.결재자, '상태', 행.상태],
  ], [80, 160, 80, 160], [0, 2]);
  _간격(body, 6);

  _소제목(body, '2. 작업 내용');
  _표(body, [[행.작업내용]], [480], 6);
  _간격(body, 6);

  _소제목(body, '3. 위험요인 및 안전조치');
  const 위험 = _표(body, [['위험요인', 행.위험요인], ['안전조치', 행.안전조치]], [80, 400], 6);
  _셀(위험.getCell(0, 0), { 크기: 9, 색: S.회색, 배경: S.라벨배경 });
  _셀(위험.getCell(1, 0), { 크기: 9, 색: S.회색, 배경: S.라벨배경 });
  _셀(위험.getCell(0, 1), { 크기: 10, 굵게: true, 색: S.빨강 });
  _셀(위험.getCell(1, 1), { 크기: 10, 색: S.본문 });
  _간격(body, 6);

  _소제목(body, '4. 필요 자재·공구');
  _표(body, [[행.필요자재]], [480], 6);
  _간격(body, 6);

  _소제목(body, '5. 확인');
  _라벨값표(body, [
    ['작업자 서명', '', '안전관리자', ''],
    ['작업 완료일', '20      .      .', '비고', ''],
  ], [80, 160, 80, 160], [0, 2]);

  return body;
}

/** 설비 점검 결과 보고서 A4 1~2장 */
function _점검결과보고서작성(body, 설비행) {
  const S = 보고서양식;
  const 태그 = String(설비행.설비태그).trim();
  const 기간시작 = _날짜문자열(설정('대장기간_시작'));
  const 기간종료 = _날짜문자열(설정('대장기간_종료'));
  const t = _설비통계(태그, 기간시작, 기간종료);
  const 상세 = _이상상세(기간시작, 기간종료, 태그);

  _첫문단정리(body);
  _제목결재표(body, '설비 점검 결과 보고서',
    설비행._번호 + '  ·  집계기간 ' + 기간시작 + ' ~ ' + 기간종료);
  const 회사 = body.appendParagraph(String(설정('회사명')) + '  ' + String(설정('부서명')));
  _글자(회사, 10, false, S.회색);
  _문단간격(회사, 6, 4);
  _구분선(body, S.진남);
  _간격(body, 8);

  // 1. 설비 정보
  _소제목(body, '1. 설비 정보');
  _라벨값표(body, [
    ['설비태그', 태그, '설비명', 설비행.설비명],
    ['설비종류', 설비행.설비종류, '구역', 설비행.구역],
    ['담당자', 설비행.담당자, '담당부서', 설비행.담당부서],
    ['설치연도', 설비행.설치연도, '점검주기', 설비행['점검주기(일)'] + ' 일'],
  ], [80, 160, 80, 160], [0, 2]);
  _간격(body, 6);

  // 2. 요약 카드
  _소제목(body, '2. 점검 요약');
  _요약카드(body, [
    ['점검 횟수', String(t.점검횟수) + '회', false],
    ['이상 건수', String(t.이상건수) + '건', t.이상건수 > 0],
    ['이상률', t.이상률.toFixed(1) + '%', t.이상률 >= 15],
    ['미조치', String(t.미조치) + '건', t.미조치 > 0],
  ]);
  _간격(body, 6);

  // 3. 항목별 이상 건수
  _소제목(body, '3. 항목별 이상 건수');
  const 항목머리 = ['점검항목', '이상', '상', '중', '하', '미조치', '최근 발생일'];
  const 항목표 = _표(body, [항목머리].concat(항목목록.map((a) => {
    const o = t.항목별[a] || { 이상: 0, 상: 0, 중: 0, 하: 0, 미조치: 0, 최근: '' };
    return [a, o.이상, o.상, o.중, o.하, o.미조치, o.최근 || '—'];
  })), [110, 50, 40, 40, 40, 60, 140], 4);
  항목머리.forEach((_, c) => _셀(항목표.getCell(0, c), { 크기: 9.5, 굵게: true, 색: '#FFFFFF', 배경: S.진남, 가운데: true }));
  항목목록.forEach((a, i) => {
    const o = t.항목별[a];
    항목머리.forEach((_, c) => {
      const 옵션 = { 크기: 9.5, 색: o ? S.본문 : '#BBBBBB', 가운데: c > 0 };
      if (c === 1 && o && o.이상 > 0) 옵션.굵게 = true;
      if (c === 5 && o && o.미조치 > 0) 옵션.색 = S.빨강;
      _셀(항목표.getCell(i + 1, c), 옵션);
    });
  });
  _간격(body, 6);

  // 4. 이상 상세
  _소제목(body, '4. 이상 상세');
  if (!상세.전체) {
    const 안내 = _표(body, [['집계기간 중 기준 초과 항목이 없습니다.']], [480], 10, S.정상배경);
    _셀(안내.getCell(0, 0), { 크기: 10, 색: S.정상글, 배경: S.정상배경, 가운데: true });
  } else {
    const 머리 = ['점검일', '항목', '측정값', '기준', '심각도', '조치상태'];
    const 가운데 = [0, 4, 5];
    const 표 = _표(body, [머리].concat(상세.행.map((x) => [x.점검일, x.항목, x.측정값, x.기준, x.심각도, x.조치상태])),
      [85, 105, 80, 80, 60, 70], 4);
    머리.forEach((_, c) => _셀(표.getCell(0, c), { 크기: 9.5, 굵게: true, 색: '#FFFFFF', 배경: S.진남, 가운데: true }));
    상세.행.forEach((x, i) => {
      const r = i + 1;
      const 줄배경 = r % 2 === 0 ? S.음영 : null;
      머리.forEach((_, c) => {
        const o = { 크기: 9, 색: S.본문, 배경: 줄배경, 가운데: 가운데.indexOf(c) > -1 };
        if (c === 4) { o.배경 = S.심각도배경[x.심각도] || 줄배경; o.굵게 = x.심각도 === '상'; }
        if (c === 5 && x.조치상태 === '미조치') { o.색 = S.빨강; o.굵게 = true; }
        _셀(표.getCell(r, c), o);
      });
    });
    if (상세.전체 > 상세.행.length) {
      const 외 = body.appendParagraph('외 ' + (상세.전체 - 상세.행.length) + '건 — 점검기록 탭 참조');
      _글자(외, 9, false, S.회색);
      _문단간격(외, 4, 0);
      _서식시도('외 N건 정렬', () => 외.setAlignment(DocumentApp.HorizontalAlignment.RIGHT));
    }
  }
  _간격(body, 6);

  // 5. 조치 현황
  _소제목(body, '5. 조치 현황');
  const 조치표 = _표(body, [['완료', String(t.조치.완료) + '건', '진행중', String(t.조치.진행중) + '건', '미조치', String(t.조치.미조치) + '건']],
    [60, 100, 60, 100, 60, 100], 6);
  [0, 2, 4].forEach((c) => _셀(조치표.getCell(0, c), { 크기: 9, 색: S.회색, 배경: S.라벨배경, 가운데: true }));
  [1, 3, 5].forEach((c) => _셀(조치표.getCell(0, c), { 크기: 11, 굵게: true, 가운데: true }));
  _간격(body, 6);

  // 6. 종합 의견
  _소제목(body, '6. 종합 의견');
  const 의견 = _표(body, [[_종합의견(t)]], [480], 8);
  _셀(의견.getCell(0, 0), { 크기: 10, 색: S.본문 });

  // 7. 확인
  _소제목(body, '7. 확인');
  _서명표(body, ['작성자', '검토자', '승인자'], [160, 160, 160]);

  return body;
}

/* ============================ 저장·발행 ============================ */

/**
 * 문서를 구글독스 폴더에 보관하고 출력 폴더에 PDF·DOCX 를 만든다.
 * 문서를 코드로 그리므로 템플릿 사본이 필요 없고, 원본 독스는 재편집·검증용으로 남긴다.
 */
function _문서저장(doc, 종류, 파일명, 옵션) {
  const o = 옵션 || {};
  doc.saveAndClose();
  const 파일 = DriveApp.getFileById(doc.getId());
  const 출력 = DriveApp.getFolderById(String(설정('출력폴더ID')));
  const 독스폴더 = DriveApp.getFolderById(String(설정('구글독스폴더ID')));

  _안전('독스 보관', () => {
    독스폴더.addFile(파일);
    DriveApp.getRootFolder().removeFile(파일);
  });

  const pdf파일 = 출력.createFile(파일.getAs('application/pdf').setName(파일명 + '.pdf'));
  let docx링크 = '';
  if (o.docx !== false) {
    _안전('docx 변환', () => {
      const docx = 파일.getAs(MimeType.MICROSOFT_WORD).setName(파일명 + '.docx');
      docx링크 = 출력.createFile(docx).getUrl();
    });
  }
  return { 독스ID: doc.getId(), 독스링크: 파일.getUrl(), pdf: pdf파일.getUrl(), docx: docx링크, 파일명 };
}

/** 문서 1건 생성. 실패하면 임시 문서를 정리하고 오류를 그대로 올린다 */
function 문서한건(종류, 대상ID) {
  const 시작 = Date.now();
  const 행 = _대상행(종류, 대상ID);
  if (!행) throw new Error('대상이 없습니다: ' + 종류 + ' / ' + 대상ID);

  const doc = DocumentApp.create('임시_' + 대상ID);
  let 저장 = null;
  let 저장완료 = false;
  try {
    const body = doc.getBody();
    let 파일명;
    if (종류 === '작업계획서') {
      _작업계획서작성(body, 행);
      파일명 = 행.계획ID + '_' + 행.설비태그 + '_작업계획서';
    } else {
      _점검결과보고서작성(body, 행);
      파일명 = 행._번호 + '_' + 행.설비태그 + '_점검결과보고서';
    }
    _바닥글(doc, '생성 ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') + ' · ' + _문서출처() + ' · 데이터: 합성(가상)');
    저장 = _문서저장(doc, 종류, 파일명);
    저장완료 = true;
  } catch (e) {
    if (!저장완료) _안전('임시 문서 정리', () => DriveApp.getFileById(doc.getId()).setTrashed(true));
    throw e;
  }

  const 소요초 = (Date.now() - 시작) / 1000;
  _안전('발행이력', () => _발행이력기록(종류, 대상ID, {
    문서명: 저장.파일명, 독스ID: 저장.독스ID, pdf: 저장.pdf, docx: 저장.docx,
    소요초: 소요초.toFixed(1), 상태: '완료',
  }));
  return { 파일명: 저장.파일명, 소요초, pdf: 저장.pdf };
}

/**
 * 미발행 건을 한 번에 N건씩 만든다.
 * 시간 제한(설정 시간제한초)에 가까워지면 남은 건은 다음 실행으로 넘긴다 —
 * 6분 제한에 걸리면 그때까지 만든 문서가 통째로 날아가기 때문.
 */
function 양식_일괄생성(종류) {
  const 시작 = Date.now();
  const 전체 = _대상목록(종류);
  const 완료 = _발행이력완료집합(종류);
  const 미발행 = 전체.filter((d) => !완료[d.ID]);
  const 한계 = Number(설정('시간제한초')) || 270;
  const 한번에 = Number(설정('문서한번에')) || 8;

  _로그(종류 + ': 전체 ' + 전체.length + '건 / 이미 발행 ' + (전체.length - 미발행.length) + '건 / 이번 대상 ' + Math.min(미발행.length, 한번에) + '건');

  const 성공 = [];
  const 실패 = [];
  let 중단 = false;
  미발행.slice(0, 한번에).forEach((d) => {
    if ((Date.now() - 시작) / 1000 > 한계) {
      중단 = true;
      return;
    }
    try {
      문서한건(종류, d.ID);
      성공.push(d.ID);
      _로그('  완료 ' + d.ID);
    } catch (e) {
      실패.push(d.ID);
      _안전('발행이력(오류)', () => _발행이력기록(종류, d.ID, { 상태: '오류', 오류: e.message }));
      _로그('  실패 ' + d.ID + ': ' + e.message);
    }
  });

  const 소요 = _측정로그('자동', 종류 + ' 생성', 성공.length, (Date.now() - 시작) / 1000, '자동',
    '성공 ' + 성공.length + ' / 실패 ' + 실패.length);
  const 남음 = 미발행.length - (성공.length + 실패.length);
  _로그('결과: 성공 ' + 성공.length + '건 / 실패 ' + 실패.length + '건 / 소요 ' + 소요.toFixed(1) + '초' +
    (남음 > 0 ? ' / 남은 ' + 남음 + '건은 한 번 더 실행하세요' + (중단 ? ' (시간 제한)' : '') : ''));
  return { 성공: 성공, 실패 };
}

/** 발행한 문서들을 한 문서로 이어 붙여 합본 PDF 1개로 만든다 */
function 합본PDF(종류) {
  const 시작 = Date.now();
  const 이력 = _발행이력완료행(종류);
  if (!이력.length) throw new Error('발행이력에 완료 건이 없습니다. 양식_일괄생성을 먼저 실행하세요.');

  const doc = DocumentApp.create('임시_합본_' + 종류);
  let 저장 = null;
  try {
    const body = doc.getBody();
    let 건수 = 0;
    이력.forEach((h) => {
      const 행 = _대상행(종류, String(h.대상ID).trim());
      if (!행) return;
      if (건수 > 0) body.appendPageBreak();
      if (종류 === '작업계획서') _작업계획서작성(body, 행);
      else _점검결과보고서작성(body, 행);
      건수++;
    });
    if (!건수) throw new Error('붙일 문서가 없습니다(대상 행을 찾지 못함)');
    _바닥글(doc, '합본 ' + 건수 + '건 · 생성 ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') +
      ' · ' + _문서출처() + ' · 데이터: 합성(가상)');
    const 파일명 = 종류 + '_합본_' + 건수 + '건_' + _날짜(new Date());
    저장 = _문서저장(doc, 종류, 파일명, { docx: false });
    저장.건수 = 건수;
  } finally {
    if (!저장) _안전('임시 문서 정리', () => DriveApp.getFileById(doc.getId()).setTrashed(true));
  }

  const 소요 = _측정로그('자동', '합본 PDF 생성', 저장.건수, (Date.now() - 시작) / 1000, '자동', 저장.파일명);
  _로그('합본 PDF: ' + 저장.건수 + '건 / ' + 저장.pdf + ' / ' + 소요.toFixed(1) + '초');
  return 저장.pdf;
}

/**
 * 원본 데이터와 문서 본문을 기계적으로 대조한다.
 * 눈으로 20부를 읽는 대신 문자열 포함 여부만 본다(누락 0건을 주장할 근거).
 */
function 검증_문서대조(종류) {
  const 이력 = _발행이력완료행(종류);
  if (!이력.length) {
    _로그('대조: 완료된 문서가 없습니다. 양식_일괄생성을 먼저 실행하세요.');
    return { 항목: 0, 누락: 0 };
  }
  let 항목수 = 0;
  let 누락수 = 0;
  const 문제 = [];

  이력.forEach((h) => {
    const 본문 = _안전('문서 읽기', () => DocumentApp.openById(String(h.구글독스ID)).getBody().getText());
    const 원본 = _대상행(종류, String(h.대상ID).trim());
    if (!본문 || !원본) return;

    const 빠진 = [];
    Object.keys(원본).forEach((k) => {
      if (k.charAt(0) === '_') return;
      let v = 원본[k];
      if (Object.prototype.toString.call(v) === '[object Date]') v = _날짜문자열(v);
      v = String(v === null || v === undefined ? '' : v).trim();
      if (!v) return;
      항목수++;
      if (본문.indexOf(v) < 0) 빠진.push(k + '=' + v);
    });

    누락수 += 빠진.length;
    if (빠진.length) {
      문제.push(h.대상ID + ': ' + 빠진.join(', '));
      _안전('발행이력(오류)', () => _발행이력기록(종류, String(h.대상ID).trim(), {
        상태: '오류', 오류: '누락 ' + 빠진.length + '건',
      }));
    }
  });

  _로그('대조 ' + 이력.length + '부 / ' + 항목수 + '항목 중 누락 ' + 누락수 + '건');
  문제.slice(0, 5).forEach((m) => _로그('  · ' + m));
  if (누락수) _로그('누락이 있으면 양식_일괄생성을 다시 실행하세요(발행이력이 오류로 바뀌어 재생성됩니다).');
  return { 항목: 항목수, 누락: 누락수 };
}


/* ===================== 설치.js ===================== */

/**
 * 실무 서류 자동 생성 세트 — 설치·샘플 데이터·메뉴
 * ------------------------------------------------------------
 * 설치_전체: 탭·설정·폴더를 만들고 수식·서식을 적용한다(두 번 돌려도 안전).
 * 설치_샘플데이터: 합성 데이터 5,040행 + 작업계획 20건을 만든다.
 *
 * 처음 한 번:  설치_전체  →  설치_샘플데이터
 * 문서 만들기: 메뉴 [서류자동화] 또는 함수 직접 실행
 */

const 점검자목록 = ['김민수', '이서연', '박지훈', '최유진', '정도현'];

/** 설비 20대 (전부 가상) */
const 설비정의 = (() => {
  const 부서 = ['기계팀', '운전팀', '안전팀'];
  const 종류별 = [['AC', '공기압축기', 5], ['PU', '펌프', 8], ['TK', '저장탱크', 4], ['BL', '송풍기', 3]];
  const 목록 = [];
  let n = 0;
  종류별.forEach(([접두, 이름, 수]) => {
    for (let i = 1; i <= 수; i++) {
      const 번호 = ('0' + i).slice(-2);
      목록.push({
        설비태그: 접두 + '-' + 번호,
        설비명: 이름 + ' ' + 번호,
        설비종류: 이름,
        구역: 접두 === 'TK' ? '탱크구역' : '유틸리티동',
        설치연도: 2015 + (n % 8),
        담당자: 점검자목록[n % 점검자목록.length],
        담당부서: 부서[n % 부서.length],
        '점검주기(일)': 1,
        비고: '',
      });
      n++;
    }
  });
  return 목록;
})();

/** 설비 종류별로 해당 있는 항목만 점검한다(없는 항목은 측정값 '-') */
const 적용항목 = {
  공기압축기: ['토출압력', '오일레벨', '진동소음', '베어링온도', '배관누설', '밸브잠금', '안전커버', '윤활급유'],
  펌프: ['토출압력', '오일레벨', '진동소음', '베어링온도', '배관누설', '밸브잠금', '안전커버', '윤활급유'],
  저장탱크: ['탱크액위', '배관누설', '밸브잠금', '안전커버', '윤활급유'],
  송풍기: ['오일레벨', '진동소음', '베어링온도', '배관누설', '밸브잠금', '안전커버', '윤활급유'],
};

const 항목정의 = {
  토출압력: { 단위: 'bar', 기준: '7.0 이하', 확률: 0.04, 심각도: '상', 정상: (r) => (5.8 + r() * 1.2).toFixed(1), 이상: (r) => (7.1 + r() * 0.7).toFixed(1) },
  오일레벨: { 단위: '', 기준: '정상', 확률: 0.05, 심각도: '중', 정상: () => '정상', 이상: () => '보충필요' },
  진동소음: { 단위: '', 기준: '정상', 확률: 0.07, 심각도: null, 정상: () => '정상', 이상: (r) => (r() < 0.6 ? '주의' : '이상') },
  베어링온도: { 단위: '℃', 기준: '70 이하', 확률: 0.05, 심각도: '상', 정상: (r) => (45 + r() * 25).toFixed(0), 이상: (r) => (71 + r() * 21).toFixed(0) },
  탱크액위: { 단위: '%', 기준: '20 이상', 확률: 0.06, 심각도: '중', 정상: (r) => (22 + r() * 70).toFixed(0), 이상: (r) => (5 + r() * 14).toFixed(0) },
  배관누설: { 단위: '', 기준: '없음', 확률: 0.03, 심각도: '상', 정상: () => '없음', 이상: () => '있음' },
  밸브잠금: { 단위: '', 기준: '정상', 확률: 0.04, 심각도: '상', 정상: () => '정상', 이상: () => '해제됨' },
  안전커버: { 단위: '', 기준: '정상', 확률: 0.02, 심각도: '상', 정상: () => '정상', 이상: () => '파손' },
  윤활급유: { 단위: '', 기준: '완료', 확률: 0.09, 심각도: '하', 정상: () => '완료', 이상: () => '미실시' },
};

/** 이상이 반복되는 설비 5대(상위 순위가 생기도록 확률을 올린다) */
const 취약설비 = {};
[3, 9, 11, 0, 17].forEach((i) => (취약설비[설비정의[i].설비태그] = true));

const 조치문구 = {
  토출압력: '언로더 밸브 조정', 오일레벨: '오일 보충', 진동소음: '베어링 정렬 재조정',
  베어링온도: '그리스 재주입', 탱크액위: '공급 밸브 개방', 배관누설: '패킹 교체',
  밸브잠금: '잠금장치 복구', 안전커버: '커버 재설치', 윤활급유: '급유 실시',
};

const 작업유형정의 = {
  수리: {
    작업: '베어링 교체', 조인원: 3,
    내용: '기준 초과가 반복된 부위를 분해 점검하고 베어링·패킹을 교체한다. 조립 후 진동·온도를 재측정해 기준 이내인지 확인한다.',
    위험: '회전체 협착, 잔류 압력, 고온부 접촉',
    안전: 'LOTO(전원 차단·잠금) 후 작업, 잔류 압력 방출 확인, 보호장갑·안전화 착용, 작업 반경 통제선 설치',
    자재: '베어링 6205 2EA, 패킹 세트, 그리스, 토크렌치, 진동계',
  },
  예방정비: {
    작업: '정밀 점검·정렬', 조인원: 2,
    내용: '커플링 정렬 상태와 체결 토크를 점검하고, 급유 상태와 마모부를 확인해 필요 부품을 사전 교체한다.',
    위험: '회전체 접촉, 미끄러짐, 협소 공간',
    안전: '운전 정지 확인 후 작업, 안전 커버 해체 시 임시 방호 설치, 2인 1조 작업',
    자재: '다이얼 게이지, 토크렌치, 그리스건, 세정제, 걸레',
  },
  교체: {
    작업: '소모품 교체', 조인원: 2,
    내용: '필터·패킹·벨트 등 소모품을 규정 주기에 맞춰 교체하고 교체 이력을 기록한다.',
    위험: '잔류 압력, 절단면 접촉',
    안전: '압력 방출 및 밸브 잠금 확인, 절단 보호장갑 착용, 폐기물 분리 보관',
    자재: '필터 엘리먼트, 패킹 세트, V벨트, 공구 세트',
  },
  청소: {
    작업: '내부 청소·점검', 조인원: 2,
    내용: '내부 이물·스케일을 제거하고 배수 상태와 부식 진행을 확인한다.',
    위험: '분진, 미끄러짐, 협소 공간',
    안전: '환기 후 작업, 방진마스크·보안경 착용, 미끄럼 방지 매트 설치',
    자재: '브러시, 산업용 세정제, 방진마스크, 보안경, 폐기물 봉투',
  },
};

/* ============================ 탭 정의 ============================ */

/**
 * 탭 정의. 함수로 감싸 호출 시점에 만든다 —
 * Apps Script 는 파일 실행 순서를 보장하지 않아, 최상위 배열이 이렇게 다른 파일(Code.js)의
 * SH·항목목록을 로드 시점에 참조하면 "SH is not defined" 로 터진다.
 */
function 탭정의() {
  return [
    { 이름: SH.설비, 헤더: ['설비태그', '설비명', '설비종류', '구역', '설치연도', '담당자', '담당부서', '점검주기(일)', '비고'] },
    { 이름: SH.점검, 헤더: ['점검ID', '점검일', '설비태그', '설비명', '구역', '담당자', '점검자', '항목명', '측정값', '기준', '단위', '판정', '심각도', '조치상태', '조치내용', '비고'] },
    { 이름: SH.계획, 헤더: ['계획ID', '작업일', '설비태그', '설비명', '구역', '작업명', '작업유형', '작업내용', '담당자', '작업조인원', '예상소요(h)', '위험요인', '안전조치', '필요자재', '상태', '결재자'] },
    { 이름: SH.설정, 헤더: ['키', '값', '메모'] },
    { 이름: SH.요약설비, 헤더: ['설비태그', '설비명', '구역'].concat(항목목록, ['이상합계', '점검횟수', '이상률(%)']) },
    { 이름: SH.요약항목, 헤더: ['항목명', '이상합계', '상', '중', '하', '미조치', '최근발생일'] },
    { 이름: SH.발행, 헤더: ['발행일시', '문서종류', '대상ID', '문서명', '구글독스ID', 'PDF링크', 'DOCX링크', '소요초', '상태', '오류', '재생성'] },
    { 이름: SH.측정, 헤더: ['측정일시', '구분', '작업', '건수', '소요초', '건당초', '방식', '비고'] },
    { 이름: SH.안내, 헤더: ['안내'] },
  ];
}

const 설정기본값 = () => {
  const 오늘 = new Date();
  const 종료 = _날짜(오늘);
  const 시작 = _날짜(new Date(오늘.getTime() - 27 * 86400000));
  return [
    ['회사명', '(주)가상케미칼', '실제 회사명이 아닙니다(포트폴리오용)'],
    ['부서명', '생산1팀 유틸리티파트', ''],
    ['승인자', '김현우 팀장', '문서 결재자'],
    ['조회함수', 'XLOOKUP', 'XLOOKUP 또는 INDEXMATCH'],
    ['문서번호접두어_작업계획', 'WP', '작업계획서 번호 접두어'],
    ['문서번호접두어_결과보고', 'IR', '결과보고서 번호 접두어'],
    ['대장기간_시작', 시작, '집계 시작일 (기본: 오늘-27일)'],
    ['대장기간_종료', 종료, '집계 종료일 (기본: 오늘)'],
    ['문서한번에', 8, '한 번 실행에 만들 문서 수(6분 제한 대비)'],
    ['시간제한초', 270, '이 시간을 넘으면 다음 실행으로 넘김'],
    ['피벗시도', 'TRUE', 'Sheets API 고급 서비스가 있으면 피벗 탭 생성'],
    ['합성데이터고지', '이 문서의 모든 데이터는 포트폴리오 시연용 합성(가상) 데이터입니다.', ''],
    ['출력폴더ID', '', '설치_전체가 자동 입력'],
    ['구글독스폴더ID', '', '설치_전체가 자동 입력'],
    ['PDF폴더ID', '', '설치_전체가 자동 입력(출력 폴더와 동일)'],
    ['시트버전', '1.0', ''],
  ];
};

/* ============================ 설치 ============================ */

function _설정안전(key, 기본) {
  try {
    const v = 설정(key);
    return v === '' || v === null || v === undefined ? 기본 : v;
  } catch (e) {
    return 기본;
  }
}

function _설정읽기(ss) {
  const sh = ss.getSheetByName(SH.설정);
  const 값 = {};
  if (!sh || sh.getLastRow() < 2) return 값;
  sh.getDataRange().getValues().slice(1).forEach((r) => {
    const k = String(r[0]).trim();
    if (k) 값[k] = r[1];
  });
  return 값;
}

function _설정쓰기_시트(ss, key, 값) {
  const sh = ss.getSheetByName(SH.설정);
  const v = sh.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(값);
      _설정캐시[key] = 값;
      return;
    }
  }
  sh.appendRow([key, 값, '']);
  _설정캐시[key] = 값;
}

function _서식헤더(sh, 열수) {
  _안전('헤더 서식', () => {
    sh.getRange(1, 1, 1, 열수).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  });
}

/** 탭이 없으면 만들고, 헤더가 다를 때만 헤더를 다시 쓴다(데이터는 건드리지 않는다) */
function _탭보장(ss, 이름, 헤더) {
  let sh = ss.getSheetByName(이름);
  if (!sh) {
    sh = ss.insertSheet(이름);
    sh.getRange(1, 1, 1, 헤더.length).setValues([헤더]);
    _서식헤더(sh, 헤더.length);
    return true;
  }
  const 현재 = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 헤더.length)).getValues()[0].map(String);
  const 같음 = 헤더.every((h, i) => String(현재[i] || '').trim() === h);
  if (!같음) {
    sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 헤더.length)).clearContent();
    sh.getRange(1, 1, 1, 헤더.length).setValues([헤더]);
    _서식헤더(sh, 헤더.length);
    _로그('[주의] ' + 이름 + ' 탭 헤더를 다시 썼습니다(데이터 행은 그대로 두었습니다)');
  }
  return false;
}

function 설치_전체() {
  로그비우기();
  const 시작 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _로그('=== 설치 시작: ' + ss.getName() + ' ===');

  _안전('시간대·로케일', () => {
    ss.setSpreadsheetTimeZone(TZ);
    ss.setSpreadsheetLocale('ko_KR');
  });

  const 만든탭 = [];
  탭정의().forEach((t) => {
    if (_탭보장(ss, t.이름, t.헤더)) 만든탭.push(t.이름);
  });
  _안전('빈 기본 시트 정리', () => {
    ss.getSheets()
      .filter((s) => /^Sheet\d*$/.test(s.getName()) && s.getLastRow() === 0)
      .forEach((s) => ss.deleteSheet(s));
  });
  _로그('탭 ' + ss.getSheets().length + '개 (새로 만든 탭: ' + (만든탭.join(', ') || '없음') + ')');

  // 설정: 기존 값은 절대 덮지 않고, 없는 키만 채운다
  const 기존 = _설정읽기(ss);
  let 채운수 = 0;
  설정기본값().forEach(([k, v]) => {
    if (!(k in 기존) || String(기존[k]).trim() === '') {
      _설정쓰기_시트(ss, k, v);
      채운수++;
    }
  });
  _로그('설정 ' + Object.keys(기존).length + '개 중 ' + 채운수 + '개 채움(기존 값은 유지)');

  // 폴더: 루트/출력/구글독스
  const 루트 = _폴더보장(DriveApp.getRootFolder(), '서류자동화_포트폴리오');
  const 출력 = _폴더보장(루트, '출력');
  const 독스 = _폴더보장(루트, '구글독스');
  [['출력폴더ID', 출력.getId()], ['구글독스폴더ID', 독스.getId()], ['PDF폴더ID', 출력.getId()]].forEach(([k, v]) => {
    const 지금 = _설정읽기(ss)[k];
    if (!지금 || String(지금).trim() === '') _설정쓰기_시트(ss, k, v);
  });
  _로그('드라이브 폴더: ' + 루트.getName() + ' / 출력 · 구글독스');

  _안내작성(ss);

  대장서식적용(_시트(SH.점검), 1);
  요약갱신();

  if (String(_설정안전('피벗시도', 'TRUE')).toUpperCase() === 'TRUE') _안전('피벗', 피벗만들기);

  _로그('설치 완료: ' + ((Date.now() - 시작) / 1000).toFixed(1) + '초');
  _로그('※ 다음: 설치_샘플데이터 실행');
  return 로그전체();
}

function 설치_피벗() {
  로그비우기();
  피벗만들기();
  return 로그전체();
}

function 설치_월간트리거() {
  로그비우기();
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === 'monthlyLedger') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('monthlyLedger').timeBased().onMonthDay(1).atHour(7).create();
  _로그('월간 트리거: 매월 1일 07시 대장내보내기');
  return 로그전체();
}

/* ============================ 안내 탭 ============================ */

function _안내작성(ss) {
  const sh = ss.getSheetByName(SH.안내) || ss.insertSheet(SH.안내, 0);
  sh.clear();
  const 줄 = [
    ['실무 서류 자동 생성 세트', '제목'],
    ['구글시트 + Apps Script (무료 계정)', '부제'],
    ['', ''],
    ['■ 무엇을 하는 예제인가', '소제목'],
    ['설비 점검 기록을 쌓으면, 그 기록에서 ① 요약표 ② 점검대장(xlsx) ③ 작업계획서 ④ 점검결과보고서를 자동으로 만듭니다.', '본문'],
    ['회사 업무에 실제로 쓰는 자료가 아니라, 자동화 실력을 보여주기 위한 포트폴리오입니다.', '본문'],
    ['데이터는 전부 합성(가상)입니다 — 실제 회사 자료, 설비 태그, 인명을 쓰지 않았습니다.', '본문'],
    ['', ''],
    ['■ 사용 순서', '소제목'],
    ['1. 설치_전체 — 탭·설정·폴더를 만들고 수식·서식을 적용합니다. 두 번 돌려도 데이터가 지워지지 않습니다.', '본문'],
    ['2. 설치_샘플데이터 — 점검기록 5,040행(20대 × 28일 × 9항목)과 작업계획 20건을 만듭니다.', '본문'],
    ['3. 점검대장(.xlsx) 내보내기 — 엑셀에서 열리는 대장 파일을 드라이브 [출력] 폴더에 만듭니다.', '본문'],
    ['4. 작업계획서 일괄 생성 / 점검결과보고서 일괄 생성 — 한 번에 8부씩 만듭니다(3회 실행).', '본문'],
    ['5. 합본 PDF 만들기 — 만들어진 문서들을 한 파일로 이어 붙입니다.', '본문'],
    ['6. 발행 문서 검증 — 원본 데이터와 문서 본문을 기계적으로 대조해 누락을 찾습니다.', '본문'],
    ['', ''],
    ['■ 시트 구성', '소제목'],
    ['설비마스터 — 설비 20대 기준 정보', '본문'],
    ['점검기록 — 점검 1건당 1행(5,040행). 설비명·구역·담당자는 XLOOKUP 수식으로 채웁니다.', '본문'],
    ['작업계획 — 설비별 이상 건수 상위부터 작업계획 20건', '본문'],
    ['요약_설비별 · 요약_항목별 — COUNTIFS 수식(엑셀에서도 그대로 계산됩니다)', '본문'],
    ['발행이력 — 만든 문서의 링크·소요 시간·상태', '본문'],
    ['측정로그 — 자동/수작업 소요 시간 비교 기록', '본문'],
    ['설정 — 회사명·기간·한 번에 만들 문서 수 등', '본문'],
    ['', ''],
    ['■ 참고', '소제목'],
    ['대장의 수식은 엑셀 호환이 되는 것만 씁니다(COUNTIFS·IF·TEXT·XLOOKUP). QUERY·FILTER·ARRAYFORMULA 는 쓰지 않았습니다.', '본문'],
    ['수식 범위는 전체 열(100만 행) 대신 실제 데이터 범위(예: $2:$5041)로 잡아 계산을 가볍게 했습니다.', '본문'],
    ['그래서 점검기록에 행을 새로 추가했다면 메뉴 [요약 다시 계산] 을 한 번 실행해 범위를 다시 잡아 주세요.', '본문'],
    ['무료 계정 할당량: 스크립트 실행 6분/회, 트리거 90분/일. 그래서 문서는 8부씩 나눠 만듭니다.', '본문'],
  ];
  const 값 = 줄.map((r) => [r[0]]);
  sh.getRange(1, 1, 값.length, 1).setValues(값);
  줄.forEach((r, i) => {
    const p = sh.getRange(i + 1, 1);
    if (r[1] === '제목') p.setFontSize(16).setFontWeight('bold').setFontColor('#1F3864');
    else if (r[1] === '부제') p.setFontSize(11).setFontColor('#666666');
    else if (r[1] === '소제목') p.setFontSize(12).setFontWeight('bold').setFontColor('#1F3864');
    else p.setFontSize(10);
    sh.setRowHeight(i + 1, r[1] === '제목' ? 26 : 18);
  });
  sh.setColumnWidth(1, 620);
  _안전('안내 탭 위치', () => ss.setActiveSheet(sh));
  return sh;
}

/* ============================ 샘플 데이터 ============================ */

/** 같은 결과가 나오도록 고정된 난수(재현 가능) */
function _난수기계(seed) {
  let s = seed >>> 0;
  return function () {
    s = (s * 1664525 + 1013904223) >>> 0;
    return s / 4294967296;
  };
}

function 설치_샘플데이터() {
  로그비우기();
  const 시작 = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 설치를 건너뛰고 실행해도 "탭이 없습니다" 로 죽지 않게 먼저 확인한다
  if (!ss.getSheetByName(SH.설비) || !ss.getSheetByName(SH.점검) || !ss.getSheetByName(SH.설정)) {
    _로그('[알림] 탭이 아직 없어 설치_전체를 먼저 실행합니다');
    설치_전체();
  }

  const r = _난수기계(20260925);

  // 1) 설비마스터
  const 설비탭 = _시트(SH.설비);
  설비탭.getRange(2, 1, Math.max(설비탭.getLastRow() - 1, 1), 설비탭.getLastColumn()).clearContent();
  const 설비헤더 = 탭정의()[0].헤더;
  설비탭.getRange(2, 1, 설비정의.length, 설비헤더.length)
    .setValues(설비정의.map((d) => 설비헤더.map((h) => d[h])));
  _로그('가상 설비 ' + 설비정의.length + '대 등록 (AC/PU/TK/BL)');

  // 2) 점검기록 — 20대 × 28일 × 9항목
  const 기간시작 = _설정안전('대장기간_시작', '');
  const 기간종료 = _설정안전('대장기간_종료', '');
  const 시작일 = 기간시작 ? new Date(_날짜문자열(기간시작) + 'T00:00:00+09:00') : new Date(Date.now() - 27 * 86400000);
  const 날짜들 = [];
  for (let i = 0; i < 28; i++) 날짜들.push(_날짜(new Date(시작일.getTime() + i * 86400000)));

  const 점검탭 = _시트(SH.점검);
  점검탭.getRange(2, 1, Math.max(점검탭.getLastRow() - 1, 1), 점검탭.getLastColumn()).clearContent();

  const 행들 = [];
  let 이상수 = 0;
  const 설비별이상 = {};
  날짜들.forEach((날, di) => {
    설비정의.forEach((설비) => {
      const 점검자 = 점검자목록[di % 점검자목록.length];
      const 적용 = 적용항목[설비.설비종류];
      항목목록.forEach((항목, ai) => {
        const 정의 = 항목정의[항목];
        const 해당없음 = 적용.indexOf(항목) < 0;
        let 측정값 = '-';
        let 판정 = '정상';
        let 심각도 = '';
        let 조치상태 = '';
        let 조치내용 = '';

        if (!해당없음) {
          const 확률 = 정의.확률 * (취약설비[설비.설비태그] ? 2.2 : 1);
          const 이상 = r() < 확률;
          측정값 = String((이상 ? 정의.이상 : 정의.정상)(r));
          if (이상) {
            판정 = '이상';
            심각도 = 정의.심각도 || (측정값 === '이상' ? '상' : '하');
            const 조치난수 = r();
            if (조치난수 < 0.55) { 조치상태 = '완료'; 조치내용 = 조치문구[항목]; }
            else if (조치난수 < 0.85) 조치상태 = '진행중';
            else 조치상태 = '미조치';
            이상수++;
            설비별이상[설비.설비태그] = (설비별이상[설비.설비태그] || 0) + 1;
          }
        }

        행들.push([
          'CK-' + 날.replace(/-/g, '') + '-' + 설비.설비태그 + '-' + (ai + 1),
          날, 설비.설비태그, '', '', '', 점검자, 항목, 측정값, 정의.기준, 정의.단위,
          판정, 심각도, 조치상태, 조치내용, '',
        ]);
      });
    });
  });
  점검탭.getRange(2, 1, 행들.length, 행들[0].length).setValues(행들);
  _점검캐시비우기(); // 방금 쓴 데이터를 다시 읽도록
  // 집계기간을 실제 생성한 날짜로 확정한다(문서·대장의 기간과 데이터가 어긋나지 않게)
  _설정쓰기_시트(ss, '대장기간_시작', 날짜들[0]);
  _설정쓰기_시트(ss, '대장기간_종료', 날짜들[날짜들.length - 1]);
  _로그('가상 점검기록 ' + 행들.length + '행 생성 (' + 날짜들[0] + ' ~ ' + 날짜들[날짜들.length - 1] + ')');
  _로그('이상 판정 ' + 이상수 + '건 (전체의 ' + ((이상수 / 행들.length) * 100).toFixed(1) + '%)');

  // 3) 작업계획 — 이상 건수 상위부터
  const 순위 = 설비정의.slice().sort((a, b) => (설비별이상[b.설비태그] || 0) - (설비별이상[a.설비태그] || 0));
  const 유형순서 = [];
  순위.forEach((d, i) => {
    유형순서.push(i < 5 ? '수리' : i < 10 ? '예방정비' : i < 14 ? '교체' : '청소');
  });
  const 계획탭 = _시트(SH.계획);
  계획탭.getRange(2, 1, Math.max(계획탭.getLastRow() - 1, 1), 계획탭.getLastColumn()).clearContent();
  const 승인자 = String(_설정안전('승인자', '김현우 팀장'));
  const 계획행들 = 순위.map((설비, i) => {
    const 유형 = 유형순서[i];
    const 정의 = 작업유형정의[유형];
    const 작업일 = _날짜(new Date(Date.now() + (i + 1) * 86400000));
    return [
      _문서번호('WP', i + 1), 작업일, 설비.설비태그, 설비.설비명, 설비.구역,
      설비.설비태그 + ' ' + 정의.작업, 유형, 정의.내용, 설비.담당자, 정의.조인원,
      (정의.조인원 * 1.5).toFixed(1), 정의.위험, 정의.안전, 정의.자재, '계획', 승인자,
    ];
  });
  계획탭.getRange(2, 1, 계획행들.length, 계획행들[0].length).setValues(계획행들);
  _로그('작업계획 ' + 계획행들.length + '건 생성 (수리 5 / 예방정비 5 / 교체 4 / 청소 6)');

  // 4) 수식·서식·요약
  대장서식적용(점검탭, 1);
  요약갱신();

  _로그('샘플 데이터 완료: ' + ((Date.now() - 시작) / 1000).toFixed(1) + '초');
  _로그('※ 다음: 메뉴 [서류자동화] → 점검대장(.xlsx) 내보내기');
  return 로그전체();
}

/* ============================ 설치 확인 ============================ */

/**
 * 항목별 자체 진단. 무엇이 잘못됐는지 한 줄씩 ○/× 로 보여준다 —
 * 사용자가 로그를 그대로 붙여주면 원인을 바로 찾을 수 있다.
 */
function 설치_확인() {
  로그비우기();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let 통과 = 0;
  let 문제 = 0;

  const 검사 = (이름, fn) => {
    let 결과;
    try {
      결과 = fn();
    } catch (e) {
      결과 = e.message;
    }
    if (결과 === true || 결과 === undefined || 결과 === null) {
      통과++;
      _로그('  ○ ' + 이름);
    } else {
      문제++;
      _로그('  × ' + 이름 + ' — ' + 결과);
    }
  };

  _로그('=== 진단: ' + ss.getName() + ' ===');

  검사('탭 9개', () => {
    const 빠진것 = 탭정의().map((t) => t.이름).filter((n) => !ss.getSheetByName(n));
    return 빠진것.length ? '없음: ' + 빠진것.join(', ') : true;
  });

  검사('탭 헤더', () => {
    const 틀린것 = [];
    탭정의().forEach((t) => {
      if (t.이름 === SH.안내) return;
      const sh = ss.getSheetByName(t.이름);
      if (!sh) return;
      const 현재 = _헤더행(sh, 1);
      if (t.헤더.some((hd, i) => String(현재[i] || '').trim() !== hd)) 틀린것.push(t.이름);
    });
    return 틀린것.length ? '헤더 다름: ' + 틀린것.join(', ') : true;
  });

  검사('설비마스터 20행', () => {
    const sh = ss.getSheetByName(SH.설비);
    const 행 = sh ? sh.getLastRow() - 1 : 0;
    return 행 === 20 ? true : 행 + '행 (20행이어야 함)';
  });

  검사('점검기록 5,040행', () => {
    const sh = ss.getSheetByName(SH.점검);
    const 행 = sh ? sh.getLastRow() - 1 : 0;
    return 행 === 5040 ? true : 행 + '행 (5,040행이어야 함)';
  });

  검사('점검기록 조회 수식 3열', () => {
    const sh = ss.getSheetByName(SH.점검);
    if (!sh || sh.getLastRow() < 2) return '점검기록이 비어 있음';
    const 수 = sh.getRange(2, 4, 1, 3).getFormulas()[0].filter((x) => x).length;
    return 수 === 3 ? true : 'D:F 수식 ' + 수 + '/3 (설치_샘플데이터를 다시 실행하세요)';
  });

  검사('점검기록 드롭다운 4열', () => {
    const sh = ss.getSheetByName(SH.점검);
    if (!sh || sh.getLastRow() < 2) return '점검기록이 비어 있음';
    const h = _헤더(sh);
    const 있음 = ['설비태그', '판정', '심각도', '조치상태'].filter((n) => {
      const c = h.indexOf(n);
      return c >= 0 && sh.getRange(2, c + 1).getDataValidation() !== null;
    });
    return 있음.length === 4 ? true : 있음.length + '/4열만 있음';
  });

  검사('요약 수식 결과', () => {
    const sh = ss.getSheetByName(SH.요약설비);
    if (!sh || sh.getLastRow() < 2) return '요약_설비별이 비어 있음';
    const 값 = sh.getRange(2, 13, Math.min(sh.getLastRow() - 1, 20), 3).getDisplayValues();
    const 오류 = 값.filter((r) => r.some((x) => String(x).charAt(0) === '#')).length;
    return 오류 ? 오류 + '행에 수식 오류(#REF! 등)' : true;
  });

  검사('설정 키 16개', () => {
    const 값 = _설정읽기(ss);
    const 빠진것 = 설정기본값().map(([k]) => k).filter((k) => !(k in 값));
    return 빠진것.length ? '없는 키: ' + 빠진것.join(', ') : true;
  });

  검사('드라이브 폴더', () => {
    const ids = ['출력폴더ID', '구글독스폴더ID'].map((k) => String(_설정안전(k, '')));
    const 없는것 = ids.filter((id) => !id);
    if (없는것.length) return '설정에 폴더 ID가 없음 → 설치_전체를 실행하세요';
    const 실패 = ids.filter((id) => !_안전('폴더 확인', () => DriveApp.getFolderById(id).getName()));
    return 실패.length ? '접근 실패: ' + 실패.join(', ') : true;
  });

  검사('작업계획 20행', () => {
    const sh = ss.getSheetByName(SH.계획);
    const 행 = sh ? sh.getLastRow() - 1 : 0;
    return 행 === 20 ? true : 행 + '행 (20행이어야 함)';
  });

  검사('발행이력', () => {
    const 전체 = Math.max(_시트(SH.발행).getLastRow() - 1, 0);
    if (!전체) {
      _로그('      · 아직 문서를 만들지 않았습니다(정상)');
      return true;
    }
    const 완료 = _발행이력완료행('작업계획서').length + _발행이력완료행('점검결과보고서').length;
    _로그('      · 완료 ' + 완료 + ' / 전체 ' + 전체 + '건');
    return 완료 ? true : '완료 건이 없습니다 — 양식_일괄생성을 실행하세요';
  });

  검사('최근 측정로그', () => {
    const sh = ss.getSheetByName(SH.측정);
    if (!sh || sh.getLastRow() < 2) return true;
    const 최근 = sh.getDataRange().getValues().slice(-3);
    최근.forEach((r) => _로그('      · ' + r[1] + ' / ' + r[2] + ' / ' + r[3] + '건 / ' + r[4] + '초'));
    return true;
  });

  _로그('=== 결과: 정상 ' + 통과 + ' / 문제 ' + 문제 + ' ===');
  if (문제) _로그('※ × 항목을 그대로 복사해서 보내주시면 원인을 잡습니다.');
  return 로그전체();
}

/* ============================ 수작업 측정 ============================ */

function 측정_수작업시작() {
  로그비우기();
  PropertiesService.getDocumentProperties().setProperty('측정_시작', String(Date.now()));
  _로그('수작업 측정 시작. 엑셀·한글로 같은 문서를 직접 만들어 보세요.');
  _로그('끝나면 측정_수작업종료 를 실행하고 만든 건수를 입력합니다.');
  return 로그전체();
}

function 측정_수작업종료(건수) {
  로그비우기();
  const p = PropertiesService.getDocumentProperties();
  const 시작 = Number(p.getProperty('측정_시작') || 0);
  if (!시작) throw new Error('측정_수작업시작을 먼저 실행하세요.');
  const 초 = (Date.now() - 시작) / 1000;
  let n = Number(건수);
  if (!n) {
    // 편집기 실행은 인수를 못 넣으므로 물어본다(대화상자가 안 되면 1건으로 기록)
    n = 1;
    try {
      const 답 = SpreadsheetApp.getUi().prompt('수작업 측정', '직접 만든 문서가 몇 건인가요?', SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
      if (답.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) n = Number(답.getResponseText()) || 1;
    } catch (e) {}
  }
  _측정로그('수작업', '문서 작성', n, 초, '수작업', '엑셀/한글 수작업 기준');
  p.deleteProperty('측정_시작');
  _로그('수작업 ' + n + '건 / ' + 초.toFixed(1) + '초 / 건당 ' + (초 / n).toFixed(1) + '초');
  return 로그전체();
}

/* ============================ 초기화 ============================ */

function 초기화_출력폴더() {
  로그비우기();
  const 폴더ID = String(_설정안전('출력폴더ID', ''));
  if (!폴더ID) throw new Error('출력폴더ID 가 설정되지 않았습니다. 설치_전체를 먼저 실행하세요.');
  const 폴더 = DriveApp.getFolderById(폴더ID);
  let n = 0;
  const 파일들 = 폴더.getFiles();
  while (파일들.hasNext()) {
    파일들.next().setTrashed(true);
    n++;
  }
  const 발행 = _시트(SH.발행);
  if (발행.getLastRow() > 1) 발행.getRange(2, 1, 발행.getLastRow() - 1, 발행.getLastColumn()).clearContent();
  _로그('출력 폴더 파일 ' + n + '개 정리, 발행이력 초기화');
  return 로그전체();
}

/* ============================ 메뉴 ============================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('서류자동화')
    .addItem('① 전체 설치', '메뉴_설치전체')
    .addItem('② 샘플 데이터 생성', '메뉴_샘플데이터')
    .addSeparator()
    .addItem('점검대장(.xlsx) 내보내기', '메뉴_대장내보내기')
    .addItem('요약 다시 계산', '메뉴_요약갱신')
    .addItem('합본 PDF 만들기', '메뉴_합본')
    .addSeparator()
    .addItem('작업계획서 일괄 생성(8부)', '메뉴_작업계획서')
    .addItem('점검결과보고서 일괄 생성(8부)', '메뉴_결과보고서')
    .addItem('발행 문서 검증', '메뉴_검증')
    .addSeparator()
    .addItem('설치 상태 확인', '메뉴_설치확인')
    .addToUi();
}

/** 메뉴 실행 공통. 오류가 나면 이유를 대화상자로 보여준다(원시 예외 문구만 던지지 않는다) */
function _메뉴실행(제목, 일) {
  로그비우기();
  try {
    const 본문 = 일();
    _알림(제목 + ' 완료', 본문 || _마지막줄(8));
  } catch (e) {
    _로그('[오류] ' + e.message);
    _알림(제목 + ' 오류', e.message + '\n\n' + _마지막줄(6));
  }
}

function _알림(제목, 본문) {
  const 글 = String(본문 === null || 본문 === undefined ? '' : 본문);
  try {
    SpreadsheetApp.getUi().alert(제목, 글, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log(제목 + ' / ' + 글);
  }
  return 글;
}

/** 설치·생성처럼 시간이 걸리는 일은 실행 로그 대신 마지막 몇 줄만 보여준다 */
function _마지막줄(n) {
  return 로그전체().split('\n').slice(-Number(n || 8)).join('\n');
}

function 메뉴_설치전체() {
  _메뉴실행('설치', () => {
    설치_전체();
    return _마지막줄(8);
  });
}

function 메뉴_샘플데이터() {
  let 답 = 'YES';
  try {
    답 = SpreadsheetApp.getUi().alert('샘플 데이터', '기존 점검기록·작업계획을 지우고 합성 데이터 5,040행을 다시 만듭니다. 계속할까요?',
      SpreadsheetApp.getUi().ButtonSet.YES_NO);
  } catch (e) {}
  // V8 에서 alert(title, message, buttonSet) 은 'YES'/'NO' 문자열을 돌려준다
  const 답글 = String(답).toUpperCase();
  if (답글.indexOf('YES') < 0 && 답글.indexOf('예') < 0) return;
  _메뉴실행('샘플 데이터', () => {
    설치_샘플데이터();
    return _마지막줄(8);
  });
}

function 메뉴_대장내보내기() {
  _메뉴실행('점검대장 내보내기', () => {
    const url = 대장내보내기();
    return _마지막줄(3) + '\n\n' + url;
  });
}

function 메뉴_요약갱신() {
  _메뉴실행('요약 다시 계산', () => {
    요약갱신();
    return _마지막줄(3);
  });
}

function 메뉴_작업계획서() {
  _메뉴실행('작업계획서 생성', () => {
    양식_일괄생성('작업계획서');
    return _마지막줄(5);
  });
}

function 메뉴_결과보고서() {
  _메뉴실행('점검결과보고서 생성', () => {
    양식_일괄생성('점검결과보고서');
    return _마지막줄(5);
  });
}

function 메뉴_합본() {
  _메뉴실행('합본 PDF', () => {
    const pdf = 합본PDF('점검결과보고서');
    return _마지막줄(2) + '\n\n' + pdf;
  });
}

function 메뉴_검증() {
  _메뉴실행('발행 문서 검증', () => {
    검증_문서대조('점검결과보고서');
    return _마지막줄(6);
  });
}

function 메뉴_설치확인() {
  _메뉴실행('진단', () => 설치_확인());
}
