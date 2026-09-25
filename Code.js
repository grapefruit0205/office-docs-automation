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