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

/** 작업계획 탭 값(헤더 포함). 문서 20부를 만들 때 매번 다시 읽지 않도록 담아 둔다 */
let _계획값캐시 = null;

function _계획값() {
  if (!_계획값캐시) _계획값캐시 = _재시도('작업계획 읽기', () => _시트(SH.계획).getDataRange().getValues());
  return _계획값캐시;
}

function _대상목록(종류) {
  const v = (종류 === '작업계획서' ? _계획값() : _설비값()).slice(1);
  return v.filter((r) => String(r[0]).trim()).map((r) => ({ ID: String(r[0]).trim() }));
}

function _대상행(종류, ID) {
  if (종류 === '작업계획서') {
    const v = _계획값().slice();
    const h = v.shift().map(String);
    const r = v.filter((x) => String(x[0]).trim() === ID)[0];
    return r ? _행객체(h, r) : null;
  }
  const v = _설비값().slice();
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
  if (!행) throw new Error('내부 함수입니다. 메뉴 [서류자동화] → 작업계획서 일괄 생성 으로 실행하세요.');
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
  if (!설비행) throw new Error('내부 함수입니다. 메뉴 [서류자동화] → 점검결과보고서 일괄 생성 으로 실행하세요.');
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

  let pdf링크 = '';
  const pdf = _내보내기(파일, 'application/pdf', 'pdf');
  if (pdf) pdf링크 = 출력.createFile(pdf.setName(파일명 + '.pdf')).getUrl();
  else _로그('[주의] PDF 변환이 막혀 문서 링크만 남깁니다: ' + 파일.getUrl());

  let docx링크 = '';
  if (o.docx !== false) {
    _안전('docx 변환', () => {
      const docx = _내보내기(파일, MimeType.MICROSOFT_WORD, 'docx');
      if (docx) docx링크 = 출력.createFile(docx.setName(파일명 + '.docx')).getUrl();
    });
  }
  return { 독스ID: doc.getId(), 독스링크: 파일.getUrl(), pdf: pdf링크, docx: docx링크, 파일명 };
}

/** 문서 1건 생성. 실패하면 임시 문서를 정리하고 오류를 그대로 올린다 */
function 문서한건(종류, 대상ID) {
  if (!종류 || !대상ID) throw new Error('사용법: 문서한건("작업계획서" 또는 "점검결과보고서", 대상ID) — 보통은 양식_일괄생성 또는 메뉴를 쓰세요.');
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
  if (종류 !== '작업계획서' && 종류 !== '점검결과보고서') {
    throw new Error('사용법: 양식_일괄생성("작업계획서") 또는 양식_일괄생성("점검결과보고서") — 편집기에서 인수를 넣기 어려우면 시트 메뉴 [서류자동화] 를 쓰세요.');
  }
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