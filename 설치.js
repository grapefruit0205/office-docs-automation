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