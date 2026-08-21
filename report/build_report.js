// 국내 자율실험실(SDL) 현황 보고서 — docx 생성 스크립트
const path = require('path');
const SCRATCH = '/tmp/claude-0/-home-user-kist/f7349fbc-9d61-5776-85de-581602fe2d75/scratchpad';
const docx = require(path.join(SCRATCH, 'node_modules', 'docx'));
const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, convertMillimetersToTwip,
} = docx;
const fs = require('fs');

// ---------- 공통 설정 ----------
const KO = { ascii: '맑은 고딕', eastAsia: '맑은 고딕', hAnsi: '맑은 고딕', cs: '맑은 고딕' };
const NAVY = '1D3567';
const LINE = 'B9C2D4';
const HEADFILL = '1D3567';
const BOXFILL = 'EEF2F9';
const ALTFILL = 'F7F9FC';

const PAGE_W = convertMillimetersToTwip(210);
const MARGIN = convertMillimetersToTwip(14);
const CONTENT_W = PAGE_W - 2 * MARGIN; // 10276 dxa 내외

// 출처 목록 (본문 하이퍼링크와 공유)
const S = {
  hellodd:  { n: 1,  url: 'https://www.hellodd.com/news/articleView.html?idxno=112789',
              t: '헬로디디, "AI가 가설 세우고 로봇이 24시간 실험…韓 첫 \'자율실험실\' 연합 출범"(2026.8.11)' },
  sedaily80:{ n: 2,  url: 'https://www.sedaily.com/article/20077980',
              t: '서울경제, "정부 \'자율실험실 생태계\' 시동…삼성·SK 등 80여개 기관 뭉쳤다"(2026.8)' },
  zdnet:    { n: 3,  url: 'https://zdnet.co.kr/view/?no=20260811095548',
              t: 'ZDNet Korea, "로봇이 알아서 연구…자율실험실 구축 \'시동\'"(2026.8.11)' },
  mt:       { n: 4,  url: 'https://www.mt.co.kr/tech/2026/04/30/2026042917513071754',
              t: '머니투데이, "AI가 신소재 개발부터 상용화까지…과기정통부, 소재 특화 독자 AI 만든다"(2026.4.30)' },
  herald:   { n: 5,  url: 'https://biz.heraldcorp.com/article/10770134',
              t: '헤럴드경제, "\'AI가 설계하고 로봇이 실험\' 자율실험실 구축…K-문샷 신약개발 속도"' },
  bioin:    { n: 6,  url: 'https://bioin.or.kr/board.do?bid=notice&cmd=view&num=332132',
              t: '생명공학정책연구센터, "2026년 AI-네이티브 첨단바이오 자율실험실사업 신규과제 공모"' },
  segye:    { n: 7,  url: 'https://www.segye.com/newsView/20240404509280',
              t: '세계일보, "나노소재 개발 실험 횟수 500분의 1로…AI 설계 플랫폼 개발"' },
  sedailyAI:{ n: 8,  url: 'https://m.sedaily.com/article/14159501',
              t: '서울경제, "1만 번 실험을 24번으로…\'AI 동료 과학자\', 연구실 패러다임을 바꾸다"' },
  octopus:  { n: 9,  url: 'https://www.inews24.com/view/1793219',
              t: '아이뉴스24, "\'문어 실험실\'을 아시나요 [지금은 과학]" (KIST OCTOPUS 원격 자율실험 시스템)' },
  kims:     { n: 10, url: 'https://www.kims.re.kr/v17/bbx/board.php?bx_table=05_01&wr_id=715',
              t: '한국재료연구원, "재료硏, 인공지능 기반 전주기 자동화 연구 시스템 개발"' },
  escience: { n: 11, url: 'https://www.e-science.co.kr/news/articleView.html?idxno=107275',
              t: '이코노미사이언스, "\'AI가 자율적으로 실험한다\' 재료硏, 오토노머스 랩 개발"' },
  krict:    { n: 12, url: 'https://chemworld.kcsnet.or.kr/post/%ED%99%94%ED%95%99%EB%8D%B0%EC%9D%B4%ED%84%B0%EA%B8%B0%EB%B0%98%EC%97%B0%EA%B5%AC%EC%84%BC%ED%84%B0',
              t: '대한화학회 Chemworld, "한국화학연구원 화학데이터기반연구센터" 소개' },
  kaistmse: { n: 13, url: 'https://mse.kaist.ac.kr/index.php?mid=mse_research_highlight_en&document_srl=387700',
              t: 'KAIST 신소재공학과, "서동화 교수 연구팀-포스코홀딩스, 연구자 없이 로봇팔·AI로 소재 혁신 실현"' },
  asiae:    { n: 14, url: 'https://www.asiae.co.kr/article/2025080309221574066',
              t: '아시아경제, "\'AI·자동화 기술 중심\' KAIST-포스코홀딩스, 자율탐색 실험실 구축"(2025.8.3)' },
  kier:     { n: 15, url: 'https://energium.kier.re.kr/sub040101/articles/view/tableid/news/category/2/view_type/webzine/page/3/id/5665',
              t: '한국에너지기술연구원 ENERGIUM, "국내 최초 로봇 기반 무인 촉매 평가 실험실 열리다"' },
  unist:    { n: 16, url: 'https://www.industrynews.co.kr/news/articleView.html?idxno=72079',
              t: '인더스트리뉴스, "UNIST, 자동화 실험 플랫폼 개발…\'AI·로봇이 하루 1000번 화학실험\'"' },
  ibs:      { n: 17, url: 'https://www.ibs.re.kr/kor/sub02_06_01.do',
              t: '기초과학연구원(IBS), "인공지능 및 로봇 기반 합성 연구단" 연구단 소개' },
  snuaiis:  { n: 18, url: 'https://aiis.snu.ac.kr/sub1_4_22.php',
              t: '서울대학교 AI연구원(AIIS), "선도혁신연구센터"(스마트실험실 완전 자율화 과제)' },
  hitnews:  { n: 19, url: 'https://www.hitnews.co.kr/news/articleView.html?idxno=77438',
              t: '히트뉴스, "AI 신약개발 다음은 \'자율실험실\'…인프라 구축 필요"' },
  medifo:   { n: 20, url: 'https://www.medifonews.com/news/article.html?no=214458',
              t: '메디포뉴스, "제약바이오협회, 2026 제1차 AI신약개발자문위원회 개최"(SDL 실습교육장 운영)' },
  telescope:{ n: 21, url: 'https://telescopeinnovations.com/telescope-innovations-installs-koreas-first-self-driving-lab-for-pharma-rd-and-education/',
              t: 'Telescope Innovations, "Installs Korea\'s First Self-Driving Lab for Pharma R&D and Education"(2025.12.9)' },
  kribb:    { n: 22, url: 'https://www.kribb.re.kr/kor/sub02/sub02_02_01_view.jsp?b_idx=34637',
              t: '한국생명공학연구원, "K-바이오파운드리, 전 세계 합성생물학 연구실을 하나로 잇는다"' },
  lgchem:   { n: 23, url: 'https://blog.lgchem.com/2025/12/11_automation_lab/',
              t: 'LG화학, "안전은 올리고 속도는 빠르게, LG화학 자동화 실험실"(2025.12.11)' },
  ablelabs: { n: 24, url: 'https://www.thebionews.net/news/articleView.html?idxno=26684',
              t: '더바이오뉴스, "에이블랩스, 100억 투자 발판 \'자율실험실\' 속도"' },
  snujung:  { n: 25, url: 'https://eng.snu.ac.kr/snu/bbs/BMSR00005/view.do?boardId=6275&menuNo=200152',
              t: '서울대학교 공과대학, "정유성 교수팀, AI로 합성 어려운 신소재 되살렸다…LLM 기반 재설계 기술 개발"' },
  melloddy: { n: 26, url: 'https://kmelloddy.org/',
              t: 'K-MELLODDY 사업단 공식 홈페이지(연합학습 기반 신약개발 가속화 프로젝트)' },
  kisti:    { n: 27, url: 'https://dataon.kisti.re.kr/',
              t: 'KISTI 국가연구데이터플랫폼 DataON' },
  rsc:      { n: 28, url: 'https://pubs.rsc.org/en/content/articlelanding/2026/dd/d6dd00024j',
              t: 'J. Hwang et al., "Self-driving laboratories in Korea: a new era of autonomous discovery", Digital Discovery 5, 1968 (2026)' },
};

// ---------- 텍스트 헬퍼 ----------
// "일반텍스트 {키:링크텍스트} 일반텍스트" 형태를 런 배열로 변환
function runs(str, { size = 15, bold = false, color = '222222' } = {}) {
  const out = [];
  const re = /\{([a-zA-Z0-9]+):([^}]*)\}/g;
  let last = 0, m;
  while ((m = re.exec(str)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: str.slice(last, m.index), size, bold, color, font: KO }));
    const src = S[m[1]];
    if (!src) throw new Error('unknown source key: ' + m[1]);
    out.push(new ExternalHyperlink({
      link: src.url,
      children: [new TextRun({ text: m[2], size, bold, font: KO, color: '1B54B5', underline: {} })],
    }));
    last = re.lastIndex;
  }
  if (last < str.length) out.push(new TextRun({ text: str.slice(last), size, bold, color, font: KO }));
  return out;
}

function p(str, opts = {}) {
  const { size = 15, bold = false, color = '222222', align, before = 0, after = 0, indent } = opts;
  return new Paragraph({
    alignment: align,
    spacing: { before, after, line: opts.line || 220, lineRule: 'auto' },
    indent,
    children: runs(str, { size, bold, color }),
  });
}

function cell(text, { w, size = 14, bold = false, color = '222222', fill, align = AlignmentType.LEFT, pad = 60 } = {}) {
  const paras = (Array.isArray(text) ? text : [text]).map((t) =>
    new Paragraph({
      alignment: align,
      spacing: { before: 12, after: 12, line: 200, lineRule: 'auto' },
      children: runs(t, { size, bold, color }),
    }));
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    shading: fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined,
    margins: { top: 105, bottom: 105, left: pad, right: pad },
    children: paras,
  });
}

const thinBorders = {
  top:    { style: BorderStyle.SINGLE, size: 3, color: LINE },
  bottom: { style: BorderStyle.SINGLE, size: 3, color: LINE },
  left:   { style: BorderStyle.SINGLE, size: 3, color: LINE },
  right:  { style: BorderStyle.SINGLE, size: 3, color: LINE },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 3, color: LINE },
  insideVertical:   { style: BorderStyle.SINGLE, size: 3, color: LINE },
};

function dataTable(widths, header, rows) {
  const headRow = new TableRow({
    tableHeader: true,
    children: header.map((h, i) =>
      cell(h, { w: widths[i], size: 15, bold: true, color: 'FFFFFF', fill: HEADFILL, align: AlignmentType.CENTER })),
  });
  const bodyRows = rows.map((r, ri) => new TableRow({
    cantSplit: true,
    children: r.map((c, i) => {
      const isCenterCol = i === 0 || i === r.length - 1;
      return cell(c, {
        w: widths[i],
        size: 16,
        bold: i === 0,
        fill: ri % 2 === 1 ? ALTFILL : undefined,
        align: isCenterCol ? AlignmentType.CENTER : AlignmentType.LEFT,
      });
    }),
  }));
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: thinBorders,
    rows: [headRow, ...bodyRows],
  });
}

// ---------- 1. 제목 ----------
const title = new Paragraph({
  spacing: { after: 40, line: 240, lineRule: 'auto' },
  children: [new TextRun({ text: '국내 자율실험실(Self-Driving Lab) 현황 — 사례 중심', bold: true, size: 26, color: NAVY, font: KO })],
});
const subtitle = new Paragraph({
  spacing: { after: 140 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } },
  children: [new TextRun({ text: '기준일: 2026년 8월  |  공개 보도자료·기관 발표 기준 정리', size: 14, color: '666666', font: KO })],
});

// ---------- 2. 개요 박스 ----------
const boxText =
  '자율실험실(Self-Driving Lab, SDL)은 AI가 가설과 실험조건을 설계하고 로봇·자동화 장비가 실험을 수행한 뒤 그 결과를 AI가 다시 학습하는 ' +
  '{hellodd:‘설계–실험–분석–재설계’ 폐루프 연구 플랫폼}으로, 24시간 무인 운영을 통해 수년 걸리던 탐색을 수개월로 단축한다. ' +
  '국내에서는 KIST·재료연·화학연·에너지연 등 출연(연)과 KAIST·서울대·UNIST 등 대학이 소재·촉매·이차전지 분야에서 폐루프 실증에 성공했고, ' +
  '2026년에는 {herald:첨단바이오(3년간 495억 원)}와 {mt:소재 분야 자율실험센터}로 확산되고 있다. ' +
  '2026년 8월 11일에는 장비·AI 기업과 대학·출연(연) 등 {sedaily80:80여 개 기관이 참여하는 국가 차원의 ‘자율실험실 산·학·연 협의체’}가 출범해 ' +
  '(위원장 정유성 서울대 교수) 장비·데이터 표준 정립과 실증·확산을 추진 중이다.';

const boxTable = new Table({
  columnWidths: [CONTENT_W],
  width: { size: CONTENT_W, type: WidthType.DXA },
  borders: {
    top:    { style: BorderStyle.SINGLE, size: 6, color: NAVY },
    bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY },
    left:   { style: BorderStyle.SINGLE, size: 18, color: NAVY },
    right:  { style: BorderStyle.SINGLE, size: 6, color: NAVY },
  },
  rows: [new TableRow({
    children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: BOXFILL, color: 'auto' },
      margins: { top: 120, bottom: 120, left: 180, right: 160 },
      children: [
        new Paragraph({
          spacing: { after: 50, line: 220, lineRule: 'auto' },
          children: [new TextRun({ text: '한눈에 보기', bold: true, size: 16, color: NAVY, font: KO })],
        }),
        new Paragraph({ spacing: { line: 230, lineRule: 'auto' }, children: runs(boxText, { size: 16 }) }),
      ],
    })],
  })],
});

// ---------- 3. 표 1: 자율실험실(폐루프 SDL) ----------
const h1 = new Paragraph({
  spacing: { before: 220, after: 70 },
  children: [new TextRun({ text: '표 1. 국내 자율실험실(폐루프 SDL) 주요 사례', bold: true, size: 17, color: NAVY, font: KO })],
});

const W1 = [1450, 2320, 1150, 4092, 1306];
const T1_HEAD = ['기관', '실험실·조직 (담당자)', '주요 분야', '주요 성과', '자율실험 정도'];
const T1 = [
  ['KIST', '계산과학연구센터\n한상수 센터장·김동훈 박사\n(고려대 이관영 교수 공동)', '금속 나노입자·나노소재',
   '목표 물성 입력 → 로봇 합성 → 광특성 측정 → AI 재설계의 폐루프 구현. 실험 횟수를 {segye:최대 1/500로 감축}(3변수 기준 약 200회), {sedailyAI:태양전지 활성층용 은 나노입자 개발}. 원격·다중사용자 운영체계 {octopus:‘OCTOPUS’로 개발시간 5배 이상 단축}',
   '★★★\n폐루프 완전자율(24h 무인)'],
  ['한국재료연구원\n(KIMS)', '{kims:오토노머스 랩}\n(Autonomous Lab)', '금속·구조소재\n(반도체·이차전지·수소 확장)',
   '{escience:로봇팔 시료 이송 → 아크멜팅 → 튜브퍼니스 열처리 → XRD 정밀분석까지 전주기 자동화}. AI가 목표 특성으로부터 실험조건을 설계하고 결과를 즉시 분석해 다음 실험을 스스로 결정',
   '★★★\n설계–합성–분석 전주기 폐루프'],
  ['한국화학연구원\n(KRICT)', '화학플랫폼연구본부\n{krict:화학데이터기반연구센터}\n(신정호 센터장)', '태양전지·촉매 등 화학소재',
   '지능형 로봇 기반 무인 실험실을 구축해 인간 연구자 대비 고속으로 실험 데이터 생산. {krict:신소재 개발기간을 평균 10.2년 → 6.6년으로 단축}하는 효과 제시',
   '★★☆\n무인 실험실 운영·확장 중'],
  ['KAIST ×\n포스코홀딩스', '신소재공학과 {kaistmse:서동화 교수팀}\n– 포스코홀딩스 미래기술연구원\n자율탐색 실험실', '이차전지 양극재',
   '{asiae:정량·혼합·펠릿화·소성·분석을 연구자 개입 없이 수행}하는 자동화 시스템 + 분석 데이터를 해석해 최적 후보를 고르는 AI 모델 결합. {kaistmse:소재 탐색 기간 93% 단축}',
   '★★★\n산·학 공동 폐루프 실증'],
  ['한국에너지\n기술연구원(KIER)', '청정연료연구실\n박지찬 박사 연구진', '촉매 성능평가',
   '{kier:국내 최초 로봇 기반 무인 촉매 평가 실험실}. 로봇 2대가 공정을 분담해 전 과정 무인 수행, 수작업 대비 {kier:실험 속도 45배 향상·결과 변동성 32% 감소}',
   '★★☆\n평가 전과정 무인화'],
  ['UNIST · IBS', '{ibs:인공지능 및 로봇 기반 합성 연구단}\nBartosz A. Grzybowski 단장\n(UNIST 화학과)', '유기합성·화학반응 네트워크',
   '{unist:AI–로봇 플랫폼으로 하루 약 1,000회 화학반응 실험 수행}, 수천 가지 반응 조건을 동시 탐색해 반응 네트워크를 정밀 지도화하고 원하는 물질을 선택적으로 생성',
   '★★★\n고속 폐루프 탐색'],
  ['서울대학교', '화학생물공학부 최장욱 교수팀\n{snuaiis:스마트실험실(선도혁신연구센터)}', '이차전지 전해질·공정',
   '고난도 파지·조작이 가능한 로봇 매니퓰레이터로 다종 실험 프로토콜을 수행하는 완전 자율화 기술 개발. 기계학습 기반 전해질·공정 최적화 및 열화 경로 예측 연계',
   '★★☆\n완전자율화 기술 개발 중'],
  ['가톨릭대 등\n6개 기관', '{bioin:AI-네이티브 첨단바이오 자율실험실}\n범용: 가톨릭대 K-Cell 플랫폼\n특화: DGIST(액체생검)·KAIST(감염병)·고려대(유전자 전달체)·POSTECH(효소공학)·UNIST(오가노이드 약효평가)', '첨단바이오·신약개발',
   'K-문샷 신약개발 가속화의 기반 사업. {herald:3년간 총 495억 원을 투입해 범용 1개·특화 5개 자율실험실 구축}, {bioin:2026년 135억 원 규모 6개 내외 과제로 착수}',
   '★☆☆\n구축 착수 단계(2026~)'],
  ['한국제약바이오\n협회', '{medifo:AI 신약개발 SDL 실습교육장}\n(서울, 협회 운영)', '유기합성 기반 신약개발',
   '{telescope:국내 최초의 신약개발용 SDL 인프라}(캐나다 Telescope Innovations 장비, 계약 3주 만에 설치). 색상 최적화·용해도 스캐닝·합성 자동화 실습 운영, 토론토대 Acceleration Consortium과 협력',
   '★★☆\n실증·교육용 SDL 가동'],
];

const t1 = dataTable(W1, T1_HEAD, T1.map((r) => r.map((c) => c.split('\n'))));

const legend = new Paragraph({
  spacing: { before: 60, after: 0, line: 200, lineRule: 'auto' },
  children: runs('※ 자율실험 정도  ★★★ 설계–실험–분석–재설계가 폐루프로 무인 순환  |  ★★☆ 실험 전 과정 자동화 + AI 부분 적용(사람 개입 일부 잔존)  |  ★☆☆ 구축 착수 단계(설비 도입·설계 진행 중)', { size: 13, color: '666666' }),
});

// ---------- 4. 표 2: 자동화 / AI 실험설계 / 드라이랩 ----------
const h2 = new Paragraph({
  pageBreakBefore: true,
  spacing: { before: 0, after: 40 },
  children: [new TextRun({ text: '표 2. (별도 구분) 자율실험실은 아니나 실험 자동화·AI 실험설계·드라이랩에 해당하는 사례', bold: true, size: 17, color: NAVY, font: KO })],
});
const note2 = new Paragraph({
  spacing: { after: 70, line: 200, lineRule: 'auto' },
  children: runs('폐루프 자율성이 아직 확인되지 않아 표 1과 분리했다. 다만 대부분 SDL을 지향점으로 명시하고 있어 향후 표 1로 이동할 가능성이 크다.', { size: 13, color: '666666' }),
});

const W2 = [1450, 2600, 1300, 4926];
const T2_HEAD = ['구분', '기관·조직 (담당자)', '주요 분야', '내용 및 성과'];
const T2 = [
  ['자동화\n(로봇 DBTL)', '한국생명공학연구원(KRIBB)\n국가바이오파운드리사업단(김하성 박사)\n— KAIST 공동 K-바이오파운드리', '합성생물학',
   '설계–제작–시험–학습(DBTL) 사이클을 로봇으로 자동화. {kribb:바이오파운드리 실험 전 과정을 4단계로 표준화한 국제 공동 운영체계 프레임워크를 한국 주도로 개발}(미·영·싱가포르 등 참여)'],
  ['자동화\n(기업)', 'LG화학 자동화 실험실 / 삼성 SAIT', '화학·소재',
   '{lgchem:LG화학은 반복 실험을 자동화한 실험 환경을 구축}하고 AI 기반 데이터 해석까지 가능한 지능형 융합 실험실을 목표로 제시. {sedaily80:SAIT는 ‘Self-Driving Lab’을 연구실의 미래상으로 제시}하고 협의체 실증·확산 분과에 참여'],
  ['자동화\n(장비기업)', '에이블랩스', '라이프사이언스 실험 자동화',
   '{ablelabs:액체핸들링 로봇 NOTABLE·NOTABLE96·SUITABLE과 로봇암 연동 랩 자동화 플랫폼을 공급}, 100억 원 투자를 유치하며 SDL 구축을 장기 목표로 추진'],
  ['AI 실험설계', '서울대 화학생물공학부 정유성 교수팀\n(자율실험실 산·학·연 협의체 위원장)', 'AI 소재 역설계',
   '{snujung:LLM을 활용해 합성이 어려웠던 신소재를 실제 합성 가능한 형태로 재설계하는 기술 개발}. 앞서 AI 역설계로 신물질 4종을 발굴하는 등 실험 전 단계의 후보 탐색을 대체'],
  ['드라이랩\n(AI 모델)', '한국화학연구원 주관\nK-MELLODDY 사업단', 'AI 신약(ADMET) 예측',
   '{melloddy:2024.4~2028.12, 총 348억 원 규모의 연합학습 기반 신약개발 가속화 프로젝트}. 민감 데이터를 공유하지 않고 병원·제약사가 각자 학습하는 ADMET 예측모델(FAM) 개발'],
  ['드라이랩\n(인프라)', 'KISTI 한국과학기술정보연구원', '연구데이터·계산과학',
   '{kisti:국가연구데이터플랫폼 DataON}과 ScienceON LAB의 가상실험·분석 환경, AI 데이터 공유·활용 서비스를 통해 실험 없는 계산·데이터 기반 연구를 지원'],
  ['자동화\n(구축 예정)', '연세대학교 K-NIBRT 사업단', '바이오의약품',
   '{hitnews:보건복지부 AI 활용 신약개발 사업의 바이오의약품 분야 SDL 실습 인프라 구축 기관으로 지정}(유기합성 분야는 한국제약바이오협회가 담당)'],
];
const t2 = dataTable(W2, T2_HEAD, T2.map((r) => r.map((c) => c.split('\n'))));


// ---------- 4-2. 정책·생태계 동향과 현재의 한계 ----------
const h25 = new Paragraph({
  spacing: { before: 240, after: 60 },
  children: [new TextRun({ text: '참고. 정책·생태계 동향과 현재의 한계', bold: true, size: 17, color: NAVY, font: KO })],
});
const bulletLines = [
  '국가 차원 결집 — 2026년 8월 11일 출범한 {hellodd:자율실험실 산·학·연 협의체}는 파크시스템스·에이치비솔루션·에이블랩스 등 장비·자동화 기업이 참여하는 「기술·플랫폼·표준」, 기계연·재료연·생명연·KIST 등 출연(연)의 「연구·운영」, 삼성종합기술원·SK하이닉스·LG화학·현대차 등 수요기업의 「실증·확산」 3개 분과로 구성된다(공동간사: 국가과학AI연구센터·KBSI).',
  '분야별 확산 — 소재 분야는 {mt:이차전지·반도체 등 전략분야별 ‘자율실험센터’와 국가소재연구데이터통합플랫폼}을 함께 구축해 「소재 AI 모델–자율실험–데이터플랫폼」을 연결하는 방향이고, 바이오 분야는 {herald:K-문샷 신약개발과 연계한 AI-네이티브 자율실험실 6곳}이 2026년부터 착수했다.',
  '남은 과제 — 개별 실험실 단위의 성과는 축적됐지만, {hitnews:장비 간 연계성과 실험데이터 표준화는 여전히 보완이 필요하다}는 지적이 이어진다. 협의체가 표준 레퍼런스 실험실과 공통 인터페이스 정립을 우선 과제로 삼은 이유다.',
];
const bullets = bulletLines.map((t) => new Paragraph({
  spacing: { before: 30, after: 30, line: 210, lineRule: 'auto' },
  indent: { left: 220, hanging: 160 },
  children: [new TextRun({ text: '▪  ', size: 14, color: NAVY, font: KO }), ...runs(t, { size: 15 })],
}));

// ---------- 5. 출처 ----------
const h3 = new Paragraph({
  spacing: { before: 240, after: 60 },
  children: [new TextRun({ text: '출처', bold: true, size: 17, color: NAVY, font: KO })],
});

const srcList = Object.values(S).sort((a, b) => a.n - b.n);
const half = Math.ceil(srcList.length / 2);
const colA = srcList.slice(0, half), colB = srcList.slice(half);
const srcRows = [];
for (let i = 0; i < half; i++) {
  const mk = (s) => {
    if (!s) return new TableCell({ width: { size: CONTENT_W / 2, type: WidthType.DXA }, children: [new Paragraph('')] });
    return new TableCell({
      width: { size: Math.floor(CONTENT_W / 2), type: WidthType.DXA },
      margins: { top: 10, bottom: 10, left: 0, right: 120 },
      children: [new Paragraph({
        spacing: { before: 14, after: 14, line: 205, lineRule: 'auto' },
        children: [
          new TextRun({ text: `[${s.n}] `, size: 13, color: '444444', font: KO }),
          new ExternalHyperlink({ link: s.url, children: [new TextRun({ text: s.t, size: 13, color: '1B54B5', underline: {}, font: KO })] }),
        ],
      })],
    });
  };
  srcRows.push(new TableRow({ children: [mk(colA[i]), mk(colB[i])] }));
}
const srcTable = new Table({
  columnWidths: [Math.floor(CONTENT_W / 2), Math.floor(CONTENT_W / 2)],
  width: { size: Math.floor(CONTENT_W / 2) * 2, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
    left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
    insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
  },
  rows: srcRows,
});

const footer = new Paragraph({
  spacing: { before: 100 },
  children: runs('본 보고서의 사례·수치는 위 공개 자료에 보도·발표된 내용을 그대로 인용한 것이다. 국내 SDL 전반의 학술적 개관은 {rsc:Digital Discovery(2026) 리뷰 논문}을 참고할 수 있다.', { size: 12, color: '666666' }),
});

// ---------- 문서 조립 ----------
const doc = new Document({
  styles: {
    default: { document: { run: { font: KO, size: 15, color: '222222' }, paragraph: { spacing: { line: 220, lineRule: 'auto' } } } },
    characterStyles: [{ id: 'Hyperlink', name: 'Hyperlink', basedOn: 'DefaultParagraphFont', run: { color: '1B54B5', underline: {} } }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: convertMillimetersToTwip(297) },
        margin: { top: MARGIN, bottom: convertMillimetersToTwip(12), left: MARGIN, right: MARGIN },
      },
    },
    children: [title, subtitle, boxTable, h1, t1, legend, h2, note2, t2, h25, ...bullets, h3, srcTable, footer],
  }],
});

const out = process.argv[2] || '/home/user/kist/report/국내_자율실험실_현황_사례중심.docx';
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log('written:', out, buf.length, 'bytes'); });
