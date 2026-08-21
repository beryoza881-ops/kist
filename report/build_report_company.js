// 국내 자율실험실(SDL) 현황 — 기업 중심 보고서 생성 스크립트
const path = require('path');
const SCRATCH = '/tmp/claude-0/-home-user-kist/f7349fbc-9d61-5776-85de-581602fe2d75/scratchpad';
const docx = require(path.join(SCRATCH, 'node_modules', 'docx'));
const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, convertMillimetersToTwip,
} = docx;
const fs = require('fs');

const KO = { ascii: '맑은 고딕', eastAsia: '맑은 고딕', hAnsi: '맑은 고딕', cs: '맑은 고딕' };
const NAVY = '1D3567';
const LINE = 'B9C2D4';
const HEADFILL = '1D3567';
const BOXFILL = 'EEF2F9';
const ALTFILL = 'F7F9FC';

const PAGE_W = convertMillimetersToTwip(210);
const MARGIN = convertMillimetersToTwip(14);
const CONTENT_W = PAGE_W - 2 * MARGIN;

const S = {
  hellodd:   { n: 1,  url: 'https://www.hellodd.com/news/articleView.html?idxno=112789',
               t: '헬로디디, "韓 첫 \'자율실험실\' 산·학·연 협의체 출범"(2026.8.11)' },
  sedaily80: { n: 2,  url: 'https://www.sedaily.com/article/20077980',
               t: '서울경제, "\'자율실험실 생태계\' 시동…80여개 기관 참여"(2026.8)' },
  zdnet:     { n: 3,  url: 'https://zdnet.co.kr/view/?no=20260811095548',
               t: 'ZDNet Korea, "로봇이 알아서 연구…자율실험실 구축 \'시동\'"(분과별 참여기업)' },
  kaistmse:  { n: 4,  url: 'https://mse.kaist.ac.kr/index.php?mid=mse_research_highlight_en&document_srl=387700',
               t: 'KAIST 신소재공학과, "서동화 교수팀-포스코홀딩스, 로봇팔·AI로 소재 혁신"' },
  hankyung12:{ n: 5,  url: 'https://www.hankyung.com/article/2025080386861',
               t: '한국경제, "연구자 없이 AI·로봇이 실험…데이터 확보 12배 늘었다"(2025.8.3)' },
  newsis:    { n: 6,  url: 'https://www.newsis.com/view/NISX20250801_0003275713',
               t: '뉴시스, "KAIST·포스코, \'자율탐색 연구실\' 공개"(2025.8.1)' },
  mtlg:      { n: 7,  url: 'https://www.mt.co.kr/industry/2025/09/22/2025092208113590509',
               t: '머니투데이, "LG화학, 화학업계 최초 \'로봇 자동화 실험실\' 구축"(2025.9.22)' },
  lgblog:    { n: 8,  url: 'https://blog.lgchem.com/2025/12/11_automation_lab/',
               t: 'LG화학, "안전은 올리고 속도는 빠르게, LG화학 자동화 실험실"(2025.12.11)' },
  lgrelease: { n: 9,  url: 'https://www.lg.co.kr/media/release/29390',
               t: 'LG, "LG화학 로봇 자동화 실험실 구축" 보도자료' },
  jw:        { n: 10, url: 'https://www.economytalk.kr/news/articleView.html?idxno=422019',
               t: '이코노미톡뉴스, "AI가 설계하고 로봇이 합성…JW중외제약 신약개발"' },
  dailypharm:{ n: 11, url: 'https://www.dailypharm.com/user/news/339565',
               t: '데일리팜, "AI가 찾고 로봇이 만든다…제약사 신약개발 새 공식"' },
  news1:     { n: 12, url: 'https://www.news1.kr/bio/pharmaceutical-bio/5986333',
               t: '뉴스1, "제약바이오 R&D, AI·로봇 기반 \'자율실험실\' 시대 열렸다"' },
  platum:    { n: 13, url: 'https://platum.kr/archives/232778',
               t: '플래텀, "실험실로 들어온 \'자율주행\'… 셀프 드라이빙 랩을 아시나요?"' },
  etnewssk:  { n: 14, url: 'https://www.etnews.com/20250715000111',
               t: '전자신문, "세계 최고 자동화 팹 만든다…SK하이닉스, 용인 클러스터에 AI 도입"(2025.7.15) — 양산 팹 사례' },
  ecopro1:   { n: 15, url: 'https://www.etnews.com/20260604000155',
               t: '전자신문, "에코프로, AX 전면 추진…2028년 전 부문 AI 도입"(2026.6.4)' },
  ecopro2:   { n: 16, url: 'https://zdnet.co.kr/view/?no=20260604140329',
               t: 'ZDNet Korea, "에코프로, AI로 제품 개발-양산 기간 절반 감축 목표"(2026.6.4)' },
  cj:        { n: 17, url: 'https://www.foodnews.co.kr/news/articleView.html?idxno=93781',
               t: '식품저널, "CJ제일제당 바이오파운드리 구축 현장 방문"' },
  seoulbio:  { n: 16, url: 'https://www.seoul.co.kr/news/plan/seoulK-bioweek2026/2026/06/09/20260609015001',
               t: '서울신문, "AI 유전체 설계하면 로봇이 검증"[2026 서울 K-바이오 위크]' },
  ablelabs:  { n: 17, url: 'https://www.thebionews.net/news/articleView.html?idxno=26684',
               t: '더바이오뉴스, "에이블랩스, 100억 투자 발판 \'자율실험실\' 속도"' },
  ablehit:   { n: 18, url: 'https://www.hitnews.co.kr/news/articleView.html?idxno=47443',
               t: '히트뉴스, "에이블랩스, \'노터블\' 사업화 박차…美 시장 정조준"' },
  nf1:       { n: 19, url: 'https://www.mt.co.kr/future/2025/10/17/2025101715192885814',
               t: '머니투데이, "\'AI로 신소재개발\' 나노포지에이아이, K-딥테크 왕중왕전 부총리상"(2025.10.17)' },
  nf2:       { n: 20, url: 'https://www.unicornfactory.co.kr/article/2025101710104547890',
               t: '유니콘팩토리, "나노포지에이아이 \'10년 걸릴 신소재 개발, AI 통해 한달로 단축\'"' },
  nf3:       { n: 21, url: 'https://core.asiae.co.kr/article/2025071508274502133',
               t: '아시아경제, "\'디지털 소재 연구\' 나노포지에이아이, 퓨처플레이·매쉬업벤처스 시드투자 유치"' },
  om1:       { n: 22, url: 'https://dealsite.co.kr/articles/158012',
               t: '딜사이트, "신동명 오믹스AI 대표 \'단백체 지도가 신약개발 방향 바꾼다\'"' },
  om2:       { n: 23, url: 'https://v.daum.net/v/20260421091346223',
               t: '"오믹스에이아이, Revvity와 \'자율주행 실험실\' 구축 협력 확대"(2026.4.21)' },
  sbl:       { n: 24, url: 'https://samsungbiologics.com/kr/media/bio-story/redefining-the-standard-with-advanced-technology-samsung-biologics-plant-5',
               t: '삼성바이오로직스, "첨단 기술로 다시 쓰는 표준 | 제5공장"' },
  sblg:      { n: 20, url: 'https://www.g-enews.com/article/Bio-Pharma/2025/05/20250521085942743d7a510102_1',
               t: '글로벌이코노믹, "삼성바이오로직스, AI·로봇으로 생산 효율 극대화"' },
  standigm:  { n: 21, url: 'https://www.biospectator.com/news/view/6065',
               t: '바이오스펙테이터, "스탠다임, AI기반 신약개발 혁신성 입증"' },
  hits:      { n: 22, url: 'https://kidd.co.kr/news/245880',
               t: '산업일보(KIDD), "\'30초면 결합 예측 끝\'…히츠, AI 신약개발 속도전"' },
  galux:     { n: 23, url: 'https://www.hankyung.com/article/202608201810i',
               t: '한국경제, "갤럭스, 항체 발굴 넘어 AI로 설계…글로벌 경쟁력 입증"(2026.8.20)' },
  virtuallab:{ n: 24, url: 'https://www.etnews.com/20230912000154',
               t: '전자신문, "버추얼랩, 소재 데이터 플랫폼 \'D3스퀘어\' 출시"' },
  lotte:     { n: 25, url: 'https://www.smarttoday.co.kr/ko-kr/articles/108746',
               t: '스마트투데이, "화학에도 분 AI 바람…LG·롯데·금호 앞장서" — 롯데 AI 활용' },
  hitnews:   { n: 26, url: 'https://www.hitnews.co.kr/news/articleView.html?idxno=77438',
               t: '히트뉴스, "AI 신약개발 다음은 \'자율실험실\'…인프라 구축 필요"' },
  domestic:  { n: 27, url: 'https://www.hellodd.com/news/articleView.html?idxno=112666',
               t: '헬로디디, "\'외산 장비만 바라볼 수 없다\'···韓 \'자율랩\' 생태계 만든다"' },
  hellodd2:  { n: 28, url: 'https://www.hellodd.com/news/articleView.html?idxno=111104',
               t: '헬로디디, "AI가 설계, 로봇은 밤샘 실험…\'자율랩\' 전쟁"(2026.3.12)' },
  bizhankook:{ n: 28, url: 'https://www.bizhankook.com/bk/article/32660',
               t: '비즈한국, "제약·바이오도 피지컬 AI 열풍"' },
  kidd:      { n: 29, url: 'https://kidd.co.kr/news/245879',
               t: '산업일보(KIDD), "실험실 자동화, \'자율 랩\'으로 향한다"' },
};

Object.values(S).forEach((s, i) => { s.n = i + 1; });  // 정의 순서대로 자동 번호 부여

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

// ---------- 제목 ----------
const title = new Paragraph({
  spacing: { after: 40, line: 240, lineRule: 'auto' },
  children: [new TextRun({ text: '국내 자율실험실(Self-Driving Lab) 현황 — 기업 편', bold: true, size: 26, color: NAVY, font: KO })],
});
const subtitle = new Paragraph({
  spacing: { after: 140 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } },
  children: [new TextRun({ text: '기준일: 2026년 8월  |  공개 보도자료·기업 발표 기준 정리  |  대학·출연(연) 사례는 별도 보고서 참조', size: 14, color: '666666', font: KO })],
});

// ---------- 개요 박스 ----------
const boxText =
  '기업은 “위험·반복 실험의 무인화”와 “개발기간 단축”이라는 사업 목적에서 오래전부터 자동화 실험실을 쌓아 왔고, ' +
  '자율실험실은 그 위에 AI 의사결정을 얹는 방식으로 진행되는 경우가 다수다. ' +
  '{kaistmse:포스코홀딩스}(이차전지 양극재)와 {jw:JW중외제약}(항암 신약)은 이미 AI가 다음 실험을 결정하는 폐루프에 도달했고, ' +
  '{mtlg:LG화학}·{cj:CJ제일제당}은 분석·합성 공정의 무인화 단계에 있다. ' +
  '2026년 8월 출범한 {sedaily80:자율실험실 산·학·연 협의체}에는 장비·자동화·AI 기업과 반도체·소재 수요기업을 합쳐 50개사가 참여하며, ' +
  '{zdnet:삼성종합기술원·SK하이닉스·LG화학·현대차는 「실증·확산」 분과, 파크시스템스·에이블랩스 등 장비기업은 「기술·플랫폼·표준」 분과}에 이름을 올렸다.';

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

// ---------- 표 1 ----------
const h1 = new Paragraph({
  spacing: { before: 220, after: 70 },
  children: [new TextRun({ text: '표 1. 기업 자율실험실(폐루프 SDL) 주요 사례', bold: true, size: 17, color: NAVY, font: KO })],
});

const W1 = [1450, 2320, 1150, 4092, 1306];
const T1_HEAD = ['기업', '실험실·조직 (담당 조직)', '주요 분야', '주요 성과', '자율실험 정도'];
const T1 = [
  ['포스코홀딩스', '미래기술연구원\nLIB소재연구센터 자율탐색 실험실\n(KAIST 서동화 교수팀 공동)', '이차전지 양극재',
   '{newsis:정량·혼합·소성·분석 전 과정을 연구자 개입 없이 수행}. {hankyung12:기존 연구자 기반 실험 대비 12배 많은 소재 데이터 확보·소재 탐색 시간 93% 단축}, 고속 소결 도입으로 합성 속도 50배 개선. {kaistmse:2026년 이후 업그레이드판을 자체 연구소에 적용 예정}',
   '★★★\n폐루프 자율탐색(24h)'],
  ['LG화학', '대전 기술연구원 분석연구소\n{lgrelease:로봇 자동화 실험실}\n(Autonomous Smart Lab)', '배터리 양극재 원료 정밀분석',
   '{mtlg:국내 화학업계 최초}. 시료 출고→전처리→분석→폐기까지 로봇이 일괄 수행하고 데이터가 자동 입력되며, 고온·고농도 산 처리 등 위험 공정을 무인화해 24시간 365일 실험. {lgblog:2026년 마곡 연구소로 확대, AI 데이터 해석까지 가능한 지능형 융합 실험실 목표}',
   '★★☆\n분석 전과정 무인화'],
  ['JW중외제약', 'AI 신약 플랫폼 ‘제이웨이브(JWave)’\n+ 로봇 합성 자동화 자율 연구실', '항암 신약 후보물질',
   '{jw:4만여 개 화합물·500여 종 세포주·오가노이드 데이터를 기반으로 AI가 후보물질을 설계 → 로봇이 자동 합성 → 결과를 AI가 재학습}하는 폐루프 구조를 갖췄다. {dailypharm:보건복지부 구조기반 AI 신약개발 지원사업 주관기관(3년)으로 선정}되며 본격 가동에 들어간 단계',
   '★★☆\n폐루프 구성, 사업 착수(2026~)'],
  ['삼성전자 SAIT\n(삼성종합기술원)', '{platum:소재 자율 합성 시스템}\n/ Self-Driving Lab', '첨단 소재',
   'AI라는 ‘뇌’와 로봇이라는 ‘손발’을 결합해 {sedaily80:24시간 스스로 가설을 세우고 실험·데이터를 쌓는 SDL을 연구실의 미래상으로 제시}. 국내 SDL 적용 영역이 합성생물학 DBTL에서 소재 자율 합성으로 확대된 대표 사례로 인용됨',
   '★★☆\n소재 자율합성 적용'],
  ['CJ제일제당', 'BIO연구소 {cj:바이오파운드리}', '균주 개발·발효 소재',
   '균주 개발 사이클(설계–제작–시험–학습)을 로봇으로 자동화한 연구 시설. 내부 데이터를 학습한 {seoulbio:‘바이오 에이전트 AI’ 도입으로 균주 평가 기간을 5개월 → 1개월로 단축}. 개발된 균주는 자사 발효 생산라인으로 연결',
   '★★☆\nDBTL 자동화 + AI 에이전트'],
  ['에코프로', '{ecopro1:전사 AX 로드맵 중 자율실험실}\n(별도로 자율제조 공장도 추진)', '이차전지 양극재·전구체',
   '{ecopro1:피지컬 AI를 도입해 자율실험실 구축을 추진}, 위험한 실험은 로봇이 대신하고 365일 24시간 상시 실험 체계를 목표. {ecopro2:AI가 축적된 실험 데이터를 학습해 소재 물성을 예측하고 최적 실험 조건을 도출, 연구개발~양산 기간 50% 단축 목표}(2028년 전 부문 AI 도입)',
   '★☆☆\n구축 추진 단계(~2028)'],
  ['나노포지에이아이', '디지털 소재 연구소\n(소재 설계 AI + 로보틱스 합성 실험실)\n김동현 대표·배재원 CTO', '신소재 설계·합성',
   '{nf3:AI 예측모델과 로보틱스 자동화를 결합해 소재 설계–합성–공정 최적화 전 과정을 자동화}. {nf2:기존 7년 이상 걸리던 R&D를 1개월 이내로 단축하고 비용 최대 80% 절감이 목표}. {nf1:2025 K-딥테크 스타트업 왕중왕전 대상(부총리상)}, 대주전자재료와 이차전지용 고체 전극 신소재 PoC',
   '★★☆\n소규모 SDL 구축(2025 창업)'],
  ['오믹스에이아이', '피지컬 AI 자율주행 실험실\n(Revvity 공동 구축)\n신동명 대표', '단백체(프로테오믹스) 분석',
   '{om1:사람 개입 없이 24시간 돌아가는 단백체 분석 자동화 인프라는 이미 구축을 마쳐} 분석 기간을 6주 → 2주 이내로 단축했다. 여기에 {om2:Revvity의 자동화 장비와 자사 AI 분석기술을 결합해 실험 설계–데이터 생성–분석–해석까지 지능화하는 AI 피드백을 도입하는 중}',
   '★★☆\n24시간 자동화 완료, AI 피드백 도입 중'],
  ['에이블랩스', '자율실험실 장비·플랫폼 사업\n(인천 송도)', '라이프사이언스 실험 자동화',
   '{ablehit:액체핸들링 로봇 노터블·노터블96·수터블과 로봇암·분석장비를 연결한 실험실 전(全)자동화 사업}. {ablelabs:100억 원 규모 시리즈A 마무리, 미국 대형 제약사 PoC 진행, 매출 전년 대비 3배 전망}',
   '★★☆\nSDL 공급자(자체 SDL 목표)'],
];
const t1 = dataTable(W1, T1_HEAD, T1.map((r) => r.map((c) => c.split('\n'))));

const legend = new Paragraph({
  spacing: { before: 60, after: 0, line: 200, lineRule: 'auto' },
  children: runs('※ 자율실험 정도  ★★★ 설계–실험–분석–재설계가 폐루프로 무인 순환  |  ★★☆ 실험 전 과정 자동화 + AI 부분 적용(사람 개입 일부 잔존)  |  ★☆☆ 구축 추진 단계', { size: 13, color: '666666' }),
});

const scopeNote = new Paragraph({
  spacing: { before: 40, after: 0, line: 200, lineRule: 'auto' },
  children: runs('※ 범위  이 보고서의 ‘실험실’은 R&D 단계의 실험·분석 공간을 가리키며, 양산 라인(팹·공장)의 자동화는 성격이 달라 표에서 제외했다. {etnewssk:SK하이닉스 용인 클러스터의 적응형 계측 샘플링(AMS)·추적 자동화 품질(TAQ)}, {sbl:삼성바이오로직스 제5공장의 AI·자율주행로봇}, {lotte:롯데케미칼의 AI 품질검사}가 여기에 해당한다. 이들 기업은 자율실험실 협의체 「실증·확산」 분과의 수요기업으로 참여 중이다(아래 참고).', { size: 13, color: '666666' }),
});

// ---------- 표 2 ----------
const h2 = new Paragraph({
  pageBreakBefore: true,
  spacing: { before: 0, after: 40 },
  children: [new TextRun({ text: '표 2. (별도 구분) 자율실험실은 아니나 실험 자동화·AI 실험설계·드라이랩에 해당하는 기업', bold: true, size: 17, color: NAVY, font: KO })],
});
const note2 = new Paragraph({
  spacing: { after: 70, line: 200, lineRule: 'auto' },
  children: runs('AI가 다음 실험을 결정하는 폐루프가 확인되지 않아 표 1과 분리했다. 드라이랩 기업은 실험 자체가 없는 계산·예측 전문 사업모델이다.', { size: 13, color: '666666' }),
});

const W2 = [1450, 2600, 1300, 4968];
const T2_HEAD = ['구분', '기업 (조직)', '주요 분야', '내용 및 성과'];
const T2 = [
  ['AI 실험설계', '대웅제약\n(제약업계 최초 AI 전담조직)', '신약 후보물질 탐색',
   '{dailypharm:8억 종 규모 화합물 데이터베이스를 기반으로 한 AI 신약개발 플랫폼 ‘데이지(DAISY)’ 운영}. AI 기반 활성물질 탐색과 선도물질 확보 체계를 24시간 가동해 연구 효율을 높임'],
  ['드라이랩', '스탠다임(Standigm)', 'AI 신약 탐색',
   '{standigm:타깃 발굴 ‘Standigm ASK’, 신규 물질 생성 ‘Standigm BEST’} 등으로 타깃 발굴–유효물질 탐색–선도물질 최적화–전임상 후보 확보까지 신약 탐색 전주기를 AI 워크플로우로 포괄'],
  ['드라이랩', '히츠(HITS)', 'AI 신약 플랫폼',
   '{hits:클라우드 플랫폼 ‘하이퍼랩(HyperLab)’ — 물리 기반 딥러닝으로 30초 이내에 약물–표적 단백질 상호작용을 예측}하고 타깃 정의·가상탐색·분자설계·ADME/T 예측을 통합 지원'],
  ['드라이랩', '갤럭스(Galux)', 'AI 단백질·항체 설계',
   '{galux:결합력·안정성·면역원성을 갖춘 단백질을 데이터 의존 없이 설계하는 디노보 플랫폼 ‘갤럭스디자인’}. AACR 2026에서 AI 설계 이중항체 후보물질의 전임상 결과 발표'],
  ['드라이랩\n(소재)', '버추얼랩', '소재 시뮬레이션·데이터',
   '{virtuallab:클라우드 소재 시뮬레이션 ‘머터리얼스 스퀘어’와 데이터 기반 소재 R&D 플랫폼 ‘D3스퀘어’ 운영}. KIST 출신 연구진이 창업, 반도체·배터리·금속 소재 연구자에게 계산 환경을 제공'],
  ['기술·장비 개발', '파크시스템스·에이치비솔루션\n·로봇앤드디자인 (+ 에이블랩스)', '계측·실험장비 자동화',
   '{zdnet:자율실험실 협의체 「기술·플랫폼·표준」 분과(32명)에 참여}해 장비 인터페이스와 데이터 표준을 정립. {domestic:한국기계연구원(KIMM)과 함께 로봇이 직접 다룰 수 있는 국산 실험장비·제어 소프트웨어를 개발해 ‘한국형 자율랩 생태계’를 만드는 것이 목표}'],
  ['AI 실험설계\n(연구지원)', '롯데정밀화학', '정밀화학 신소재',
   '{lotte:특허·논문을 분석·추천하는 AI 플랫폼을 신소재 연구지원에 활용}. 계열사 롯데케미칼의 AI 품질검사·컬러매칭은 양산 공정 적용 사례로, 실험실 자동화와는 구분된다'],
];
const t2 = dataTable(W2, T2_HEAD, T2.map((r) => r.map((c) => c.split('\n'))));

// ---------- 참고 ----------
const h25 = new Paragraph({
  spacing: { before: 240, after: 60 },
  children: [new TextRun({ text: '참고. 기업 생태계 구도와 현재의 한계', bold: true, size: 17, color: NAVY, font: KO })],
});
const bulletLines = [
  '수요–공급 구도 — {zdnet:협의체 「실증·확산」 분과에는 삼성종합기술원·SK하이닉스·LG화학·현대자동차 등 수요기업이, 「기술·플랫폼·표준」 분과에는 파크시스템스·에이치비솔루션·에이블랩스·로봇앤드디자인 등 장비·자동화 기업 32명이 참여}한다(분과 주도: 서정주 한국기초과학지원연구원 본부장). 참여 기업은 장비·자동화·AI 기업과 반도체·소재 수요기업을 합쳐 50개사이며, 「실증·확산」 분과에는 18명이 포진했다.',
  '제약·바이오가 가장 빠르다 — {kidd:보건복지부·한국보건산업진흥원이 SDL 이론·실습 교육(4억 원)과 SDL 실습 인프라 구축(30억 원) 두 트랙으로 지원}하며, 유기합성은 한국제약바이오협회, 바이오의약품은 연세대 K-NIBRT가 맡는다. {hellodd2:AI가 설계하고 로봇이 밤샘 합성하는 ‘자율랩’이 차세대 항암제 개발의 핵심 인프라로 부상}했다는 평가다.',
  '남은 과제 — 기업 사례 대부분은 아직 ‘분석·합성 자동화’ 단계이고, AI가 다음 실험을 스스로 결정하는 폐루프까지 간 사례는 일부다. {hitnews:개별 장비 간 연계성과 실험데이터 표준화가 공통 과제}로 지적된다. 공급 측에서는 {domestic:자율실험실 장비를 만들 수 있는 국내 기업이 있으나 규모가 영세해, 정부 R&D 구매를 통한 시장 개방이 필요하다}는 지적과 함께 외산 장비 의존 탈피가 과제로 꼽힌다.',
];
const bullets = bulletLines.map((t) => new Paragraph({
  spacing: { before: 30, after: 30, line: 210, lineRule: 'auto' },
  indent: { left: 220, hanging: 160 },
  children: [new TextRun({ text: '▪  ', size: 14, color: NAVY, font: KO }), ...runs(t, { size: 15 })],
}));

// ---------- 출처 ----------
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
    if (!s) return new TableCell({ width: { size: Math.floor(CONTENT_W / 2), type: WidthType.DXA }, children: [new Paragraph('')] });
    return new TableCell({
      width: { size: Math.floor(CONTENT_W / 2), type: WidthType.DXA },
      margins: { top: 10, bottom: 10, left: 0, right: 120 },
      children: [new Paragraph({
        spacing: { before: 6, after: 6, line: 200, lineRule: 'auto' },
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
  children: runs('본 보고서의 사례·수치는 위 공개 자료에 보도·발표된 내용을 그대로 인용한 것이다. 기업 내부 구축 현황은 비공개인 경우가 많아, 표에 없는 기업이 자율실험실을 운영하지 않는다는 뜻은 아니다.', { size: 12, color: '666666' }),
});

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
    children: [title, subtitle, boxTable, h1, t1, legend, scopeNote, h2, note2, t2, h25, ...bullets, h3, srcTable, footer],
  }],
});

const out = process.argv[2] || '/home/user/kist/report/국내_자율실험실_현황_기업편.docx';
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log('written:', out, buf.length, 'bytes'); });
