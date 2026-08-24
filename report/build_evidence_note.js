// "장비는 있어도 운영이 불가능하다" 주장별 근거 검증 — docx 생성
const path = require('path');
const SCRATCH = '/tmp/claude-0/-home-user-kist/f7349fbc-9d61-5776-85de-581602fe2d75/scratchpad';
const docx = require(path.join(SCRATCH, 'node_modules', 'docx'));
const { Document, Packer, Paragraph, TextRun, ExternalHyperlink, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, convertMillimetersToTwip } = docx;
const fs = require('fs');

const KO = { ascii: '맑은 고딕', eastAsia: '맑은 고딕', hAnsi: '맑은 고딕', cs: '맑은 고딕' };
const NAVY = '1D3567', LINE = 'B9C2D4', HEADFILL = '1D3567', BOXFILL = 'EEF2F9', ALTFILL = 'F7F9FC';
const OKFILL = 'EAF3EA', WARNFILL = 'FDF3E7', BADFILL = 'FBECEC';
const GREEN = '2E7D32', ORANGE = 'B26A00', RED = 'B3261E';
const PAGE_W = convertMillimetersToTwip(210), MARGIN = convertMillimetersToTwip(14);
const CONTENT_W = PAGE_W - 2 * MARGIN;

const S = {
  hd90:   { url: 'https://www.hellodd.com/news/articleView.html?idxno=110899',
            t: '헬로디디, "연구실 기본장비 90% 이상 수입품···정부, 2~3년 내 국산 대체 나선다"' },
  brief:  { url: 'https://www.korea.kr/news/policyNewsView.do?newsId=148959754',
            t: '대한민국 정책브리핑, "필수인데 수입에 의존?…범용 연구장비 \'국산화\' 시동"' },
  dt:     { url: 'https://www.digitaltoday.co.kr/news/articleView.html?idxno=632408',
            t: '디지털투데이, "과기정통부, 수입 의존도 높은 범용 연구장비 국산화…전담 분과 신설"' },
  bio90:  { url: 'https://www.hellodd.com/news/articleView.html?idxno=98970',
            t: '헬로디디, "미래 먹거리 \'바이오장비\'···국산화 생태계 조성 澤 아닌 必"(한국기계연구원 정책포럼)' },
  eco:    { url: 'https://www.hellodd.com/news/articleView.html?idxno=112666',
            t: '헬로디디, "\'외산 장비만 바라볼 수 없다\'···韓 \'자율랩\' 생태계 만든다"' },
  sisa:   { url: 'https://www.sisajournal-e.com/news/articleView.html?idxno=407214',
            t: '시사저널e, "신상 에이블랩스 대표 \'클라우드 기반 전자동화 실험실 플랫폼 꿈꾼다\'"' },
  hit:    { url: 'https://www.hitnews.co.kr/news/articleView.html?idxno=47443',
            t: '히트뉴스, "에이블랩스, 액체 핸들링 로봇 \'노터블\' 사업화 박차…美 시장 정조준"' },
  zdnet:  { url: 'https://zdnet.co.kr/view/?no=20260811095548',
            t: 'ZDNet Korea, "로봇이 알아서 연구…자율실험실 구축 \'시동\'"(협의체 분과별 과제)' },
  nist:   { url: 'https://www.nist.gov/programs-projects/development-standards-support-modular-and-autonomous-laboratory-ecosystem',
            t: 'NIST, "Development of Standards to Support a Modular and Autonomous Laboratory Ecosystem"' },
  slasnew:{ url: 'https://www.selectscience.net/product-news/ansi-accredits-new-slas-microplate-standard-for-well-bottom-elevation/?artID=26395',
            t: 'SelectScience, "ANSI accredits new SLAS microplate standard for well-bottom elevation"' },
  slas:   { url: 'https://www.slas.org/education/ansi-slas-microplate-standards/', t: 'SLAS, ANSI/SLAS 마이크로플레이트 표준' },
  rsc:    { url: 'https://pubs.rsc.org/dd/article/5/5/1968/1244548/Self-driving-laboratories-in-Korea-a-new-era-of',
            t: '"Self-driving laboratories in Korea: a new era of autonomous discovery", Digital Discovery 5, 1968 (2026) — 동료평가 리뷰' },
  kmds:   { url: 'https://kmds.re.kr/en/', t: 'K-MDS(Korea Materials Data Station) — 국가 R&D 소재 데이터 저장소' },
  oecd:   { url: 'https://www.oecd.org/en/publications/access-to-public-research-data-toolkit_a12e8998-en/korea-materials-data-station-k-mds_7b9b0814-en.html',
            t: 'OECD, Access to Public Research Data Toolkit — K-MDS 사례' },
  mordor: { url: 'https://www.mordorintelligence.com/industry-reports/analytical-instrumentation-market',
            t: 'Mordor Intelligence, Analytical Instrumentation Market — 상위 5개사 합계 약 45%' },
  kita:   { url: 'https://www.koreatimes.co.kr/business/companies/20221103/korea-highly-dependent-on-foreign-chip-equipment',
            t: 'The Korea Times / KITA, "Korea highly dependent on foreign chip equipment"' },
  kiria:  { url: 'https://www.kiria.org/portal/policysut/portalPlcyInquiry.do',
            t: '한국로봇산업진흥원, 로봇산업 실태조사(2006년부터 매년 실시되는 정부승인통계)' },
  cobot:  { url: 'https://www.newspost.kr/news/articleView.html?idxno=213502',
            t: '뉴스포스트, "두산로보틱스, 중복 상장 속 AI·가격 경쟁력 우려"(협동로봇 점유율)' },
  cobot2: { url: 'https://www.news2day.co.kr/article/20230707500185',
            t: '뉴스투데이, "12조원 협동로봇 시장 놓고 \'다윗\' 레인보우로보틱스, \'골리앗\' 두산에 도전장"' },
  lhmkt:  { url: 'https://www.globenewswire.com/news-release/2026/07/10/3325342/0/en/south-korea-automated-liquid-handling-system-market-report-published-breaks-down-demand-across-drug-discovery-genomic-research-and-clinical-diagnostics.html',
            t: 'GlobeNewswire, "South Korea Automated Liquid Handling System Market Report"(2026.7)' },
  labmkt: { url: 'https://www.marketresearchfuture.com/reports/south-korea-laboratory-automation-market-50056',
            t: 'MRFR, "South Korea Laboratory Automation Market"' },
  top5:   { url: 'https://www.rootsanalysis.com/key-insights/top-automated-liquid-handlers-manufacturing-companies.html',
            t: 'Roots Analysis, "Top 5 Manufacturers of Automated Liquid Handlers"' },
  bioin:  { url: 'https://www.ibric.org/bric/trend/bio-report.do?articleNo=8693867',
            t: 'BRIC Bio리포트, "바이오산업 소부장 국내현황 및 경쟁력과 발전 방향"' },
  nistp:  { url: 'https://www.sciencedirect.com/science/article/abs/pii/S2590238526001190',
            t: 'H. Joress et al., "Toward a composable, modular laboratory ecosystem for autonomous materials R&D", Matter(2026)' },
  lads:   { url: 'https://reference.opcfoundation.org/specs/OPC-30500-1/1', t: 'OPC Foundation, OPC UA for LADS Part 1(OPC 30500-1)' },
  sila:   { url: 'https://sila-standard.com/faq/', t: 'SiLA Standard FAQ — SiLA 2 개요' },
  emmc:   { url: 'https://emmc.eu/focus-areas/digitalisation-interoperability/', t: 'EMMC, Digitalisation & Interoperability(EMMO 기반 소재 온톨로지)' },
  hk:     { url: 'https://www.hankyung.com/article/2023082171811',
            t: '한국경제, "장비 공짜로 줄게…계약 물량까지 뺏는 외국계 바이오"(2023.8)' },
};
Object.values(S).forEach((x, i) => { x.n = i + 1; });

function emitPlain(out, text, size, bold, color) {
  const re = /\*\*([^*]+)\*\*/g; let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), size, bold, color, font: KO }));
    out.push(new TextRun({ text: m[1], size, bold: true, color, font: KO }));
    last = re.lastIndex;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), size, bold, color, font: KO }));
}
function runs(str, { size = 15, bold = false, color = '222222' } = {}) {
  const out = []; const re = /\{([a-zA-Z0-9]+):([^}]*)\}/g; let last = 0, m;
  while ((m = re.exec(str)) !== null) {
    if (m.index > last) emitPlain(out, str.slice(last, m.index), size, bold, color);
    const src = S[m[1]]; if (!src) throw new Error('unknown key: ' + m[1]);
    const label = m[2].replace(/\*\*/g, ''); const strong = bold || /\*\*/.test(m[2]);
    out.push(new ExternalHyperlink({ link: src.url, children: [new TextRun({ text: label, size, bold: strong, font: KO, color: '1B54B5', underline: {} })] }));
    last = re.lastIndex;
  }
  if (last < str.length) emitPlain(out, str.slice(last), size, bold, color);
  return out;
}
function cell(text, { w, size = 15, bold = false, color = '222222', fill, align = AlignmentType.LEFT } = {}) {
  const paras = (Array.isArray(text) ? text : [text]).map((t) =>
    new Paragraph({ alignment: align, spacing: { before: 12, after: 12, line: 205, lineRule: 'auto' }, children: runs(t, { size, bold, color }) }));
  return new TableCell({ width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    shading: fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined,
    margins: { top: 80, bottom: 80, left: 75, right: 75 }, children: paras });
}
const B = { top: { style: BorderStyle.SINGLE, size: 3, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 3, color: LINE },
  left: { style: BorderStyle.SINGLE, size: 3, color: LINE }, right: { style: BorderStyle.SINGLE, size: 3, color: LINE },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 3, color: LINE }, insideVertical: { style: BorderStyle.SINGLE, size: 3, color: LINE } };
function dataTable(widths, header, rows, rowFills) {
  return new Table({ columnWidths: widths, width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, borders: B,
    rows: [new TableRow({ tableHeader: true, children: header.map((hh, i) => cell(hh, { w: widths[i], size: 15, bold: true, color: 'FFFFFF', fill: HEADFILL, align: AlignmentType.CENTER })) }),
      ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) =>
        cell(c, { w: widths[i], size: 15, bold: i === 0, fill: (rowFills && rowFills[ri]) || (ri % 2 === 1 ? ALTFILL : undefined),
          align: i === 0 ? AlignmentType.CENTER : AlignmentType.LEFT })) }))] });
}
function h(t, color) { return new Paragraph({ spacing: { before: 280, after: 80 }, children: [new TextRun({ text: t, bold: true, size: 18, color: color || NAVY, font: KO })] }); }
function para(t, size) { return new Paragraph({ spacing: { before: 40, after: 40, line: 215, lineRule: 'auto' }, children: runs(t, { size: size || 16 }) }); }
function quote(t) {
  return new Table({ columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
    borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
               left: { style: BorderStyle.SINGLE, size: 14, color: NAVY }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
    rows: [new TableRow({ children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: 'F7F9FC', color: 'auto' }, margins: { top: 100, bottom: 100, left: 170, right: 140 },
      children: [new Paragraph({ spacing: { line: 215, lineRule: 'auto' }, children: runs(t, { size: 16 }) })] })] })] });
}
function box(titleText, lines, fill, accent) {
  return new Table({ columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
    borders: { top: { style: BorderStyle.SINGLE, size: 6, color: accent || NAVY }, bottom: { style: BorderStyle.SINGLE, size: 6, color: accent || NAVY },
               left: { style: BorderStyle.SINGLE, size: 18, color: accent || NAVY }, right: { style: BorderStyle.SINGLE, size: 6, color: accent || NAVY } },
    rows: [new TableRow({ children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: fill || BOXFILL, color: 'auto' }, margins: { top: 120, bottom: 120, left: 180, right: 160 },
      children: [new Paragraph({ spacing: { after: 50, line: 220, lineRule: 'auto' }, children: [new TextRun({ text: titleText, bold: true, size: 16, color: accent || NAVY, font: KO })] }),
        ...lines.map((l) => new Paragraph({ spacing: { before: 22, after: 22, line: 225, lineRule: 'auto' }, children: runs(l, { size: 16 }) }))] })] })] });
}

const c = [];
c.push(new Paragraph({ spacing: { after: 40, line: 240, lineRule: 'auto' },
  children: [new TextRun({ text: '외산 벤더 의존과 생태계 고착화 — 근거 자료집', bold: true, size: 25, color: NAVY, font: KO })] }));
c.push(new Paragraph({ spacing: { after: 150 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } },
  children: [new TextRun({ text: '2026년 8월 기준  |  자율실험실 관련 장비 우선  |  ✅ 확실한 근거   🔶 간접·부분 근거   ❌ 근거 못 찾음', size: 14, color: '666666', font: KO })] }));

c.push(box('한눈에 보기', [
  '자율실험실 관련 장비만 따로 집계한 국산·외산 점유율 통계는 국내에도 영문 자료에도 **없다.** ' +
  '따라서 현재 확보 가능한 근거 중 **자율실험 워크플로에 직결되는 품목**을 골라 쓰는 것이 가장 설득력 있다. ' +
  '가장 강한 카드는 {hd90:마이크로플레이트 리더 외산 비중 100%}와 {eco:개발률 36.4% vs 도입률 1%} 두 가지다.',
  '고착화 논거는 점유율이 아니라 **구조**로 세워야 한다. 자율실험실은 기성 장비를 엮는 것이므로 경쟁력이 ' +
  '규격과 소프트웨어에서 갈리는데, 그 규격이 모두 국외에 있다. 이 진단은 {rsc:동료평가 리뷰가 직접 뒷받침}하므로 반박이 어렵다.',
]));

// 종합 판정
c.push(h('1. 핵심 근거 — 자율실험실 관련 장비의 외산 비중', GREEN));
c.push(para('{hd90:한국표준과학연구원 첨단혁신장비기술정책센터}가 국가연구시설장비 구매현황(2019~2023년)을 분석한 결과다. 자율실험 워크플로와의 관련도를 함께 표시했다.'));
c.push(dataTable([3400, 1500, 1500, 3918], ['품목', '외산 비중', 'SDL 관련도', '자율실험에서의 역할'], [
  ['**마이크로플레이트 리더**', '**100%**', '★★★', '플레이트 기반 자동화의 판독 장비. 리퀴드핸들러가 분주하면 결과를 읽는 쪽이 이 장비다. **자율실험 폐루프에서 “분석” 단계를 담당**'],
  ['가스 크로마토그래피', '91.0%', '★★★', '자동 시료주입기와 결합해 무인 분석에 직결. 촉매·합성 워크플로의 표준 측정 장비'],
  ['증류·농축기', '93.6%', '★★', '시료 전처리 자동화 구간'],
  ['시료절편기', '95.8%', '★★', '시료 준비 자동화 구간'],
  ['오실로스코프 · 스펙트럼 분석기', '**100%**', '★', '전자·물성 측정. 소재 자율실험에서 사용'],
]));
c.push(quote('**인용 포인트** — “연구장비 90% 외산”보다 **“자율실험의 판독 장비인 마이크로플레이트 리더는 100% 외산”** 이 훨씬 강하다. 품목이 특정되고, 자율실험과의 연결이 자명하며, 정부 통계가 출처이기 때문이다.'));

c.push(h('2. 자동화 전용 장비 — 리퀴드핸들러와 로봇팔', GREEN));
c.push(dataTable([2200, 8118], ['품목', '현황'], [
  ['리퀴드핸들러\n(자동분주기)', '{top5:글로벌 상위 5개사(Agilent·Beckman Coulter·Eppendorf·Hamilton·Thermo Fisher)가 모두 외산}. {sisa:국내에서는 2021년 설립된 에이블랩스가 “유일하게 액체 핸들링 자동화 로봇을 만들었다”}고 밝힌다 — 그 전까지는 사실상 100% 외산이었다는 뜻. {lhmkt:한국 시장은 2025년 2,350만 달러}, 국산·외산 분리 점유율 통계는 없다'],
  ['로봇팔\n(협동로봇)', '**여기는 국내 역량이 있다.** {kiria:로봇밀도 1,012대로 세계 최고 수준}, {cobot:두산로보틱스는 중국 제외 글로벌 4위권이면서 국내 1위}({cobot2:글로벌 1위는 덴마크 유니버설로봇}). “장비가 다 외산”이라고 쓰면 이 대목에서 반박당한다'],
  ['분석기기 전반', '{mordor:글로벌 상위 5개사(Agilent·Thermo Fisher·Shimadzu·Danaher·Waters)가 2025년 매출의 약 45%}를 차지. 한국 시장 내 국산 비중을 집계한 자료는 없다'],
]));

c.push(h('3. 가장 강한 한 줄 — 개발률 36.4% vs 도입률 1%', GREEN));
c.push(quote('{eco:국내 바이오장비 개발률은 36.4%인데, 연구현장의 국산 장비 실제 도입률은 1%}다.'));
c.push(para('“못 만든다”가 아니라 **“만들어도 안 쓰인다”** 는 뜻이다. 성능의 문제가 아니라 기존 시스템에 붙지 못하는 **연결의 문제**이므로, 생태계 고착화를 말하려는 논지에 90%보다 정확하게 들어맞는다. 같은 기사는 연구자들이 외산으로 첫 연구를 시작해 성과를 낸 탓에 실험 환경 변화에 거부감이 있다는 구조적 원인도 지적한다.'));

c.push(h('4. 고착화가 일어나는 세 가지 구조', ORANGE));
c.push(dataTable([2200, 8118], ['메커니즘', '근거'], [
  ['① 번들과 소모품 계약', '분석 소프트웨어는 장비를 사면 따라오므로 **장비 점유율이 곧 소프트웨어 점유율**이 된다. 여기에 {hk:외국계 기업은 장비를 무상 제공하는 대신 5~10년간 소모품 구매를 조건으로 거는 영업 방식}을 쓴다. 장비·SW·소모품이 함께 묶여 교체 비용이 커진다'],
  ['② 규격 주도권 부재', '{nistp:자율실험실은 상용 기성품(COTS)을 엮는 구조}라 경쟁력이 **장비를 잇는 규격**에서 갈린다. 그런데 장비 통신은 {lads:OPC UA LADS}(독일 주도)와 {sila:SiLA 2}, 데이터는 Allotrope·AnIML, 소재 온톨로지는 {emmc:EU의 EMMO}가 쥐고 있다. 국내에도 {kmds:K-MDS}({oecd:OECD 사례로 등재})·ModuFlow(서울대)·OCTOPUS OS(KIST) 같은 시도가 있으나 **개별 프로젝트 수준**이며, {zdnet:협의체 「기술·플랫폼·표준」 분과는 2026년 8월 출범}했다'],
  ['③ 연결의 실패', '개발률 36.4% 대비 도입률 1%. 국산 장비가 만들어져도 기존 워크플로에 편입되지 못한다'],
]));

c.push(h('5. 반박하기 가장 어려운 근거 — 동료평가 리뷰', GREEN));
c.push(para('{rsc:Digital Discovery(2026)에 실린 한국 자율실험실 리뷰}는 국내 SDL 확산의 병목을 직접 진단한다. 시장조사 보고서와 달리 동료평가를 거친 자료여서 인용 시 반박 여지가 적다.'));
c.push(quote('한국은 반도체·이차전지·화학 산업 기반 덕에 자율실험실 도입 여건이 좋지만, **높은 초기 투자비, 표준 프로토콜의 부재, 제한적인 산업 투자**가 걸림돌이다. 상호운용성 병목은 **실험 이력(provenance)·워크플로 의미체계(semantics)·예외 및 이벤트 로깅** 수준에 남아 있으며, **오케스트레이션 시스템마다 같은 동작 레이블을 다르게 해석해 워크플로 이식성이 떨어진다.**'));
c.push(para('요지는 분명하다 — **병목은 장비 소유가 아니라 상호운용성이다.** “장비 국산화”가 아니라 “시스템·표준”이 문제라는 논지를 학술 근거로 세울 수 있는 지점이다.'));

c.push(h('6. 바로 쓸 수 있는 문장 — “외산 벤더 장비·시스템 의존, 생태계 고착화 우려”', NAVY));
c.push(para('아래 문장은 위 근거를 그대로 반영해 작성했다. 밑줄 친 부분에 출처 링크가 걸려 있다.', 14));

c.push(new Paragraph({ spacing: { before: 160, after: 50 }, children: [new TextRun({ text: '의존의 실태', bold: true, size: 16, color: NAVY, font: KO })] }));
[
  '국내 연구실의 기본 연구장비는 **90% 이상이 수입품**이며, 특히 {hd90:자율실험의 판독을 담당하는 마이크로플레이트 리더는 외산 비중이 100%}다.',
  '바이오 분야는 {bioin:수요기업의 90% 이상이 해외 소부장에 의존}하고 있다.',
  '실험 자동화의 핵심인 자동분주기(리퀴드핸들러)는 {top5:글로벌 상위 5개사가 모두 외산}이며, {sisa:국내에서는 2021년 설립된 1개사가 유일한 제조사}다.',
].forEach((t, i) => c.push(new Paragraph({ spacing: { before: 30, after: 30, line: 215, lineRule: 'auto' }, indent: { left: 300, hanging: 240 },
  children: [new TextRun({ text: `${i + 1}. `, size: 16, bold: true, color: NAVY, font: KO }), ...runs(t, { size: 16 })] })));

c.push(new Paragraph({ spacing: { before: 150, after: 50 }, children: [new TextRun({ text: '고착화가 일어나는 구조 — 핵심', bold: true, size: 16, color: NAVY, font: KO })] }));
[
  '자율실험실은 새 장비를 만드는 것이 아니라 기성 장비를 엮는 것이다. {nistp:미국 NIST는 자율실험실 생태계를 “상용 기성품(COTS) 부품으로 조립 가능한 모듈형 구조”로 정의}한다. 따라서 경쟁력은 개별 장비가 아니라 **장비를 잇는 규격과 소프트웨어**에서 결정된다.',
  '그런데 그 규격은 국외에서 정해진다. 장비 통신은 {lads:독일 주도의 OPC UA LADS}와 {sila:SiLA 2}, 데이터는 제약업계 컨소시엄의 Allotrope·AnIML, 소재 온톨로지는 {emmc:EU의 EMMO}가 쥐고 있다. 국내에도 {kmds:K-MDS}·ModuFlow·OCTOPUS OS 같은 시도가 있으나 개별 프로젝트 수준에 머물러 있고, {zdnet:협의체 「기술·플랫폼·표준」 분과는 2026년 8월에야 출범}했다.',
  '소프트웨어는 장비에 묶여 들어온다. 분석 소프트웨어는 장비를 사면 따라오는 구조여서 장비 점유율이 곧 소프트웨어 점유율이 된다. 여기에 {hk:외국계 기업은 장비를 무상 제공하는 대신 5~10년간 소모품 구매를 조건으로 거는 영업 방식}을 쓴다. 한 번 들어온 벤더는 장비·소프트웨어·소모품을 함께 묶어 교체 비용을 키운다.',
].forEach((t, i) => c.push(new Paragraph({ spacing: { before: 30, after: 30, line: 215, lineRule: 'auto' }, indent: { left: 300, hanging: 240 },
  children: [new TextRun({ text: `${i + 4}. `, size: 16, bold: true, color: NAVY, font: KO }), ...runs(t, { size: 16 })] })));

c.push(new Paragraph({ spacing: { before: 150, after: 50 }, children: [new TextRun({ text: '고착화의 결과', bold: true, size: 16, color: NAVY, font: KO })] }));
[
  '{eco:국내 바이오장비 개발률은 36.4%인데 연구현장의 국산 장비 실제 도입률은 1%}에 그친다. 만들 수 있는데도 쓰이지 않는다는 뜻이며, 이는 개별 장비의 성능 문제가 아니라 **기존 시스템에 붙지 못하는 연결의 문제**다.',
  '역설적으로 한국은 부품 역량이 없는 나라가 아니다. {kiria:로봇밀도는 1,012대로 세계 최고 수준}이고 {cobot:협동로봇에서는 국내 기업이 글로벌 4위권}에 있다. 그럼에도 그 로봇을 자국 실험장비에 붙여 쓰는 자율실험실은 만들지 못하고 있다. **부품이 없어서가 아니라, 부품을 엮는 규격과 소프트웨어가 모두 밖에 있기 때문이다.**',
].forEach((t, i) => c.push(new Paragraph({ spacing: { before: 30, after: 30, line: 215, lineRule: 'auto' }, indent: { left: 300, hanging: 240 },
  children: [new TextRun({ text: `${i + 7}. `, size: 16, bold: true, color: NAVY, font: KO }), ...runs(t, { size: 16 })] })));

c.push(new Paragraph({ spacing: { before: 150, after: 50 }, children: [new TextRun({ text: '왜 지금이 분기점인가', bold: true, size: 16, color: NAVY, font: KO })] }));
c.push(new Paragraph({ spacing: { before: 30, after: 30, line: 215, lineRule: 'auto' }, indent: { left: 300, hanging: 240 },
  children: [new TextRun({ text: '9. ', size: 16, bold: true, color: NAVY, font: KO }),
    ...runs('자율실험실 시장은 이제 형성되는 중이다. {labmkt:한국 실험실 자동화 시장은 2024년 약 1억 2,575만 달러} 규모이고, {lhmkt:자동분주기 시장은 2025년 2,350만 달러에서 2036년 5,100만 달러로 성장}할 전망이다. **지금 표준과 시스템 계층을 확보하지 못하면, 지금의 장비 의존이 자율실험실 시대의 시스템 의존으로 그대로 이월된다.**', { size: 16 })] }));

c.push(box('2문장으로 압축한 판', [
  '**①** 국내 연구실 기본 장비의 90% 이상이 수입품이고 {hd90:자율실험의 판독장비인 마이크로플레이트 리더는 외산 비중이 100%}이며, 국산 바이오장비는 {eco:개발률 36.4%에도 현장 도입률이 1%}에 그친다.',
  '**②** 자율실험실의 경쟁력은 개별 장비가 아니라 장비를 잇는 규격과 소프트웨어에서 결정되는데, 국내에도 {kmds:K-MDS}·ModuFlow·OCTOPUS OS 같은 시도가 있으나 개별 프로젝트 수준이고 장비 통신({lads:OPC UA LADS}·{sila:SiLA 2})과 데이터 표준(Allotrope·{emmc:EMMO})은 모두 국외에서 정해져, 지금의 장비 의존이 자율실험실 시대의 시스템 의존으로 고착될 우려가 있다.',
]));

c.push(h('부록 A. 원문 주장별 판정 — 무엇을 쓰고 무엇을 빼야 하는가'));
c.push(dataTable([2900, 900, 6518], ['원문 주장', '판정', '어떻게 쓸 것인가'], [
  ['자동화 장비 90% 이상 외산', '✅', '근거 있음. 다만 통계 기준이 “연구장비 일반”이므로 **“연구실 기본장비 90% 이상 수입, 마이크로플레이트 리더 등은 100%”** 로 구체화해 쓸 것'],
  ['(추가 권장)', '✅', '**“개발률 36.4% vs 도입률 1%”** 를 첫 줄로 올릴 것. 가장 강력한 한 줄'],
  ['소프트웨어로 실험 조건 설정이 어렵다', '🔶', '국내 1차 근거 없음. 국산 업체 인터뷰를 인용하는 형태로 우회'],
  ['플레이트 제조사만 바꿔도 엔지니어 호출', '🔶', '**“웰 바닥 높이조차 최근에야 표준화됐다”** 로 치환하면 근거가 확실해짐'],
  ['현장 엔지니어 부재 · 해외 엔지니어 대기\n· 출장비 수백만 원 · 며칠씩 다운타임', '❌', '**공개 근거 없음.** 출처를 붙이지 말고 현장 사례로 제시하거나 “현장에서는 ~라는 지적이 나온다” 수준으로 수위 조정'],
  ['(추가 검토)', '✅', '외국계의 **장비 무상 제공 + 소모품 5~10년 구매 조건** lock-in 전략. “왜 못 바꾸는가”를 설명'],
], [OKFILL, OKFILL, WARNFILL, WARNFILL, BADFILL, OKFILL]));

// ①
c.push(h('① “자동화 장비 90% 이상 외산”', GREEN));
c.push(para('가장 강한 근거는 정부 통계다. {hd90:헬로디디 보도}에 따르면 **한국표준과학연구원 첨단혁신장비기술정책센터**가 국가연구시설장비 구매현황(**2019~2023년**)을 분석한 결과는 다음과 같다.'));
c.push(dataTable([6500, 3818], ['품목', '외산 비중'], [
  ['오실로스코프 · **마이크로플레이트 리더** · 스펙트럼 분석기', '**100%**'],
  ['시료절편기', '95.8%'],
  ['증류·농축기', '93.6%'],
  ['가스 크로마토그래피', '91.0%'],
]));
c.push(quote('**마이크로플레이트 리더가 100% 외산**이라는 대목은 실험실 자동화 문맥에 그대로 쓸 수 있다. 플레이트를 다루는 장비가 국산이 하나도 없다는 뜻이기 때문이다.'));
c.push(para('뒷받침 자료는 셋이다. {brief:과기정통부는 1억 원 이하 범용장비 국산화를 위해 전담 분과를 신설}하고 2~3년 내 국산화 방침을 밝혔다({dt:관련 보도}). {bio90:한국기계연구원이 연 바이오장비 정책포럼}에서는 **“상용화된 장비의 90%가 모두 수입된 외산”** 이라는 진단이 나왔다. 그리고 {sisa:에이블랩스 신상 대표}는 **“국내에서 유일하게 액체 핸들링 자동화 로봇을 만들었다”** 고 말한다 — 리퀴드핸들러가 사실상 전량 외산이었다는 뜻이다.'));
c.push(box('그대로 쓰면 걸릴 수 있는 지점', [
  '위 통계는 **“연구장비 일반 / 범용 연구장비 / 바이오장비”** 기준이지 **“실험실 자동화 장비”만 따로 집계한 것이 아니다.** ' +
  '자동화 장비 카테고리의 공식 점유율 통계는 존재하지 않는다. “자동화 장비 90%”라고 쓰면 출처와 어긋나므로, ' +
  '**“연구실 기본장비 90% 이상이 수입품이며, 마이크로플레이트 리더처럼 자동화 실험에 직결되는 품목은 외산 비중이 100%”** 로 쓰는 편이 안전하다.',
], WARNFILL, ORANGE));

// 도입률
c.push(h('② 더 센 숫자 — 개발률 36.4% vs 도입률 1%', GREEN));
c.push(para('{eco:헬로디디 자율랩 생태계 기사}에 이런 대목이 있다.'));
c.push(quote('국내 바이오장비 **개발률은 36.4%** 인데, 연구현장의 **국산 장비 실제 도입률은 1%** 다. 연구자들이 계속 외산을 선호하기 때문이다.'));
c.push(para('만들 수는 있는데 안 쓴다는 뜻이다. **“장비가 없다”가 아니라 “생태계가 없다”** 는 원문의 논지에는 90%보다 이 숫자가 훨씬 정확하게 들어맞는다. 같은 기사에서 연구자들이 외산으로 첫 연구를 시작해 성과를 낸 탓에 실험 환경 변화에 거부감이 있다는 구조적 원인도 함께 지적된다.'));

// ③
c.push(h('③ “소프트웨어로 실험 조건 설정이 어렵다”', ORANGE));
c.push(para('이 문장을 직접 뒷받침하는 국내 자료는 찾지 못했다. 대신 **반대 방향의 방증** 세 가지가 있다.'));
c.push(dataTable([2600, 7718], ['자료', '내용'], [
  ['{sisa:에이블랩스 인터뷰}', '국산 제품의 강점으로 **“외산 제품이 갖추지 못했던 유저 소프트웨어의 편의성과 빠른 기술지원”** 을 꼽는다. 국산이 이것을 차별점으로 내세운다는 것은 외산에서 그 부분이 문제였다는 뜻이다'],
  ['{zdnet:자율실험실 협의체}', '「기술·플랫폼·표준」 분과의 우선 과제가 **장비 인터페이스와 데이터 표준 정립**이다'],
  ['{nist:미국 NIST}', '자율실험 표준 4영역 중 하나가 **instrument control and communication**. 같은 문제를 미국도 “표준 부재”로 규정하고 있다'],
]));
c.push(para('논지 자체는 국제적으로 인정된 문제다. 다만 “한국에서 이렇다”는 1차 근거가 없으므로, 국산 업체 인터뷰를 근거로 달거나 현장 사례로 보강하는 편이 안전하다.'));

// ④
c.push(h('④ “플레이트 제조사만 바꿔도 엔지니어를 불러야 한다”', ORANGE));
c.push(para('이 구체적 사례를 보도한 자료는 없다. 그러나 **제조사 간 플레이트 편차가 실재하는 문제**라는 근거는 있다.'));
c.push(para('{slas:기존 ANSI/SLAS 1~4 표준}은 외형 치수와 풋프린트를 규정했지만 **웰 바닥 높이(well-bottom elevation)** 는 표준화하지 않았고, {slasnew:최근에야 별도 표준이 ANSI 인증을 받았다}. 제조사를 바꾸면 로봇 좌표와 광학 초점이 어긋날 수 있다는 뜻으로, 원문 주장의 메커니즘을 그대로 뒷받침한다.'));
c.push(box('권장 표현', ['“웰 바닥 높이조차 최근에야 국제 표준이 만들어졌을 만큼 제조사 간 편차가 컸다” — 이렇게 바꿔 쓰면 검증 가능한 근거가 붙는다.'], WARNFILL, ORANGE));

// ⑤
c.push(h('⑤ “엔지니어 대기 · 출장비 수백만 원 · 며칠씩 다운타임”', RED));
c.push(para('**공개 자료에서 확인하지 못했다.** 국정감사 자료·언론 보도·정책 보고서를 여러 각도로 검색했으나, 출장비 규모나 다운타임 일수 같은 구체 수치를 제시한 자료는 없었다. 간접 방증은 ③과 같다 — 국산 업체가 “빠른 기술지원”을 핵심 차별점으로 내세운다는 점.'));
c.push(box('권장 처리', [
  '이 항목은 **출처를 붙이지 말고 현장 인터뷰·자체 사례로 제시**하거나, “현장에서는 ~라는 지적이 나온다” 정도로 수위를 낮추는 것이 안전하다. ' +
  '수치를 그대로 쓰면 출처를 요구받았을 때 답할 수 없다.',
], BADFILL, RED));

// ⑥
c.push(h('부록 B. 계층별 국내 역량 — 부품은 있는데 시스템이 없다', GREEN));
c.push(para('“자율실험실에 들어가는 자동화 장비”만 따로 집계한 통계는 없다. 이유는 두 가지다. 첫째, {nistp:NIST가 자율실험실을 “상용 기성품(COTS) 부품으로 조립 가능한 모듈형 생태계”로 정의}하듯 자율실험실은 새 장비를 만드는 것이 아니라 **기성 장비를 엮는 것**이라 별도 품목이 성립하지 않는다. 둘째, KSIC·HS코드와 국가연구시설장비 분류체계에 해당 품목 자체가 없다.'));
c.push(para('그래서 계층을 나눠 보면 국내 역량의 분포가 드러난다.'));
c.push(dataTable([2300, 1500, 6518], ['계층', '국내 역량', '근거'], [
  ['로봇팔(협동로봇)', '**있음**', '{kiria:로봇산업 실태조사}는 2006년부터 매년 나오는 정부승인통계다. 한국의 로봇밀도는 **1,012대로 세계 최고 수준**이고, {cobot:두산로보틱스는 중국을 제외한 글로벌 시장 4위권이면서 국내 1위}다({cobot2:글로벌 1위는 덴마크 유니버설로봇})'],
  ['리퀴드핸들러', '**1개사**', '{top5:글로벌 상위 5개사(Agilent·Beckman Coulter·Eppendorf·Hamilton·Thermo Fisher)가 모두 외산}. {sisa:국내에서는 2021년 설립된 에이블랩스가 “유일하게 액체 핸들링 자동화 로봇을 만들었다”}고 밝힌다. 국산·외산 분리 점유율 통계는 없다'],
  ['분석장비', '거의 없음', '{hd90:마이크로플레이트 리더·오실로스코프·스펙트럼 분석기 외산 비중 100%}. {bioin:바이오 분야는 수요기업의 90% 이상이 해외 소부장에 의존}'],
  ['연결 규격·소프트웨어', '**지분 없음**', '장비 통신은 {lads:OPC UA LADS}(독일 주도)와 {sila:SiLA 2}, 데이터는 Allotrope·AnIML, 소재 온톨로지는 {emmc:EU의 EMMO}가 쥐고 있다. 한국은 {zdnet:2026년 8월에야 협의체 「기술·플랫폼·표준」 분과를 출범}시켰다'],
]));
c.push(quote('세계 최고 로봇밀도에 협동로봇 국내 1위 기업까지 있는 나라가, 그 로봇을 자국 실험장비에 붙여 쓰는 자율실험실은 만들지 못하고 있다. **부품이 없어서가 아니라, 부품을 엮는 규격과 소프트웨어가 전부 밖에 있기 때문이다.**'));
c.push(box('주의 — “로봇팔도 외산”이라고 쓰면 안 된다', [
  '두산로보틱스·레인보우로보틱스·뉴로메카가 있고 **국내 협동로봇 시장 1위는 두산로보틱스**다. ' +
  '“장비가 다 외산”이라고 뭉뚱그리면 이 대목에서 반박당하고, 그 순간 문단 전체의 신뢰가 흔들린다. ' +
  '오히려 **“부품은 있는데 시스템이 없다”** 로 써야 논지가 강해진다.',
], WARNFILL, ORANGE));

c.push(h('부록 C. 보조 근거', GREEN));
c.push(dataTable([2600, 7718], ['자료', '내용'], [
  ['{hk:한국경제(2023.8)}', '외국계 기업이 **장비를 무상 제공하는 대신 5~10년간 소모품 구매를 조건**으로 거는 영업 전략. “왜 국산으로 바꾸지 못하는가”를 설명하는 구조적 요인(lock-in)으로, 종속 논지에 직결된다'],
  ['{eco:헬로디디}', '한국기계연구원 임현의 연구단장이 **AI와 국산 실험장비를 연결하는 플랫폼**을 구축 중이라고 밝혔다. 문제 제기에 그치지 않고 대응이 진행 중이라는 점을 함께 쓸 수 있다'],
]));

// 출처
c.push(h('출처'));
const srcList = Object.values(S).sort((a, b) => a.n - b.n);
const half = Math.ceil(srcList.length / 2);
const rows = [];
for (let i = 0; i < half; i++) {
  const mk = (s) => s ? new TableCell({ width: { size: Math.floor(CONTENT_W / 2), type: WidthType.DXA }, margins: { top: 10, bottom: 10, left: 0, right: 120 },
      children: [new Paragraph({ spacing: { before: 8, after: 8, line: 200, lineRule: 'auto' },
        children: [new TextRun({ text: `[${s.n}] `, size: 13, color: '444444', font: KO }),
          new ExternalHyperlink({ link: s.url, children: [new TextRun({ text: s.t, size: 13, color: '1B54B5', underline: {}, font: KO })] })] })] })
    : new TableCell({ width: { size: Math.floor(CONTENT_W / 2), type: WidthType.DXA }, children: [new Paragraph('')] });
  rows.push(new TableRow({ children: [mk(srcList[i]), mk(srcList[i + half])] }));
}
c.push(new Table({ columnWidths: [Math.floor(CONTENT_W / 2), Math.floor(CONTENT_W / 2)], width: { size: Math.floor(CONTENT_W / 2) * 2, type: WidthType.DXA },
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
             insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } }, rows }));
c.push(new Paragraph({ spacing: { before: 80 },
  children: runs('조사 환경에서 대부분의 언론사 도메인이 차단돼 원문을 직접 열람하지 못했다. 검색 결과에 인용된 문구와 복수 매체 교차 확인으로 정리했으므로, 인용 전 링크를 직접 확인하시기를 권한다.', { size: 12, color: '666666' }) }));

const doc = new Document({
  styles: { default: { document: { run: { font: KO, size: 16, color: '222222' }, paragraph: { spacing: { line: 220, lineRule: 'auto' } } } },
    characterStyles: [{ id: 'Hyperlink', name: 'Hyperlink', basedOn: 'DefaultParagraphFont', run: { color: '1B54B5', underline: {} } }] },
  sections: [{ properties: { page: { size: { width: PAGE_W, height: convertMillimetersToTwip(297) },
    margin: { top: MARGIN, bottom: convertMillimetersToTwip(12), left: MARGIN, right: MARGIN } } }, children: c }],
});
const out = process.argv[2] || '/home/user/kist/report/외산장비_의존_근거검증.docx';
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(out, b); console.log('written:', out, b.length); });
