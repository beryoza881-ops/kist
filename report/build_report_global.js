// 해외 자율실험실(SDL) 대표사례 보고서 생성 스크립트
const path = require('path');
const SCRATCH = '/tmp/claude-0/-home-user-kist/f7349fbc-9d61-5776-85de-581602fe2d75/scratchpad';
const docx = require(path.join(SCRATCH, 'node_modules', 'docx'));
const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, convertMillimetersToTwip,
} = docx;
const fs = require('fs');

const KO = { ascii: '맑은 고딕', eastAsia: '맑은 고딕', hAnsi: '맑은 고딕', cs: '맑은 고딕' };
const NAVY = '1D3567', LINE = 'B9C2D4', HEADFILL = '1D3567', BOXFILL = 'EEF2F9', ALTFILL = 'F7F9FC';
const PAGE_W = convertMillimetersToTwip(210);
const MARGIN = convertMillimetersToTwip(14);
const CONTENT_W = PAGE_W - 2 * MARGIN;

const S = {
  alab:      { url: 'https://www.nature.com/articles/s41586-023-06734-w',
               t: 'N. J. Szymanski et al., "An autonomous laboratory for the accelerated synthesis of novel materials", Nature 624, 86 (2023)' },
  alabdoubt: { url: 'https://www.chemistryworld.com/news/new-analysis-raises-doubts-over-autonomous-labs-materials-discoveries/4018791.article',
               t: 'Chemistry World, "New analysis raises doubts over autonomous lab\'s materials discoveries"' },
  alabfix:   { url: 'https://cen.acs.org/research-integrity/Nature-robot-chemist-paper-corrected/104/web/2026/01',
               t: 'C&EN, "Nature robot chemist paper corrected"(2026.1)' },
  liv:       { url: 'https://www.nature.com/articles/s41586-020-2442-2',
               t: 'B. Burger et al., "A mobile robotic chemist", Nature 583, 237 (2020)' },
  livcw:     { url: 'https://www.chemistryworld.com/news/your-new-labmate-does-700-reactions-in-eight-days-and-its-a-robot/4012125.article',
               t: 'Chemistry World, "Your new labmate does 700 reactions in eight days – and it\'s a robot"' },
  cosci:     { url: 'https://www.nature.com/articles/s41586-023-06792-0',
               t: 'D. A. Boiko et al., "Autonomous chemical research with large language models", Nature 624, 570 (2023)' },
  cmu:       { url: 'https://www.cmu.edu/chemistry/news/2023/1220_ai-coscientist-automates-discovery.html',
               t: 'Carnegie Mellon University, "CMU-Designed Artificially Intelligent Coscientist Automates Scientific Discovery"' },
  uoft:      { url: 'https://www.utoronto.ca/news/u-t-receives-200-million-grant-support-acceleration-consortium-s-self-driving-labs-research',
               t: 'University of Toronto, "U of T receives $200-million grant to support Acceleration Consortium\'s \'self-driving labs\' research"' },
  uoftlaser: { url: 'https://www.artsci.utoronto.ca/news/acceleration-consortium-and-global-collaboration-self-driving-labs-discovers-new-molecules',
               t: 'U of T Arts & Science, "Acceleration Consortium and global collaboration of self-driving labs discovers new molecules for organic solid-state lasers"' },
  argonne:   { url: 'https://www.anl.gov/article/selfdriving-lab-transforms-materials-discovery',
               t: 'Argonne National Laboratory, "Self-driving lab transforms materials discovery"' },
  rainbow:   { url: 'https://news.ncsu.edu/2025/08/rainbow-multi-robot-lab/',
               t: 'NC State News, "Meet Rainbow: The Multi-Robot Lab Racing to Discover the Next Quantum Dots"(2025.8)' },
  speed:     { url: 'https://news.ncsu.edu/2026/06/speeding-up-scientific-discovery/',
               t: 'NC State News, "Speeding Up Scientific Discovery"(NSF SPEED 사업, 2026.6)' },
  ustc:      { url: 'https://www.nature.com/articles/s44160-023-00424-1',
               t: 'Q. Zhu et al., "Automated synthesis of oxygen-producing catalysts from Martian meteorites by a robotic AI chemist", Nature Synthesis (2023)' },
  cas:       { url: 'https://english.cas.cn/newsroom/cas_media/202311/t20231115_643207.shtml',
               t: 'Chinese Academy of Sciences, "China\'s AI Robotic Chemist Synthesizes Catalysts for Oxygen Production on Mars"' },
  chemify:   { url: 'https://www.chem.gla.ac.uk/cronin/news/chemify-opens-world-first-chemputation-facility',
               t: 'The Cronin Group, "Chemify opens world-first chemputation facility"(Chemifarm, Glasgow)' },
  chemifycen:{ url: 'https://cen.acs.org/business/informatics/AI-drug-synthesizer-Chemify-raises/103/web/2025/10',
               t: 'C&EN, "AI drug synthesizer Chemify raises $50 million"(2025.10)' },
  nimsos:    { url: 'https://www.nims.go.jp/eng/press/2023/07/202307200.html',
               t: 'NIMS, "Development of NIMS-OS: General-Purpose Software Enabling Autonomous, Automated Experiments"(2023.7)' },
  rscjp:     { url: 'https://pubs.rsc.org/dd/article/4/6/1384/846229/Self-driving-laboratories-in-Japan',
               t: '"Self-driving laboratories in Japan", Digital Discovery 4, 1384 (2025) — RIKEN Mahoro 등 일본 사례 개관' },
  gnome:     { url: 'https://www.nature.com/articles/s41586-023-06735-9',
               t: 'A. Merchant et al., "Scaling deep learning for materials discovery", Nature 624, 80 (2023) — GNoME' },
  dmblog:    { url: 'https://deepmind.google/blog/millions-of-new-materials-discovered-with-deep-learning/',
               t: 'Google DeepMind, "Millions of new materials discovered with deep learning"' },
  lbl:       { url: 'https://newscenter.lbl.gov/2023/11/29/google-deepmind-new-compounds-materials-project/',
               t: 'Berkeley Lab, "Google DeepMind Adds Nearly 400,000 New Compounds to Berkeley Lab\'s Materials Project"' },
  dupes:     { url: 'https://cen.acs.org/research-integrity/Duplicate-structures-haunt-crystallography-databases/103/web/2025/12',
               t: 'C&EN, "Duplicate structures haunt crystallography databases"(2025.12)' },
  ecl:       { url: 'https://www.emeraldcloudlab.com/',
               t: 'Emerald Cloud Lab 공식 홈페이지 — 원격 접속 실험실' },
  eclwiki:   { url: 'https://en.wikipedia.org/wiki/Emerald_Cloud_Lab',
               t: 'Emerald Cloud Lab 개요(장비 200종 이상, 24/365 운영)' },
  lila:      { url: 'https://www.flagshippioneering.com/news/press-release/flagship-pioneering-unveils-lila-sciences-to-build-superintelligence-in-science',
               t: 'Flagship Pioneering, "Flagship Pioneering Unveils Lila Sciences to Build Superintelligence in Science"' },
  lilaa:     { url: 'https://www.fiercebiotech.com/biotech/flagships-lila-sciences-lands-235m-expand-ai-powered-autonomous-research-labs',
               t: 'Fierce Biotech, "Flagship\'s Lila Sciences lands $235M to expand AI-powered autonomous research labs"' },
  periodic:  { url: 'https://techcrunch.com/2025/09/30/former-openai-and-deepmind-researchers-raise-whopping-300m-seed-to-automate-science',
               t: 'TechCrunch, "Former OpenAI and DeepMind researchers raise whopping $300M seed to automate science"(2025.9.30)' },
  insilico:  { url: 'https://www.eurekalert.org/news-releases/975722',
               t: 'Insilico Medicine, "Insilico Medicine launches 6th generation Intelligent Robotics Lab"(Life Star, Suzhou)' },
  mtr:       { url: 'https://www.technologyreview.com/2025/12/15/1129210/ai-materials-science-discovery-startups-investment/',
               t: 'MIT Technology Review, "AI materials discovery now needs to move into the real world"(2025.12)' },
};
Object.values(S).forEach((x, i) => { x.n = i + 1; });

function runs(str, { size = 15, bold = false, color = '222222' } = {}) {
  const out = []; const re = /\{([a-zA-Z0-9]+):([^}]*)\}/g; let last = 0, m;
  while ((m = re.exec(str)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: str.slice(last, m.index), size, bold, color, font: KO }));
    const src = S[m[1]]; if (!src) throw new Error('unknown source key: ' + m[1]);
    out.push(new ExternalHyperlink({ link: src.url, children: [new TextRun({ text: m[2], size, bold, font: KO, color: '1B54B5', underline: {} })] }));
    last = re.lastIndex;
  }
  if (last < str.length) out.push(new TextRun({ text: str.slice(last), size, bold, color, font: KO }));
  return out;
}

function cell(text, { w, size = 14, bold = false, color = '222222', fill, align = AlignmentType.LEFT, pad = 60 } = {}) {
  const paras = (Array.isArray(text) ? text : [text]).map((t) =>
    new Paragraph({ alignment: align, spacing: { before: 12, after: 12, line: 200, lineRule: 'auto' }, children: runs(t, { size, bold, color }) }));
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    shading: fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined,
    margins: { top: 78, bottom: 78, left: pad, right: pad }, children: paras,
  });
}

const thinBorders = {
  top: { style: BorderStyle.SINGLE, size: 3, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 3, color: LINE },
  left: { style: BorderStyle.SINGLE, size: 3, color: LINE }, right: { style: BorderStyle.SINGLE, size: 3, color: LINE },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 3, color: LINE }, insideVertical: { style: BorderStyle.SINGLE, size: 3, color: LINE },
};

function dataTable(widths, header, rows) {
  const headRow = new TableRow({ tableHeader: true,
    children: header.map((h, i) => cell(h, { w: widths[i], size: 15, bold: true, color: 'FFFFFF', fill: HEADFILL, align: AlignmentType.CENTER })) });
  const bodyRows = rows.map((r, ri) => new TableRow({ cantSplit: true,
    children: r.map((c, i) => cell(c, { w: widths[i], size: 16, bold: i === 0,
      fill: ri % 2 === 1 ? ALTFILL : undefined,
      align: (i === 0 || i === r.length - 1) ? AlignmentType.CENTER : AlignmentType.LEFT })) }));
  return new Table({ columnWidths: widths, width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: thinBorders, rows: [headRow, ...bodyRows] });
}

const title = new Paragraph({ spacing: { after: 40, line: 240, lineRule: 'auto' },
  children: [new TextRun({ text: '해외 자율실험실(Self-Driving Lab) 대표사례', bold: true, size: 26, color: NAVY, font: KO })] });
const subtitle = new Paragraph({ spacing: { after: 140 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } },
  children: [new TextRun({ text: '기준일: 2026년 8월  |  학술지·기관 발표 기준 정리  |  국내 사례는 별도 보고서 참조', size: 14, color: '666666', font: KO })] });

const boxText =
  '해외 자율실험실은 2020년 영국 리버풀대의 이동형 로봇 화학자를 기점으로 “AI가 다음 실험을 결정하는” 폐루프가 학술적으로 입증된 뒤, ' +
  '{alab:버클리연구소 A-Lab}·{argonne:아르곤 Polybot} 같은 국립연구소 플랫폼과 {ustc:중국과기대 로봇 AI 화학자}로 확산됐다. ' +
  '국가 차원 투자({uoft:캐나다 CFREF 2억 달러}, {speed:미국 NSF SPEED 2,000만 달러})와 민간 자본({periodic:Periodic Labs 3억 달러}, {lilaa:Lila Sciences 총 5.5억 달러})이 ' +
  '2025년을 기점으로 동시에 유입되면서, SDL은 실험실 규모의 연구 주제에서 산업 카테고리로 넘어가는 국면에 있다. ' +
  '다만 {alabdoubt:A-Lab 결과에 대한 검증 논란과 논문 정정}이 보여주듯, 속도 지표와 별개로 결과의 신뢰성 검증은 아직 해결되지 않은 과제다.';

const boxTable = new Table({ columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
  borders: { top: { style: BorderStyle.SINGLE, size: 6, color: NAVY }, bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY },
             left: { style: BorderStyle.SINGLE, size: 18, color: NAVY }, right: { style: BorderStyle.SINGLE, size: 6, color: NAVY } },
  rows: [new TableRow({ children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill: BOXFILL, color: 'auto' }, margins: { top: 120, bottom: 120, left: 180, right: 160 },
    children: [
      new Paragraph({ spacing: { after: 50, line: 220, lineRule: 'auto' }, children: [new TextRun({ text: '한눈에 보기', bold: true, size: 16, color: NAVY, font: KO })] }),
      new Paragraph({ spacing: { line: 230, lineRule: 'auto' }, children: runs(boxText, { size: 16 }) }),
    ] })] })] });

const h1 = new Paragraph({ spacing: { before: 220, after: 70 },
  children: [new TextRun({ text: '표 1. 해외 자율실험실(폐루프 SDL) 대표사례', bold: true, size: 17, color: NAVY, font: KO })] });

const W1 = [900, 2400, 1200, 4518, 1300];
const T1_HEAD = ['국가', '기관·실험실 (담당자)', '주요 분야', '주요 성과', '자율실험 정도'];
const T1 = [
  ['영국', 'University of Liverpool\n{liv:Mobile Robotic Chemist}\nAndrew Cooper 교수', '광촉매\n(수소 생산)',
   '실험실을 자유롭게 돌아다니는 이동형 로봇이 고체 촉매 계량·액체 분주·광조사·수소 측정을 수행하고 다음 실험을 스스로 결정. {livcw:10개 변수 공간에서 8일간 688회 실험}(배치 베이지안 탐색)으로 초기 조성 대비 활성 6배인 광촉매 조합 발견 — 사람이면 수개월. {liv:Nature(2020)} 게재, SDL 폐루프의 최초 실증',
   '★★★\n폐루프 완전자율'],
  ['미국', 'Lawrence Berkeley National Lab\n{alab:A-Lab}\nGerbrand Ceder·Yan Zeng\n(Google DeepMind 협력)', '무기소재 합성',
   '{alab:17일간 58개 목표 중 41종의 신규 무기화합물을 합성}(성공률 71%, 하루 21건). 계산·문헌 데이터·능동학습으로 합성 경로를 설계하고 실패를 학습해 조건을 수정. 다만 {alabdoubt:발표 직후 UCL 연구자가 실험 분석 품질에 “매우 심각한 문제”를 제기}했고 {alabfix:2026년 1월 논문이 정정}됐다',
   '★★★\n폐루프 완전자율\n(결과 검증 논란)'],
  ['미국', 'Carnegie Mellon University\n{cosci:Coscientist}\nGabe Gomes 교수팀', '유기합성',
   '{cmu:LLM(GPT-4·Claude)이 문헌 검색·코드 실행·장비 제어를 도구로 삼아 실험을 스스로 설계·계획·수행}한 최초 사례. {cosci:팔라듐 촉매 교차결합 반응의 최적화를 자율 수행}하는 등 6가지 과제를 시연, Nature(2023) 게재. “로봇이 조건을 바꾸는” 단계에서 “언어모델이 연구를 지휘하는” 단계로 넘어간 전환점',
   '★★★\nLLM 주도 폐루프'],
  ['캐나다', 'University of Toronto\n{uoft:Acceleration Consortium}\nAlán Aspuru-Guzik 소장', '소재·신약 전반',
   '{uoft:CFREF에서 2억 달러를 지원받은 캐나다 대학 사상 최대 규모의 연방 연구비}로 SDL 거점을 구축. 자율실험실로 {uoft:항암 후보물질을 30일 만에 도출}(통상 수년~수십 년), {uoftlaser:전 세계 SDL을 연결한 국제 협업으로 유기 고체 레이저용 신규 분자 발견}',
   '★★★\n다기관 SDL 네트워크'],
  ['미국', 'Argonne National Laboratory\n{argonne:Polybot}\n(Center for Nanoscale Materials)', '전자 고분자 박막',
   '{argonne:공정 조합이 약 100만 가지에 이르는 전자 고분자 박막 공정을 자율 탐색해, 현재 최고 수준에 필적하는 평균 전도도의 박막과 대량생산용 레시피를 도출}. 모든 데이터가 자동 기록·기계학습 분석을 거쳐 AI에 전달되고, AI가 다음 실험을 지시하는 구조',
   '★★★\n폐루프 완전자율'],
  ['미국', 'NC State University\n{rainbow:Rainbow}(다중로봇 SDL)\nMilad Abolhasani 교수', '퀀텀닷·용액 소재',
   '{rainbow:여러 로봇이 협업해 하루 최대 1,000회의 실험을 사람 개입 없이 수행·분석}. 가동 하루 만에 최고 수준(best-in-class) 퀀텀닷을 찾아냈는데, 이는 사람 중심 실험으로 축적된 약 7년치 문헌에 해당. {speed:NSF가 4년간 2,000만 달러 규모의 SPEED 사업으로 지원}',
   '★★★\n폐루프 완전자율'],
  ['중국', '중국과학기술대(USTC)\n·심우주탐사실험실\n{ustc:로봇 AI 화학자}', '화성 산소 발생 촉매',
   '{cas:화성 운석을 원료로 산소발생반응(OER) 촉매를 6주 만에 자동 합성·최적화}. 기계학습으로 376만 가지 가능한 조성 중 최적 조성을 탐색했는데, 사람이 하면 약 2,000년이 걸릴 규모. {ustc:Nature Synthesis(2023)} 게재',
   '★★★\n폐루프 완전자율'],
  ['일본', 'NIMS(물질·재료연구기구)\n{nimsos:NIMS-OS}', '소재 탐색\n(범용 미들웨어)',
   '개별 실험실이 아니라 {nimsos:어떤 소재탐색 AI와 어떤 로봇 실험 시스템이든 짝지어 사람 개입 없는 폐루프를 만들어 주는 범용 미들웨어(Python 라이브러리)}를 공개. 자율실험을 “한 연구실의 장치”가 아니라 “표준 소프트웨어 계층”으로 다룬 접근',
   '★★★\n폐루프 표준화 도구'],
  ['일본', 'RIKEN(이화학연구소)\n인간형 실험로봇 {rscjp:Mahoro}', '세포배양·생명과학',
   '{rscjp:최적화 알고리즘과 결합해 세포배양 조건을 자동으로 반복 조정}. 별도 연구에서는 자연어로 쓰인 실험 절차를 LLM이 로봇 동작 코드로 자동 변환하는 것을 시연 — 사람이 쓰는 프로토콜과 로봇 사이의 번역 문제를 겨냥',
   '★★☆\n생명과학 폐루프'],
  ['영국\n(기업)', '{chemify:Chemify}\n(Glasgow대 스핀아웃)\nLee Cronin 창업', '유기합성·의약품',
   '화학 프로그래밍 언어 χDL로 분자 설계를 실행 코드로 바꾸고 모듈형 로봇 ‘Chemputer’가 수행. {chemify:2025년 세계 최초의 대규모 자동합성 시설 Chemifarm을 글래스고에 개소}해 {chemify:나이톨·루피나미드·실데나필 3종을 사람 개입 없이 합성}(수율은 수작업과 같거나 그 이상). {chemifycen:2025년 시리즈B 5,000만 달러}',
   '★★★\n무인 합성 시설 가동'],
  ['중국\n(기업)', '{insilico:Insilico Medicine}\nLife Star(쑤저우 BioBAY)', 'AI 신약개발',
   '{insilico:자율주행 운반로봇(AGV)과 영상장비를 갖춘 6세대 지능형 로봇 실험실로, 사람 개입 없이 표적 발굴·화합물 스크리닝·정밀의료·중개연구를 수행}. 자사 AI 플랫폼 Pharma.AI와 완전 자동화 생물실험 모듈을 결합해 폐루프를 구성하고, 자체 실험실이 없는 기업에도 개방',
   '★★★\n신약 폐루프 가동'],
  ['미국\n(기업)', '{lila:Lila Sciences}\n(Flagship Pioneering)', '생명·화학·소재',
   '가설 수립 → 실험 수행 → 측정 → 다음 실험 결정을 무인으로 반복하는 ‘AI Science Factories’를 표방. {lilaa:2025년 한 해에 시드 2억 달러·시리즈A 2.35억 달러·추가 1.15억 달러(엔비디아 벤처 참여) 등 총 5.5억 달러를 조달}, 기업가치 약 12억 달러',
   '★★☆\n기업 자체 운영\n(외부 검증 제한)'],
  ['미국\n(기업)', '{periodic:Periodic Labs}', '고온 초전도체\n·기능성 소재',
   '{periodic:구글 딥마인드 소재·화학팀 리드와 전 OpenAI 리서치 VP가 2025년 9월 창업, a16z 주도로 시드 3억 달러 조달}(엔비디아·베이조스·슈밋 등 참여). AI가 실험을 제안하고 로봇이 합성·소성, 장비가 물성을 측정해 다시 AI가 설계하는 폐루프로 월 수천 건 무인 실험을 목표',
   '★☆☆\n설립 초기(2025~)'],
];
const t1 = dataTable(W1, T1_HEAD, T1.map((r) => r.map((c) => c.split('\n'))));

const legend = new Paragraph({ spacing: { before: 60, after: 0, line: 200, lineRule: 'auto' },
  children: runs('※ 자율실험 정도  ★★★ 설계–실험–분석–재설계가 폐루프로 무인 순환  |  ★★☆ 폐루프를 갖췄으나 적용 범위가 제한적이거나 외부 검증이 어려움  |  ★☆☆ 구축 초기 단계', { size: 13, color: '666666' }) });

const h2 = new Paragraph({ spacing: { before: 260, after: 40 },
  children: [new TextRun({ text: '표 2. (별도 구분) 자율실험실은 아니나 AI 실험설계·드라이랩·원격 자동화에 해당하는 사례', bold: true, size: 17, color: NAVY, font: KO })] });
const note2 = new Paragraph({ spacing: { after: 70, line: 200, lineRule: 'auto' },
  children: runs('해외 사례는 처음부터 폐루프를 목표로 설계·보도되는 경우가 많아 이 표가 국내 보고서보다 짧다. 여기 실린 것은 실험을 하지 않는 예측 모델, 또는 사람이 실험을 지시하되 실행만 자동화한 인프라다.', { size: 13, color: '666666' }) });

const W2 = [1450, 2600, 1300, 4968];
const T2_HEAD = ['구분', '기관 (국가)', '주요 분야', '내용 및 성과'];
const T2 = [
  ['드라이랩\n(AI 예측)', 'Google DeepMind — {gnome:GNoME}\n(영국·미국)', '신소재 탐색',
   '{gnome:그래프 신경망을 대규모로 학습시켜 220만 개 결정 구조를 예측하고 그중 38.1만 개를 새로운 안정 물질로 제시}, Nature(2023). {dmblog:이 가운데 736개는 전 세계 연구자들이 독립적으로 실제 합성}해 예측의 실현 가능성을 보였다. 실험은 하지 않는 순수 예측 모델로, A-Lab 등 SDL에 탐색 후보를 공급하는 위치'],
  ['데이터 인프라', 'Berkeley Lab — {lbl:Materials Project}\n(미국)', '소재 데이터베이스',
   '{lbl:GNoME이 예측한 화합물 약 40만 개가 편입}되며 공개 소재 데이터베이스가 크게 확장. 자율실험실이 “무엇을 만들어 볼지” 후보를 얻는 원천으로, SDL과 짝을 이루는 계산·데이터 계층'],
  ['원격 자동화\n(클라우드 랩)', '{ecl:Emerald Cloud Lab}\n(미국)', '생명과학·화학 실험 대행',
   '{eclwiki:200종이 넘는 장비를 단일 소프트웨어(ECL Command Center)로 원격 제어하는 실험실을 24시간 365일 운영}. 연구자가 인터넷으로 실험을 의뢰하면 실행·측정·데이터 반환이 자동으로 이뤄진다. 다만 실험 설계와 다음 실험 결정은 사람 몫이라 폐루프가 아니다'],
];
const t2 = dataTable(W2, T2_HEAD, T2.map((r) => r.map((c) => c.split('\n'))));

const h25 = new Paragraph({ spacing: { before: 240, after: 60 },
  children: [new TextRun({ text: '참고. 흐름과 쟁점', bold: true, size: 17, color: NAVY, font: KO })] });
const bulletLines = [
  '국가가 인프라로 투자한다 — {uoft:캐나다는 CFREF 2억 달러를 토론토대 Acceleration Consortium에 몰아주며 자국 대학 사상 최대 연방 연구비를 SDL에 배정}했고, {speed:미국 NSF는 4년 2,000만 달러 규모의 SPEED 사업}으로 용액 화학·소재 SDL을 지원한다. 일본은 개별 장치가 아니라 {nimsos:AI와 로봇을 잇는 범용 미들웨어(NIMS-OS)를 공개}해 표준 계층을 선점하는 방식을 택했다.',
  '2025년, 민간 자본이 한꺼번에 들어왔다 — {lilaa:Lila Sciences 총 5.5억 달러}, {periodic:Periodic Labs 시드 3억 달러}, {chemifycen:Chemify 시리즈B 5,000만 달러}가 같은 해에 몰렸다. 딥마인드·OpenAI 출신이 소재·물리로 옮겨 온 것이 특징이며, {mtr:AI 소재 발굴이 예측 단계를 넘어 실물 실험으로 이동해야 한다는 문제의식}이 이 흐름의 배경이다.',
  '속도보다 검증이 쟁점이 됐다 — {alabdoubt:A-Lab의 “41종 신규 화합물”은 발표 직후 실험 분석의 품질 문제가 제기}됐고 {alabfix:2026년 1월 논문이 정정}됐다. 예측 쪽에서도 {dupes:결정학 데이터베이스의 중복 구조 문제}가 지적됐다. 실험 횟수·기간 단축 같은 속도 지표는 쉽게 커지지만, 나온 결과가 정말 새로운 물질인지 확인하는 체계는 아직 따라가지 못하고 있다.',
];
const bullets = bulletLines.map((t) => new Paragraph({
  spacing: { before: 30, after: 30, line: 210, lineRule: 'auto' }, indent: { left: 220, hanging: 160 },
  children: [new TextRun({ text: '▪  ', size: 14, color: NAVY, font: KO }), ...runs(t, { size: 15 })] }));

const h3 = new Paragraph({ spacing: { before: 240, after: 60 },
  children: [new TextRun({ text: '출처', bold: true, size: 17, color: NAVY, font: KO })] });
const srcList = Object.values(S).sort((a, b) => a.n - b.n);
const half = Math.ceil(srcList.length / 2);
const colA = srcList.slice(0, half), colB = srcList.slice(half);
const srcRows = [];
for (let i = 0; i < half; i++) {
  const mk = (s) => s ? new TableCell({ width: { size: Math.floor(CONTENT_W / 2), type: WidthType.DXA },
      margins: { top: 10, bottom: 10, left: 0, right: 120 },
      children: [new Paragraph({ spacing: { before: 6, after: 6, line: 200, lineRule: 'auto' },
        children: [new TextRun({ text: `[${s.n}] `, size: 13, color: '444444', font: KO }),
          new ExternalHyperlink({ link: s.url, children: [new TextRun({ text: s.t, size: 13, color: '1B54B5', underline: {}, font: KO })] })] })] })
    : new TableCell({ width: { size: Math.floor(CONTENT_W / 2), type: WidthType.DXA }, children: [new Paragraph('')] });
  srcRows.push(new TableRow({ children: [mk(colA[i]), mk(colB[i])] }));
}
const srcTable = new Table({ columnWidths: [Math.floor(CONTENT_W / 2), Math.floor(CONTENT_W / 2)],
  width: { size: Math.floor(CONTENT_W / 2) * 2, type: WidthType.DXA },
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE },
             right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } },
  rows: srcRows });

const footer = new Paragraph({ spacing: { before: 100 },
  children: runs('본 보고서의 사례·수치는 위 학술논문 및 기관 공식 발표에 기재된 내용을 인용한 것이다. 기업 사례(Lila·Periodic·Insilico)는 자사 발표에 근거하며 동료평가를 거친 검증 자료가 아직 제한적이다.', { size: 12, color: '666666' }) });

const doc = new Document({
  styles: { default: { document: { run: { font: KO, size: 15, color: '222222' }, paragraph: { spacing: { line: 220, lineRule: 'auto' } } } },
    characterStyles: [{ id: 'Hyperlink', name: 'Hyperlink', basedOn: 'DefaultParagraphFont', run: { color: '1B54B5', underline: {} } }] },
  sections: [{ properties: { page: { size: { width: PAGE_W, height: convertMillimetersToTwip(297) },
        margin: { top: MARGIN, bottom: convertMillimetersToTwip(12), left: MARGIN, right: MARGIN } } },
    children: [title, subtitle, boxTable, h1, t1, legend, h2, note2, t2, h25, ...bullets, h3, srcTable, footer] }],
});

const out = process.argv[2] || '/home/user/kist/report/해외_자율실험실_대표사례.docx';
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log('written:', out, buf.length, 'bytes'); });
