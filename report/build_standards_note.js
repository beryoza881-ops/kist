// 자율실험실 표준화 동향 노트 — docx 생성
const path = require('path');
const SCRATCH = '/tmp/claude-0/-home-user-kist/f7349fbc-9d61-5776-85de-581602fe2d75/scratchpad';
const docx = require(path.join(SCRATCH, 'node_modules', 'docx'));
const { Document, Packer, Paragraph, TextRun, ExternalHyperlink, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, convertMillimetersToTwip } = docx;
const fs = require('fs');

const KO = { ascii: '맑은 고딕', eastAsia: '맑은 고딕', hAnsi: '맑은 고딕', cs: '맑은 고딕' };
const NAVY = '1D3567', LINE = 'B9C2D4', HEADFILL = '1D3567', BOXFILL = 'EEF2F9', ALTFILL = 'F7F9FC', WARNFILL = 'FDF3E7';
const PAGE_W = convertMillimetersToTwip(210), MARGIN = convertMillimetersToTwip(14);
const CONTENT_W = PAGE_W - 2 * MARGIN;

const S = {
  prog:    { url: 'https://www.nist.gov/programs-projects/development-standards-support-modular-and-autonomous-laboratory-ecosystem',
             t: 'NIST, "Development of Standards to Support a Modular and Autonomous Laboratory Ecosystem"' },
  paper:   { url: 'https://www.sciencedirect.com/science/article/abs/pii/S2590238526001190',
             t: 'H. Joress et al., "Toward a composable, modular laboratory ecosystem for autonomous materials R&D", Matter (2026.5.6)' },
  rg:      { url: 'https://www.researchgate.net/publication/394529622',
             t: '위 논문 프리프린트(ResearchGate)' },
  sp1320:  { url: 'https://tsapps.nist.gov/publication/get_pdf.cfm?pub_id=958246',
             t: 'NIST SP 1320, "Driving U.S. Innovation in Materials and Manufacturing using AI and Autonomous Labs"' },
  ws25:    { url: 'https://www.nist.gov/news-events/events/2025/09/workshop-towards-autonomous-materials-research-ecosystem',
             t: 'NIST Workshop, "Towards an Autonomous Materials Research Ecosystem"(2025.9)' },
  ws26:    { url: 'https://www.nist.gov/news-events/events/2026/08/towards-best-practice-guide-designing-equipment-autonomous-materials-rd',
             t: 'NIST Workshop, "Best Practice Guide for Designing Equipment for Autonomous Materials R&D"(2026.8)' },
  hub:     { url: 'https://www.nist.gov/autonomous-laboratories', t: 'NIST 자율실험실 허브 페이지' },
  csis:    { url: 'https://www.csis.org/blogs/perspectives-innovation/self-driving-labs-ai-and-robotics-accelerating-materials-innovation',
             t: 'CSIS, "Self-Driving Labs: AI and Robotics Accelerating Materials Innovation"' },
  eo:      { url: 'https://www.whitehouse.gov/presidential-actions/2025/11/launching-the-genesis-mission/',
             t: '백악관, "Launching the Genesis Mission" 행정명령 원문(2025.11.24)' },
  eoapp:   { url: 'https://www.presidency.ucsb.edu/documents/executive-order-14363-launching-the-genesis-mission',
             t: 'Executive Order 14363 전문(American Presidency Project)' },
  mofo:    { url: 'https://www.mofo.com/resources/insights/251211-executive-order-establishes-genesis-mission',
             t: 'Morrison Foerster, Genesis Mission 행정명령 해설 — 단계별 기한' },
  doe800:  { url: 'https://www.energy.gov/undersecretaryforscience/articles/us-department-energy-announces-more-800-million-partner',
             t: 'DOE, "$800M+ in Partner Commitments to the Genesis Mission"' },
  doefoa:  { url: 'https://science.osti.gov/grants/FOAs/FOAs/2026/DE-FOA-0003612',
             t: 'DOE, Genesis Mission 공모(DE-FOA-0003612, 2026.3)' },
  doeproj: { url: 'https://www.energy.gov/articles/secretary-energy-chris-wright-announces-first-genesis-mission-projects-selected-accelerate',
             t: 'DOE, "First Genesis Mission Projects Selected"' },
  wh5b:    { url: 'https://www.whitehouse.gov/releases/2026/07/45502/',
             t: '백악관, "More Than $5 Billion for the Genesis Mission"(2026.7)' },
  hpc:     { url: 'https://www.hpcwire.com/2026/08/07/doe-launches-open-model-initiative-for-genesis-mission/',
             t: 'HPCwire, "DOE Launches Open Model Initiative"(Genesis-Science-1, 2026.8)' },
  slas:    { url: 'https://www.slas.org/education/ansi-slas-microplate-standards/', t: 'SLAS, ANSI/SLAS 마이크로플레이트 표준' },
  lads:    { url: 'https://reference.opcfoundation.org/specs/OPC-30500-1/1', t: 'OPC Foundation, OPC UA for LADS Part 1: Basics (OPC 30500-1)' },
  ladsnews:{ url: 'https://opcfoundation.org/news/press-releases/new-international-spectaris-standard-for-laboratory-equipment-communication/',
             t: 'OPC Foundation, "New international SPECTARIS standard for lab equipment communication"' },
  sila:    { url: 'https://sila-standard.com/faq/', t: 'SiLA Standard FAQ — SiLA 2 개요' },
  allo:    { url: 'https://opcfoundation.org/news/press-releases/breakthrough-in-smarter-labs-spectaris-lads-showcases-integration-of-opc-ua-with-allotrope-standards/',
             t: 'OPC Foundation, "Spectaris LADS × Allotrope Standards 통합 시연"(2025.4)' },
  cwa:     { url: 'https://www.cencenelec.eu/media/CEN-CENELEC/CWAs/RI/2025/cwa17815_2025.pdf', t: 'CEN-CENELEC, CWA 17815 (2025)' },
  emmc:    { url: 'https://emmc.eu/focus-areas/digitalisation-interoperability/', t: 'EMMC, Digitalisation & Interoperability (EMMO 기반 소재 온톨로지)' },
  eu:      { url: 'https://roboticsandautomationnews.com/2025/12/02/europe-begins-major-push-to-standardize-materials-terminology-across-industry/97254/',
             t: 'Robotics & Automation News, "Europe begins major push to standardize materials terminology"(2025.12)' },
  think:   { url: 'https://www.cell.com/matter/fulltext/S2590-2385(26)00120-7',
             t: '"ThinkFactory 2025: harmonizing and accelerating self-driving laboratories", Matter (2026)' },
  nimsos:  { url: 'https://www.nims.go.jp/eng/press/2023/07/202307200.html', t: 'NIMS, "Development of NIMS-OS"(2023.7)' },
  zdnet:   { url: 'https://zdnet.co.kr/view/?no=20260811095548', t: 'ZDNet Korea, 자율실험실 협의체 분과별 참여기업(2026.8.11)' },
};
Object.values(S).forEach((x, i) => { x.n = i + 1; });

// 평문 조각 안의 **강조**를 굵은 런으로 분해
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
    const label = m[2].replace(/\*\*/g, '');
    const strong = bold || /\*\*/.test(m[2]);
    out.push(new ExternalHyperlink({ link: src.url, children: [new TextRun({ text: label, size, bold: strong, font: KO, color: '1B54B5', underline: {} })] }));
    last = re.lastIndex;
  }
  if (last < str.length) emitPlain(out, str.slice(last), size, bold, color);
  return out;
}
function cell(text, { w, size = 15, bold = false, color = '222222', fill, align = AlignmentType.LEFT } = {}) {
  const paras = (Array.isArray(text) ? text : [text]).map((t) =>
    new Paragraph({ alignment: align, spacing: { before: 10, after: 10, line: 200, lineRule: 'auto' }, children: runs(t, { size, bold, color }) }));
  return new TableCell({ width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    shading: fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined,
    margins: { top: 68, bottom: 68, left: 70, right: 70 }, children: paras });
}
const B = {
  top: { style: BorderStyle.SINGLE, size: 3, color: LINE }, bottom: { style: BorderStyle.SINGLE, size: 3, color: LINE },
  left: { style: BorderStyle.SINGLE, size: 3, color: LINE }, right: { style: BorderStyle.SINGLE, size: 3, color: LINE },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 3, color: LINE }, insideVertical: { style: BorderStyle.SINGLE, size: 3, color: LINE } };
function dataTable(widths, header, rows) {
  return new Table({ columnWidths: widths, width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, borders: B,
    rows: [new TableRow({ tableHeader: true, children: header.map((h, i) => cell(h, { w: widths[i], size: 15, bold: true, color: 'FFFFFF', fill: HEADFILL, align: AlignmentType.CENTER })) }),
      ...rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, i) =>
        cell(c, { w: widths[i], size: 15, bold: i === 0, fill: ri % 2 === 1 ? ALTFILL : undefined, align: i === 0 ? AlignmentType.CENTER : AlignmentType.LEFT })) }))] });
}
function h(t, before = 260) { return new Paragraph({ spacing: { before, after: 70 }, children: [new TextRun({ text: t, bold: true, size: 18, color: NAVY, font: KO })] }); }
function para(t, opts = {}) { return new Paragraph({ spacing: { before: 30, after: 30, line: 215, lineRule: 'auto' }, indent: opts.indent, children: runs(t, { size: opts.size || 16 }) }); }
function bullet(t) { return new Paragraph({ spacing: { before: 26, after: 26, line: 212, lineRule: 'auto' }, indent: { left: 240, hanging: 170 },
  children: [new TextRun({ text: '▪  ', size: 15, color: NAVY, font: KO }), ...runs(t, { size: 16 })] }); }
function calloutBox(titleText, lines, fill) {
  return new Table({ columnWidths: [CONTENT_W], width: { size: CONTENT_W, type: WidthType.DXA },
    borders: { top: { style: BorderStyle.SINGLE, size: 6, color: NAVY }, bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY },
               left: { style: BorderStyle.SINGLE, size: 18, color: NAVY }, right: { style: BorderStyle.SINGLE, size: 6, color: NAVY } },
    rows: [new TableRow({ children: [new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: fill || BOXFILL, color: 'auto' }, margins: { top: 120, bottom: 120, left: 180, right: 160 },
      children: [new Paragraph({ spacing: { after: 50, line: 220, lineRule: 'auto' }, children: [new TextRun({ text: titleText, bold: true, size: 16, color: NAVY, font: KO })] }),
        ...lines.map((l) => new Paragraph({ spacing: { before: 20, after: 20, line: 225, lineRule: 'auto' }, children: runs(l, { size: 16 }) }))] })] })] });
}

const children = [];
children.push(new Paragraph({ spacing: { after: 40, line: 240, lineRule: 'auto' },
  children: [new TextRun({ text: '자율실험실 표준화 동향 — NIST 4영역과 주요국 현황', bold: true, size: 26, color: NAVY, font: KO })] }));
children.push(new Paragraph({ spacing: { after: 150 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } },
  children: [new TextRun({ text: '조사일: 2026년 8월  |  공개 자료 기준 정리', size: 14, color: '666666', font: KO })] }));

children.push(calloutBox('한눈에 보기', [
  '미국 {prog:NIST}는 소재 R&D에 표준화된 자율실험 생태계가 없다는 점을 산업 확산의 장애물로 보고 ' +
  '시료관리·장비 제어통신·데이터지식관리·알고리즘모델 통합의 네 영역 표준을 개발하고 있다. ' +
  '다만 현재 단계는 표준 제정이 아니라 범위 정의와 커뮤니티 합의 형성이며, 발행된 표준 번호는 아직 없다.',
  '장비 통신·데이터 계층은 이미 유럽과 제약업계가 {lads:OPC UA LADS}·{sila:SiLA 2}·{allo:Allotrope}로 선점해 경쟁 구도까지 형성했다. ' +
  'NIST가 노리는 빈 곳은 소재 R&D 특유의 시료관리와 알고리즘·모델 통합이다.',
]));

// 1. NIST
children.push(h('1. NIST 자료 — 무엇을 인용할 것인가'));
children.push(dataTable([1500, 3200, 5618], ['성격', '자료', '내용과 인용 포인트'], [
  ['근거 원문', '{prog:Development of Standards to Support a Modular and Autonomous Laboratory Ecosystem}\n(NIST 프로그램 페이지)',
   '네 영역을 명시한 직접 출처. ① sample management ② instrument control and communication ③ data and knowledge management ④ algorithm and model integration. 표준 기반 생태계가 플랫폼 엔지니어링 비용과 진부화 위험을 줄이고, 장비·SW 벤더가 자율통합을 전제로 제품을 설계하게 만든다는 논리. 소재과학 특유의 난점으로 고체 시료 이송을 지목'],
  ['동료평가 논문', '{paper:Toward a composable, modular laboratory ecosystem for autonomous materials research and development}\nH. Joress, B. DeCost, K. Jones, F. Tavazza 외, Matter (2026.5.6)',
   '같은 구상의 학술판. 상용 기성품(COTS) 부품으로 조립 가능한 모듈형 SDL 생태계 비전. 보고서·논문에 인용할 때는 웹페이지보다 이쪽이 적절하다({rg:프리프린트}도 공개)'],
  ['정책 문서', '{sp1320:NIST SP 1320 — Driving U.S. Innovation in Materials and Manufacturing using AI and Autonomous Labs}\n(DOI 10.6028/NIST.SP.1320)',
   '표준 문서가 아니라 국가 전략 제안서. **국가자율소재과학센터(National Center for Autonomous Materials Science)** 설립 비전을 제시하고, 미국이 국가적 노력으로 이 패러다임을 채택해야 기술 리더십을 유지한다고 주장. 4영역 표준과 함께 자주 인용되는 자료'],
  ['추진 흔적', '{ws25:워크숍(2025.9)} — 생태계 비전 공유, 표준 범위 정의, 워킹그룹·컨소시엄 설립 논의\n{ws26:워크숍(2026.8)} — 장비 설계 베스트프랙티스 가이드\n{hub:NIST 자율실험실 허브}',
   '프로그램이 문서상 선언에 그치지 않았음을 보여주는 활동 기록. NIST는 MOF 자율합성 로봇 플랫폼 안에서 선행표준(precursor standards)을 먼저 만들고 있다고 밝힌다. 정책 해설로는 {csis:CSIS 분석}이 있다'],
]));

// 2. 판단
children.push(h('2. 실제로 추진 중이라고 볼 수 있는가'));
children.push(para('결론부터: **추진 중이 맞다. 단, “표준을 만들고 있다”가 아니라 “무엇을 표준화할지 정하는 단계”다.** 이 구분을 흐리면 인용이 부정확해진다.'));
children.push(dataTable([2100, 4100, 4118], ['', '확인되는 것', '아직 확인되지 않는 것'], [
  ['NIST 4영역 표준', '프로그램 페이지 개설 · 워크숍 2회(2025.9 / 2026.8) · 동료평가 논문 게재(Matter 2026.5) · MOF 로봇 플랫폼 내 선행표준 개발 · 정책문서 SP 1320 발간',
   '발행된 표준 번호(ANSI·ISO·ASTM 등) 없음 · 공식 표준화기구 절차 착수 여부 확인 안 됨 · 워킹그룹/컨소시엄이 실제로 발족했는지 불명'],
]));
children.push(para('따라서 정확한 표현은 “NIST가 자율실험 표준을 제정했다”가 아니라 **“NIST가 네 영역을 표준화 대상으로 정의하고 커뮤니티 합의를 형성하는 단계에 있다”** 이다.', { size: 16 }));

// 3. Genesis Mission
children.push(h('3. 제네시스 미션 행정명령을 근거로 쓸 수 있는가'));
children.push(para('{eoapp:행정명령 14363 “Launching the Genesis Mission”}(2025.11.24 서명)은 DOE 주도의 AI 과학 국가 미션이다. 인용하신 (e)항은 원문 그대로이며, 단계별 기한의 일부다.'));
children.push(dataTable([2000, 8318], ['기한', '내용'], [
  ['90일', '{mofo:DOE 컴퓨팅 자원 인벤토리}'],
  ['120일', '초기 데이터·모델 자산 식별'],
  ['240일', '**국립연구소 및 참여 연방 연구시설의 로봇 실험실·생산시설 역량 검토 — AI 주도 실험·제조 수행 능력, 자동화 및 AI 증강 워크플로, 그리고 이에 필요한 기술·운영 표준을 포함**'],
  ['270일', '최소 1개 국가 과학기술 챌린지에 대한 초기운용능력(IOC) 시연'],
]));
children.push(new Paragraph({ spacing: { before: 120, after: 40 }, children: [new TextRun({ text: '쓸 수 있는 근거', bold: true, size: 16, color: NAVY, font: KO })] }));
children.push(bullet('**“AI 주도 실험에 필요한 기술·운영 표준”이 대통령 명령의 문언에 직접 들어가 있다.** 자율실험 표준화가 개별 기관의 연구 과제가 아니라 연방 차원 의제로 격상됐다는 근거로는 이보다 강한 자료가 드물다.'));
children.push(bullet('행정명령이 문서로만 남지 않았다는 실행 증거가 있다 — {doe800:파트너 약정 8억 달러 이상}, {doefoa:2026년 3월 2억 9,376만 달러 규모 공모(21개 분야)}, {doeproj:첫 과제 선정(신청 5,000건 초과)}, {wh5b:2026년 7월 50억 달러 이상 발표}, {hpc:2026년 8월 개방형 모델 Genesis-Science-1 착수}.'));

children.push(new Paragraph({ spacing: { before: 140, after: 40 }, children: [new TextRun({ text: '그대로 쓰면 논리가 새는 지점 — 셋', bold: true, size: 16, color: 'A0522D', font: KO })] }));
children.push(dataTable([1300, 9018], ['', '주의할 점'], [
  ['부처가 다르다', '제네시스 미션은 **에너지부(DOE)**, 네 영역 표준은 **상무부 산하 NIST**다. 행정명령은 NIST 프로그램의 추진 근거가 아니다. “미국 정부가 자율실험 표준을 중요하게 본다”는 입증하지만, “NIST가 실제로 하고 있다”의 근거로 쓰면 부처를 뒤섞은 것이 된다'],
  ['명령 내용은 ‘검토’다', '(e)항이 지시한 것은 표준 제정이 아니라 **역량 검토(review)와 필요한 표준의 식별**이다. 산출물은 표준이 아니라 검토 결과다. “행정명령으로 표준 제정에 착수했다”는 과장이 된다'],
  ['기한이 이미 지났다', '2025년 11월 24일 기준 240일은 **2026년 7월 22일**에 도래했다. 그러나 이 검토의 결과물이 공개된 흔적을 찾지 못했다. 인용할 때는 “검토가 지시됐다”까지만 쓰고, 결과를 본 것처럼 쓰지 않는 편이 안전하다'],
]));
children.push(calloutBox('권장 인용 방식', [
  '두 자료를 **각각 다른 역할로** 쓰는 것이 정확하다. {prog:NIST 4영역}은 “무엇을 표준화하려 하는가”(표준의 내용) 근거로, ' +
  '{eo:제네시스 미션 행정명령}은 “왜 지금이고 미국이 얼마나 진지한가”(정책 우선순위) 근거로 쓴다. ' +
  '둘을 묶어 “미국이 자율실험 표준을 제정하고 있다”로 뭉뚱그리면 부처·단계·산출물이 모두 어긋난다.',
], WARNFILL));

// 4. 주요국
children.push(h('4. 주요국·국제 표준 현황'));
children.push(dataTable([1250, 2500, 1450, 5118], ['주체', '표준·이니셔티브', '계층', '내용'], [
  ['미국\nNIST', '모듈형·자율 실험실 생태계 표준(4영역)', '생태계 전반', '시료·통신·데이터·알고리즘을 하나로 묶으려는 유일한 시도. 범위 정의 단계'],
  ['미국\nSLAS/ANSI', '{slas:ANSI/SLAS 1~4 마이크로플레이트 표준}', '물리 규격', '1995~96년 논의 시작, 2004년 승인. 랩웨어 치수·풋프린트의 사실상 기반 — 이것이 있어서 로봇 핸들링이 성립한다'],
  ['독일 주도\n국제', '{lads:OPC UA LADS (OPC 30500)}', '장비 제어·통신', 'OPC Foundation + SPECTARIS + VDMA 공동. {ladsnews:2024년 1월 Part 1(Base System) 공개}. 제조사 독립 개방형 표준'],
  ['국제', '{sila:SiLA 2}', '장비 제어·통신', 'gRPC 기반. LADS와 기술·접근이 달라 경쟁·병존 관계'],
  ['국제\n컨소시엄', '{allo:Allotrope(ADF/ADM)} · AnIML', '데이터·온톨로지', '제약 중심. 2025년 4월 해커톤에서 Allotrope 온톨로지·Simple Model의 OPC UA 통합 시연'],
  ['EU', '{cwa:CEN/CENELEC CWA} · {emmc:EMMC/EMMO}', '소재 온톨로지·용어', '{eu:전문가 23명이 ①용어 정의 ②검토·승인 워크플로 ③기술 구현 3개 소그룹}으로 운영. CHAMEO(특성분석 방법론 온톨로지), CWA 17815, DiMAT 프로젝트'],
  ['캐나다', '{think:ThinkFactory 2025}\n(Acceleration Consortium + CMAC)', '커뮤니티 합의', 'SDL 조화·가속 논의. AI/ML·데이터·오케스트레이션·로보틱스 4개 트랙, 50명 이상 참여. 상호운용성·안전성·재현성 표준을 우선과제로 도출, Matter 게재'],
  ['일본', '{nimsos:NIMS-OS}', '사실상 표준 계층', '공식 표준 제정은 아니나, 어떤 탐색 AI와 어떤 로봇이든 폐루프로 잇는 범용 미들웨어를 공개해 소프트웨어 계층을 선점'],
  ['한국', '{zdnet:자율실험실 산·학·연 협의체\n「기술·플랫폼·표준」 분과}', '생태계 전반', '32명 참여(파크시스템스·에이치비솔루션·에이블랩스·로봇앤드디자인 등). 장비 인터페이스·데이터 표준 정립, 레퍼런스 자율실험실 구축'],
]));
children.push(new Paragraph({ spacing: { before: 80, line: 200, lineRule: 'auto' },
  children: runs('※ 일본·중국의 SDL 전용 국가표준 제정 활동, 그리고 Acceleration Consortium과 CSA·ISO 등 공식 표준기구의 워킹그룹은 공개 자료에서 확인되지 않았다.', { size: 13, color: '666666' }) }));

// 5. 구도
children.push(h('5. 계층 구도와 빈 곳'));
children.push(dataTable([2400, 4200, 3718], ['계층', '이미 있는 표준', '상태'], [
  ['물리 규격', 'ANSI/SLAS 마이크로플레이트', '확립됨 (미국, 1996~)'],
  ['장비 통신', 'OPC UA LADS ↔ SiLA 2', '경쟁 구도 (독일·유럽 주도)'],
  ['데이터·의미', 'Allotrope / AnIML / EMMO·CHAMEO', '제약 컨소시엄과 EU가 분점'],
  ['시료 관리', '— 없음 —', '**빈 곳** → NIST 4영역에 포함'],
  ['알고리즘·모델 통합', '— 없음 —', '**빈 곳** → NIST 4영역에 포함'],
]));
children.push(para('통신과 데이터 계층은 이미 남이 만들어 놓았고 경쟁까지 붙었다. NIST가 네 영역에 통신·데이터뿐 아니라 **시료관리와 알고리즘·모델 통합**을 넣은 것은, 소재 R&D 자율실험에 고유하면서 아직 주인이 없는 두 계층을 겨냥한 것으로 읽힌다. 한국 협의체의 「기술·플랫폼·표준」 분과가 어느 계층을 목표로 삼을지 정할 때 참고할 지점이다.'));

// 출처
children.push(h('출처'));
const srcList = Object.values(S).sort((a, b) => a.n - b.n);
const half = Math.ceil(srcList.length / 2);
const srcRows = [];
for (let i = 0; i < half; i++) {
  const mk = (s) => s ? new TableCell({ width: { size: Math.floor(CONTENT_W / 2), type: WidthType.DXA }, margins: { top: 10, bottom: 10, left: 0, right: 120 },
      children: [new Paragraph({ spacing: { before: 8, after: 8, line: 200, lineRule: 'auto' },
        children: [new TextRun({ text: `[${s.n}] `, size: 13, color: '444444', font: KO }),
          new ExternalHyperlink({ link: s.url, children: [new TextRun({ text: s.t, size: 13, color: '1B54B5', underline: {}, font: KO })] })] })] })
    : new TableCell({ width: { size: Math.floor(CONTENT_W / 2), type: WidthType.DXA }, children: [new Paragraph('')] });
  srcRows.push(new TableRow({ children: [mk(srcList[i]), mk(srcList[i + half])] }));
}
children.push(new Table({ columnWidths: [Math.floor(CONTENT_W / 2), Math.floor(CONTENT_W / 2)], width: { size: Math.floor(CONTENT_W / 2) * 2, type: WidthType.DXA },
  borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
             insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE } }, rows: srcRows }));
children.push(new Paragraph({ spacing: { before: 70 },
  children: runs('원문 접근 제약: 조사 환경에서 nist.gov·sciencedirect.com·whitehouse.gov 등이 차단돼 직접 열람하지 못했다. 검색 결과에 인용된 원문 문구와 복수 출처 교차 확인으로 정리했으므로 인용 전 링크를 직접 확인하시기를 권한다.', { size: 12, color: '666666' }) }));

const doc = new Document({
  styles: { default: { document: { run: { font: KO, size: 16, color: '222222' }, paragraph: { spacing: { line: 220, lineRule: 'auto' } } } },
    characterStyles: [{ id: 'Hyperlink', name: 'Hyperlink', basedOn: 'DefaultParagraphFont', run: { color: '1B54B5', underline: {} } }] },
  sections: [{ properties: { page: { size: { width: PAGE_W, height: convertMillimetersToTwip(297) },
    margin: { top: MARGIN, bottom: convertMillimetersToTwip(12), left: MARGIN, right: MARGIN } } }, children }],
});
const out = process.argv[2] || '/home/user/kist/report/자율실험실_표준화_동향.docx';
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(out, b); console.log('written:', out, b.length); });
