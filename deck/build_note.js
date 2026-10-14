// 심층 학습 노트: cIMPACT-NOW update 12 / PDGFRA amplification in IDH-mutant astrocytoma
// Style follows the existing 학습노트 docx files (Malgun Gothic, navy headers, tinted summary boxes, per-section sources).
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, LevelFormat, BorderStyle, HeadingLevel, PageBreak, Footer, PageNumber,
  TableOfContents,
} = require("docx");

const FONT = "Malgun Gothic";
const NAVY = "1F3864", BLUE = "2E75B6", GRAY = "7F7F7F", TINT = "DEEAF6", LIGHT = "F2F2F2",
  YELLOW = "FFF2CC", GREEN = "E2EFDA";
const PAGE_W = 11906, MARGIN = 1134, CONTENT_W = PAGE_W - 2 * MARGIN; // A4, 2 cm margins

// ---------- inline markup: **bold**, *italic*, [보충]/[원문 확인 권장] tags highlighted ----------
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*|\[보충[^\]]*\]|\[원문[^\]]*\]|\[제안[^\]]*\])/g;
  let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, ...base }));
    const t = m[0];
    if (t.startsWith("**")) out.push(new TextRun({ text: t.slice(2, -2), bold: true, font: FONT, ...base }));
    else if (t.startsWith("*")) out.push(new TextRun({ text: t.slice(1, -1), italics: true, font: FONT, ...base }));
    else out.push(new TextRun({ text: t, bold: true, color: BLUE, font: FONT, ...base }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, ...base }));
  return out;
}
const P = (t, o = {}) => new Paragraph({ children: runs(t, o.run), spacing: { after: 100, line: 300 }, alignment: o.align, ...o.p });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT })],
  spacing: { before: 360, after: 160 } });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: FONT })],
  spacing: { before: 240, after: 120 } });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: t, font: FONT })],
  spacing: { before: 180, after: 80 } });
// bullets: array of strings; nested level via leading "  " (2 spaces per level)
function B(items) {
  return items.map((s) => {
    const lvl = Math.floor((s.match(/^ */)[0].length) / 2);
    return new Paragraph({ numbering: { reference: "bul", level: Math.min(lvl, 2) },
      children: runs(s.trim()), spacing: { after: 60, line: 290 } });
  });
}
const SRC = (t) => new Paragraph({ children: [new TextRun({ text: "출처: " + t, italics: true, color: GRAY, size: 18, font: FONT })],
  spacing: { before: 60, after: 200 } });

const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(content, w, o = {}) {
  const paras = (Array.isArray(content) ? content : [content]).flatMap((c) =>
    typeof c === "string" ? [new Paragraph({ children: runs(c, o.run), spacing: { after: 40, line: 270 }, alignment: o.align })] : [c]);
  return new TableCell({ children: paras, width: { size: w, type: WidthType.DXA }, borders,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 70, bottom: 70, left: 110, right: 110 }, columnSpan: o.span });
}
// table(headers, rows, widths(ratios)) — header navy w/ white bold text; first col shaded light
function T(headers, rows, ratios, o = {}) {
  const sum = ratios.reduce((a, b) => a + b, 0);
  const w = ratios.map((r) => Math.round((r / sum) * CONTENT_W));
  w[w.length - 1] += CONTENT_W - w.reduce((a, b) => a + b, 0);
  const head = new TableRow({ tableHeader: true, children: headers.map((h, i) =>
    cell(h, w[i], { fill: NAVY, run: { bold: true, color: "FFFFFF" } })) });
  const body = rows.map((r) => new TableRow({ children: r.map((c, i) => {
    const hi = o.highlight && o.highlight(r);
    return cell(c, w[i], { fill: hi ? YELLOW : (i === 0 ? LIGHT : undefined) });
  }) }));
  return [new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: w, rows: [head, ...body] }),
    new Paragraph({ children: [], spacing: { after: 80 } })];
}
// tinted box with title + bullet list
function BOX(title, items, fill = TINT) {
  const kids = [new Paragraph({ children: [new TextRun({ text: title, bold: true, font: FONT, color: NAVY, size: 22 })], spacing: { after: 80 } }),
    ...B(items)];
  return [new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [new TableCell({ children: kids, width: { size: CONTENT_W, type: WidthType.DXA },
      shading: { fill, type: ShadingType.CLEAR, color: "auto" }, borders: { top: { style: BorderStyle.SINGLE, size: 4, color: fill },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: fill }, left: { style: BorderStyle.SINGLE, size: 4, color: fill },
        right: { style: BorderStyle.SINGLE, size: 4, color: fill } },
      margins: { top: 140, bottom: 140, left: 200, right: 200 } })] })] }),
    new Paragraph({ children: [], spacing: { after: 120 } })];
}

// =====================================================================
const C = [];

// Q&A block: question header row (navy), then 답변 / 근거 / 인정할 한계 rows
function QA(no, q, ans, ev, lim) {
  const w1 = Math.round(CONTENT_W * 0.16), w2 = CONTENT_W - w1;
  const row = (label, content, fill) => new TableRow({ children: [
    cell(label, w1, { fill: fill || LIGHT, run: { bold: true } }),
    cell(content, w2, { fill: fill ? undefined : undefined }),
  ] });
  const rows = [
    new TableRow({ children: [cell("Q" + no + ". " + q, CONTENT_W, { fill: NAVY, run: { bold: true, color: "FFFFFF" }, span: 2 })] }),
    row("짧은 답", ans, TINT),
    row("근거", ev),
  ];
  if (lim) rows.push(row("인정할 한계", lim, YELLOW));
  return [new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [w1, w2], rows }),
    new Paragraph({ children: [], spacing: { after: 140 } })];
}

// ---------- Title ----------
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 600, after: 120 },
  children: [new TextRun({ text: "cIMPACT-NOW와 IDH-mutant Glioma Grading", bold: true, size: 44, color: NAVY, font: FONT })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
  children: [new TextRun({ text: "Update 1–12 전체 · Update 12의 PDGFRA amplification → grade 4 · 영상 · 치료 · 공격적 질문 대비", size: 26, font: FONT })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 },
  children: [new TextRun({ text: "Disease review 발표 대비 심층 학습 노트", bold: true, size: 24, font: FONT })] }));
C.push(P("**핵심 원저**: Brat DJ, Aldape K, French PJ, Louis DN, Nasrallah MP, Reifenberger G, Reuss DE, Touat M, van den Bent MJ, Wesseling P, Figarella-Branger D. *cIMPACT-NOW Update 12: Refining pathology-based risk stratification and grading for IDH-mutant gliomas.* Neuro-Oncology 2026;28(9):2069–2089. doi:10.1093/neuonc/noag108", { align: AlignmentType.CENTER }));
C.push(P("**발표 Figure 출처**: Park YW, Wagner A, Karschnia P, Vollmuth P, Sahm F, Ahn SS, Kim HS, Park JE. *Imaging of IDH-Mutant Gliomas in the Era of IDH Inhibitors: What Radiologists Need to Know.* Korean J Radiol 2026 (in press), Fig. 2", { align: AlignmentType.CENTER, p: { spacing: { after: 360 } } }));

C.push(...BOX("이 노트의 출처와 한계 (발표 전 반드시 숙지)", [
  "**Update 12 본문은 초록만 확인**했다. PMC 전문 본문이 이 환경에서 열람되지 않았음 → 원문 Table의 세부 정의·인용 수치는 **[원문 확인 권장]**. Update 12의 내용은 ① PubMed 초록 ② Update 12를 직접 인용한 2차 문헌(Richardson 2026 전문, Tralongo 2026 리뷰 전문, Tauziède-Espariat 2026 IHC 전문·Ann Pathol 초록) ③ Park YW KJR 2026 Fig. 2로 재구성.",
  "**전문을 직접 읽은 문헌**: Update 3, Update 5, Update 10, Update 11, Virata 2026, Richardson 2026, CATNON methylation 2021(Tesileanu), PDGFRA IHC 2026, Tralongo 2026(methylation 리뷰), Kurokawa 2026(Jpn J Radiol), Ozono 2026.",
  "**초록만 확인한 문헌**: WHO CNS5 요약(Louis 2021), Update 6–9, Yang 2019, CATNON 2021·2026, INDIGO, ASCO-SNO 2022, EANO 2021, Alfattal, Yang AJNR, Singh, Imai, Jo, Wong 등.",
  "**Update 1, 2, 4는 PubMed에 초록이 없음** → 제목 수준의 내용만 적고 그 이상은 **[보충]** 으로 표시.",
  "PMC 텍스트에서 이탤릭 유전자명이 빠지는 문제가 있어, Virata 2026의 유전자별 수치는 **본문 나열 순서로 대응** → **[원문 확인 권장]**.",
  "표기: **[보충]** = 다른 문헌·일반 지식, **[제안]** = 필자의 실무 제안, **[원문 확인 권장]** = 원문 대조 필요.",
], YELLOW));

C.push(new Paragraph({ children: [new PageBreak()] }));
C.push(new TableOfContents("목차", { hyperlink: true, headingStyleRange: "1-2" }));
C.push(new Paragraph({ children: [new PageBreak()] }));

// =============== 0. How to use ===============
C.push(H1("0. 발표 준비 전략"));
C.push(...B([
  "**발표 흐름(PPT 5장)**: ① WHO CNS5 이후 cIMPACT-NOW → ② Update 12 grading 알고리즘 → ③ PDGFRA amp 근거 → ④ 영상 → ⑤ 왜 중요한가(이 증례의 치료 변경)",
  "**질문은 크게 6가지 방향**으로 올 가능성이 높음 → 11장에 유형별로 정리",
  "  개념(cIMPACT-NOW 위상, WHO와의 관계) / 근거(연구 질, 이견) / 기술(검출법, cut-off) / 영상(왜 enhancement가 없나) / 치료(CCRT의 근거) / 증례(63세, 검체 대표성)",
  "**방어 원칙**",
  "  모르는 수치는 추측하지 말고 \"원문 Table에서 확인하겠다\"로 답함",
  "  근거 수준을 구분해서 말함: 무작위 시험(CATNON, INDIGO) vs 후향적 코호트(Yang, Virata, Richardson) vs 전문가 합의(cIMPACT-NOW)",
  "  이견이 있는 부분(Virata 2026의 intermediate 분류, concurrent TMZ)은 **먼저 인정**하고 우리 결정의 논리를 제시",
]));
C.push(...BOX("30초 요약 (발표 첫머리·마무리용)", [
  "WHO CNS5 이후에도 cIMPACT-NOW가 grading을 계속 다듬고 있고, 2026년 Update 12에서 **PDGFRA amplification이 IDH-mutant astrocytoma의 grade 4 기준**이 됐다.",
  "근거는 여러 후향적 코호트와 CATNON 분자 분석이며, 1,140명 validation에서 PDGFRA amp으로 grade 4가 된 환자의 median OS는 51.4개월로 grade 3 유지군(126.2개월)보다 짧았다.",
  "PDGFRA amp 자체의 영상 소견은 없지만, 이 환자처럼 **ADC·rCBV가 조직 grade보다 나쁘면** 영상의학과의 피드백이 분자 재검토와 grade 상향, 치료 변경으로 이어질 수 있다.",
  "English: \"Update 12 made PDGFRA amplification a grade 4 criterion. Imaging cannot detect it, but discordant ADC/rCBV can trigger the molecular work-up that changes grade and treatment.\"",
], GREEN));

// =============== 1. Summary ===============
C.push(H1("1. 한눈에 보는 핵심 요약"));
C.push(...BOX("Take-home messages", [
  "WHO CNS5(2021): **Astrocytoma, IDH-mutant = 단일 type, CNS WHO grade 2–4**. Grade 4 = MVP 또는 necrosis 또는 **CDKN2A/B homozygous deletion(HD)**.",
  "Update 12(2026): **PDGFRA amplification → grade 4** (\"highly aggressive clinical behavior, consistent with CNS WHO grade 4\").",
  "Grade 2/3 경계: **mitotic count ≥3 / 2.4 mm²** 제안(\"could be considered\"). 이전(Update 5)에는 검증된 수치가 없었고 관행상 전체 검체 ≥2 mitoses.",
  "Grade 3 시사 분자 변화: PIK3CA/PIK3R1 mutation, EGFR alteration, MYCN amp, RB pathway(CDK4·CDK6·CCND2 amp, RB1 HD/mutation, CDKN2A mutation), methylation **G-CIMP-low / A_IDH_HG**(다른 기준에 따라 grade 3 또는 4).",
  "Oligodendroglioma: grading을 바꿀 새 근거 없음. 단 mitosis만 높고 necrosis·MVP·CDKN2A/B HD·**MRI 조영증강이 없는** grade 3은 생존이 길 수 있음.",
  "PDGFRA amp → grade 4 근거: Yang 2019(18.8%, 최고위험군), CATNON(독립 예후인자), Virata 2026(단독 alteration HR 4.5), **Richardson 2026 validation(51.4 vs 126.2개월, HR 1.9)**.",
  "이견: Virata 2026은 PDGFRA alteration을 grade 4가 아닌 **intermediate** 군으로 봄. 다변량에서는 독립 유의성 미달.",
  "검출: FISH, methylation array CNV, NGS, SNP/CGH array. **PDGFRA IHC**(≥10% 강양성)는 민감도 100%, 특이도 89%의 screening 도구.",
  "영상: PDGFRA amp 특이 소견은 **없음**. T2-FLAIR mismatch 소실, poorly defined enhancing margin, ADC↓·rCBV↑는 고위험 단서.",
  "치료: grade 3 = RT → adjuvant TMZ(CATNON), grade 4 IDH-mut = **grade 3 또는 GBM 권고 중 선택 가능**(ASCO-SNO 2022), vorasidenib은 grade 2 근거(INDIGO).",
]));
C.push(SRC("Brat 2026 (abstract); Brat 2020 Update 5 (PMC8443062); Richardson 2026 (PMC13391642); Virata 2026 (PMC13437890); Tauziède-Espariat, Brain Pathol 2026 (PMC13535536); Mohile, J Clin Oncol 2022 (abstract)"));

// =============== 2. Biology ===============
C.push(H1("2. 기초 생물학 — 질문의 출발점"));
C.push(H2("2-1. IDH mutation과 D-2HG"));
C.push(...B([
  "IDH1/IDH2 mutation은 adult lower-grade diffuse glioma 대부분에서 **초기(early) 사건**",
  "가장 흔한 변이: **IDH1 p.R132H**. 그 외 IDH1 R132 다른 치환, IDH2 R172는 드묾",
  "정상 IDH: isocitrate → α-ketoglutarate(α-KG). 변이 IDH: **α-KG → D-2-hydroxyglutarate(D-2HG)** (gain-of-function)",
  "고농도 D-2HG가 α-KG 의존 효소들을 방해 → **DNA·histone hypermethylation** → glioma CpG island methylator phenotype(**G-CIMP**)",
  "IDH variant 종류는 **현재 독립적인 grading 기준도, 치료 선택 biomarker도 아님**",
  "임상 연결: MRS의 **2HG peak**로 IDH 변이 비침습 추정, IDH inhibitor(vorasidenib)의 표적",
]));
C.push(SRC("Tralongo, Neuropathol Appl Neurobiol 2026 (PMC13498940 전문); Park YW, Korean J Radiol 2026 (in press) p.2; Kurokawa 2026 (PMC13526116)"));

C.push(H2("2-2. Astrocytoma의 분자 진화"));
C.push(...B([
  "계통(lineage) 결정: **IDH + ATRX + TP53** (1p/19q codeletion 없음) — 이것들은 lineage를 지지할 뿐 **grade를 결정하지 않음**",
  "진행(progression) 과정에서 추가되는 변화: **cell-cycle(RB pathway), RTK, PI3K** 신호 이상 → Update 12의 grade 3/4 marker가 바로 이 축",
  "  RB pathway: CDKN2A/B HD, CDK4/CDK6/CCND2 amp, RB1 HD·mutation",
  "  RTK: **PDGFRA amp**, EGFR alteration",
  "  PI3K: PIK3CA, PIK3R1 mutation",
  "Tesileanu 2022 systematic review(44편): 예후와의 연관이 **여러 연구에서 재현된** marker — non-canonical IDH mutation, PI3K mutation, CDKN2A/B HD, **PDGFRA amp**, CN-neutral 17p loss, 19q loss, G-CIMP-low, high total CNV, high TMB",
  "  영향을 받는 경로: RB 신호, RTK-PI3K-mTOR 신호, genomic stability, (epi)genetic gene regulation",
  "Yang 2019: PDGFRA amp, CDKN2A HD, CDK4 amp가 **서로 배타적** → 같은 '증식 신호 활성화'에 이르는 서로 다른 경로로 해석 가능 [보충]",
]));
C.push(SRC("Tralongo 2026 (PMC13498940); Tesileanu, Eur J Cancer 2022 (10.1016/j.ejca.2022.08.016, abstract); Yang 2019 (abstract)"));

C.push(H2("2-3. G-CIMP와 methylation class — 헷갈리기 쉬운 용어"));
C.push(...T(["용어", "정의·유래", "Grading에서의 위치"], [
  ["G-CIMP-high", "IDH 변이에 따른 광범위 hypermethylation (대부분의 IDH-mut glioma)", "기본 상태. 예후 양호군"],
  ["**G-CIMP-low**", "연구 기반 분류. 상대적 **hypomethylation**, RB pathway 이상 동반(>75%), 예후 불량", "Update 12: grade 3–4 시사. **단독 grade 4 기준 아님**"],
  ["A_IDH_LG", "Methylation classifier의 저등급 subclass", "—"],
  ["**A_IDH_HG**", "Classifier 기반 고등급 subclass. CNV 부담·cell-cycle 이상 많음", "Update 12: grade 3–4 시사. **단독 grade 4 기준 아님**"],
], [1.6, 4.2, 3.4]));
C.push(...B([
  "**G-CIMP-low와 A_IDH_HG는 생물학적으로 겹치지만 같은 말이 아님** — 분석 방법이 다름. CATNON에서는 G-CIMP-low가 모두 A_IDH_HG로 분류됨",
  "Methylation high-grade signature만 있고 확립된 grade 4 기준이 없으면 → **\"adverse molecular feature\"로 기술**하고 자동으로 grade 4로 바꾸지 말 것",
  "반대로 necrosis·MVP·명확한 CDKN2A/B HD가 있으면 methylation이 저등급이어도 grade 4",
  "Update 5 당시: G-CIMP-low는 1,122예 연구에서 IDH-mut astrocytic glioma의 작은 subset, 절반이 grade 4·절반이 grade II/III 조직 → 검증 cohort 필요하다고 보류",
]));
C.push(SRC("Tralongo 2026 (PMC13498940); Tesileanu 2021 CATNON (PMC8408862); Update 5 (PMC8443062)"));

// =============== 3. WHO CNS5 ===============
C.push(H1("3. WHO CNS5 (2021) — 기준점"));
C.push(H2("3-1. 일반 원칙"));
C.push(...B([
  "6번째 판(5th edition) WHO CNS 분류, 2016 revised 4th edition과 **cIMPACT-NOW 작업 위에 구축**",
  "분자진단의 역할 확대 + 조직·IHC 유지 → **integrated diagnosis, layered report** 강조",
  "Grade는 **tumor type 안에서** 매김 ('CNS WHO grade'), **아라비아 숫자** 사용(Update 5가 제안: II vs III 오타 방지)",
  "DNA methylome profiling 같은 신기술 기반의 새 type·subtype 도입",
]));
C.push(SRC("Louis DN, Neuro-Oncol 2021;23:1231 (10.1093/neuonc/noab106, abstract); Update 5 (PMC8443062)"));

C.push(H2("3-2. Adult-type diffuse glioma"));
C.push(...T(["Type", "정의", "Grade"], [
  ["Astrocytoma, IDH-mutant (AST)", "IDH mutation, 1p/19q codeletion 없음 (ATRX·TP53 이상 흔함)", "2, 3, 4"],
  ["Oligodendroglioma, IDH-mutant and 1p/19q-codeleted (ODG)", "IDH mutation + **whole-arm** 1p/19q codeletion", "2, 3"],
  ["Glioblastoma, IDH-wildtype (GBM)", "IDH-wt diffuse astrocytic glioma + (MVP/necrosis 또는 EGFR amp / TERTp mutation / +7/−10)", "4"],
], [3.0, 4.6, 1.2]));
C.push(...B([
  "폐지: \"anaplastic astrocytoma\", \"**glioblastoma, IDH-mutant**\" → IDH-mut 고등급은 \"Astrocytoma, IDH-mutant, CNS WHO grade 4\"",
  "\"Molecular GBM\"은 공식 명칭이 아님 — 조직학적 grade 2/3이지만 분자 기준으로 GBM인 경우의 비공식 용어",
  "한 코호트(1,211예): 조직학적 grade 4 GBM 64.2%, molecular GBM 3.3%, AST 16.8%, ODG 15.7%",
]));
C.push(SRC("Kurokawa, Jpn J Radiol 2026 (PMC13526116); Virata 2026 Introduction"));

C.push(H2("3-3. IDH-mutant astrocytoma의 grading (CNS5, Update 5 기반)"));
C.push(...T(["Grade", "기준 (Update 5 원문 기반)", "비고"], [
  ["2", "Significant mitotic activity, anaplasia, MVP, necrosis, CDKN2A/B HD **모두 없음**", "Median OS >10년"],
  ["3", "Mitotic activity·anaplasia 증가, MVP·necrosis·CDKN2A/B HD 없음", "**검증된 mitosis cut-off 없음**. 관행: 전체 검체 ≥2개, 작은 생검 1개"],
  ["4", "MVP **또는** necrosis **또는** CDKN2A/B HD (조합 가능)", "'GBM, IDH-mutant' 대신 'Astrocytoma, IDH-mutant, grade 4'"],
], [0.8, 5.2, 3.2]));
C.push(...B([
  "CDKN2A/B HD 빈도(Update 5 인용 문헌): grade II **0–12%**, III **6–20%**, IV **16–34%**",
  "CDKN2A/B HD 검출: FISH, qPCR, MLPA, microarray, NGS — **p16 IHC는 deletion과 상관이 낮음**",
  "Update 5의 경고: 대부분 후향적 코호트(나이·치료 차이로 교란 가능), 추적 기간이 median OS(>10년)에 비해 짧음",
  "9p21 HD는 CDKN2A/B 외 주변 유전자도 함께 소실시킬 수 있음",
]));
C.push(SRC("Brat DJ et al. Update 5. Acta Neuropathol 2020;139:603 (PMC8443062 전문)"));

C.push(H2("3-4. 'Integrated grade가 histologic grade보다 우선'이라는 원칙 (Update 3)"));
C.push(...B([
  "Update 3(2018): IDH-wt diffuse astrocytic glioma가 조직학적 grade II/III이어도 **EGFR high-level amp, +7/−10(whole chromosome), TERTp mutation** 중 하나면 grade IV에 해당",
  "근거: 이런 종양의 생존이 IDH-wt GBM과 비슷 / methylation에서도 GBM과 같은 cluster",
  "**핵심 문장**: \"grade는 임상 경과를 예측하기 위한 것이므로, 이 경우 **integrated histological and molecular grade가 strictly histological grade보다 우선**한다\"",
  "추가 논리: **조직 소견은 종양 안에서 공간적으로 다를 수 있음**(sampling 문제)",
  "주의(Update 3 원문): TERTp mutation은 PXA, ganglioglioma 등 다른 IDH-wt 종양에도 있을 수 있음 → 임상·영상·조직이 diffuse astrocytic glioma에 확실히 부합할 때만 적용. CDKN2A HD는 PXA 등에도 있어 단독 기준이 아님",
  "→ **Update 12의 PDGFRA amp → grade 4도 같은 철학의 연장선**(IDH-mut에 적용)",
]));
C.push(SRC("Brat DJ et al. Update 3. Acta Neuropathol 2018;136:805 (PMC6204285 전문)"));

// =============== 4. cIMPACT-NOW 1-12 ===============
C.push(H1("4. cIMPACT-NOW 전체 정리 (Update 1–12)"));
C.push(H2("4-1. cIMPACT-NOW란"));
C.push(...B([
  "**Consortium to Inform Molecular and Practical Approaches to CNS Tumor Taxonomy – Not Official WHO**, 2016년 설립",
  "목적: WHO 개정 사이에 분자 지식을 빠르게 진료에 반영",
  "구성: 신경병리, 신경영상, 분자유전, 신경종양 등 국제 전문가. **Working committee** 단위로 주제별 논의(teleconference, 대면 회의 — 예: 2019 Utrecht 회의)",
  "위상: **WHO 공식 정의가 아님**. 다음 WHO 판의 기초 자료가 됨 → Update 1–7 대부분이 CNS5에 반영",
  "Update 11 원문: 권고는 \"WHO 정의·진단기준의 공식 변경이 아니라 WHO CNS6 이전의 진단 가이드\"",
]));
C.push(SRC("Update 3 (PMC6204285) Introduction; Update 5 (PMC8443062); Update 11 abstract; Kurokawa 2026"));

C.push(H2("4-2. 전체 목록"));
C.push(...T(["#", "연도·저널", "주제", "핵심 내용"], [
  ["1", "2018 Acta Neuropathol", "NOS와 NEC", "진단 용어 NOS(검사 불가·불충분)와 NEC(검사했지만 기존 분류에 안 맞음)의 정의 [보충: 제목 외 내용]"],
  ["2", "2018 Acta Neuropathol", "DMG H3 K27M, IDH-mut astrocytoma 진단 명확화", "[보충: 초록 없음, 제목 수준]"],
  ["3", "2018 Acta Neuropathol", "IDH-wt, molecular features of GBM", "EGFR amp, +7/−10, TERTp → grade IV. Integrated grade 우선 원칙"],
  ["4", "2019 Acta Neuropathol", "MYB/MYBL1, FGFR1, BRAF V600E diffuse glioma", "[보충: 초록 없음] 주로 소아형 저등급 diffuse glioma"],
  ["5", "2020 Acta Neuropathol", "**IDH-mut astrocytoma grading·용어**", "**CDKN2A/B HD → grade 4**, 'GBM, IDH-mutant' 폐지 권고, 아라비아 숫자. PDGFRA amp은 \"추가 근거 시 grading 기준 가능\""],
  ["6", "2020 Brain Pathol", "Utrecht 회의: 새 entity·진단 원칙", "다수의 새 type·subtype을 향후 분류에 포함 권고, 분류·명명·grading 원칙"],
  ["7", "2020 Brain Pathol", "Ependymal tumor", "**부위별** 분리(PF, ST, spinal), ST: ZFTA(C11orf95)·YAP1 fusion, PF: PFA/PFB, spinal MYCN amp. Classic vs anaplastic 구분 폐지"],
  ["8", "2025 Neuro-Oncol", "Meningioma", "Brain invasion 판정 표준화, grade 1 형태 + **1p del + 22q del and/or NF2 → grade 2**"],
  ["9", "2025 Neuro-Oncol Adv", "DNA methylation profiling", "Classifier = 통합 진단의 한 도구, layered report 예시, backward compatibility"],
  ["10", "2025 Brain Pathol", "새 tumor type 정의 기준", "Clinical phenotype(영상 포함) 필수 + cluster·driver·조직, ≥2 독립 보고. 미충족 시 NEC"],
  ["11", "2025 Brain Pathol", "IDH/H3-wt HGG, PF ependymoma", "<40세 EGFR amp·+7/−10 단독 GBM 금지, TERTp-only 신중, PDGFRA 등은 <25세 pedHGG marker, EZHIP IHC"],
  ["**12**", "**2026 Neuro-Oncol**", "**IDH-mutant glioma grading**", "**PDGFRA amp → grade 4**, mitosis ≥3/2.4 mm² → grade 3, grade 3 분자 marker, oligo 개정 없음"],
], [0.5, 1.6, 2.4, 4.7], { highlight: (r) => r[0] === "**12**" || r[0] === "5" }));
C.push(SRC("PubMed metadata/abstract/full text — Update 1 (10.1007/s00401-018-1808-0), 2 (10.1007/s00401-018-1826-y), 3 (PMC6204285), 4 (10.1007/s00401-019-01987-0), 5 (PMC8443062), 6 (10.1111/bpa.12832), 7 (10.1111/bpa.12866), 8 (10.1093/neuonc/noae170), 9 (10.1093/noajnl/vdae228), 10 (PMC12488253), 11 (PMC12695693), 12 (10.1093/neuonc/noag108)"));

C.push(H2("4-3. Update 6 — Utrecht 회의 (2020)"));
C.push(...B([
  "Working Committee 3가 Utrecht에서 회의 → 향후 WHO 회의에 앞서 **새로 인정된 다수의 type·subtype을 포함하도록 권고**",
  "분류 범주, 분류 접근법, **명명(nomenclature)과 grading 원칙**을 승인 → CNS5 구조의 기초",
]));
C.push(SRC("Louis DN, Brain Pathol 2020;30:844 (abstract)"));

C.push(H2("4-4. Update 7 — Ependymal tumor (2020)"));
C.push(...B([
  "Methylome 자료에 근거해 **해부학적 부위(PF, supratentorial, spinal)별로 분리**",
  "Supratentorial: **C11orf95(현 ZFTA)** fusion, **YAP1** fusion / PF: **PFA, PFB**(methylation group) / spinal: **MYCN amp** 공격적 ependymoma",
  "Myxopapillary ependymoma(grade 2 권고), subependymoma는 조직 기반 type 유지",
  "Classic vs anaplastic ependymoma 구분 폐지. 분자 정의 ependymoma의 grade 근거는 아직 미성숙",
]));
C.push(SRC("Ellison DW, Brain Pathol 2020;30:863 (abstract)"));

C.push(H2("4-5. Update 8 — Meningioma (2024 온라인, 2025 지면)"));
C.push(...B([
  "배경: CNS5가 meningioma 최초의 분자 grading 도입 — **TERTp mutation, CDKN2A/B HD → grade 3**",
  "Update 8의 내용",
  "  **Brain invasion**의 형태학적 평가·해석 표준화",
  "  고급 분자 검사가 권장되는 상황 제시",
  "  **Grade 1 형태 + 1p deletion과 22q deletion 그리고/또는 NF2 oncogenic variant → grade 2** 제안",
  "  근거가 아직 부족한 영역은 권고 보류",
  "영상의학 연결: 형태상 grade 1 수막종의 grade가 분자 결과로 올라갈 수 있음 → 추적 간격·보조 RT 논의에 영향 [보충]",
]));
C.push(SRC("Sahm F, Neuro-Oncol 2025;27:319 (abstract)"));

C.push(H2("4-6. Update 9 — DNA methylation profiling (2025)"));
C.push(...B([
  "2018년 CNS tumor methylation classifier 이후 진단 도구로 확산 → 사용 원칙 정리",
  "Classifier는 **기존 진단 도구와 함께 완전히 통합해서** 쓰는 도구",
  "WHO·ICCR layered report에 methylation 결과를 넣는 예시",
  "향후 array의 **backward compatibility** 필요(축적된 자료 호환)",
  "Methylation class와 CNS WHO tumor type의 대응 관계 정리",
  "실무 포인트(Tralongo 2026): low-confidence 결과는 'methylation negative'가 아니라 **'not classifiable'**, 최고 점수 class가 진단적 진실은 아님, grade를 바꾸는 애매한 focal CNV(특히 CDKN2A/B HD)는 **다른 방법으로 확인**",
]));
C.push(SRC("Aldape K, Neuro-Oncol Adv 2025;7:vdae228 (abstract); Tralongo 2026 (PMC13498940)"));

C.push(H2("4-7. Update 10 — 새 tumor type의 정의 (2025)"));
C.push(...B([
  "필수: **clinical phenotype** — 전형적 예후/치료 반응, 호발 연령, 호발 위치, **전형적 영상 소견**",
  "가급적 충족: ① 대규모 profiling에서 고유 cluster(위치 의존 signature만으로는 불충분) ② driver alteration(하나 또는 조합) ③ 연관 조직·IHC 소견",
  "형식: **≥2개 독립 publication**(가급적 각 >10명) + cIMPACT-NOW 위원회 승인",
  "Field test: 13개 후보 + 3개 provisional + 2개 기존 type 시험",
  "  FET-CREB fusion-positive intracranial mesenchymal tumor는 type 기준 충족",
  "  DGONC, CRINET은 여전히 provisional",
  "  기준 미달 시 형태·추정 조직기원 + **NEC** + 분자 소견으로 layered 보고",
  "**Grade 철학 논쟁**(Update 10 원문): grade를 '현재 치료 하 결과'로 볼지 '자연 경과(natural history)'로 볼지 → 예: WNT-activated medulloblastoma는 장기 생존 >90%인데도 grade 4 유지",
  "  → \"치료 후 결과\" 기준이면 치료가 바뀔 때마다 grade가 바뀌는 문제, \"자연 경과\" 기준이면 실제 결과와 괴리 문제",
]));
C.push(SRC("Hawkins C, Brain Pathol 2025;35:e70018 (PMC12488253 전문)"));

C.push(H2("4-8. Update 11 — IDH/H3-wt HGG와 PF ependymal tumor (2025)"));
C.push(...B([
  "Working group \"Clarification\"이 CNS5를 검토해 두 주제를 선정",
  "**GBM, IDH-wt vs diffuse pediatric-type HGG, H3-wt and IDH-wt**",
  "  TERTp mutation **단독**의 조직학적 저등급 IDH-wt 종양에 grade 4(GBM) 부여는 신중히",
  "  **<40세**: EGFR amp 또는 +7/−10 **단독**으로 GBM 진단 금지",
  "  성인, 특히 <40세에서 pedHGG H3/IDH-wt 감별 고려, 확진은 methylation profiling",
  "  **PDGFRA alteration**, EGFR alteration, MYCN amp은 **<25세**에서만 pedHGG의 key molecular feature",
  "  >55세 non-midline 종양은 H3 K27 검사 생략 가능, ATRX 유지 + OLIG2 양성은 H3 G34-wt 지지",
  "**PF ependymal tumor**: nuclear EZHIP IHC → PFA 지지 / H3 K27me3 유지 + EZHIP 음성 + methylation 미시행 → **PF ependymoma, NOS** / CNS5에 없는 새 종양은 NEC 부가",
  "배경 맥락: TERTp-only molecular GBM 중 일부(HGG-F 등)는 예후가 더 좋을 수 있다는 우려",
  "**PDGFRA가 Update 11과 12에 모두 등장하지만 맥락이 다름** — 11은 IDH-wt 젊은 환자의 **진단** marker, 12는 IDH-mut astrocytoma의 **grading** marker",
]));
C.push(SRC("Wesseling P, Brain Pathol 2025;36:e70035 (PMC12695693 전문, abstract); Kurokawa 2026"));

// =============== 5. Update 12 ===============
C.push(H1("5. cIMPACT-NOW Update 12 상세"));
C.push(H2("5-1. 목적과 저자"));
C.push(...B([
  "Working group이 IDH-mutant glioma 문헌을 검토해 **pathology-based risk stratification·grading 개선** 기회를 평가",
  "특히 **grade 2와 3의 경계** — \"치료 결정이 risk assessment에 의존하기 때문\"",
  "새로 기술된 IDH-mutant glioma subgroup 평가, 치료 결정의 **다학제적 측면** 고려",
  "저자: Brat(1저자, Update 3·5도 1저자), Aldape, French, Louis, Nasrallah, Reifenberger, Reuss, Touat, van den Bent(CATNON PI), Wesseling, Figarella-Branger — 신경병리 + 신경종양",
]));
C.push(SRC("Brat 2026 (abstract, author list)"));

C.push(H2("5-2. Grading 알고리즘 (Park YW, KJR 2026, Fig. 2)"));
C.push(...T(["Tumor", "Grade", "기준"], [
  ["Astrocytoma, IDH-mut", "**4**", "MVP, necrosis, **CDKN2A/B HD** 그리고/또는 **PDGFRA amp** (하나라도 있으면)"],
  ["Astrocytoma, IDH-mut", "**3**", "Grade 4 기준 없음 + **mitotic activity ≥3 / 2.4 mm²**"],
  ["", "(3 시사)", "CDK4/CDK6/CCND2 amp, CDKN2A mutation, RB1 HD or mutation · PIK3CA/PIK3R1 mutation · EGFR alteration, MYCN amp · methylation **A_IDH_HG** 또는 **G-CIMP-low** (Fig.에는 IDH_A_HG로 표기)"],
  ["Astrocytoma, IDH-mut", "**2**", "Grade 4 기준 없음 + mitosis <3 / 2.4 mm² (+ grade 3 시사 분자 변화 없음)"],
  ["Oligodendroglioma", "**3**", "MVP, necrosis 그리고/또는 **mitotic activity ≥6 / 2.4 mm²**"],
  ["Oligodendroglioma", "**2**", "위 기준 없음"],
], [2.2, 1.0, 6.0], { highlight: (r) => r[1] === "**4**" }));
C.push(...B([
  "Fig. 2는 grade 3 분자 변화를 \"that **may** suggest WHO grade 3\"로 표현 → 제안 수준 [원문 확인 권장: 정확한 적용 문구]",
  "Methylation(G-CIMP-low, A_IDH_HG): \"다른 grading 기준에 따라 **grade 3 또는 4**와 부합\"",
  "Mitosis 기준: \"**could be considered**\"",
  "Tralongo 2026 해석: Update 12의 권고는 **해석 지침**이며 모든 사건을 독립적인 WHO 기준으로 만드는 것은 아님. PDGFRA amp은 'grade 4-like', RB/PI3K/receptor 변화는 'intermediate risk'",
]));
C.push(SRC("Brat 2026 (abstract); Park YW, KJR 2026 Fig. 2; Tralongo 2026 (PMC13498940)"));

C.push(H2("5-3. Grade 4: PDGFRA amplification"));
C.push(...B([
  "초록 원문: \"PDGFRA amplification was associated with **highly aggressive clinical behavior** of IDH-mutant astrocytomas, **consistent with CNS WHO grade 4**.\"",
  "→ MVP, necrosis, CDKN2A/B HD와 **같은 grade 4 기준**으로 알고리즘에 배치",
  "조직학적 grade(2 또는 3)와 무관하게 적용",
  "Update 5(2020)에서 \"추가 근거가 쌓이면 grading 기준이 될 수 있다\"고 보류했던 marker를 6년 뒤 채택",
]));
C.push(SRC("Brat 2026 (abstract); Update 5 (PMC8443062)"));

C.push(H2("5-4. Grade 3: mitosis 기준과 분자 marker"));
C.push(...B([
  "**Mitotic count ≥3 / 2.4 mm²** 제안",
  "  2.4 mm²는 고배율 시야 약 10개에 해당하는 면적 [보충 — 현미경 시야 직경에 따라 다름]",
  "  이전: 검증된 cut-off 없음, 관행상 전체 검체 ≥2 mitoses",
  "  PDGFRA IHC 논문: \"grade 2/3 cut-off가 낮다(2.4 mm²당 3개)\" + 생검이 **조영증강 밖**에서 얻어지거나 세포 밀도가 낮으면 grading이 어려움",
  "중간 위험(grade 3에 가장 부합) 분자 변화 — **PI3K**(PIK3CA·PIK3R1 mut), **EGFR** alteration, **MYCN** amp, **RB pathway**(CDK4·CDK6·CCND2 amp, RB1 HD/mut, CDKN2A **mutation**)",
  "CDKN2A: **HD = grade 4**, **point mutation = grade 3 시사** — 혼동 주의",
  "Methylation G-CIMP-low / A_IDH_HG: 다른 기준에 따라 grade 3 또는 4",
]));
C.push(SRC("Brat 2026 (abstract); Tauziède-Espariat, Brain Pathol 2026 (PMC13535536); Update 5"));

C.push(H2("5-5. Oligodendroglioma"));
C.push(...B([
  "Grading을 개선할 새 feature를 **찾지 못함**",
  "단, mitosis가 높아 grade 3이지만 **necrosis, MVP, CDKN2A/B HD, MRI 조영증강이 없는** 종양은 생존이 길 수 있음",
  "  → 이 문장은 oligodendroglioma grade 3에서 CDKN2A/B HD와 조영증강을 **위험 인자로 함께 고려**한다는 의미로 읽힘",
  "  → 영상의학적 의미: **조영증강 유무가 예후 정보**로 언급된 드문 사례",
  "Fig. 2 기준: MVP, necrosis 그리고/또는 mitosis ≥6 / 2.4 mm² → grade 3",
  "Tralongo 2026: Update 12는 oligodendroglioma grading을 더 정교하게 할 만큼 검증된 genomic·epigenomic 기준을 찾지 못함, 근거는 astrocytoma에서 가장 강하며 oligodendroglioma로 그대로 옮기면 안 됨",
]));
C.push(SRC("Brat 2026 (abstract); Tralongo 2026 (PMC13498940)"));

C.push(H2("5-6. Update 5 → Update 12 비교"));
C.push(...T(["항목", "Update 5 / WHO CNS5", "Update 12 (2026)"], [
  ["Grade 4 분자 기준", "CDKN2A/B HD", "CDKN2A/B HD + **PDGFRA amp**"],
  ["Grade 2 vs 3", "\"significant\" mitotic activity (수치 없음)", "**≥3 mitoses / 2.4 mm²** 제안"],
  ["Grade 3 분자 marker", "없음 (근거 부족으로 보류)", "PI3K, EGFR, MYCN, RB pathway, G-CIMP-low / A_IDH_HG"],
  ["PDGFRA amp", "\"추가 근거 시 grading 기준 가능\"", "**Grade 4 기준**"],
  ["G-CIMP-low", "\"검증 cohort 필요\"", "Grade 3–4 시사 (단독 grade 4 아님)"],
  ["Oligodendroglioma", "MVP·necrosis·mitosis·CDKN2A/B HD", "개정 근거 없음, mitosis ≥6 / 2.4 mm²"],
  ["필요한 검사", "CDKN2A/B 검출(FISH, MLPA, NGS 등)", "**NGS + CNV(methylation array 등)** 사실상 필요"],
], [2.0, 3.4, 3.8]));
C.push(SRC("Update 5 (PMC8443062); Brat 2026 (abstract); Tauziède-Espariat, Ann Pathol 2026 (abstract)"));

// =============== 6. PDGFRA ===============
C.push(H1("6. PDGFRA amplification 심화"));
C.push(H2("6-1. 기본 개념과 다른 종양에서의 의미"));
C.push(...B([
  "Platelet-derived growth factor receptor alpha — **receptor tyrosine kinase(RTK)** [보충]",
  "유전자 위치: 4번 염색체(FISH에서 chromosome 4 centromere probe와 비교), [보충] 4q12",
  "IDH-mut astrocytoma: **grading marker**(Update 12)",
  "pedHGG H3/IDH-wt, 특히 **RTK1 subgroup**: 흔하고 **진단적** 의미(Update 11, <25세)",
  "GBM IDH-wt: 관찰됨(IHC 코호트 11/45)",
  "Diffuse hemispheric glioma, H3 G34-mutant: amplification 보고",
  "Myxoid glioneuronal tumor: **PDGFRA mutation**이 진단 marker",
  "IDH-mutant primary GBM(Wong 2021, 2016 분류): PDGFRA amp 또는 mutation **27.1%(16/59)**, G-CIMP-low와 연관",
]));
C.push(SRC("Tauziède-Espariat, Brain Pathol 2026 (PMC13535536); Le Rhun, Neuro-Oncol 2025 (abstract); Wong, Mod Pathol 2021 (10.1038/s41379-021-00778-x, abstract); Update 11"));

C.push(H2("6-2. 근거 연구 정리"));
C.push(...T(["연구", "설계·대상", "주요 결과", "약점"], [
  ["**Yang RR** Brain Pathol 2019", "후향, IDH-mut grade II/III 160명, FISH", "PDGFRA amp **18.8%**, CDKN2A HD 15.0%, CDK4 amp 18.1%, 상호 배타. PDGFRA amp → PFS p=0.0003, OS p<0.0001. Mean OS 40.5 / 62.9 / 71.5개월", "단일 기관, 추적 짧음(저자 인정), mean OS 사용"],
  ["**Tesileanu** Neuro-Oncol 2021 (CATNON)", "RCT 내 분자 분석, IDH-mut anaplastic astrocytoma 432명", "다변량 독립 예후인자: 나이, MMSE, TMZ, classifier, **PDGFRA amp**, CDKN2A/B HD, PI3K mut, CNV load. RPA 첫 분기 = CDKN2A HD", "Grade 3 조직만 포함, 사후 분석"],
  ["**van den Bent** Lancet Oncol 2026 (CATNON final)", "IDH-mut 444명, 추적 10.9년", "PDGFRA·CDK4 amp, CDKN2A HD, CNV load → 예후 불량. **TMZ 효과 predictive 아님**", "탐색적 분석"],
  ["**Virata** Neuro-Oncol 2026", "후향, GENIE 840 + cBioPortal 367명", "PDGFRA alteration 6.5%. Cohort 1 median OS 35.5개월(HR 4.4). **다변량 독립 유의성 미달**, 단독 alteration HR 4.5. 'Intermediate' 군에 배치", "Amp + mut 합산, 치료 변수 없음, 소수"],
  ["**Richardson** Acta Neuropathol 2026", "후향 validation, 1,140명", "PDGFRA amp으로 2→4 **7례**, 3→4 **15례**. Median OS **51.4 vs 126.2개월**(HR 1.9, p=0.0284)", "Letter, mitosis 직접 확인 불가, methylation 260례만"],
  ["**Tauziède-Espariat** Ann Pathol 2026", "실제 진료, 50례", "**24%** grade 변경(2→3 7, 2→4 2, **3→4 3**), CNV 기반 8례", "단일 기관, 생존 자료 없음"],
], [1.9, 2.0, 3.6, 1.7]));
C.push(SRC("Yang 2019 (10.1111/bpa.12801); Tesileanu 2021 (PMC8408862); van den Bent 2026 (10.1016/S1470-2045(25)00614-X); Virata 2026 (PMC13437890); Richardson 2026 (PMC13391642); Tauziède-Espariat 2026 (10.1016/j.annpat.2026.09.001)"));

C.push(H2("6-3. Validation (Richardson 2026) 상세"));
C.push(...B([
  "코호트: AACR GENIE + TCGA + MSKCC, IDH-mut astrocytoma **1,140명**",
  "2016 → 2021 WHO: CDKN2A/B HD로 grade 2의 7.9%(40례), grade 3의 12.2%(45례)가 grade 4로",
  "2021 WHO → Update 12",
  "  PDGFRA amp: grade 2→4 **7례**(grade 2의 1.5%), grade 3→4 **15례**(grade 3의 4.6%)",
  "  Grade 3 분자 기준: grade 2→3 **77례**(grade 2의 16.6%)",
  "  2021 기준 grade 2/3 중 **12.5%(99/789)** upgrade",
  "생존(median OS): grade 2 유지 146.0 / 2→3 upgrade 109.6(HR 2.6) / **PDGFRA amp grade 4 51.4 vs grade 3 유지 126.2(HR 1.9)**",
  "**Upgrade 군과 원래 같은 grade 군 사이 차이 없음**(grade 4: p=0.1277) → 재분류가 타당함을 지지",
  "Update 12 기준 전체: grade 2 146.0 / grade 3 120.4 / grade 4 **37.7개월**",
  "한계: methylation은 TCGA 260례만, 여러 코호트·시대 혼합, mitosis 밀도 직접 확인 불가(일부 grade 3가 grade 2로 내려갈 수 있음)",
]));
C.push(SRC("Richardson 2026 (PMC13391642 전문)"));

C.push(H2("6-4. Virata 2026 상세 — 반론의 핵심 자료"));
C.push(...B([
  "Cohort 1(GENIE 840) reference(grade 2/3, 변화 없음) median OS **140.7개월**",
  "Univariate median OS(Cohort 1, 유전자명은 순서 대응 [원문 확인 권장]): CDKN2A HD 25.4, CDK4 amp 29.1, CCND2 amp 29.1, **PDGFRA alteration 35.5(HR 4.4)**, PIK3R1 32.7, PIK3CA 48.0, MYCN 30.2, EGFR 29.8, TERTp 59.0, grade 4 histology 35.2개월",
  "다변량(Cohort 1+2, 나이·성별·코호트 보정): grade 4 histology HR 2.24, CDKN2A HD 2.82, CDK4 1.91, CCND2 1.45, PIK3R1 1.74, MYCN 1.59, EGFR 1.59 유의 / **PDGFRA alteration, PIK3CA, TERTp는 독립 유의성 없음**",
  "그러나 **단독(isolated) PDGFRA alteration → HR 4.5**로 reference보다 나쁨 → 저자들: \"alteration alone is sufficient for poor outcome\"",
  "6개 marker(CDK4, CCND2, PDGFRA, PIK3R1, MYCN, EGFR) 있는 grade 2/3 → median OS **67.3–82.0개월**(reference 135.0–140.7, grade 4 35.3–45.0) → '**intermediate grade**'",
  "**단독 marker 각각과 CDKN2A HD 단독 사이 OS 차이 없음** → 저자들도 저/고 2-tier grading 가능성 언급",
  "Grade 4 내에서: CDKN2A HD만 vs 조직학적 grade 4만 → OS 차이 없음 / **둘 다 있으면 더 나쁨**",
  "한계: 대부분 marker가 <5%로 소수, 절제 범위·RT·항암 변수 없음, 생존 측정·검사 방법이 코호트마다 다름",
]));
C.push(SRC("Virata MC et al. Neuro-Oncol 2026 (PMC13437890 전문)"));

C.push(H2("6-5. 검출 방법"));
C.push(...T(["방법", "Amplification 정의 / 특징", "비고"], [
  ["FISH", "PDGFRA locus **>8 copies in ≥10% 종양세포**", "**Clonal(부분) amplification** 검출에 가장 유리"],
  ["Methylation array CNV", "log ratio >0.8 (IHC 논문) / ±0.60 log2 (CATNON)", "Methylation class와 동시 획득. Clonal amp 10례 중 **4례만** CNV에서 보임. 순도 낮으면 진폭 감소"],
  ["NGS", "Panel 기반 CNV call", "Grade 3 marker(PI3K 등 mutation)와 동시 평가"],
  ["SNP / CGH array", "Copy-number 분석", "Ann Pathol 2026에서 grade 3 검체에 권장 기법으로 언급"],
  ["**PDGFRA IHC**", ["Method 1: ≥50% 강양성 → 민감도 **54%**, 특이도 **99%**", "Method 2: **≥10% 강양성** → 민감도 **100%**, 특이도 **89%**"], "**Screening**. Gain(≤6 copies)·mutation에서도 양성 가능 → FISH 등으로 확진. Mutation 검출 불가"],
], [1.8, 4.2, 3.2]));
C.push(...B([
  "IHC 논문 제안 알고리즘: **>10% 강양성 → amplification 강력 의심**, 국소(<10%) 강양성 → FISH 등 필수",
  "IHC 코호트: HGG 114례 중 PDGFRA amp 27례 — pedRTK1A 9/14, **high-grade AST IDH-mut 5/10**, GBM IDH-wt 11/45",
  "Clonal amplification에서는 IHC 양성이 증폭된 부분에만 국한 → **종양 내 이질성**을 시각적으로 보여줌",
  "Tralongo 2026: 순도가 낮으면 focal event는 arm-level event보다 신뢰도 낮음, grade를 바꾸는 애매한 결과는 **orthogonal 확인**",
]));
C.push(SRC("Tauziède-Espariat, Brain Pathol 2026;e70142 (PMC13535536 전문); Tesileanu 2021 (PMC8408862); Tralongo 2026 (PMC13498940)"));

C.push(H2("6-6. 쟁점 정리"));
C.push(...T(["쟁점", "Grade 4 지지", "반대·유보"], [
  ["Grade 4인가 intermediate인가", "Update 12, Richardson 2026(재분류 군 OS = 원래 grade 4, p=0.1277), Yang 2019(최고위험군)", "Virata 2026(intermediate, 다변량 비유의) — 단, 단독 alteration은 CDKN2A HD 단독과 OS 차이 없음"],
  ["독립 예후인자인가", "CATNON 다변량 독립 인자(grade 3 조직 432명)", "Virata 다변량 비유의(amp+mut 합산, 공동 변이 많음)"],
  ["예측(predictive) marker인가", "—", "CATNON final: TMZ 효과 예측 못함"],
  ["Clonal amp도 grade 4인가", "—", "**기준 없음**(필자가 찾은 문헌 범위). FISH ≥10% 정의가 실무 기준"],
  ["WHO 공식 기준인가", "—", "**아님**. cIMPACT-NOW 권고, CNS6에서 결정"],
], [2.2, 3.6, 3.4]));

// =============== 7. Molecular grading pitfalls ===============
C.push(H1("7. 분자 grading의 함정과 보고서 원칙"));
C.push(...B([
  "**CDKN2A/B HD vs hemizygous loss**: HD만 grade 4 기준. 단, 재발 IDH-mut astrocytoma 32명 연구(MLPA, HD <0.4 / hemizygous <0.7)에서 **hemizygous loss도** 재발 후 생존과 독립적으로 연관 → HD 정의가 표준화돼 있지 않다는 문제 제기",
  "**CDKN2A/B HD가 항상 예후를 가르는가**: IDH-mut primary GBM 67예(Wong 2021)에서는 CDKN2A/B HD가 OS·PFS와 **무관**(p=0.197) — 코호트에 따라 결과가 다름",
  "**Methylation high-grade 단독은 grade 4 아님**: adverse feature로 기술",
  "**Classifier 결과 보고**: platform·classifier version·confidence score·CNV 소견·형태와의 일치 여부·한계를 명시. Low confidence는 'not classifiable'",
  "**Sampling bias**: 단일 block이 모든 subclone을 대표하지 못함. Macrodissection은 순도를 높이지만 한 부위만 과대 선택 가능",
  "**순도(purity)**: 비종양 조직 혼입 → classifier confidence·CNV 진폭 감소",
  "**치료 효과**: 재발 검체의 methylation 변화가 치료 영향인지 subclone 확장인지 구분 어려움",
  "Discordance(형태 vs 분자) → 조직 질, sampling, 치료 효과, 형태, 유전 결과, reference class 대표성을 **다시 평가**",
]));
C.push(SRC("Yamashita, Brain Tumor Pathol 2026 (10.1007/s10014-026-00541-7, abstract); Wong 2021 (abstract); Tralongo 2026 (PMC13498940)"));

// =============== 8. Imaging ===============
C.push(H1("8. 영상 소견"));
C.push(H2("8-1. 영상이 여전히 중요한 이유"));
C.push(...B([
  "분자 정보가 중요해졌다고 영상의 역할이 줄어든 것이 아님 — 세 type은 예후, 수술 전략, 수술 후 CCRT, 표적치료 적격성이 모두 다름",
  "영상의 역할: 생검 회피가 아니라 **조직 획득 최적화**",
  "  가장 aggressive한 부위를 **biopsy target**으로: focal enhancement, **diffusion restriction, perfusion 상승**, 비조영 CT 고음영, 대사 이상",
  "  절제 예정이어도 high-grade 예상 부위를 **따로 표시해 제출**(spatially annotated sampling)",
  "  분자 subtype 예측 → 임상시험 screening, 예후 상담",
  "INDIGO 이후 수술 전 IDH 상태 예측의 중요성 증가",
  "Update 10: **전형적 영상 소견이 새 tumor type 정의의 구성 요소**",
]));
C.push(SRC("Kurokawa 2026 (PMC13526116) Introduction, cIMPACT-NOW section"));

C.push(H2("8-2. IDH-mutant astrocytoma vs oligodendroglioma"));
C.push(...T(["소견", "시사", "성능", "주의점"], [
  ["Classic T2-FLAIR mismatch", "AST", "Pooled 민감도 **0.42**, 특이도 **0.99**", "낭성 변화·괴사는 양성으로 판정 금지"],
  ["Partial mismatch (grade 4 astrocytic tumor)", "AST (IDH-mut)", "IDH-mut 26.4%(32/121) vs IDH-wt 0.4%(8/2044) → 특이도 99.6%", "Near-complete가 아니어도 의미 있음"],
  ["Redefined mismatch (partial 포함, subcortical-only 제외)", "AST", "AST vs ODG 민감도 87.5%, 특이도 80.0%", "단일 기관"],
  ["Coarse calcification, cortical-subcortical, frontal", "ODG", "석회화 ODG 50.0% vs AST 0%(한 코호트)", "CT 중요"],
  ["ASL cortical high-flow sign", "ODG", "민감도 54.5%, 특이도 92.9%", "피질 혈관 밀도와 상관"],
  ["Subcortical FLAIR signal drop", "ODG", "민감도 44.0%, 특이도 93.8%", "경미하면 partial mismatch와 혼동"],
  ["ODG 소견 ≥2개", "ODG", "민감도 48.0%, 특이도 **100%**", "—"],
  ["MRS 2HG / cystathionine", "IDH-mut / ODG", "Cystathionine AUC 0.69(ODG vs AST)", "보조 정보"],
], [2.6, 1.2, 3.0, 2.4]));
C.push(...B([
  "Mismatch의 병리 배경: microcystic change, myxoid matrix",
  "Mismatch가 보이는 다른 종양: DMG H3 K27-altered, MYB/MYBL1-altered diffuse astrocytoma, angiocentric glioma, MAPK-altered LGG, PA, DNET, 탈수초 질환 → **adult-type diffuse glioma 맥락에서만** AST 시사",
]));
C.push(SRC("Kurokawa 2026 (PMC13526116) Table 1"));

C.push(H2("8-3. Grade·고위험 분자 변화와 영상"));
C.push(...T(["연구", "대상", "결과"], [
  ["Jo, Neurooncol Pract 2025", "CNS5 재분류 736명 (A2 69, A3 37, **A4 32**, O2 60, O3 23, GBM 515)", "조영증강: IDH-mut 전체 48% vs GBM 93%, **A4 96%**(GBM과 차이 없음). Min ADC는 IDH-mut > GBM. A4 평균 39.8세(GBM 63.1세)"],
  ["Ozono, J Neurooncol 2026", "Non-enhancing AST 38명", "Mismatch(>50%) 55.3%, (>33%) 73.7%. 양성은 낮은 grade·CDKN2A HD 없음과 연관. **CDKN2A HD 5례 모두 mismatch 음성**. HD 제외 시 작은 종양에서 mismatch 음성(250.7 vs 1842.9 mm²)"],
  ["Alfattal, Int J Surg Pathol 2026", "≥55세 IDH-mut AST 22명", "Mismatch: grade 2의 25%, grade 3·4 0%. Grade 2의 **50% 조영증강**. Grade 분포 2:3:4 = 55:18:27%"],
  ["Yang, AJNR 2025", "IDH-mut AST 122 + 외부 18명", "**Poorly defined enhancing margin** → CDKN2A HD 독립 예측(OR 13.19), AUC 0.84 / 외부 0.96"],
  ["Singh, Neuroradiology 2025", "IDH-mut AST 72명", "A_IDH_HG 예측: midline crossing, ependymal extension, poorly defined enhancement 등, 정확도 75%(AUC 0.77). CDKN2A/B 상태와는 부피 차이 없음"],
  ["Imai, J Clin Med 2025", "Non-enhancing glioma 106명", "rCBV–MET PET 상관(r=0.5). rCBV는 AST에서 가장 낮음, **grade 4 구분 AUC 0.718**"],
], [2.1, 2.6, 4.5]));
C.push(SRC("Jo 2025 (10.1093/nop/npaf067); Ozono 2026 (PMC13002771); Alfattal 2026 (10.1177/10668969261489616); Yang 2025 (10.3174/ajnr.A8861); Singh 2025 (10.1007/s00234-025-03766-z); Imai 2025 (10.3390/jcm14196777) — PubMed abstract/전문"));

C.push(H2("8-4. 감별: Molecular GBM (63세 환자에서 특히 중요)"));
C.push(...B([
  "고령 환자의 비조영·침윤성 종양은 **IDH-wt molecular GBM**도 감별해야 함 (≥55세는 IDH-wt가 많음 — Alfattal 2026 서론)",
  "Molecular GBM 영상: 경계 불명확한 침윤성, 조영증강 없음 또는 미약, 대부분 diffusion restriction 없고 rCBV 낮음",
  "  그러나 **국소 고음영(CT), diffusion restriction, rCBV 상승** 부위가 공존할 수 있음 → targeted sampling",
  "  AST·ODG를 지지하는 소견이 **없음**이 단서",
  "  Gliomatosis cerebri pattern 30%, 다발성 9%",
  "  시간이 지나며 조직학적 grade가 올라가는 경우 있음, 생검에서 grade 4 성분이 빠질 수 있음",
]));
C.push(SRC("Kurokawa 2026 (PMC13526116) Molecular GBM; Alfattal 2026 (abstract)"));

C.push(H2("8-5. PDGFRA amplification과 영상"));
C.push(...B([
  "**IDH-mut astrocytoma에서 PDGFRA amp에 특이적인 영상 소견은 찾지 못함**(PubMed 검색 기준)",
  "Perfusion과 여러 유전자 marker(PDGFRA 포함)의 연관을 본 연구는 있으나 1.5T, 67명, tumor type 혼합 → 일반화 어려움",
  "영상의 역할은 PDGFRA amp을 맞히는 것이 아니라 **조직 grade보다 aggressive한 부위를 찾아 sampling·분자 검사·재판독을 촉발**하는 것",
  "Radiogenomics 연구 주제로 남아 있음 [제안]",
]));
C.push(SRC("Śledzińska-Bebyn, Magn Reson Imaging 2024 (10.1016/j.mri.2024.110318, abstract); PubMed 검색"));

C.push(H2("8-6. 판독 체크리스트"));
C.push(...BOX("[제안] IDH-mutant (추정) astrocytoma 판독 시", [
  "T2-FLAIR mismatch(완전/부분) — 있으면 AST 시사, **없거나 추적 중 소실되면** 고위험 분자 변화 가능성",
  "조영증강: 유무, **경계(poorly defined margin)**, 새 증강",
  "DWI/ADC: 최저 ADC 부위 위치·값",
  "Perfusion(DSC/DCE/ASL): rCBV 최고 부위 위치·값",
  "Midline crossing, ependymal extension",
  "고령이면 molecular GBM 감별 언급",
  "**조직 grade와 영상 불일치** 시 보고서·MDT에서 **CNV(PDGFRA, CDKN2A/B)와 NGS 확인 요청**",
  "Biopsy/절제 계획 시 aggressive focus 위치를 명시(spatially annotated sampling)",
  "추적: INDIGO 대상(grade 2, 비조영 잔존)인지, 조영증강 새로 생겼는지 구분",
], GREEN));

// =============== 9. Treatment ===============
C.push(H1("9. 치료 — 왜 grade가 중요한가"));
C.push(H2("9-1. CATNON (EORTC 26053-22054)"));
C.push(...B([
  "대상: 새로 진단된 **1p/19q non-codeleted anaplastic glioma**, ≥18세, WHO PS 0–2, 137개 기관",
  "4군 무작위(1:1:1:1, 751명): RT 단독(59.4 Gy/33 fx) / RT + **concurrent** TMZ(75 mg/m²/일) / RT + **adjuvant** TMZ(12 cycles, 150–200 mg/m² day 1–5) / 둘 다",
  "2nd interim(2021, 추적 55.7개월)",
  "  Concurrent: median OS 66.9 vs 60.4개월, HR 0.97 → **futility**",
  "  Adjuvant: 82.3 vs 46.9개월, **HR 0.64** (p<0.0001)",
  "  Grade 3–4 독성: 혈액학적, RT 단독 0, concurrent 9%, adjuvant 포함군 15%",
  "Final(2026, 추적 10.9년)",
  "  ITT: adjuvant HR 0.65 / concurrent HR 0.91(비유의)",
  "  **IDH-mut 444명**: adjuvant TMZ median OS **12.5년 vs 6.0년(HR 0.54)** / concurrent 9.7 vs 7.2년(HR 0.81, 95% CI 0.63–1.04, 비유의)",
  "  IDH-wt: TMZ 이득 없음",
  "  PDGFRA·CDK4 amp, CDKN2A HD, CNV load → 예후 불량, **TMZ 효과 predictive 아님**",
  "결론: 공격적 IDH-mut astrocytoma에서 **RT → adjuvant TMZ 12 cycles(concurrent 없이)**",
]));
C.push(SRC("van den Bent, Lancet Oncol 2021;22:813 (10.1016/S1470-2045(21)00090-5); van den Bent, Lancet Oncol 2026;27:45 (10.1016/S1470-2045(25)00614-X) — PubMed abstracts"));

C.push(H2("9-2. INDIGO (vorasidenib)"));
C.push(...B([
  "대상: 수술 외 치료가 없는 **잔존/재발 grade 2 IDH-mutant glioma**, 331명, 이중맹검 phase 3",
  "  참가자는 **비조영 또는 최소 조영** 병변, 즉시 RT·항암 대신 관찰이 적절한 환자",
  "PFS **27.7 vs 11.1개월** (HR 0.39), 다음 치료까지 기간 HR 0.26, grade ≥3 ALT 상승 9.6%",
  "허가: **FDA 2024년 8월, 국내 식약처 2026년 1월**",
  "한계: grade 3–4, 빠르게 진행, 뚜렷한 조영증강, 이전 치료 많은 환자로 **일반화 불가**. RT·항암과 직접 비교 안 함. **OS 이득은 아직 확립되지 않음**",
  "→ Upgrade로 grade 3/4가 되면 vorasidenib 'active delay' 전략의 대상에서 벗어남",
]));
C.push(SRC("Mellinghoff, NEJM 2023;389:589 (10.1056/NEJMoa2304194); Park YW, KJR 2026 (in press) p.1–2; Tralongo 2026 (PMC13498940)"));

C.push(H2("9-3. 가이드라인"));
C.push(...T(["Tumor", "ASCO-SNO 2022 권고 (초록)"], [
  ["AST IDH-mut grade 2", "RT + adjuvant chemotherapy (TMZ 또는 PCV)"],
  ["AST IDH-mut grade 3", "RT + adjuvant TMZ"],
  ["**AST IDH-mut grade 4**", "**Grade 3 권고 또는 GBM IDH-wt 권고 중 어느 쪽을 따라도 됨**"],
  ["GBM IDH-wt grade 4", "Concurrent TMZ + RT → adjuvant TMZ 6개월, TTFields 고려, bevacizumab 비권고"],
  ["ODG grade 2/3", "RT + PCV (TMZ는 대안)"],
], [3.0, 6.2], { highlight: (r) => r[0].startsWith("**AST") }));
C.push(...B([
  "**즉, IDH-mut grade 4의 치료는 무작위 근거가 따로 없고**, 가이드라인이 두 선택지를 모두 허용 → CCRT(GBM-type)는 허용된 선택지 중 하나",
  "EANO 2021 가이드라인도 존재(초록만 확인, 세부 권고는 [원문 확인 권장])",
  "Vorasidenib 시대의 grade 2 관리: 최대 안전 절제 → 저위험은 관찰, 잔존 비조영 안정 병변은 vorasidenib, 고위험은 RT + 항암 [보충]",
]));
C.push(SRC("Mohile NA, J Clin Oncol 2022;40:403 (10.1200/JCO.21.02036, abstract); Weller M, Nat Rev Clin Oncol 2021;18:170 (abstract); Wu, Curr Treat Options Oncol 2026 (abstract)"));

C.push(H2("9-4. Grade 상향의 임상적 의미"));
C.push(...B([
  "치료 강도·시점, **임상시험 적격성**(grade별 eligibility), 예후 상담",
  "Virata 2026: 확장된 분자 기준이 \"systemic underestimation of true tumor grade\"를 막고, 저등급 조직이지만 고위험 분자 변화를 가진 환자에서 **치료 escalation** 고려 가능",
  "Update 3의 논리: molecular grade IV 지정은 \"combined chemo- and radiotherapy 권고를 뒷받침\", 임상시험 포함 범위 확대",
  "**그러나** 예후 marker가 곧 치료 반응 marker는 아님(CATNON final) — 치료 강화가 생존을 늘린다는 직접 근거는 없음",
]));
C.push(SRC("Virata 2026 Discussion; Update 3 (PMC6204285); van den Bent 2026"));

// =============== 10. Case ===============
C.push(H1("10. 증례 적용 — 63/F, Astrocytoma, IDH-mutant"));
C.push(...T(["단계", "내용", "근거·논리"], [
  ["① 초기 병리", "Astrocytoma, IDH-mutant, **grade 3**(조직)", "Mitosis 기반"],
  ["② 영상", "**ADC·rCBV에서만** aggressive", "Imai 2025, Kurokawa 2026"],
  ["③ 피드백", "영상의학과 → 병리과", "Spatially annotated sampling 원칙"],
  ["④ 분자", "**PDGFRA amplification**", "Update 12"],
  ["⑤ 통합 진단", "**Grade 4**", "Update 12 알고리즘"],
  ["⑥ 치료", "**CCRT**", "ASCO-SNO: grade 4 IDH-mut은 GBM 권고도 선택 가능"],
], [1.6, 3.6, 4.0]));
C.push(...B([
  "**63세의 의미**: IDH-mut AST 평균 진단 나이 37.8세, ≥55세 8.6%(Virata 2026) / A4 평균 39.8세(Jo 2025) → 드문 고령. ≥55세는 IDH-wt가 많음(Alfattal 2026)",
  "**조영증강이 뚜렷하지 않았다면**: Jo 2025의 A4는 96% 조영증강이지만 대부분 조직학적 grade 4로 보임 [보충]. 분자적으로만 grade 4인 종양은 조영증강·괴사가 없을 수 있음(Ozono 2026 non-enhancing cohort에도 grade 4 존재)",
  "발표 전 차트에서 확인할 것 [원문 확인 권장]",
  "  IDH 확인 방법(IDH1 R132H IHC vs sequencing), ATRX·p53, 1p/19q",
  "  PDGFRA 검출 방법(FISH/NGS/methylation), copy 수, clonal 여부",
  "  CDKN2A/B 상태, MGMT promoter, methylation class(시행 시)",
  "  조영증강 양상, 최저 ADC 값, rCBV 값, 생검/절제 부위와 영상 이상 부위의 일치 여부",
  "  절제 범위, CCRT regimen(RT 선량, TMZ 일정), 이후 경과",
]));
C.push(SRC("Virata 2026 Table 1; Jo 2025; Alfattal 2026; Ozono 2026; Mohile 2022"));

// =============== 11. Aggressive Q&A ===============
C.push(H1("11. 공격적 질문 대비 Q&A"));
C.push(P("각 질문은 **짧은 답 → 근거 → 인정할 한계** 순서. 질문자가 반박하면 '인정할 한계'를 먼저 말하고 우리 논리로 돌아오는 것이 안전하다."));

C.push(H2("11-A. 개념·분류"));
C.push(...QA(1, "cIMPACT-NOW는 WHO가 아닌데, 공식 진단을 바꿔도 되나?",
  "공식 WHO 기준은 아니다. 다만 CNS5의 많은 내용(CDKN2A/B HD → grade 4 등)이 cIMPACT-NOW에서 먼저 나왔고, 진료에서는 WHO 개정 전 진단 가이드로 쓰인다.",
  ["이름 자체가 'Not Officially WHO'", "Update 11: 권고는 WHO 공식 변경이 아니라 CNS6 이전 가이드", "Update 1–7 대부분이 CNS5에 반영(Kurokawa 2026)", "Update 12 저자 중 Louis, Brat, Wesseling, Reifenberger, Figarella-Branger는 WHO CNS5 요약 논문(Louis 2021)의 공저자"],
  "기관마다 적용 여부가 다를 수 있다. 보고서에 'cIMPACT-NOW update 12에 따른 grade'라고 명시하는 것이 안전하다 [제안]."));
C.push(...QA(2, "왜 조직학적 grade보다 분자 grade를 우선하나?",
  "Grade의 목적은 임상 경과 예측이고, 분자 grade가 생존을 더 잘 예측한다는 자료가 있기 때문이다.",
  ["Update 3 원문: \"integrated histological and molecular grade should take precedence over a strictly histological grade\"", "조직 소견은 종양 안에서 공간적으로 다르고 생검이 대표하지 못할 수 있음", "Richardson 2026: upgrade 군의 OS가 해당 grade 군과 같음"],
  "Grade를 '자연 경과'로 볼지 '현재 치료 하 결과'로 볼지는 아직 논쟁 중이다(Update 10)."));
C.push(...QA(3, "Grade 4 IDH-mutant astrocytoma는 GBM과 같은 병인가?",
  "아니다. CNS5는 'GBM, IDH-mutant'라는 이름을 없앴다. IDH-mut grade 4는 IDH-wt GBM과 임상·유전적으로 다르고 grade 2/3 IDH-mut astrocytoma와 가깝다.",
  ["Update 5 원문", "Wong 2021: IDH-mut primary GBM median OS 39.4개월로 IDH-wt GBM보다 나음", "Jo 2025: A4 평균 39.8세 vs GBM 63.1세"],
  null));
C.push(...QA(4, "Update 11의 PDGFRA와 Update 12의 PDGFRA는 같은 의미인가?",
  "다르다. Update 11은 IDH-wt, H3-wt HGG에서 <25세일 때 pediatric-type HGG를 시사하는 진단 marker, Update 12는 IDH-mut astrocytoma의 grading marker다.",
  ["Update 11 abstract", "Update 12 abstract"], null));

C.push(H2("11-B. Update 12의 근거와 방법론"));
C.push(...QA(5, "PDGFRA amp → grade 4의 근거는 결국 후향적 연구 아닌가?",
  "맞다. 무작위 시험에서 grade 기준을 검증한 연구는 없다. 다만 여러 독립 코호트와 무작위 시험(CATNON)의 분자 분석에서 일관되게 나쁜 예후와 연관됐고, 대규모 validation도 있다.",
  ["Yang 2019(160명), CATNON 432명 다변량 독립 인자, Tesileanu 2022 systematic review에서 여러 연구 재현", "Richardson 2026(1,140명): 51.4 vs 126.2개월, HR 1.9", "Update 5(2020)는 근거 부족으로 보류했다가 2026년 채택 — 근거 축적 후 결정"],
  "대부분 후향적, 치료 이질성, 추적 기간이 IDH-mut 생존(>10년)에 비해 짧은 문제는 Update 5도 지적했다."));
C.push(...QA(6, "Virata 2026은 PDGFRA를 intermediate로 봤는데, 왜 grade 4인가?",
  "연구마다 결론이 다르다는 것을 인정한다. 다만 Virata에서도 PDGFRA alteration 단독은 HR 4.5였고, 단독 marker와 CDKN2A HD 단독 사이 OS 차이는 없었다. Update 12와 Richardson validation은 grade 4를 지지한다.",
  ["Virata: 다변량 비유의, 단독 alteration HR 4.5, 2-tier 가능성 언급", "Richardson: PDGFRA amp grade 4와 원래 grade 4 OS 차이 없음(p=0.1277)", "Virata는 amplification과 mutation을 합친 'alteration'으로 분석"],
  "PDGFRA alteration 환자 수가 적고(6.5%), 다른 변이와 함께 있는 경우가 많아 독립 효과를 가리기 어렵다."));
C.push(...QA(7, "51.4개월이면 grade 4치고 너무 길지 않나? IDH-wt GBM은 훨씬 짧은데.",
  "IDH-mut grade 4는 IDH-wt GBM보다 원래 예후가 좋다. 비교 대상은 IDH-wt GBM이 아니라 IDH-mut grade 4이고, 이 군과 차이가 없었다.",
  ["Richardson: Update 12 기준 grade 4 전체 median OS 37.7개월, PDGFRA amp 재분류 군과 원래 grade 4 차이 없음", "Virata: grade 4 35.3–45.0개월", "Wong 2021: IDH-mut primary GBM 39.4개월"],
  "51.4개월은 grade 4 전체(37.7)보다 숫자상 길다. 표본이 22례로 작아 신뢰구간이 넓을 것이다 [원문 확인 권장]."));
C.push(...QA(8, "Mitosis 3개 / 2.4 mm²의 근거는? 왜 2.4 mm²인가?",
  "Update 12 초록은 '고려할 수 있다(could be considered)'는 제안 수준이다. 세부 근거 연구는 원문 Table에서 확인해야 한다.",
  ["Update 5: 검증된 cut-off가 없어 'significant' mitotic activity 유지, 관행상 전체 검체 ≥2", "IHC 논문: cut-off가 낮아 생검에서 grading이 어렵다고 지적"],
  "2.4 mm²의 선택 근거와 근거 연구는 직접 확인하지 못했다 [원문 확인 권장]."));
C.push(...QA(9, "Methylation A_IDH_HG만 나오면 grade 4로 해도 되나?",
  "안 된다. Grade 3 또는 4를 시사하는 adverse feature로 기술하고, 확립된 grade 4 기준(MVP, necrosis, CDKN2A/B HD, PDGFRA amp)이 있을 때만 grade 4.",
  ["Update 12 abstract: 'grade 3 or 4, depending upon other grading criteria'", "Tralongo 2026: A_IDH_HG·G-CIMP-low는 독립 grade 4 기준 아님, 두 용어는 서로 바꿔 쓰면 안 됨"], null));

C.push(H2("11-C. 검출과 기술"));
C.push(...QA(10, "PDGFRA amplification의 정의는? 몇 copy부터인가?",
  "검사법마다 다르다. FISH는 >8 copies in ≥10% 종양세포, methylation CNV는 log ratio >0.8(IHC 논문) 또는 ±0.60 log2(CATNON)를 썼다.",
  ["Tauziède-Espariat, Brain Pathol 2026", "Tesileanu 2021"],
  "Update 12가 정한 통일 정의는 확인하지 못했다 [원문 확인 권장]. Gain(≤6 copies)은 amplification이 아니다."));
C.push(...QA(11, "종양 일부에만 증폭(clonal amplification)이 있어도 grade 4인가?",
  "명확한 기준은 찾지 못했다. FISH의 실무 정의(≥10% 세포)를 만족하면 amplification으로 보고된다.",
  ["IHC 논문: clonal amp 10례 중 FISH로 모두 검출, CNV로는 4례만 검출", "Tralongo: 단일 block은 모든 subclone을 대표하지 못함"],
  "Clonal amp의 예후 의미는 근거가 부족하다."));
C.push(...QA(12, "PDGFRA IHC만으로 진단해도 되나?",
  "안 된다. Screening 도구다. ≥10% 강양성은 민감도 100%, 특이도 89%로, gain이나 mutation에서도 양성이 나올 수 있어 FISH 등으로 확진해야 한다.",
  ["Tauziède-Espariat, Brain Pathol 2026: 강양성 >10% → 강력 의심, 국소 → 유전 검사 필수"], null));
C.push(...QA(13, "우리 병원에서 모든 grade 2/3에 NGS·methylation을 해야 하나?",
  "Update 12를 적용하려면 사실상 CNV 검사가 필요하다. 실제 진료에서 24%의 grade가 바뀌었다.",
  ["Ann Pathol 2026: grade 2는 NGS + methylation, grade 3는 CDKN2A/B HD·PDGFRA amp을 볼 수 있는 기법(NGS, methylation, SNP/CGH array) 권장, 필요 시 전문 기관 의뢰"],
  "비용·조직량·turnaround time 문제는 남아 있다(Tralongo 2026)."));

C.push(H2("11-D. 영상"));
C.push(...QA(14, "영상으로 PDGFRA amp을 예측할 수 있나?",
  "현재로서는 아니다. 특이 소견이 확립되지 않았다. 영상은 aggressive focus를 찾아 분자 검사를 유도하는 역할이다.",
  ["PubMed 검색에서 IDH-mut astrocytoma 대상 PDGFRA 영상 연구를 찾지 못함", "Kurokawa 2026: 영상은 targeted sampling과 분자 검사 우선순위 결정에 핵심"], null));
C.push(...QA(15, "Grade 4인데 조영증강이나 괴사가 없는 게 말이 되나?",
  "된다. 분자적으로만 grade 4인 종양은 조직학적 grade 4 소견(괴사·MVP)이 없고, 영상에서도 조영증강이 없을 수 있다.",
  ["Ozono 2026: non-enhancing AST 38명 중 grade 4 존재", "Imai 2025: non-enhancing glioma에서 rCBV로 grade 4 구분(AUC 0.718)", "Alfattal 2026: grade 2의 50%가 조영증강 → 조영증강과 grade는 일대일 대응이 아님"],
  "Jo 2025에서 A4의 96%가 조영증강을 보였으므로, 비조영 grade 4는 드문 편이다."));
C.push(...QA(16, "ADC·rCBV 이상만으로 grade를 올리자고 한 것은 과잉 아닌가?",
  "영상으로 grade를 바꾼 것이 아니라, 불일치를 근거로 분자 재검토를 요청했고 grade는 병리 소견(PDGFRA amp)으로 바뀌었다.",
  ["Kurokawa 2026: diffusion restriction·perfusion 상승 부위가 biopsy target이고 분자 검사 우선순위를 정하는 근거"],
  "ADC·rCBV 단독으로 IDH-mut astrocytoma grade를 정하는 검증된 cut-off는 없다."));
C.push(...QA(17, "T2-FLAIR mismatch가 있으면 저등급 아닌가?",
  "Mismatch는 IDH-mut astrocytoma를 시사할 뿐 grade를 보장하지 않는다. Partial mismatch는 grade 4에서도 보인다. 다만 CDKN2A HD 증례는 mismatch 음성이었다는 보고가 있다.",
  ["Kurokawa 2026: grade 4 IDH-mut의 26.4%가 partial mismatch", "Ozono 2026: CDKN2A HD 5례 모두 mismatch 음성"], null));

C.push(H2("11-E. 치료"));
C.push(...QA(18, "CATNON에서 concurrent TMZ는 이득이 없었는데 왜 CCRT를 했나?",
  "CATNON은 grade 3 조직(anaplastic) 환자 대상이었다. 분자적 grade 4 IDH-mut astrocytoma에 대한 무작위 근거는 없고, ASCO-SNO는 grade 4 IDH-mut에서 grade 3 권고와 GBM 권고 중 어느 쪽이든 허용한다. CCRT는 그 선택지 중 하나다.",
  ["CATNON final: IDH-mut에서 concurrent HR 0.81(95% CI 0.63–1.04, 비유의), adjuvant HR 0.54", "Mohile 2022(ASCO-SNO)"],
  "Concurrent TMZ가 IDH-mut에서 이득이 있다는 직접 근거는 없다. PDGFRA amp은 TMZ 효과를 예측하지 않는다(CATNON final). 치료 결정은 MDT의 판단이다."));
C.push(...QA(19, "Grade만 올리고 치료 효과는 증명 못 한 것 아닌가?",
  "맞는 지적이다. Grade 상향은 예후 분류의 개선이고, 치료 강화가 생존을 늘린다는 직접 근거는 아직 없다.",
  ["CATNON final: 분자 marker는 prognostic, not predictive", "Virata 2026: 확장 기준이 치료 escalation과 임상시험 층화를 돕는다고 제안"],
  "전향적 검증이 필요하다."));
C.push(...QA(20, "Grade 3로 뒀으면 vorasidenib을 쓸 수 있었나?",
  "아니다. INDIGO는 grade 2만 포함했다. Grade 3였어도 vorasidenib 근거는 없다.",
  ["INDIGO: 잔존/재발 grade 2, 비조영 또는 최소 조영", "Tralongo 2026: grade 3–4, 뚜렷한 조영증강 환자로 일반화 불가"], null));
C.push(...QA(21, "그럼 upgrade가 실제로 치료를 바꾼 지점은?",
  "Grade 3 기준이면 RT → adjuvant TMZ가 표준이다(CATNON). Grade 4가 되면서 GBM-type 요법(concurrent + adjuvant TMZ)도 가이드라인상 선택지가 됐고, 이 환자는 CCRT를 받았다. 예후 상담과 임상시험 적격성도 달라진다.",
  ["ASCO-SNO 2022", "CATNON"], null));

C.push(H2("11-F. 증례 특이"));
C.push(...QA(22, "63세면 IDH-wt GBM일 가능성은 배제했나?",
  "IDH-mutant로 진단됐다. 다만 고령 IDH-mut astrocytoma는 드물어 IDH 확인 방법을 차트에서 확인해 두는 것이 좋다.",
  ["Virata 2026: IDH-mut AST ≥55세 8.6%", "Alfattal 2026: ≥55세는 IDH-wt가 많지만 IDH-mut AST 22명 보고, IDH1 R132H가 95%"],
  "검사 방법(IHC vs sequencing)은 증례 정보에서 확인 필요 [원문 확인 권장]."));
C.push(...QA(23, "PDGFRA amp이 나온 검체가 영상의 aggressive 부위와 같은 곳인가?",
  "차트·병리 보고서에서 확인해야 한다. 같은 부위라면 영상-병리 상관의 좋은 예이고, 다르다면 종양 내 이질성을 보여준다.",
  ["Tralongo 2026: sampling bias, 지역별 차이", "IHC 논문: clonal amp은 증폭 부위에만 IHC 양성"],
  "[원문 확인 권장]"));
C.push(...QA(24, "CDKN2A/B 상태는? HD가 같이 있었다면 PDGFRA는 의미가 없지 않나?",
  "CDKN2A/B HD가 있었다면 그것만으로 이미 grade 4다. 이 증례의 의미는 PDGFRA amp이 grade 4의 근거였다는 것이므로, CDKN2A/B 상태를 확인해 발표해야 한다.",
  ["Yang 2019: PDGFRA amp과 CDKN2A HD는 상호 배타적인 경향"],
  "[원문 확인 권장] 차트 확인"));

// =============== 12. Evidence level ===============
C.push(H1("12. 주장별 근거 수준"));
C.push(...T(["주장", "근거 유형", "강도 [제안]"], [
  ["CDKN2A/B HD → grade 4", "다수 후향 코호트 + WHO CNS5 공식 기준", "높음 (공식)"],
  ["PDGFRA amp → grade 4", "후향 코호트 + RCT 내 분자 분석 + 대규모 후향 validation + 전문가 합의", "중간 (cIMPACT-NOW, 이견 존재)"],
  ["Mitosis ≥3 / 2.4 mm²", "전문가 합의 ('could be considered')", "낮음–중간 [원문 확인 권장]"],
  ["Grade 3 분자 marker", "후향 코호트 (Virata, Richardson)", "중간"],
  ["Adjuvant TMZ (grade 3 IDH-mut)", "Phase 3 RCT (CATNON)", "높음"],
  ["Concurrent TMZ (IDH-mut)", "Phase 3 RCT — 유의한 이득 없음", "이득 근거 없음"],
  ["Vorasidenib (grade 2)", "Phase 3 RCT (INDIGO), PFS", "높음 (PFS), OS 미확립"],
  ["Grade 4 IDH-mut 치료", "가이드라인 (선택지 2개)", "낮음 (전용 RCT 없음)"],
  ["T2-FLAIR mismatch → AST", "Meta-analysis", "특이도 높음, 민감도 낮음"],
  ["PDGFRA amp 영상 소견", "없음", "근거 없음"],
], [3.2, 4.0, 2.0]));

// =============== 13. Glossary ===============
C.push(H1("13. 용어·약어"));
C.push(...T(["약어", "의미"], [
  ["AST / ODG / GBM", "Astrocytoma IDH-mut / Oligodendroglioma IDH-mut 1p/19q-codeleted / Glioblastoma IDH-wt"],
  ["HD", "Homozygous deletion (양쪽 allele 소실). Hemizygous = 한쪽 소실"],
  ["amp", "Amplification (gain보다 높은 copy 증가; FISH >8 copies 등)"],
  ["MVP", "Microvascular proliferation"],
  ["G-CIMP", "Glioma CpG island methylator phenotype (high / low)"],
  ["A_IDH_LG / A_IDH_HG", "Methylation classifier의 IDH-mut astrocytoma 저등급 / 고등급 subclass"],
  ["CNV", "Copy number variation"],
  ["NOS / NEC", "Not otherwise specified / Not elsewhere classified"],
  ["pedHGG", "Diffuse pediatric-type high-grade glioma, H3-wt and IDH-wt"],
  ["RTK", "Receptor tyrosine kinase"],
  ["D-2HG", "D-2-hydroxyglutarate, IDH 변이의 oncometabolite"],
  ["CCRT", "Concurrent chemoradiotherapy"],
  ["MFDS", "식품의약품안전처"],
], [2.4, 6.8]));

// =============== 14. Summary table ===============
C.push(H1("14. 최종 요약표"));
C.push(...T(["항목", "핵심", "출처"], [
  ["WHO CNS5", "AST IDH-mut grade 2–4, CDKN2A/B HD → grade 4", "Louis 2021; Update 5"],
  ["Update 3", "IDH-wt + EGFR amp/+7−10/TERTp → grade IV, integrated grade 우선", "Brat 2018"],
  ["Update 5", "CDKN2A/B HD → grade 4, PDGFRA는 보류", "Brat 2020"],
  ["Update 8", "Meningioma: 1p del + 22q del/NF2 → grade 2", "Sahm 2025"],
  ["Update 9", "Methylation classifier는 통합 진단의 한 도구", "Aldape 2025"],
  ["Update 10", "새 type 정의(영상 포함 clinical phenotype 필수)", "Hawkins 2025"],
  ["Update 11", "<40세 EGFR amp·+7/−10 단독 GBM 금지, PDGFRA는 <25세 pedHGG marker", "Wesseling 2025"],
  ["**Update 12**", "**PDGFRA amp → grade 4**, mitosis ≥3/2.4 mm², grade 3 분자 marker", "Brat 2026; Park YW KJR 2026 Fig. 2"],
  ["Validation", "51.4 vs 126.2개월 (HR 1.9)", "Richardson 2026"],
  ["반론", "Intermediate, 다변량 비유의, 단독 HR 4.5", "Virata 2026"],
  ["실제 진료", "24% grade 변경", "Tauziède-Espariat 2026"],
  ["검출", "FISH, NGS, methylation CNV, SNP/CGH, IHC screening", "Tauziède-Espariat 2026"],
  ["영상", "PDGFRA 특이 소견 없음, ADC↓/rCBV↑·mismatch 소실·poorly defined margin = 고위험 단서", "Ozono, Yang, Imai, Kurokawa"],
  ["치료", "G2 vorasidenib / G3 RT→adjuvant TMZ / G4 grade 3 또는 GBM 권고", "INDIGO, CATNON, ASCO-SNO"],
], [1.8, 5.2, 2.2], { highlight: (r) => r[0] === "**Update 12**" }));

// ---------- References ----------
C.push(H1("참고문헌"));
const refs = [
  "Brat DJ, et al. cIMPACT-NOW Update 12: Refining pathology-based risk stratification and grading for IDH-mutant gliomas. Neuro-Oncol 2026;28(9):2069–2089. doi:10.1093/neuonc/noag108",
  "Park YW, Wagner A, Karschnia P, Vollmuth P, Sahm F, Ahn SS, Kim HS, Park JE. Imaging of isocitrate dehydrogenase-mutant gliomas in the era of isocitrate dehydrogenase inhibitors: what radiologists need to know. Korean J Radiol 2026 (in press). [Fig. 2]",
  "Louis DN, et al. The 2021 WHO Classification of Tumors of the Central Nervous System: a summary. Neuro-Oncol 2021;23(8):1231–1251. doi:10.1093/neuonc/noab106",
  "Louis DN, et al. cIMPACT-NOW update 1. Acta Neuropathol 2018;135:481–484. doi:10.1007/s00401-018-1808-0",
  "Louis DN, et al. cIMPACT-NOW update 2. Acta Neuropathol 2018;135:639–642. doi:10.1007/s00401-018-1826-y",
  "Brat DJ, et al. cIMPACT-NOW update 3. Acta Neuropathol 2018;136:805–810. doi:10.1007/s00401-018-1913-0",
  "Ellison DW, et al. cIMPACT-NOW update 4. Acta Neuropathol 2019;137:683–687. doi:10.1007/s00401-019-01987-0",
  "Brat DJ, et al. cIMPACT-NOW update 5. Acta Neuropathol 2020;139:603–608. doi:10.1007/s00401-020-02127-9",
  "Louis DN, et al. cIMPACT-NOW update 6. Brain Pathol 2020;30:844–856. doi:10.1111/bpa.12832",
  "Ellison DW, et al. cIMPACT-NOW update 7. Brain Pathol 2020;30:863–866. doi:10.1111/bpa.12866",
  "Sahm F, et al. cIMPACT-NOW update 8. Neuro-Oncol 2025;27(2):319–330. doi:10.1093/neuonc/noae170",
  "Aldape K, et al. cIMPACT-NOW update 9. Neuro-Oncol Adv 2025;7(1):vdae228. doi:10.1093/noajnl/vdae228",
  "Hawkins C, et al. cIMPACT-NOW update 10. Brain Pathol 2025;35(6):e70018. doi:10.1111/bpa.70018",
  "Wesseling P, et al. cIMPACT-NOW update 11. Brain Pathol 2025;36(1):e70035. doi:10.1111/bpa.70035",
  "Richardson TE, et al. Validation of proposed cIMPACT-NOW update 12 molecular grading criteria for IDH-mutant astrocytoma. Acta Neuropathol 2026;152. doi:10.1007/s00401-026-03058-7",
  "Virata MC, et al. Expanding the molecular grading criteria in IDH-mutant astrocytoma. Neuro-Oncol 2026. doi:10.1093/neuonc/noag069",
  "Tauziède-Espariat A, et al. [Update 12 in routine practice]. Ann Pathol 2026. doi:10.1016/j.annpat.2026.09.001",
  "Tauziède-Espariat A, et al. Immunohistochemistry as a tool for identifying PDGFRA amplification in CNS tumors. Brain Pathol 2026:e70142. doi:10.1111/bpa.70142",
  "Tralongo P, et al. DNA methylation profiling in IDH-mutant gliomas: biological evolution, molecular grading and clinical implementation. Neuropathol Appl Neurobiol 2026. doi:10.1111/nan.70100",
  "Yang RR, et al. IDH mutant lower grade astrocytomas can be stratified by CDKN2A, CDK4 and PDGFRA copy number alterations. Brain Pathol 2019. doi:10.1111/bpa.12801",
  "Tesileanu CMS, et al. Prognostic significance of genome-wide DNA methylation profiles within the CATNON trial. Neuro-Oncol 2021. doi:10.1093/neuonc/noab088",
  "Tesileanu CMS, et al. Molecular markers related to outcome in IDH-mutant astrocytomas grade 2–4: a systematic review. Eur J Cancer 2022. doi:10.1016/j.ejca.2022.08.016",
  "Wong QH, et al. Molecular landscape of IDH-mutant primary astrocytoma grade IV/glioblastomas. Mod Pathol 2021. doi:10.1038/s41379-021-00778-x",
  "Yamashita S, et al. CDKN2A/B status and histological grade predict post-recurrence survival in recurrent IDH-mutant astrocytoma. Brain Tumor Pathol 2026. doi:10.1007/s10014-026-00541-7",
  "van den Bent MJ, et al. CATNON second interim analysis. Lancet Oncol 2021;22:813–823. doi:10.1016/S1470-2045(21)00090-5",
  "van den Bent MJ, et al. CATNON final and exploratory analyses. Lancet Oncol 2026;27:45–56. doi:10.1016/S1470-2045(25)00614-X",
  "Mellinghoff IK, et al. Vorasidenib in IDH1- or IDH2-mutant low-grade glioma (INDIGO). N Engl J Med 2023;389:589–601. doi:10.1056/NEJMoa2304194",
  "Mohile NA, et al. Therapy for diffuse astrocytic and oligodendroglial tumors in adults: ASCO-SNO guideline. J Clin Oncol 2022;40(4):403–426. doi:10.1200/JCO.21.02036",
  "Weller M, et al. EANO guidelines on the diagnosis and treatment of diffuse gliomas of adulthood. Nat Rev Clin Oncol 2021;18:170–186. doi:10.1038/s41571-020-00447-z",
  "Wu J, et al. Treatment options for adult grade 2 IDH-mutant diffuse gliomas in the vorasidenib era. Curr Treat Options Oncol 2026. doi:10.1007/s11864-026-01401-4",
  "Kurokawa R, et al. Imaging diagnosis of adult-type diffuse gliomas in the era of WHO CNS5 and cIMPACT-NOW updates 8–11. Jpn J Radiol 2026;44:1534–1542. doi:10.1007/s11604-026-02013-6",
  "Jo J, et al. Presenting clinical and radiographic features of diffuse glioma molecular subtypes (CNS5). Neurooncol Pract 2025. doi:10.1093/nop/npaf067",
  "Ozono I, et al. T2-FLAIR mismatch sign in nonenhancing astrocytoma, IDH-mutant. J Neurooncol 2026. doi:10.1007/s11060-026-05519-1",
  "Alfattal R, et al. IDH-mutant astrocytomas in patients older than 55 years. Int J Surg Pathol 2026. doi:10.1177/10668969261489616",
  "Yang H, et al. Clinical and VASARI features to predict CDKN2A homozygous deletion in IDH-mutant astrocytomas. AJNR 2025. doi:10.3174/ajnr.A8861",
  "Singh KPB, et al. MRI-based prediction of DNA methylation grade in IDH-mutant astrocytomas. Neuroradiology 2025. doi:10.1007/s00234-025-03766-z",
  "Imai N, et al. [11C]-methionine PET and CBV imaging in non-contrast-enhanced gliomas. J Clin Med 2025;14:6777. doi:10.3390/jcm14196777",
  "Śledzińska-Bebyn P, et al. Investigating glioma genetics through perfusion MRI. Magn Reson Imaging 2024. doi:10.1016/j.mri.2024.110318",
  "Le Rhun E, et al. The clinical and molecular landscape of diffuse hemispheric glioma, H3 G34-mutant. Neuro-Oncol 2025. doi:10.1093/neuonc/noaf015",
];
refs.forEach((r) => C.push(new Paragraph({ numbering: { reference: "ref", level: 0 }, children: runs(r, { size: 18 }), spacing: { after: 40 } })));

// =====================================================================
const doc = new Document({
  creator: "Study note",
  title: "cIMPACT-NOW Update 12 심층 학습 노트",
  styles: {
    default: { document: { run: { font: FONT, size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, color: NAVY, font: FONT }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 4 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 25, bold: true, color: BLUE, font: FONT }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, color: NAVY, font: FONT }, paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: "bul", levels: [
      { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 260 } } } },
      { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 800, hanging: 260 } } } },
      { level: 2, format: LevelFormat.BULLET, text: "·", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1200, hanging: 260 } } } },
    ] },
    { reference: "ref", levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 440, hanging: 440 } } } },
    ] },
  ] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: GRAY, font: FONT })] })] }) },
    children: C,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = "/home/user/test-repro/cIMPACT-NOW_update12_학습노트.docx";
  fs.writeFileSync(out, buf);
  console.log("wrote", out);
});
