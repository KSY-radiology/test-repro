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
  "**Update 12는 원문 PDF 전체(본문, Table 1·2, Figure 1, 참고문헌)를 읽고 반영**했다(Neuro-Oncol 2026;28(9):2069–2089). 5장은 원문 기준으로 다시 썼고, 이전 판에서 2차 문헌으로 추정했던 부분을 바로잡았다.",
  "**전문을 직접 읽은 다른 문헌**: Update 3, Update 5, Update 10, Update 11, Virata 2026, Richardson 2026, CATNON methylation 2021(Tesileanu), PDGFRA IHC 2026, Tralongo 2026(methylation 리뷰), Kurokawa 2026(Jpn J Radiol), Ozono 2026.",
  "**Update 12가 인용했지만 원문을 직접 보지 않은 연구**(Ghosh 2025, Kros 2018, Tran 2023, Shirahata 2018, POLA 등)는 **Update 12에 적힌 수치 그대로** 옮겼다.",
  "**초록만 확인한 문헌**: WHO CNS5 요약(Louis 2021), Update 6–9, Yang 2020, CATNON 2021·2026, INDIGO, ASCO-SNO 2022, EANO 2021, Alfattal, Yang AJNR, Singh, Imai, Jo, Wong 등.",
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
  "핵심 근거는 CATNON 분자 분석으로, grade 3 환자 중 PDGFRA amp이 있으면 median OS가 3.1년(없으면 9.3년)이었고 이는 CDKN2A/B HD와 독립적이었다.",
  "PDGFRA amp 자체의 영상 소견은 없지만, 이 환자처럼 **ADC·rCBV가 조직 grade보다 나쁘면** 영상의학과의 피드백이 분자 재검토와 grade 상향, 치료 변경으로 이어질 수 있다.",
  "English: \"Update 12 made PDGFRA amplification a grade 4 criterion: in CATNON, PDGFRA-amplified grade 3 tumors had a median OS of 3.1 vs 9.3 years, independent of CDKN2A/B. Imaging cannot detect it, but discordant ADC/rCBV can trigger the molecular work-up that changes grade and treatment.\"",
], GREEN));

// =============== 1. Summary ===============
C.push(H1("1. 한눈에 보는 핵심 요약"));
C.push(...BOX("Take-home messages", [
  "WHO CNS5(2021): **Astrocytoma, IDH-mutant = 단일 type, CNS WHO grade 2–4**. Grade 4 = MVP 또는 necrosis 또는 **CDKN2A/B homozygous deletion(HD)**.",
  "Update 12(2026): **PDGFRA amplification → grade 4** (\"highly aggressive clinical behavior, consistent with CNS WHO grade 4\").",
  "Grade 2/3 경계: **≥3 mitoses / 2.4 mm²**(= 0.24 mm² HPF 10개). 근거는 Kros 2018(455명, PFS만 유의)과 Tran 2023(100명). 원문도 \"WHO 기준으로 완전히 승인하려면 추가 연구 필요\"라고 명시.",
  "**Grade 3 정의에 분자 기준 포함(Table 1)**: CDK4·CDK6·CCND2 amp, RB1 HD/mutation, CDKN2A mutation, PIK3CA/PIK3R1 mutation, EGFR alteration, MYCN amp, methylation **G-CIMP-low / A_IDH_HG**(최소 grade 3) → 조직 grade 2라도 이 중 하나면 grade 3.",
  "Oligodendroglioma: grading을 바꿀 새 근거 없음. 단 mitosis만 높고 necrosis·MVP·CDKN2A/B HD·**MRI 조영증강이 없는** grade 3은 생존이 길 수 있음.",
  "PDGFRA amp 빈도: grade 2·3의 **4–6%**, grade 4의 **8–15%**.",
  "**핵심 근거 = CATNON**(grade 3 IDH-mut astrocytoma 440명): PDGFRA amp **OS 3.1년 vs 9.3년, HR 3.51, 다변량, CDKN2A/B HD와 독립** → grade 4 기준치(median OS <6년)에 해당. 보조: Yang 2020(HR 2.90), Ghosh 2025(HR 1.9), Shirahata 2018, Richardson & Walker 2020.",
  "이견: Virata 2026은 다변량에서 유의하지 않았고 intermediate로 봄 — **Update 12 원문도 이를 인용하며 인정**. 발표 후 독립 검증(Richardson 2026, 1,140명): 51.4 vs 126.2개월(HR 1.9).",
  "검출: FISH, methylation array CNV, NGS, SNP/CGH array. **PDGFRA IHC**(≥10% 강양성)는 민감도 100%, 특이도 89%의 screening 도구.",
  "영상: PDGFRA amp 특이 소견은 **없음**. Update 12는 **조영증강이나 빠른 진행이 보이면 분자 검사를 하라**고 명시하고, **천천히 자라는 T2 고신호 비조영 병변 + 저등급 생검이면 분자 검사 없이도 grade 2**를 진단할 수 있다고 함. 수술 전·후 종양 크기가 가장 일관된 예후 인자.",
  "치료: grade 3 = RT → adjuvant TMZ(CATNON), grade 4 IDH-mut = **grade 3 또는 GBM 권고 중 선택 가능**(ASCO-SNO 2022), vorasidenib은 grade 2 근거(INDIGO).",
]));
C.push(SRC("Brat 2026 (원문 전체); Brat 2020 Update 5 (PMC8443062); Richardson 2026 (PMC13391642); Virata 2026 (PMC13437890); Tauziède-Espariat, Brain Pathol 2026 (PMC13535536); Mohile, J Clin Oncol 2022 (abstract)"));

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
  "Yang 2020: PDGFRA amp, CDKN2A HD, CDK4 amp가 **서로 배타적** → 같은 '증식 신호 활성화'에 이르는 서로 다른 경로로 해석 가능 [보충]",
]));
C.push(SRC("Tralongo 2026 (PMC13498940); Tesileanu, Eur J Cancer 2022 (10.1016/j.ejca.2022.08.016, abstract); Yang 2020 (abstract)"));

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
  ["**12**", "**2026 Neuro-Oncol**", "**IDH-mutant glioma grading**", "**PDGFRA amp → grade 4**, ≥3 mitoses/2.4 mm² 또는 Table 2 분자 변화 → grade 3, 새 subgroup 3개, oligo 개정 없음"],
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
C.push(H1("5. cIMPACT-NOW Update 12 상세 (원문 전체 기준)"));
C.push(H2("5-1. 배경·목적·방법"));
C.push(...B([
  "Update 5와 CNS5는 주로 **grade 4 marker**를 세우는 데 집중 → **grade 2와 3의 경계**를 개선할 필요가 남음",
  "CNS5의 grade 2 vs 3 구분: astrocytoma = mitotic activity + histologic anaplasia / oligodendroglioma = mitotic activity, MVP, necrosis → \"이 기준이 최적이 아니라는 데 광범위한 합의\"",
  "임상적 의미가 커진 이유: **저위험 환자는 가이드라인상 watchful waiting 또는 IDH inhibitor** 대상 → grade 2 vs 3 구분이 치료에 직결",
  "단, 치료 결정은 grade만이 아니라 **임상·신경영상 소견**에 크게 좌우 → 다학제 결정 필요",
  "구성: 신경병리, 분자생물, 신경종양 전문가. 문헌 검토 후 **6회 teleconference**, Steering Committee·Clinical Advisory Panel·ISN Executive 검토",
  "2022년 cIMPACT-NOW steering committee 활동 재개 → **WHO CNS6**를 위한 새 update 시리즈의 하나",
]));
C.push(SRC("Brat 2026, Neuro-Oncol 28:2069, p.2069–2070"));

C.push(H2("5-2. Biomarker 채택 기준과 grade별 생존 기준치"));
C.push(...B([
  "**Grading biomarker 자격**: OS와의 연관이 **≥2개 독립 연구**에서 입증",
  "  보조 근거: 종양 grade가 높을수록 빈도 증가, 임상적 진행 시 빈도 증가",
  "  대규모 코호트, 충분한 추적, 통계적 유의성이 있는 OS 자료가 있는 연구에 집중",
  "검토한 분자 변화: CDKN2A/B hemizygous deletion·HD·mutation, CDK4/CDK6 amp, RB1 HD·mutation, CCND2 amp, PIK3CA/PIK3R1 mutation, **PDGFRA amp**, MYCN/MYC amp, EGFR alteration, NOTCH1 mutation, TERTp mutation, TP53 mutation, DNA methylation signature, genomic instability(CNV, TMB)",
  "Biomarker 결과를 IDH-mut glioma 임상 코호트의 OS와 비교해 grade를 매김",
]));
C.push(...T(["Tumor", "Grade", "Median OS 기준", "문헌상 median OS 범위"], [
  ["Astrocytoma, IDH-mut", "**4**", "**< 6년**", "4.1–5.9년"],
  ["Astrocytoma, IDH-mut", "3", "< 10년", "7.0–12.5년"],
  ["Astrocytoma, IDH-mut", "2", "≥ 10년", "11.4–11.8년"],
  ["Oligodendroglioma", "3", "< 14년", "9.8–14.2년"],
  ["Oligodendroglioma", "2", "≥ 14년", "15.6년–미도달"],
], [2.6, 1.0, 2.4, 3.2], { highlight: (r) => r[1] === "**4**" }));
C.push(...B([
  "Grade 4 marker 후보는 **CDKN2A/B HD의 생존 결과와도 비교**",
  "저자들이 인정한 한계: 대부분 임상 보고가 **장기간·이질적으로 치료된** 환자, 표준치료 변화로 생존 개선, 전향적 대규모 series는 일부만 WHO 2016/2021로 분류, 후향 series는 추적 부족 + **수술 전후 종양 부피**와 치료 패턴 차이로 교란",
]));
C.push(SRC("Brat 2026, p.2070"));

C.push(H2("5-3. 공식 Grading 정의 (Table 1)"));
C.push(...T(["Grade", "Table 1 원문 정의 (요약 번역)"], [
  ["2", "IDH1/2 mutation이 있는 diffusely infiltrative astrocytic glioma로, mitotic activity가 없거나 낮음(**<3 mitoses / 2.4 mm²**). MVP, necrosis, **grade 3 또는 4와 연관된 분자 기준 모두 없음**"],
  ["3", "Significant mitotic activity(**≥3 mitoses / 2.4 mm²**) **또는 grade 3에 부합하는 분자 변화**, 또는 그 조합. MVP, necrosis, grade 4 분자 기준 없음"],
  ["**4**", "**MVP 또는 necrosis 또는 CDKN2A/B HD 또는 PDGFRA amplification**, 또는 그 조합"],
], [0.9, 8.3], { highlight: (r) => r[0] === "**4**" }));
C.push(...B([
  "**중요(이전 노트 수정)**: grade 3 분자 변화는 '고려'가 아니라 **Table 1에서 grade 3 정의의 일부** → 형태학적 grade 2라도 Table 2 변화가 있으면 grade 3",
  "2.4 mm² = **고배율 시야(HPF) 0.24 mm² × 10개**",
]));
C.push(SRC("Brat 2026, Table 1 (p.2078), Grading Considerations (p.2077)"));

C.push(H2("5-4. Table 2 — 고등급을 시사하는 분자 변화"));
C.push(...T(["분자 변화", "Grade", "원문 결론 요지"], [
  ["CDK4 amplification", "3", "RB pathway — grade 2와 부합하지 않음. 일부 연구는 grade 4 시사했으나 일관성 부족"],
  ["CDK6 amplification", "3", "RB pathway"],
  ["CCND2 amplification", "3", "RB pathway (일부 연구 HR 8.6)"],
  ["RB1 HD 또는 mutation", "3", "RB pathway"],
  ["CDKN2A **mutation**", "3", "401명 중 mutation 15명 — HD(47명)와 비슷한 결과. 반대 allele LOH 유무와 무관"],
  ["PIK3CA 또는 PIK3R1 mutation", "3", "결과 혼재. CATNON 장기 추적에서는 연관 소실. TMZ 후 hypermutation에서 획득될 수 있음. '대세 근거상 grade 3'"],
  ["EGFR alteration (amp 또는 mutation)", "3", "드묾(grade 2 0/250, grade 3 1.1%, grade 4 4.5%), 예후 불량"],
  ["MYCN amplification", "3", "드물고 고등급에 많음, 예후 불량"],
  ["Methylation G-CIMP-low 또는 A_IDH_HG", "3 (또는 4)", "최소 grade 3. Grade 4 여부는 다른 grade 4 소견 유무로 결정. Classifier score **>0.90**만 grading의 확정 근거"],
  ["**PDGFRA amplification**", "**4**", "CATNON에서 CDKN2A/B HD와 독립적인 다변량 효과"],
  ["**CDKN2A/B homozygous deletion**", "**4**", "CNS5 기준 유지"],
], [3.0, 1.0, 5.2], { highlight: (r) => r[1] === "**4**" }));
C.push(SRC("Brat 2026, Table 2 (p.2079) 및 각 절 Summary"));

C.push(H2("5-5. PDGFRA amplification — 원문 근거 전체"));
C.push(...B([
  "**빈도**: grade 2·3 IDH-mut astrocytoma의 **4–6%**, grade 4의 **8–15%** → grade가 높을수록 빈도 증가(채택 기준 충족)",
  "다른 예후 불량 marker와 **겹침이 적음** → 독립적 영향 시사",
  "원문이 인용한 연구",
]));
C.push(...T(["연구", "대상", "결과"], [
  ["**CATNON** (van den Bent, Lancet Oncol 2026)", "Grade 3 IDH-mut astrocytoma, n=440 (CDKN2A/B HD 포함)", "**OS 3.1년 vs 9.3년**, HR **3.51** (2.21–5.58), P<.0001, **다변량**, **CDKN2A/B HD와 독립**"],
  ["Yang RR (Brain Pathol 2020)", "Grade 2/3, n=155", "PDGFRA amp 18.8%, mean OS 40.5개월, 최고위험군, 다변량 HR **2.90** (1.37–6.15), P=.005"],
  ["Richardson & Walker (J Neuropathol Exp Neurol 2020)", "TCGA·cBioPortal 저등급, n=379", "PFS·OS 단축, mean survival 77개월 (P=.0033, 단변량)"],
  ["Shirahata (Acta Neuropathol 2018)", "전 grade, n=211", "생존 단축 (P=.03, 단변량)"],
  ["Ghosh (Neuro-Oncol 2025)", "전 grade, n=998", "다변량 HR **1.9** (1.1–3.4). Grade 2/3에서는 다른 focal amp와 묶어 **intermediate risk** (CDKN2A/B HD 없을 때)"],
  ["Virata (Neuro-Oncol 2026)", "GENIE 840 + cBioPortal 367", "Grade 2/3에서 단변량 HR **4.5** (1.5–13.2), **다변량에서는 유의하지 않음**"],
], [2.6, 2.6, 4.0]));
C.push(...B([
  "**원문 Summary**: \"대부분의 연구는 PDGFRA amp이 grade 2 또는 3 임상 경과와 **맞지 않음**을 보여줌 → IDH-mut astrocytoma에서 **grade 4 기준으로 고려해야 함**\"",
  "**Conclusions**: 예후 불량 변화는 독립 효과인지 CDKN2A/B HD 등과 함께 있는지 따져야 함 → PDGFRA amp은 **여러 연구의 다변량 분석에서 독립 효과** 입증. CATNON(CDKN2A/B HD 포함 코호트)에서 median 3.1년 = grade 4 행동, CDKN2A/B HD와 독립",
  "→ 이전 노트의 'Richardson 2026 validation(51.4 vs 126.2개월)'은 **Update 12 발표 이후** 나온 독립 검증 letter로, Update 12의 근거는 아님",
]));
C.push(SRC("Brat 2026, PDGFRA Amplification (p.2071), Conclusions (p.2083); 인용 24, 26, 29, 34–39"));

C.push(H2("5-6. RB pathway — 왜 PDGFRA는 4이고 CDK4는 3인가"));
C.push(...B([
  "RB pathway 구성: p16INK4a(CDKN2A), p15INK4b(CDKN2B), CDK4, CDK6, Cyclin D2(CCND2), Rb(RB1)",
  "근거 요약",
  "  Shirahata: RB1 HD + CDK4 amp + CDK6 amp를 묶으면 CDKN2A/B HD와 비슷(n=211), 개별로는 RB1 HD만 유의",
  "  Aoki: RB pathway 변화(CDKN2A/B HD, RB1 HD/mutation, CDK4 amp) 묶음 → grade 4 행동(HR 7.1, 20.5), **개별로는 비유의**",
  "  Richardson & Walker: 개별로는 CDKN2A/B HD와 CDK4 amp만 다변량 유의",
  "  Virata: CDK4 amp 다변량 HR 1.91 / Cimino: CDK4 amp 또는 CDKN2A/B HD 단독 = intermediate(median 63개월), chromosome 14 loss 동반 시 23개월",
  "  Yang: CDKN2A/B HD 또는 CDK4 amp → intermediate(HR 2.43)",
  "**원문 결론**: 일부 연구가 CDK4·CCND2 amp의 grade 4 행동을 시사했지만 **CDKN2A/B HD만큼 일관되고 강하지 않음** → **grade 3 기준**",
  "왜 CDKN2A/B HD가 더 강한가(원문 가설): ① p16INK4a + p15INK4b 동시 소실 ② CDKN2A가 **p16INK4a와 p14ARF** 둘 다 encode → RB + p53 경로 동시 타격",
  "**CDKN2A/B hemizygous deletion**: 2개 연구는 생존 단축, 2개는 무관 → 방법·threshold 차이 → **grading 기준으로 쓰지 않음**",
  "CDKN2A/B·CDKN2C promoter methylation의 예후 의미는 불명확",
]));
C.push(SRC("Brat 2026, Retinoblastoma Pathway Alterations (p.2071–2072), Conclusions (p.2083)"));

C.push(H2("5-7. Grading에 쓰지 않기로 한 것들"));
C.push(...T(["후보", "결론", "이유"], [
  ["MYC amplification", "근거 불충분", "결과 상반, 8q gain vs focal amp 정의 불일치"],
  ["NOTCH1 alteration", "근거 불충분", "4–5%로 드묾, 생존 연관 없음, 재발 시 hypermutation과 연관 가능"],
  ["CNV load (genomic instability)", "기준 제시 불가", "고등급·생존 단축과 일관된 연관이나 측정법·threshold 제각각 → **high CNV는 grade 2와 맞지 않음**, 조직 grade 2 + high CNV면 **methylation 검사 권장**"],
  ["TMB", "근거 불충분", "IDH-mut AST 평균 0.5–2.5 mut/Mb. **7–10 mut/Mb**면 **MMR 이상 평가** 필요"],
  ["Ki-67", "기준 권고 안 함", "Grade 2 보통 <4%, 3은 4–10%, 4는 >10%이나 겹침. 5% threshold는 단변량에서만 의미"],
  ["CDKN2A/B hemizygous deletion", "기준 아님", "연구 결과 상반"],
  ["G-CIMP-high / A_IDH", "grade 2 보장 아님", "\"G-CIMP-high나 A_IDH가 항상 grade 2 행동을 뜻하지는 않음\""],
], [2.6, 1.6, 5.0]));
C.push(SRC("Brat 2026, p.2073–2076"));

C.push(H2("5-8. Mitosis 기준 (≥3 / 2.4 mm²)의 근거"));
C.push(...B([
  "CNS5: 문헌 근거 부족으로 **수치 기준 미제시**. Pre-IDH 시대 관행: 전체 검체에서 **0–1 vs ≥2**(작은 생검은 1개)",
  "WHO CNS5 방침: mitotic count는 **HPF 개수가 아닌 면적(mm²)** 기준으로 — 현미경마다 HPF 크기가 달라서, 디지털 병리에서도 필요",
  "Update 8(meningioma) 인용: H&E 3–5 μm, ×400, **hotspot**에서 연속 HPF 계수. 현재 대물렌즈 HPF는 대개 0.23–0.24 mm² → **0.16 mm² HPF 10개 ≈ 0.23–0.24 mm² HPF 7개**",
  "**근거 연구 1 — Kros 2018** (Acta Neuropathol Commun)",
  "  CATNON(EORTC 26053) + EORTC 22033-26033, IDH-mut astrocytoma **455명**, consensus panel review",
  "  최적 threshold **≥3 mitoses / 2.4 mm²** → PFS 유의(P<.001, 다변량 HR 1.62 [1.02–2.56], P=.04), **OS는 비유의**(P=.7, 자료 미성숙)",
  "  한계: 디지털 slide, 7명 중 4명 판독자 결과만 사용, 일치도 **60%**. 디지털 판독은 mitosis를 과소평가할 수 있음",
  "**근거 연구 2 — Tran 2023** (Neuropathol Appl Neurobiol)",
  "  Grade 2/3 IDH-mut astrocytoma **100명**, Update 8 방식(linear 또는 serpentine 연속 HPF)",
  "  고위험: ≥6 / 3 mm²(≥4.8 / 2.4 mm²), 중간위험: ≥3 / 3 mm²(≥2.4 / 2.4 mm²) (P=.0098), kappa 0.76",
  "  한계: 후향, 단일 기관, 100명, 수술 후 치료 다양",
  "**원문 결론**: 기준은 **낮음(mm²당 1–2개)**. ≥3 / 2.4 mm²은 **grade 2 지정의 신중한 재평가를 촉구**하고 grade 2/3 구분 threshold로 **쓸 수 있음**. **WHO 기준으로 완전히 승인하려면 추가 연구 필요**",
  "참고 생존(Shirahata/von Deimling): grade 4 median OS 3–4년, grade 2 11.4–11.8년, grade 3 7.9–9.7년 / German Glioma Network 258명: grade 2 미도달(추적 10.4년), grade 3 8.1년, grade 4 4.7년",
]));
C.push(SRC("Brat 2026, Mitotic Activity and Proliferation Indices (p.2075–2076); 인용 83(Kros 2018), 85(Tran 2023)"));

C.push(H2("5-9. DNA methylation"));
C.push(...B([
  "TCGA: G-CIMP-low 9/9 모두 grade 3(G-CIMP-high는 46%). Grade 4에서 G-CIMP-low 32–40%",
  "G-CIMP-low에 CDKN2A/B HD가 더 많은 경향(41–65% vs 25–26%), G-CIMP-high → low로 진행 가능",
  "**CATNON**: G-CIMP-low 다변량 HR **4.12** (n=428) / A_IDH_HG 다변량 유의(n=412)",
  "A_IDH_HG 비율: grade 2 0–8%, grade 3 6–48%, grade 4 74–79%",
  "German Glioma Network(258명): LINE1 methylation 감소(A_IDH_HG 대리) → grade 4·짧은 생존, 이 중 CDKN2A HD 동반 시 더 나쁨",
  "A_IDH와 A_IDH_HG의 정의는 **고정된 것이 아니며** 새 자료로 바뀔 수 있음",
  "**결론**: G-CIMP-low·A_IDH_HG는 **최소 grade 3**, grade 4는 다른 grade 4 소견 유무로. **Classifier score >0.90만 grading의 확정 근거**, 미만은 다른 소견과 함께 해석",
]));
C.push(SRC("Brat 2026, DNA Methylation Signature (p.2073–2074)"));

C.push(H2("5-10. 새 IDH-mutant astrocytoma subgroup"));
C.push(...T(["Subgroup", "특징", "원문 제안"], [
  ["**Infratentorial** IDH-mut astrocytoma", "42명(grade 2/3/4 = 12/18/12). **80%가 non-R132H**(IDH1 R132C/G/S, IDH2 R172S/G), ATRX loss <50%, MGMT methylation 60%, methylation profile 다름, **생존이 더 나쁨**", "Grading에 반영할 근거는 없음, **별도 subtype** 고려 가능. 절제 가능성 차이 때문일 수도"],
  ["**Primitive neuronal component**", "작은 원형 세포, GFAP 음성, synaptophysin·TTF1 양성(소세포 신경내분비암 전이와 혼동 주의, cytokeratin 음성). TP53·ATRX + RB1·MYCN 이상, G-CIMP-low/A_IDH_HG, **측두엽 호발**, grade 4와 비슷한 생존, **연수막·축외 전파 위험 증가**", "Grade 4에 부합, 별도 subtype 고려"],
  ["**Primary MMR-deficient**", "Lynch / constitutional MMR deficiency, 소아·청소년·젊은 성인. Hypermutation, MGMT unmethylated >60%, **TMZ 반응 불량**, median 생존 **15개월**", "**<25세 고등급 IDH-mut astrocytoma는 MMR 검사**, MMR-deficient glioma는 IDH 검사. 유전 상담 필요"],
], [2.2, 4.6, 2.4]));
C.push(...B([
  "영상의학 연결 [제안]: **후두와 IDH-mut glioma**는 R132H IHC 음성일 수 있음 → sequencing 필요. Primitive neuronal component는 **척수·연수막 전파** 확인 영상이 의미 있음",
]));
C.push(SRC("Brat 2026, Subgroups of IDH-Mutant Astrocytomas (p.2076–2077)"));

C.push(H2("5-11. Oligodendroglioma (원문 상세)"));
C.push(...B([
  "**결론: grading 기준을 바꿀 근거 불충분**",
  "**CDKN2A/B HD**: grade 3에서 5.5–6.8%, **grade 2에서는 보고된 적 없음**(40예, 109예). POLA cohort(grade 3)에서 PFS·OS 불량 인자, 그러나 **median 생존 >10년으로 grade 4 행동 아님** → grade 4로 올리지 않음. 단 발견되면 **grade 2 지정을 재고**",
  "CDKN2A/B hemizygous deletion: 15.9%(메타분석 659명), 예후와 무관 → 기준 아님",
  "NOTCH1(31%), PIK3CA(~20%)/PIK3R1(~9%), TP53: 근거 불충분",
  "**TERTp-wildtype oligodendroglioma**: 청소년, grade 2, 1p/19q 외 CNV 없음, 5년 PFS 93.7% vs 61.1% → 독립 검증 필요",
  "**Mitosis**: brisk mitotic activity = **≥2.5 / mm² = ≥6 / 10 HPF(0.24 mm²)**를 실용적 cut-off로 사용 가능",
  "**POLA cohort(grade 3)**: group 1(mitosis만 높음) 예후 최선, group 2(MVP만), group 3(MVP + necrosis) 최악",
  "  **Group 1 중 MRI 조영증강이 있으면 예후 불량**(OS P=.028, PFS P=.006)",
  "  → **mitosis만으로 grade 3이고 조영증강이 없으면 생존이 길 수 있음**",
  "Ki-67 ≥15%, MCM6 ≥50% → grade 3에서 OS 단축(POLA 220명). ≥2.5 mitoses/mm²는 MRI 성장 속도 증가·PFS 단축과 연관",
  "**Global DNA hypermethylation 소실**(classifier에서 A_IDH_HG로 분류, Heidelberg v12.8): 생존 불량, CDKN2A/B HD·CDK4 amp 많음. 단 A_IDH(저등급)로 분류된 것은 낮은 종양세포 함량 때문",
  "**Oligosarcoma**: 대부분 치료 후 재발에서 발생, NF1·TP53, CDKN2A/B HD 빈번, 진단 후 median 생존 **1.3년**",
]));
C.push(SRC("Brat 2026, Oligodendroglioma 절 (p.2079–2082)"));

C.push(H2("5-12. 다학제 고려사항 — 영상의학과에 직접 관련"));
C.push(...B([
  "Pre-IDH 시대의 예후 인자(**40세 이상**, PS, grade)는 상당 부분 의미를 잃음 → **40세 cut-off는 더 이상 유효하지 않음**",
  "  분자 분류 + 현대 치료(RT 후 항암)의 효과 때문. 예: CATNON에서 RT 단독 5.1년 vs **RT + adjuvant TMZ >12년**",
  "**치료로 결과가 좋아졌다고 grade를 낮추면 안 됨** — 치료 없을 때의 본래 행동을 덜 공격적으로 오해하게 됨",
  "Grade 3 IDH-mut astrocytoma에서 **60세 이상은 예후 불량**",
  "**수술 전·후 종양 부피와 MRI상 성장 속도**가 **주요 예후 인자** → 최대 안전 절제의 중요성(특히 astrocytoma)",
  "**MRI 조영증강의 예후 의미는 논쟁 중**(연관 있다는 연구와 없다는 연구 혼재)",
  "**가장 일관된 임상 예후 인자 = 수술 전·후 종양 크기**",
  "Early adjuvant treatment 비교 후향 연구들은 모두 고위험 환자(큰 잔존, 고등급)에 조기 치료가 몰리는 bias",
]));
C.push(SRC("Brat 2026, Multidisciplinary Considerations (p.2082)"));

C.push(H2("5-13. 분자 검사는 언제 하나 (Figure 1과 결론)"));
C.push(...B([
  "Figure 1 알고리즘: IDH1/2 → ATRX(소실 → astrocytoma 경로 / 유지 → 1p/19q 검사, 필요 시 TERTp·methylation) → 조직 분석 → **분자 진단 검사** → 통합 진단",
  "**분자 검사가 특히 필요한 경우**(원문)",
  "  조직 소견이 grade 2와 3 사이 경계",
  "  **신경영상에서 조영증강 또는 비교적 빠른 진행**",
  "  특이한 임상 양상",
  "**광범위한 분자 검사 없이도 grade 2를 꽤 높은 확신으로 진단할 수 있는 예**(원문): **천천히 자라는 T2 고신호·비조영 병변** + 생검에서 낮은 세포 밀도, 낮은 mitosis, 낮은 Ki-67, grade 3·4 소견 없음",
  "→ **영상 소견이 분자 검사 여부를 결정하는 공식 기준으로 명시**됨 — 발표의 핵심 연결고리",
  "IHC 대리 marker(PDGFRA, EGFR, MYCN, p16, MTAP, Rb, CDK4, CDK6): 가능성은 있으나 **아직 분자 검사를 대체할 단계 아님**",
  "자원이 제한된 환경에서는 기관 상황에 맞춰 결정",
]));
C.push(SRC("Brat 2026, Figure 1 (p.2078), Grading Considerations, Conclusions (p.2083–2084)"));

C.push(H2("5-14. Update 5 → Update 12 비교"));
C.push(...T(["항목", "Update 5 / WHO CNS5", "Update 12 (2026)"], [
  ["Grade 4 분자 기준", "CDKN2A/B HD", "CDKN2A/B HD + **PDGFRA amp**"],
  ["Grade 2 vs 3", "\"significant\" mitotic activity (수치 없음)", "**≥3 mitoses / 2.4 mm²** (10 HPF × 0.24 mm²)"],
  ["Grade 3 분자 기준", "없음", "**Table 2**: CDK4/CDK6/CCND2 amp, RB1, CDKN2A mut, PIK3CA/R1, EGFR, MYCN, G-CIMP-low / A_IDH_HG"],
  ["PDGFRA amp", "\"추가 근거 시 grading 기준 가능\"", "**Grade 4 기준**"],
  ["CDKN2A/B hemizygous", "—", "기준 아님"],
  ["Methylation", "G-CIMP-low 검증 필요", "최소 grade 3, classifier score >0.90"],
  ["Oligodendroglioma", "MVP·necrosis·mitosis", "개정 없음, ≥6 / 10 HPF 실용 cut-off, 비조영 mitosis-only grade 3은 예후 양호 가능"],
  ["새 subgroup", "—", "Infratentorial, primitive neuronal component, primary MMR-deficient"],
], [2.0, 3.2, 4.0]));
C.push(SRC("Update 5 (PMC8443062); Brat 2026"));

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
  ["**CATNON** (van den Bent 2026, Update 12 인용)", "RCT 내 분석, grade 3 IDH-mut AST 440명", "PDGFRA amp **OS 3.1 vs 9.3년, HR 3.51**(2.21–5.58), 다변량, **CDKN2A/B HD와 독립**", "Grade 3 조직만, 사후 분석"],
  ["**Yang RR** Brain Pathol 2020", "후향, IDH-mut grade II/III 160명(다변량 155명), FISH", "PDGFRA amp **18.8%**, CDKN2A HD 15.0%, CDK4 amp 18.1%, 상호 배타. PDGFRA amp → PFS p=0.0003, OS p<0.0001. Mean OS 40.5 / 62.9 / 71.5개월. 다변량 HR 2.90", "단일 기관, 추적 짧음(저자 인정), mean OS 사용"],
  ["**Tesileanu** Neuro-Oncol 2021 (CATNON)", "RCT 내 분자 분석, IDH-mut anaplastic astrocytoma 432명", "다변량 독립 예후인자: 나이, MMSE, TMZ, classifier, **PDGFRA amp**, CDKN2A/B HD, PI3K mut, CNV load. RPA 첫 분기 = CDKN2A HD", "Grade 3 조직만 포함, 사후 분석"],
  ["**van den Bent** Lancet Oncol 2026 (CATNON final)", "IDH-mut 444명, 추적 10.9년", "PDGFRA·CDK4 amp, CDKN2A HD, CNV load → 예후 불량. **TMZ 효과 predictive 아님**", "탐색적 분석"],
  ["**Virata** Neuro-Oncol 2026", "후향, GENIE 840 + cBioPortal 367명", "PDGFRA alteration 6.5%. Cohort 1 median OS 35.5개월(HR 4.4). **다변량 독립 유의성 미달**, 단독 alteration HR 4.5. 'Intermediate' 군에 배치", "Amp + mut 합산, 치료 변수 없음, 소수"],
  ["**Ghosh** Neuro-Oncol 2025 (Update 12 인용)", "전 grade 998명", "다변량 HR 1.9. Grade 2/3에서는 다른 focal amp와 묶어 intermediate", "원문 미확인"],
  ["**Richardson** Acta Neuropathol 2026 (**Update 12 이후** 검증)", "후향 validation, 1,140명", "PDGFRA amp으로 2→4 **7례**, 3→4 **15례**. Median OS **51.4 vs 126.2개월**(HR 1.9, p=0.0284)", "Letter, mitosis 직접 확인 불가, methylation 260례만"],
  ["**Tauziède-Espariat** Ann Pathol 2026", "실제 진료, 50례", "**24%** grade 변경(2→3 7, 2→4 2, **3→4 3**), CNV 기반 8례", "단일 기관, 생존 자료 없음"],
], [1.9, 2.0, 3.6, 1.7]));
C.push(SRC("Brat 2026 p.2071; Yang 2020 (10.1111/bpa.12801); Tesileanu 2021 (PMC8408862); van den Bent 2026 (10.1016/S1470-2045(25)00614-X); Virata 2026 (PMC13437890); Richardson 2026 (PMC13391642); Tauziède-Espariat 2026 (10.1016/j.annpat.2026.09.001)"));

C.push(H2("6-3. 발표 후 독립 검증 (Richardson 2026) 상세"));
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

C.push(H2("6-4. Virata 2026 상세 — 반론의 핵심 자료 (Update 12도 인용)"));
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
  "**Update 12 원문에는 PDGFRA amplification의 copy 수 정의가 없음** → 검사법별 기관 기준을 따름",
  "Update 12의 IHC 언급: PDGFRA·EGFR·MYCN·p16·MTAP·Rb·CDK4·CDK6 IHC는 대리 marker 가능성이 있으나 **분자 검사를 대체할 단계 아님**",
  "IHC 논문 제안 알고리즘: **>10% 강양성 → amplification 강력 의심**, 국소(<10%) 강양성 → FISH 등 필수",
  "IHC 코호트: HGG 114례 중 PDGFRA amp 27례 — pedRTK1A 9/14, **high-grade AST IDH-mut 5/10**, GBM IDH-wt 11/45",
  "Clonal amplification에서는 IHC 양성이 증폭된 부분에만 국한 → **종양 내 이질성**을 시각적으로 보여줌",
  "Tralongo 2026: 순도가 낮으면 focal event는 arm-level event보다 신뢰도 낮음, grade를 바꾸는 애매한 결과는 **orthogonal 확인**",
]));
C.push(SRC("Tauziède-Espariat, Brain Pathol 2026;e70142 (PMC13535536 전문); Tesileanu 2021 (PMC8408862); Tralongo 2026 (PMC13498940)"));

C.push(H2("6-6. 쟁점 정리"));
C.push(...T(["쟁점", "Grade 4 지지", "반대·유보"], [
  ["Grade 4인가 intermediate인가", "Update 12(CATNON 3.1년 = grade 4 기준치 <6년), Richardson 2026(재분류 군 OS = 원래 grade 4, p=0.1277), Yang 2020(최고위험군)", "Virata 2026(intermediate, 다변량 비유의) — 단, 단독 alteration은 CDKN2A HD 단독과 OS 차이 없음"],
  ["독립 예후인자인가", "CATNON 다변량 HR 3.51, **CDKN2A/B HD와 독립**; Yang HR 2.90; Ghosh HR 1.9", "Virata 다변량 비유의(amp+mut 합산, 공동 변이 많음)"],
  ["예측(predictive) marker인가", "—", "CATNON final: TMZ 효과 예측 못함"],
  ["Clonal amp도 grade 4인가", "—", "**Update 12에도 정의 없음**. FISH ≥10% 정의가 실무 기준"],
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
  "**Update 12 기준**: 조영증강 또는 빠른 진행 → 보고서에 분자 검사 필요성 명시 / 천천히 자라는 T2 고신호 비조영 병변 → grade 2 가능성 지지",
  "수술 전 종양 부피와 수술 후 잔존 부피를 정량 기술 (Update 12: 가장 일관된 예후 인자)",
  "후두와 IDH-mut glioma는 non-R132H 많음 → IHC 음성이어도 IDH sequencing 필요성 언급",
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
  "ASCO-SNO 2025 rapid recommendation update(Blakeley, JCO 2025;43:2129)가 있음 — Update 12가 vorasidenib 맥락에서 인용, 세부 내용은 미확인 [원문 확인 권장]",
  "Update 12: 가이드라인상 **저위험 환자는 watchful waiting 또는 IDH inhibitor** → grade 2 vs 3 구분이 치료에 직결",
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
  "Grade 기준을 무작위로 검증한 연구는 없다. 그러나 핵심 근거는 무작위 3상 CATNON의 분자 분석으로, grade 3 환자 440명에서 PDGFRA amp이 있으면 OS 3.1년 vs 9.3년(HR 3.51, 다변량)이었고 CDKN2A/B HD와 독립적이었다.",
  ["Update 12의 biomarker 채택 기준: OS 연관이 ≥2개 독립 연구 + grade·진행에 따라 빈도 증가", "PDGFRA amp: grade 2·3 4–6% → grade 4 8–15%", "독립 연구: CATNON(HR 3.51), Yang 2020(HR 2.90), Ghosh 2025(HR 1.9), Shirahata 2018, Richardson & Walker 2020", "Update 5(2020)는 근거 부족으로 보류 → 근거가 쌓인 2026년에 채택"],
  "CATNON도 사후(exploratory) 분석이다. 저자들 스스로 이질적 치료, 표준치료 변화, 추적 부족, 종양 부피 교란을 한계로 적었다."));
C.push(...QA(6, "Virata 2026은 다변량에서 PDGFRA가 유의하지 않았다는데?",
  "Update 12 원문도 그 결과(단변량 HR 4.5, 다변량 비유의)를 인용하고 있다. 그럼에도 grade 4로 판단한 이유는 CDKN2A/B HD가 포함된 CATNON에서 독립 효과가 다변량으로 확인됐고, 3.1년이라는 생존이 grade 4 기준치(<6년)에 해당하기 때문이다.",
  ["Update 12 원문 Summary: '대부분의 연구'가 grade 2/3 경과와 맞지 않음을 보임", "Virata는 amplification + mutation을 합친 'alteration'으로 분석, PDGFRA는 6.5%로 소수", "Virata에서도 단독 PDGFRA alteration은 HR 4.5", "발표 후 독립 검증(Richardson 2026): 재분류 grade 4와 원래 grade 4 OS 차이 없음"],
  "모든 연구가 일치하는 것은 아니다. 원문 표현도 'majority of studies'다."));
C.push(...QA(7, "Grade 4면 생존이 얼마여야 하나? 기준이 있나?",
  "Update 12는 문헌 검토로 기준치를 정했다. IDH-mut astrocytoma grade 4 = median OS <6년(문헌 4.1–5.9년), grade 3 <10년(7.0–12.5년), grade 2 ≥10년(11.4–11.8년). CATNON의 PDGFRA amp 3.1년은 grade 4 범위다.",
  ["Brat 2026 p.2070", "발표 후 검증(Richardson 2026): PDGFRA amp 재분류 51.4개월 ≈ 4.3년, Update 12 기준 grade 4 전체 37.7개월", "IDH-wt GBM과 비교하면 안 됨 — IDH-mut grade 4는 원래 예후가 더 좋음"],
  "기준치 자체가 이질적으로 치료된 과거 코호트에서 나왔다. 치료가 좋아진 지금 생존이 길어졌다고 grade를 낮추면 안 된다는 것이 원문 입장."));
C.push(...QA(8, "Mitosis 3개 / 2.4 mm²의 근거는? 왜 2.4 mm²인가?",
  "2.4 mm²는 0.24 mm² HPF 10개다. WHO CNS5가 HPF 대신 면적 기준을 요구해서다. 근거는 2개 연구: Kros 2018(CATNON + EORTC 22033, 455명, PFS만 유의 HR 1.62)과 Tran 2023(100명, ≥2.4–3 / 2.4 mm²가 중간 위험).",
  ["Brat 2026 p.2075–2076", "WHO CNS5: HPF 크기가 현미경마다 달라 mm² 기준", "Update 8 방식: hotspot에서 연속 HPF 계수"],
  "Kros는 OS 비유의(자료 미성숙), 디지털 slide, 판독자 일치도 60%. Tran은 단일 기관 100명. 원문도 'WHO 기준으로 완전히 승인하려면 추가 연구 필요'라고 명시. 이 증례의 grade 4 판정은 mitosis가 아니라 PDGFRA amp에 근거하므로 이 한계와 무관."));
C.push(...QA(9, "Methylation A_IDH_HG만 나오면 grade 4로 해도 되나?",
  "안 된다. Update 12는 G-CIMP-low·A_IDH_HG를 '최소 grade 3'으로 보고, grade 4는 다른 grade 4 소견(MVP, necrosis, CDKN2A/B HD, PDGFRA amp)이 있을 때만이다. 또 classifier score >0.90일 때만 grading의 확정 근거로 쓴다.",
  ["Brat 2026 DNA methylation Summary(p.2074)", "Tralongo 2026: 두 용어는 서로 바꿔 쓰면 안 됨"], "A_IDH와 A_IDH_HG의 정의 자체가 새 자료로 바뀔 수 있다(원문)."));

C.push(...QA(25, "왜 CDK4 amp는 grade 3이고 PDGFRA amp는 grade 4인가? 둘 다 증폭인데.",
  "근거의 일관성 차이다. CDK4·CCND2 amp도 일부 연구에서 grade 4 수준이었지만 Update 12는 'CDKN2A/B HD만큼 일관되고 강하지 않다'고 판단했다. PDGFRA amp은 CDKN2A/B HD를 포함한 CATNON에서 다변량 독립 효과가 확인됐다.",
  ["Brat 2026 Conclusions(p.2083)", "Aoki: RB pathway 묶음은 grade 4 수준이나 개별로는 비유의", "Cimino: CDK4 amp 단독 median 63개월(intermediate)"],
  "원문도 왜 RB pathway 개별 변화가 CDKN2A/B HD보다 약한지는 '명확하지 않다'고 했다(가설: p16·p15·p14ARF 동시 소실)."));
C.push(...QA(26, "CDKN2A/B hemizygous deletion도 grade를 올리나?",
  "아니다. 2개 연구는 생존 단축, 2개는 무관으로 결과가 상반돼 grading 기준에서 제외됐다. CDKN2A point mutation은 grade 3 기준이다.",
  ["Brat 2026 p.2072", "Oligodendroglioma에서도 hemizygous deletion은 예후와 무관"], "방법·threshold 차이가 원인일 수 있어 향후 바뀔 수 있다."));
C.push(H2("11-C. 검출과 기술"));
C.push(...QA(10, "PDGFRA amplification의 정의는? 몇 copy부터인가?",
  "Update 12 원문에는 copy 수 정의가 없다. 검사법별 기관 기준을 따른다. 참고로 PDGFRA IHC 논문은 FISH >8 copies in ≥10% 종양세포, CNV log ratio >0.8을 썼다.",
  ["Brat 2026 원문 전체 검토 — 증폭 threshold 미기재", "Tauziède-Espariat, Brain Pathol 2026", "CATNON methylation 분석: ±0.60 log2"],
  "정의가 통일되지 않은 것은 실제 한계다. Gain(≤6 copies)은 amplification이 아니다."));
C.push(...QA(11, "종양 일부에만 증폭(clonal amplification)이 있어도 grade 4인가?",
  "명확한 기준은 찾지 못했다. FISH의 실무 정의(≥10% 세포)를 만족하면 amplification으로 보고된다.",
  ["IHC 논문: clonal amp 10례 중 FISH로 모두 검출, CNV로는 4례만 검출", "Tralongo: 단일 block은 모든 subclone을 대표하지 못함"],
  "Clonal amp의 예후 의미는 근거가 부족하다."));
C.push(...QA(12, "PDGFRA IHC만으로 진단해도 되나?",
  "안 된다. Screening 도구다. ≥10% 강양성은 민감도 100%, 특이도 89%로, gain이나 mutation에서도 양성이 나올 수 있어 FISH 등으로 확진해야 한다.",
  ["Tauziède-Espariat, Brain Pathol 2026: 강양성 >10% → 강력 의심, 국소 → 유전 검사 필수"], null));
C.push(...QA(13, "모든 grade 2/3에 NGS·methylation을 해야 하나?",
  "Update 12는 '모든 경우 필요하지는 않다'고 한다. 분자 검사가 특히 필요한 경우는 ① 조직이 grade 2/3 경계 ② 영상에서 조영증강 또는 비교적 빠른 진행 ③ 특이한 임상 양상. 반대로 천천히 자라는 T2 고신호 비조영 병변 + 저등급 생검이면 분자 검사 없이도 grade 2를 꽤 높은 확신으로 진단할 수 있다.",
  ["Brat 2026 Figure 1 설명, Conclusions", "Ann Pathol 2026: 실제 진료에서 24% grade 변경"],
  "이 증례는 조영증강이 뚜렷하지 않았지만 ADC·rCBV 이상이 있었다. 원문 기준(조영증강·빠른 진행)에 정확히 들어가지는 않으므로, '원문 취지(영상상 공격성)를 확장 적용한 것'이라고 설명하는 것이 정직하다 [제안]."));

C.push(H2("11-D. 영상"));
C.push(...QA(14, "영상으로 PDGFRA amp을 예측할 수 있나?",
  "현재로서는 아니다. 특이 소견이 확립되지 않았다. 영상은 aggressive focus를 찾아 분자 검사를 유도하는 역할이다.",
  ["PubMed 검색에서 IDH-mut astrocytoma 대상 PDGFRA 영상 연구를 찾지 못함", "Kurokawa 2026: 영상은 targeted sampling과 분자 검사 우선순위 결정에 핵심"], null));
C.push(...QA(15, "Grade 4인데 조영증강이나 괴사가 없는 게 말이 되나?",
  "된다. 분자적으로만 grade 4인 종양은 괴사·MVP가 없고 조영증강도 없을 수 있다. Update 12도 MRI 조영증강의 예후 의미는 '논쟁 중'이라고 했다.",
  ["Brat 2026 p.2082: 조영증강의 예후 의미는 disputed", "Ozono 2026: non-enhancing AST 중 grade 4 존재", "Imai 2025: 비조영 glioma에서 rCBV로 grade 4 구분(AUC 0.718)", "Alfattal 2026: grade 2의 50%가 조영증강"],
  "Jo 2025에서 A4의 96%가 조영증강을 보였으므로 비조영 grade 4는 드문 편이다."));
C.push(...QA(16, "ADC·rCBV 이상만으로 grade를 올리자고 한 것은 과잉 아닌가?",
  "영상으로 grade를 바꾼 것이 아니라, 불일치를 근거로 분자 재검토를 요청했고 grade는 병리 소견(PDGFRA amp)으로 바뀌었다.",
  ["Kurokawa 2026: diffusion restriction·perfusion 상승 부위가 biopsy target이고 분자 검사 우선순위를 정하는 근거"],
  "ADC·rCBV 단독으로 IDH-mut astrocytoma grade를 정하는 검증된 cut-off는 없다."));
C.push(...QA(17, "T2-FLAIR mismatch가 있으면 저등급 아닌가?",
  "Mismatch는 IDH-mut astrocytoma를 시사할 뿐 grade를 보장하지 않는다. Partial mismatch는 grade 4에서도 보인다. 다만 CDKN2A HD 증례는 mismatch 음성이었다는 보고가 있다.",
  ["Kurokawa 2026: grade 4 IDH-mut의 26.4%가 partial mismatch", "Ozono 2026: CDKN2A HD 5례 모두 mismatch 음성"], null));

C.push(...QA(27, "Update 12는 영상에 대해 뭐라고 하나?",
  "세 가지다. ① 조영증강이나 비교적 빠른 진행이 보이면 분자 검사를 하라. ② 천천히 자라는 T2 고신호 비조영 병변 + 저등급 생검이면 분자 검사 없이도 grade 2를 꽤 확신할 수 있다. ③ 수술 전·후 종양 부피와 MRI 성장 속도가 가장 일관된 예후 인자이고, 조영증강의 예후 의미는 논쟁 중이다. Oligodendroglioma에서는 mitosis만으로 grade 3이고 조영증강이 없으면 생존이 길 수 있다.",
  ["Brat 2026 Figure 1 설명, Conclusions, Multidisciplinary Considerations, Oligodendroglioma 절"], null));
C.push(H2("11-E. 치료"));
C.push(...QA(18, "CATNON에서 concurrent TMZ는 이득이 없었는데 왜 CCRT를 했나?",
  "CATNON은 grade 3(anaplastic) 환자 대상이었다. 분자적 grade 4 IDH-mut astrocytoma에 대한 무작위 근거는 없고, ASCO-SNO는 grade 4 IDH-mut에서 grade 3 권고와 GBM 권고 중 어느 쪽이든 허용한다. CCRT는 허용된 선택지 중 하나다.",
  ["CATNON final: IDH-mut에서 concurrent HR 0.81(비유의), adjuvant HR 0.54", "Update 12: CATNON에서 RT 단독 5.1년 vs RT + adjuvant TMZ >12년", "Mohile 2022(ASCO-SNO), 2025 rapid update(Blakeley)"],
  "IDH-mut에서 concurrent TMZ의 이득을 보여준 직접 근거는 없다. PDGFRA amp은 TMZ 효과를 예측하지 않는다. 치료 결정은 MDT의 판단이다."));
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
C.push(...QA(22, "63세면 IDH-wt GBM일 가능성은 배제했나? 나이 자체가 예후 인자 아닌가?",
  "IDH-mutant로 진단됐다. Update 12에 따르면 40세 cut-off는 더 이상 유효하지 않지만, grade 3 IDH-mut astrocytoma에서 60세 이상은 예후가 나쁘다. 고령이라 IDH 확인 방법을 차트에서 확인해 두는 것이 좋다.",
  ["Brat 2026 p.2082", "Virata 2026: IDH-mut AST ≥55세 8.6%", "Alfattal 2026: ≥55세 IDH-mut AST 22명, IDH1 R132H 95%"],
  "검사 방법(IHC vs sequencing)은 증례 정보에서 확인 필요 [원문 확인 권장]."));
C.push(...QA(23, "PDGFRA amp이 나온 검체가 영상의 aggressive 부위와 같은 곳인가?",
  "차트·병리 보고서에서 확인해야 한다. 같은 부위라면 영상-병리 상관의 좋은 예이고, 다르다면 종양 내 이질성을 보여준다.",
  ["Tralongo 2026: sampling bias, 지역별 차이", "IHC 논문: clonal amp은 증폭 부위에만 IHC 양성"],
  "[원문 확인 권장]"));
C.push(...QA(24, "CDKN2A/B 상태는? HD가 같이 있었다면 PDGFRA는 의미가 없지 않나?",
  "CDKN2A/B HD가 있었다면 그것만으로 이미 grade 4다. 이 증례의 의미는 PDGFRA amp이 grade 4의 근거였다는 것이므로, CDKN2A/B 상태를 확인해 발표해야 한다.",
  ["Yang 2020: PDGFRA amp과 CDKN2A HD는 상호 배타적인 경향"],
  "[원문 확인 권장] 차트 확인"));

// =============== 12. Evidence level ===============
C.push(H1("12. 주장별 근거 수준"));
C.push(...T(["주장", "근거 유형", "강도 [제안]"], [
  ["CDKN2A/B HD → grade 4", "다수 후향 코호트 + WHO CNS5 공식 기준", "높음 (공식)"],
  ["PDGFRA amp → grade 4", "RCT 내 분자 분석(CATNON HR 3.51, CDKN2A/B 독립) + 후향 코호트 다수 + 전문가 합의 + 발표 후 검증", "중간–높음 (cIMPACT-NOW, 일부 이견)"],
  ["Mitosis ≥3 / 2.4 mm²", "2개 연구(Kros 455명 PFS만, Tran 100명) + 합의", "낮음–중간 (원문: 추가 연구 필요)"],
  ["Grade 3 분자 marker (Table 2)", "후향 코호트 다수, 일부 RCT 분석(PIK3는 CATNON 장기추적서 소실)", "중간"],
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
  ["**Update 12**", "**PDGFRA amp → grade 4**(CATNON 3.1 vs 9.3년, HR 3.51), ≥3 mitoses/2.4 mm² 또는 Table 2 분자 변화 → grade 3", "Brat 2026; Park YW KJR 2026 Fig. 2"],
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
  "Yang RR, et al. IDH mutant lower grade astrocytomas can be stratified for risk by CDKN2A, CDK4 and PDGFRA copy number alterations. Brain Pathol 2020;30:541–553. doi:10.1111/bpa.12801",
  "Ghosh HS, et al. Canonical amplifications and CDKN2A/B loss refine IDH1/2-mutant astrocytoma prognosis. Neuro-Oncol 2025;27:993–1003. doi:10.1093/neuonc/noae258 (Update 12 인용)",
  "Richardson TE, Walker JM. The prognostic significance of RB and PI3K pathway alterations in IDH-mutant grade II/III astrocytomas. J Neuropathol Exp Neurol 2020;79:1019–1023. doi:10.1093/jnen/nlaa083 (Update 12 인용)",
  "Shirahata M, et al. Novel, improved grading system(s) for IDH-mutant astrocytic gliomas. Acta Neuropathol 2018;136:153–166. doi:10.1007/s00401-018-1849-4 (Update 12 인용)",
  "Kros JM, et al. Mitotic count is prognostic in IDH mutant astrocytoma without homozygous deletion of CDKN2A/B. Acta Neuropathol Commun 2018;6:95. (Update 12 인용)",
  "Tran S, et al. A threshold for mitotic activity and post-surgical residual volume defines distinct prognostic groups for astrocytoma IDH-mutant. Neuropathol Appl Neurobiol 2023;49:e12928. doi:10.1111/nan.12928 (Update 12 인용)",
  "Weller M, et al. Improved prognostic stratification of patients with IDH-mutant astrocytoma. Acta Neuropathol 2024;147:11. doi:10.1007/s00401-023-02662-1 (Update 12 인용)",
  "Figarella-Branger D, et al. Reappraisal of prognostic factors in CNS WHO grade 3 oligodendrogliomas: the French POLA cohort. Neuro-Oncol 2025;27:755–766. doi:10.1093/neuonc/noae221 (Update 12 인용)",
  "van den Bent MJ, et al. The biological significance of tumor grade, age, enhancement, and extent of resection in IDH-mutant gliomas. Neuro-Oncol 2024;26:1805–1822. doi:10.1093/neuonc/noae107 (Update 12 인용)",
  "Blakeley J, et al. ASCO-SNO guideline rapid recommendation update. J Clin Oncol 2025;43:2129–2133. doi:10.1200/JCO-25-00250 (Update 12 인용)",
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
