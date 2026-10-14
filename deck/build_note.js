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

// ---------- Title ----------
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 600, after: 120 },
  children: [new TextRun({ text: "cIMPACT-NOW Update 12", bold: true, size: 48, color: NAVY, font: FONT })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
  children: [new TextRun({ text: "IDH-mutant glioma grading의 개정 — PDGFRA amplification과 grade 4", size: 28, font: FONT })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 },
  children: [new TextRun({ text: "WHO CNS5 이후 cIMPACT-NOW update 8–12 정리 · 영상의학적 의미 · 증례 적용 — 심층 학습 노트", bold: true, size: 24, font: FONT })] }));
C.push(P("Brat DJ, Aldape K, French PJ, Louis DN, Nasrallah MP, Reifenberger G, Reuss DE, Touat M, van den Bent MJ, Wesseling P, Figarella-Branger D", { align: AlignmentType.CENTER }));
C.push(P("*cIMPACT-NOW Update 12: Refining pathology-based risk stratification and grading for IDH-mutant gliomas*", { align: AlignmentType.CENTER }));
C.push(P("Neuro-Oncology 2026;28(9):2069–2089 · doi:10.1093/neuonc/noag108 · PMID 42152226 · PMC13550680", { align: AlignmentType.CENTER, p: { spacing: { after: 360 } } }));

C.push(...BOX("이 노트의 출처와 한계", [
  "주 출처: Update 12의 **PubMed 초록**. PMC 전문은 이 작업 환경에서 본문이 열람되지 않아 **원문 Table·Figure의 세부 문구(각주, 정의, 인용 연구별 수치)는 확인하지 못함** → 해당 부분은 **[원문 확인 권장]** 으로 표시.",
  "Grading 알고리즘 그림은 발표 슬라이드에 사용한 Fig. 2(cIMPACT-NOW update 12 schematic)를 기준으로 정리. 그 Figure가 실린 2차 리뷰 논문 자체는 검색으로 특정하지 못함.",
  "PDGFRA amplification의 근거는 **전문을 직접 읽은 원저**(Update 5 전문, Virata 2026 전문, Richardson 2026 전문, CATNON methylation 2021 전문, PDGFRA IHC 2026 전문)와 초록(Yang 2019, CATNON final 2026, Tauziède-Espariat 2026)으로 구성.",
  "다른 문헌·일반 지식에서 온 내용은 **[보충]**, 필자의 실무 제안은 **[제안]** 으로 구분.",
  "PMC 텍스트 추출 과정에서 이탤릭 유전자명이 빠지는 문제가 있어, Virata 2026의 유전자별 수치는 **본문에 나열된 유전자 순서로 대응**시킴 → **[원문 확인 권장]**.",
], YELLOW));

C.push(new Paragraph({ children: [new PageBreak()] }));
C.push(new TableOfContents("목차", { hyperlink: true, headingStyleRange: "1-2" }));
C.push(new Paragraph({ children: [new PageBreak()] }));

// ---------- 1. Summary ----------
C.push(H1("1. 한눈에 보는 핵심 요약"));
C.push(...BOX("Take-home messages", [
  "WHO CNS5(2021): **Astrocytoma, IDH-mutant = 단일 type, CNS WHO grade 2–4**. Grade 4 = MVP 또는 necrosis 또는 **CDKN2A/B homozygous deletion(HD)**.",
  "Update 12(2026): **PDGFRA amplification을 grade 4 criterion으로 추가** — \"highly aggressive clinical behavior, consistent with CNS WHO grade 4\".",
  "Grade 2/3 경계: **mitotic count ≥3 / 2.4 mm²** 를 grade 3 기준으로 제안(이전에는 수치 기준이 없었음 — Update 5에서 \"significant mitotic activity\", 관행상 전체 검체 ≥2 mitoses).",
  "Grade 3을 시사하는 **분자 변화**: PIK3CA/PIK3R1 mutation, EGFR alteration, MYCN amp, RB pathway(CDK4·CDK6·CCND2 amp, RB1 HD/mutation, CDKN2A mutation), DNA methylation **G-CIMP-low / A_IDH_HG**.",
  "Oligodendroglioma: grading을 바꿀 새 근거 없음. 단 mitosis만 높고 necrosis·MVP·CDKN2A/B HD·**MRI 조영증강이 없는** grade 3은 생존이 길 수 있음.",
  "PDGFRA amp → grade 4의 근거: Validation(n=1,140)에서 PDGFRA amp으로 grade 4가 된 군의 median OS **51.4개월** vs grade 3 유지군 **126.2개월**(HR 1.9).",
  "실제 진료: 2년 반 동안 molecular 검사가 필요했던 50례 중 **24%(12/50)** 에서 grade 변경(grade 3→4: 3례).",
  "영상: PDGFRA amp에 **특이적인 영상 소견은 확립되지 않음**. ADC↓/rCBV↑ 같은 aggressive focus는 targeted sampling과 molecular 검사 요청의 근거.",
  "임상적 의미: grade는 **치료 강도(RT±TMZ, vorasidenib 적응증), 임상시험 적격성, 예후 상담**을 결정 → 이 증례(63/F)는 grade 3 → 4로 바뀌며 CCRT 시행.",
]));
C.push(SRC("Brat 2026 (PubMed abstract); Brat 2020 Update 5 (PMC8443062); Richardson 2026 (PMC13391642); Tauziède-Espariat, Ann Pathol 2026 (abstract)"));

// ---------- 2. Background ----------
C.push(H1("2. 배경 — WHO CNS5와 cIMPACT-NOW"));
C.push(H2("2-1. cIMPACT-NOW란"));
C.push(...B([
  "**Consortium to Inform Molecular and Practical Approaches to CNS Tumor Taxonomy – Not Official WHO**",
  "WHO 개정 사이(약 5년 이상)의 공백 동안 새 분자 지식을 진단에 반영하기 위한 국제 전문가 그룹",
  "  신경병리, 신경영상, 분자유전, 신경종양 등 다학제 구성",
  "  권고는 **WHO 공식 정의가 아님** → 다음 WHO(CNS6) 전까지의 진단 가이드",
  "Update 1–7(2018–2020)은 **WHO CNS5(2021)에 대부분 반영**됨",
]));
C.push(SRC("Kurokawa, Jpn J Radiol 2026;44:1534 (PMC13526116); Wesseling, Brain Pathol 2025 (Update 11 abstract)"));

C.push(H2("2-2. Update 1–7 (WHO CNS5 이전) 요약"));
C.push(...T(["Update", "연도·저널", "주제", "WHO CNS5와의 관계"], [
  ["1", "2018 Acta Neuropathol", "NOS / NEC 용어 정의", "NOS·NEC 진단 용어로 채택"],
  ["2", "2018 Acta Neuropathol", "Diffuse midline glioma H3 K27M, IDH-mut astrocytoma 진단 명확화", "DMG H3 K27-altered 정의의 기반"],
  ["3", "2018 Acta Neuropathol", "IDH-wt diffuse astrocytic glioma with molecular features of GBM (EGFR amp, +7/−10, TERTp)", "\"molecular GBM\" 기준으로 채택"],
  ["4", "2019 Acta Neuropathol", "MYB/MYBL1, FGFR1, BRAF V600E diffuse glioma", "Pediatric-type low-grade glioma 분류의 기반"],
  ["5", "2020 Acta Neuropathol", "**IDH-mut astrocytoma grading·용어**", "**CDKN2A/B HD → grade 4**, 'GBM, IDH-mutant' 폐지, 아라비아 숫자 grade"],
  ["6", "2020 Brain Pathol", "Utrecht meeting: 새 entity·진단 원칙", "CNS5의 구조·명명 원칙"],
  ["7", "2020 Brain Pathol", "Ependymal tumor 분자 분류", "부위 + 분자 기반 ependymoma 분류"],
], [1.0, 1.6, 3.1, 2.9], { highlight: (r) => r[0] === "5" }));
C.push(SRC("PubMed metadata/abstract — Louis 2018 (10.1007/s00401-018-1808-0), Louis 2018 (10.1007/s00401-018-1826-y), Brat 2018 (10.1007/s00401-018-1913-0), Ellison 2019 (10.1007/s00401-019-01987-0), Brat 2020 (10.1007/s00401-020-02127-9), Louis 2020 (10.1111/bpa.12832), Ellison 2020 (10.1111/bpa.12866). 'CNS5와의 관계' 열은 [보충]."));

C.push(H2("2-3. WHO CNS5의 adult-type diffuse glioma 체계"));
C.push(...B([
  "3개 type으로 재편",
  "  **Astrocytoma, IDH-mutant** (AST): CNS WHO grade 2, 3, 4",
  "  **Oligodendroglioma, IDH-mutant and 1p/19q-codeleted** (ODG): grade 2, 3",
  "  **Glioblastoma, IDH-wildtype** (GBM): grade 4",
  "폐지된 용어: \"anaplastic astrocytoma\", \"glioblastoma, IDH-mutant\"",
  "Grade는 **type 안에서** 매김 (아라비아 숫자, 'CNS WHO grade')",
  "IDH-wt diffuse astrocytic glioma: 조직학적 grade 4 소견이 없어도 **EGFR amp, TERTp mutation, +7/−10** 중 하나면 GBM (\"molecular GBM\")",
]));
C.push(SRC("Kurokawa, Jpn J Radiol 2026 (PMC13526116) Introduction"));

C.push(H2("2-4. Update 5 (2020) — Update 12의 출발점"));
C.push(...B([
  "**Grade 2**: significant mitotic activity, anaplasia, MVP, necrosis, CDKN2A/B HD 모두 없음 → median OS >10년",
  "**Grade 3**: mitotic activity·anaplasia 상승, MVP·necrosis·CDKN2A/B HD 없음",
  "  **mitosis cut-off는 검증된 기준이 없다**고 명시 → \"significant\" mitotic activity 유지",
  "  관행: 전체 검체에서 **≥2 mitoses**, 아주 작은 생검에서는 1개, 큰 검체는 더 많이 필요",
  "**Grade 4**: MVP 또는 necrosis 또는 **CDKN2A/B HD**",
  "  CDKN2A/B HD 빈도(문헌): grade II 0–12%, III 6–20%, IV 16–34%",
  "  검출: FISH, qPCR, MLPA, microarray, NGS. **p16 IHC는 deletion과 상관 불량**",
  "**PDGFRA amp에 대한 당시 평가**: \"여러 연구에서 나쁜 예후 marker로 확인, **추가 근거가 쌓이면 grading criterion으로 포함될 가능성**\" → 6년 뒤 Update 12에서 grade 4 criterion으로 채택",
  "기타(당시 근거 부족): RB pathway(CDK4 amp, RB1), PIK3R1/PIK3CA mutation, MYCN amp, genomic instability, G-CIMP-low",
]));
C.push(SRC("Brat DJ et al. cIMPACT-NOW Update 5. Acta Neuropathol 2020;139:603–608 (PMC8443062 전문)"));

// ---------- 3. Updates 8-11 ----------
C.push(H1("3. WHO CNS5 이후 cIMPACT-NOW Update 8–11"));
C.push(...T(["Update", "발표", "대상", "핵심 권고"], [
  ["8", "Sahm F. Neuro-Oncol 2025;27:319", "Meningioma grading", "Brain invasion 판정 표준화. Grade 1 형태라도 **1p deletion + 22q deletion and/or NF2 oncogenic variant → grade 2** 제안. 진행성 분자 검사가 권장되는 상황 제시"],
  ["9", "Aldape K. Neuro-Oncol Adv 2025;7:vdae228", "DNA methylation profiling", "Classifier는 **integrated diagnosis의 한 도구**(단독 결과로 진단 X). Layered report 예시. 향후 array 간 backward compatibility 강조"],
  ["10", "Hawkins C. Brain Pathol 2025;35:e70018", "새 tumor type 정의 기준", "Clinical phenotype(예후·연령·위치·**전형적 영상 소견**) 필수 + methylation cluster, driver alteration, 조직 소견 + ≥2개 독립 보고(가급적 각 >10명). 미충족 시 **NEC**"],
  ["11", "Wesseling P. Brain Pathol 2025;36:e70035", "IDH/H3-wt HGG, PF ependymoma", "TERTp-only 저등급 IDH-wt → grade 4 신중. **<40세: EGFR amp, +7/−10 단독으로 GBM 진단 금지**. PDGFRA·EGFR alteration, MYCN amp는 **<25세**에서만 pedHGG H3/IDH-wt의 key feature. PF EPN: EZHIP IHC"],
], [1.0, 2.0, 1.7, 4.4]));
C.push(SRC("PubMed abstract/full text — Update 8 (10.1093/neuonc/noae170), Update 9 (10.1093/noajnl/vdae228), Update 10 (PMC12488253 전문), Update 11 (10.1111/bpa.70035)"));

C.push(H2("3-1. Update 8 — Meningioma"));
C.push(...B([
  "WHO CNS5에서 처음 들어온 meningioma 분자 grading: **TERTp mutation, CDKN2A/B HD → grade 3**",
  "Update 8의 추가: **grade 1 형태 + 1p deletion과 22q deletion 그리고/또는 NF2 oncogenic variant → grade 2** 제안",
  "Brain invasion 평가의 형태학적 해석 표준화",
  "[보충] 영상의학적 의미: 형태상 grade 1로 예상되는 수막종도 분자 결과로 grade가 올라갈 수 있음 → 추적 간격·RT 결정에 영향",
]));
C.push(SRC("Sahm F, Neuro-Oncol 2025 (PubMed abstract)"));

C.push(H2("3-2. Update 9 — DNA methylation profiling"));
C.push(...B([
  "Methylation classifier는 **다른 진단 도구와 통합해서** 쓰는 도구",
  "Layered diagnosis(WHO·ICCR 형식)에 methylation 결과를 넣는 예시 제시",
  "결과 해석 시 methylation class뿐 아니라 **copy-number profile과 보조 biomarker**까지 볼 것",
  "→ Update 12에서 **CNV(PDGFRA amp, CDKN2A/B HD)를 methylation array에서 얻을 수 있는 것**과 연결됨",
]));
C.push(SRC("Aldape K, Neuro-Oncol Adv 2025 (PubMed abstract); Kurokawa 2026 (PMC13526116)"));

C.push(H2("3-3. Update 10 — 새 tumor type의 정의"));
C.push(...B([
  "필수: **clinical phenotype** — 전형적 예후/치료 반응, 호발 연령, 호발 위치, **전형적 영상 소견**",
  "가급적 모두 충족: ① 대규모 profiling(methylation 등)에서 고유한 cluster(위치 의존 signature만으로는 불충분) ② driver alteration ③ 연관된 조직·IHC 소견",
  "형식 요건: **≥2개 독립 publication**(가급적 각 >10명) + cIMPACT-NOW 위원회 승인",
  "Field test: 13개 후보 중 기준 충족은 일부. 기준 미충족 tumor는 형태 + NEC + 분자 소견으로 layered 보고",
  "영상의학적 의미: **영상 소견이 tumor type 정의의 구성 요소**로 명시됨",
]));
C.push(SRC("Hawkins C, Brain Pathol 2025 (PMC12488253 전문); Kurokawa 2026"));

C.push(H2("3-4. Update 11 — IDH/H3-wildtype HGG와 PF ependymal tumor"));
C.push(...B([
  "GBM, IDH-wt vs **diffuse pediatric-type HGG, H3-wt and IDH-wt** 구분",
  "  TERTp mutation **단독**의 조직학적 저등급 IDH-wt 종양 → grade 4(GBM) 지정 신중",
  "  **<40세**: EGFR amp 또는 +7/−10 **단독**으로 GBM 진단 금지",
  "  성인(특히 <40세)에서도 pedHGG H3/IDH-wt 감별 고려, 확진은 methylation profiling",
  "  **PDGFRA alteration**, EGFR alteration, MYCN amp → **<25세**에서만 pedHGG의 key molecular feature",
  "  >55세 non-midline 종양: H3 K27 검사 생략 가능, ATRX 유지 + OLIG2 양성은 H3 G34-wt 지지",
  "PF ependymoma: nuclear **EZHIP** IHC → PFA 지지 / H3 K27me3 유지 + EZHIP 음성 + methylation 미시행 → PF ependymoma NOS",
  "주의: PDGFRA는 Update 11(IDH-wt, 젊은 환자의 진단 marker)과 Update 12(IDH-mut astrocytoma의 grading marker)에서 **서로 다른 맥락**으로 등장",
]));
C.push(SRC("Wesseling P, Brain Pathol 2025 (PubMed abstract); Kurokawa 2026 (PMC13526116)"));

// ---------- 4. Update 12 ----------
C.push(H1("4. cIMPACT-NOW Update 12 상세"));
C.push(H2("4-1. 목적"));
C.push(...B([
  "IDH-mutant glioma의 **pathology-based risk stratification·grading 개선**, 특히 **grade 2와 3의 경계**",
  "새로 기술된 IDH-mutant glioma subgroup 평가",
  "치료 결정의 다학제적 측면 고려 — \"치료 결정이 risk assessment에 의존하기 때문\"",
]));
C.push(SRC("Brat 2026 (PubMed abstract)"));

C.push(H2("4-2. Grading 알고리즘 (Fig. 2 기준)"));
C.push(...T(["Tumor", "Grade", "기준"], [
  ["Astrocytoma, IDH-mut", "**4**", "MVP, necrosis, **CDKN2A/B HD** 그리고/또는 **PDGFRA amp** (하나라도 있으면)"],
  ["Astrocytoma, IDH-mut", "**3**", "Grade 4 기준 없음 + **mitotic activity ≥3 / 2.4 mm²**"],
  ["", "(3 시사)", "분자 변화: CDK4/CDK6/CCND2 amp, CDKN2A mutation, RB1 HD or mutation · PIK3CA/PIK3R1 mutation · EGFR alteration, MYCN amp · methylation **A_IDH_HG** 또는 **G-CIMP-low** (Fig.에는 IDH_A_HG로 표기)"],
  ["Astrocytoma, IDH-mut", "**2**", "Grade 4 기준 없음 + mitosis <3 / 2.4 mm² (+ grade 3 시사 분자 변화 없음)"],
  ["Oligodendroglioma", "**3**", "MVP, necrosis 그리고/또는 **mitotic activity ≥6 / 2.4 mm²**"],
  ["Oligodendroglioma", "**2**", "위 기준 없음"],
], [2.2, 1.0, 6.0], { highlight: (r) => r[1] === "**4**" }));
C.push(...B([
  "Grade 3 시사 분자 변화는 Fig. 2에서 \"that **may** suggest WHO grade 3\"로 표현 — 형태학적 grade 2 종양에 이 변화가 있으면 grade 3을 고려 [원문 확인 권장: 정확한 적용 문구]",
  "Methylation(G-CIMP-low, A_IDH_HG): \"다른 grading 기준에 따라 **grade 3 또는 4**와 부합\" (초록)",
  "Mitosis 기준은 \"**could be considered**\" — 제안 수준의 표현",
]));
C.push(SRC("Brat 2026 (PubMed abstract); Fig. 2 schematic(발표용 그림); Richardson 2026 (PMC13391642) Supplemental Table 기술"));

C.push(H2("4-3. Grade 4: PDGFRA amplification"));
C.push(...B([
  "초록 원문: \"PDGFRA amplification was associated with **highly aggressive clinical behavior** of IDH-mutant astrocytomas, **consistent with CNS WHO grade 4**.\"",
  "→ MVP, necrosis, CDKN2A/B HD와 **동등한 grade 4 criterion**",
  "판정 원칙(알고리즘): 조직학적 grade(2 또는 3)와 무관하게 PDGFRA amp가 있으면 grade 4",
  "근거 연구·수치는 5장 참조",
]));
C.push(SRC("Brat 2026 (PubMed abstract)"));

C.push(H2("4-4. Grade 3: mitosis 기준과 분자 marker"));
C.push(...B([
  "**Mitotic count ≥3 / 2.4 mm²** → grade 3 기준으로 제안",
  "  2.4 mm²: 고배율 시야 약 10개에 해당하는 면적 [보충 — 현미경 시야 면적에 따라 다름]",
  "  이전(Update 5, CNS5): 수치 기준 없음, 관행상 전체 검체 ≥2 mitoses",
  "  PDGFRA IHC 논문(2026)도 \"grade 2/3 cut-off가 낮다(2.4 mm²당 3개)\"고 지적 → 생검 sampling에 민감",
  "중간 위험(grade 3에 가장 부합) 분자 변화",
  "  **PI3K**: PIK3CA, PIK3R1 mutation",
  "  **EGFR** alteration",
  "  **MYCN** amplification",
  "  **RB pathway**: CDK4, CDK6, CCND2 amp / RB1 HD 또는 mutation / CDKN2A **mutation**(HD가 아닌 point mutation)",
  "  **DNA methylation**: G-CIMP-low 또는 A_IDH_HG",
]));
C.push(SRC("Brat 2026 (PubMed abstract); Tauziède-Espariat, Brain Pathol 2026 (PMC13535536)"));

C.push(H2("4-5. Oligodendroglioma"));
C.push(...B([
  "Grading을 개선할 새 feature는 **찾지 못함**",
  "단, **mitosis가 높아 grade 3이지만** necrosis, MVP, CDKN2A/B HD, **MRI contrast enhancement가 없는** 종양은 생존이 길 수 있음",
  "→ 영상의학적 의미: oligodendroglioma grade 3에서 **조영증강 유무가 예후 정보**로 언급됨",
  "Fig. 2 기준: MVP, necrosis 그리고/또는 mitosis ≥6 / 2.4 mm² → grade 3",
]));
C.push(SRC("Brat 2026 (PubMed abstract)"));

C.push(H2("4-6. Update 5 → Update 12 변화 비교"));
C.push(...T(["항목", "Update 5 / WHO CNS5", "Update 12 (2026)"], [
  ["Grade 4 분자 기준", "CDKN2A/B HD", "CDKN2A/B HD + **PDGFRA amp**"],
  ["Grade 2 vs 3", "\"significant\" mitotic activity (수치 없음)", "**≥3 mitoses / 2.4 mm²** 제안"],
  ["Grade 3 분자 marker", "없음 (근거 부족으로 보류)", "PI3K, EGFR, MYCN, RB pathway, G-CIMP-low / A_IDH_HG"],
  ["PDGFRA amp 평가", "\"추가 근거 시 grading criterion 가능\"", "**Grade 4 criterion**"],
  ["Oligodendroglioma", "MVP·necrosis·mitosis·CDKN2A/B HD", "개정 근거 없음, mitosis ≥6 / 2.4 mm² 명시"],
  ["필요한 검사", "CDKN2A/B 검출(FISH, MLPA, NGS 등)", "**NGS + CNV(methylation array 등)** 사실상 필요"],
], [2.0, 3.4, 3.8]));
C.push(SRC("Brat 2020 Update 5 (PMC8443062); Brat 2026 (abstract); Tauziède-Espariat, Ann Pathol 2026"));

// ---------- 5. PDGFRA ----------
C.push(H1("5. PDGFRA amplification 심화"));
C.push(H2("5-1. 기본 개념"));
C.push(...B([
  "PDGFRA: platelet-derived growth factor receptor alpha — receptor tyrosine kinase(RTK) [보충]",
  "유전자 위치: **4번 염색체** (FISH에서 chromosome 4 centromere probe와 비교) — [보충] 4q12",
  "다른 CNS 종양에서의 PDGFRA",
  "  pedHGG H3/IDH-wt, 특히 **RTK1 subgroup**에서 흔함 (진단적 의미, Update 11)",
  "  GBM IDH-wt에서도 관찰 (IHC 논문 코호트 11/45)",
  "  **PDGFRA mutation**: myxoid glioneuronal tumor의 진단 marker",
  "  Diffuse hemispheric glioma, H3 G34-mutant에서도 amplification 보고",
]));
C.push(SRC("Tauziède-Espariat, Brain Pathol 2026 (PMC13535536); Le Rhun, Neuro-Oncol 2025 (abstract); Update 11 abstract"));

C.push(H2("5-2. 근거 연구 정리"));
C.push(...T(["연구", "대상", "주요 결과"], [
  ["**Yang RR** Brain Pathol 2019", "IDH-mut grade II/III astrocytoma 160명 (FISH)", "PDGFRA amp **18.8%**, CDKN2A HD 15.0%, CDK4 amp 18.1% — **서로 배타적**. PDGFRA amp → 짧은 PFS(p=0.0003)·OS(p<0.0001). 3 risk group mean OS: 고위험(PDGFRA amp) **40.5** / 중간(CDKN2A HD·CDK4 amp) 62.9 / 저위험 71.5개월"],
  ["**Tesileanu CMS** Neuro-Oncol 2021 (CATNON)", "IDH-mut anaplastic astrocytoma 432명 (methylation array + NGS)", "다변량 독립 예후인자: 나이, MMSE, TMZ, epigenetic classifier, **PDGFRA amp**, CDKN2A/B HD, PI3K mutation, total CNV load. CNV는 methylation array에서 ±0.60 log2로 정의"],
  ["**van den Bent MJ** Lancet Oncol 2026 (CATNON final)", "IDH-mut 444명, 추적 10.9년", "PDGFRA·CDK4 amp, CDKN2A HD, CNV load → 나쁜 예후. **TMZ 효과의 predictive marker는 아님**"],
  ["**Virata MC** Neuro-Oncol 2026", "IDH-mut astrocytoma 840 + 367명 (GENIE, TCGA/GLASS)", "PDGFRA alteration(amp/mut) 6.5%. Cohort 1 median OS 35.5개월(HR 4.4) vs reference 140.7개월. **다변량에서는 독립 유의성 미달**, 단 **단독(isolated) PDGFRA alteration → HR 4.5**. 저자들은 'intermediate grade' 그룹에 배치 [원문 확인 권장]"],
  ["**Richardson TE** Acta Neuropathol 2026", "IDH-mut astrocytoma 1,140명 (GENIE, TCGA, MSKCC)", "Update 12 기준 적용: PDGFRA amp로 grade 2→4 **7례**, 3→4 **15례**. 재분류 grade 4 median OS **51.4개월** vs grade 3 유지 126.2개월(**HR 1.9**, p=0.0284), vs grade 2 유지(HR 4.0)"],
  ["**Tauziède-Espariat A** Ann Pathol 2026", "실제 진료 81례 중 molecular 검사 필요 50례", "**24%(12/50)** grade 변경: 2→3 7례, 2→4 2례, **3→4 3례**. CNV(methylation array) 기반 변경이 8례로 가장 많음"],
], [2.2, 2.3, 4.7]));
C.push(SRC("PubMed — Yang 2019 (10.1111/bpa.12801, abstract); Tesileanu 2021 (PMC8408862 전문); van den Bent 2026 (10.1016/S1470-2045(25)00614-X, abstract); Virata 2026 (PMC13437890 전문); Richardson 2026 (PMC13391642 전문); Tauziède-Espariat 2026 (10.1016/j.annpat.2026.09.001, abstract)"));

C.push(H2("5-3. Validation 연구 (Richardson 2026) 상세"));
C.push(...B([
  "코호트: AACR GENIE + TCGA + MSKCC, IDH-mut astrocytoma **1,140명**",
  "2016 → 2021 WHO: CDKN2A/B HD로 grade 2의 7.9%(40례), grade 3의 12.2%(45례)가 grade 4로 재분류",
  "2021 WHO → Update 12",
  "  PDGFRA amp: grade 2→4 **7례**(grade 2의 1.5%), grade 3→4 **15례**(grade 3의 4.6%)",
  "  Grade 3 분자 기준: grade 2→3 **77례**(grade 2의 16.6%)",
  "  2021 기준 grade 2/3 중 **12.5%(99/789)** 가 upgrade",
  "생존(median OS)",
  "  Grade 2 유지 146.0개월 / grade 2→3 upgrade 109.6개월(HR 2.6, p<0.0001)",
  "  PDGFRA amp로 grade 4 재분류 **51.4개월** vs grade 3 유지 **126.2개월** (HR 1.9, p=0.0284)",
  "  Upgrade된 군과 원래 같은 grade 군 사이 유의한 차이 없음 → 재분류가 적절함을 지지",
  "  Update 12 기준 전체: grade 2 146.0 / grade 3 120.4 / grade 4 **37.7개월**",
  "한계",
  "  Methylation 데이터는 TCGA 260례만 → methylation 기준은 대부분 반영 못함",
  "  여러 코호트·시대 혼합, **mitosis 밀도 직접 확인 불가**",
]));
C.push(SRC("Richardson TE et al. Acta Neuropathol 2026;152 (PMC13391642 전문)"));

C.push(H2("5-4. 검출 방법"));
C.push(...T(["방법", "Amplification 정의 / 특징", "비고"], [
  ["FISH", "PDGFRA locus **>8 copies in ≥10% 종양세포** (IHC 논문 기준)", "**Clonal(부분) amplification** 검출에 유리"],
  ["Methylation array CNV", "log ratio >0.8 (IHC 논문) / ±0.60 log2 (CATNON)", "Methylation class와 동시 획득. Clonal amp은 놓칠 수 있음(10례 중 4례만 CNV에서 보임)"],
  ["NGS", "Panel CNV call", "Mutation(PIK3CA 등 grade 3 marker)과 동시 평가"],
  ["SNP / CGH array", "Copy-number 분석", "Ann Pathol 2026에서 grade 3 검체에 권장 기법으로 언급"],
  ["**PDGFRA IHC**", ["Method 1: ≥50% 강양성 → 민감도 **54%**, 특이도 **99%**", "Method 2: **≥10% 강양성** → 민감도 **100%**, 특이도 **89%**"], "**Screening 도구**. Gain(≤6 copies)·mutation에서도 양성 가능 → 확진은 FISH 등. PDGFRA mutation 검출은 불가"],
], [1.8, 4.2, 3.2]));
C.push(...B([
  "제안 알고리즘(IHC 논문): **>10% 종양세포 강양성 → amplification 강력 의심**, 강양성이 국소(<10%)면 FISH 등 유전 검사 필요",
  "IHC 논문 코호트: high-grade AST IDH-mut 10례 중 **5례**에서 PDGFRA amp",
  "의미: 생검이 **조영증강 밖**에서 얻어지거나 세포 밀도가 낮으면 mitosis 기반 grading이 어려움 → IHC가 보완",
]));
C.push(SRC("Tauziède-Espariat, Brain Pathol 2026;e70142 (PMC13535536 전문); Tesileanu 2021 (PMC8408862); Tauziède-Espariat, Ann Pathol 2026 (abstract)"));

C.push(H2("5-5. 쟁점과 해석"));
C.push(...B([
  "**PDGFRA amp는 grade 4인가, intermediate인가?**",
  "  Update 12·Richardson 2026: grade 4와 부합(재분류 군 OS가 원래 grade 4와 차이 없음, p=0.1277)",
  "  Virata 2026: PDGFRA alteration을 CDK4, CCND2, PIK3R1, MYCN, EGFR와 묶어 **intermediate** 군으로 제시. 단, 이들이 **단독일 때 CDKN2A HD 단독과 OS 차이 없음** → 2-tier(저/고) grading도 가능하다고 언급",
  "  차이 원인 [보충]: Virata는 amplification + mutation을 합친 'alteration'으로 분석, 코호트·CNV call 방식 차이",
  "**예후(prognostic) vs 예측(predictive)**",
  "  CATNON final: PDGFRA amp는 나쁜 예후와 연관되지만 **TMZ 효과를 예측하지는 않음**",
  "  → grade 상향의 의미는 \"이 약이 듣는다\"가 아니라 \"**공격적 질환이므로 치료 강도를 높인다**\"",
  "**동반 변화**: PDGFRA amp, CDKN2A HD, CDK4 amp는 서로 배타적인 경향(Yang 2019) — 같은 'RTK/RB pathway' 축의 다른 경로 [보충]",
  "**분자 검사 접근성**: Update 12 적용에는 NGS + CNV가 필요 → 기관별 차이, 외부 의뢰 필요성(Ann Pathol 2026 결론)",
]));
C.push(SRC("Virata 2026 (PMC13437890); Richardson 2026 (PMC13391642); van den Bent 2026 (abstract); Yang 2019 (abstract); Tauziède-Espariat, Ann Pathol 2026 (abstract)"));

// ---------- 6. Imaging ----------
C.push(H1("6. 영상 소견"));
C.push(H2("6-1. IDH-mutant astrocytoma의 기본 영상 소견"));
C.push(...B([
  "**T2-FLAIR mismatch sign**",
  "  정의: T2에서 (거의) 전체 고신호 + FLAIR에서 (거의) 전체 저신호, **말초 고신호 rim**",
  "  Meta-analysis: pooled **민감도 0.42, 특이도 0.99** (IDH-mut, 1p/19q non-codeleted)",
  "  병리적 배경: microcystic change, myxoid matrix",
  "  **Partial mismatch**: grade 4 astrocytic tumor 중 IDH-mut 26.4%(32/121) vs IDH-wt 0.4%(8/2044) → 특이도 99.6%",
  "  Redefined mismatch(partial 포함, subcortical-only 제외): AST vs ODG 민감도 87.5%, 특이도 80.0%",
  "  함정: **낭성 변화·괴사를 mismatch로 판정 금지** / DMG H3 K27-altered, MYB/MYBL1-altered, angiocentric glioma, PA, DNET, 탈수초 질환에서도 보일 수 있음",
  "ODG를 시사(감별): coarse calcification, cortical-subcortical 위치, frontal 우세, **ASL cortical high-flow sign**(민감도 54.5%, 특이도 92.9%), **subcortical FLAIR signal drop**(44.0% / 93.8%), MRS cystathionine",
  "IDH 변이 추정: MRS **2-hydroxyglutarate (2HG)**",
]));
C.push(SRC("Kurokawa, Jpn J Radiol 2026;44:1534 (PMC13526116 전문, Table 1)"));

C.push(H2("6-2. Grade·고위험 분자 변화와 영상"));
C.push(...T(["연구", "대상", "영상 관련 결과"], [
  ["Ozono, J Neurooncol 2026", "Non-enhancing AST 38명", "Mismatch(>50% 기준) 55.3%, (>33%) 73.7%. Mismatch 양성은 **낮은 grade, CDKN2A HD 없음**과 연관. **CDKN2A HD 5례는 모두 mismatch 음성**. HD 제외 시 mismatch 음성은 작은 종양 크기와 연관(median 250.7 vs 1842.9 mm²)"],
  ["Alfattal, Int J Surg Pathol 2026", "≥55세 IDH-mut AST 22명", "Mismatch: grade 2의 **25%**, grade 3·4 **0%**. Grade 2의 **50%가 조영증강**"],
  ["Yang, AJNR 2025", "IDH-mut AST 122 + 외부 18명", "**Enhancing margin의 경계 불명확(poorly defined)** 이 CDKN2A HD 독립 예측인자(OR 13.19). 모델 AUC 0.84 / 외부 0.96"],
  ["Singh, Neuroradiology 2025", "IDH-mut AST 72명 (methylation)", "Methylation high-grade(A_IDH_HG) 예측: midline crossing, ependymal extension, poorly defined enhancing margin 등. 정확도 75%, AUC 0.77"],
  ["Imai, J Clin Med 2025", "Non-enhancing glioma 106명", "rCBV·MET PET 상관(r=0.5). rCBV는 AST에서 가장 낮음, **grade 4 구분 AUC 0.718**(MET 0.617보다 높은 경향)"],
], [2.2, 2.0, 5.0]));
C.push(SRC("PubMed abstract/전문 — Ozono 2026 (PMC13002771), Alfattal 2026 (10.1177/10668969261489616), Yang 2025 (10.3174/ajnr.A8861), Singh 2025 (10.1007/s00234-025-03766-z), Imai 2025 (10.3390/jcm14196777)"));

C.push(H2("6-3. PDGFRA amplification과 영상"));
C.push(...B([
  "**IDH-mut astrocytoma에서 PDGFRA amp에 특이적인 영상 소견은 확립된 근거를 찾지 못함** (PubMed 검색 기준)",
  "Perfusion(rCBV/rCBF)과 PDGFRA를 포함한 유전자 marker의 연관을 본 소규모 연구는 있음(1.5T, 67명, 여러 tumor type 혼합) → IDH-mut astrocytoma grading에 일반화하기 어려움",
  "따라서 영상의 역할은 **PDGFRA amp 자체를 맞히는 것이 아니라**, 조직 grade보다 aggressive한 부위를 찾아 **sampling·분자 검사·재판독을 촉발**하는 것",
]));
C.push(SRC("Śledzińska-Bebyn, Magn Reson Imaging 2024 (10.1016/j.mri.2024.110318, abstract); PubMed 검색 결과"));

C.push(H2("6-4. 영상 기반 tissue sampling"));
C.push(...B([
  "Biopsy target: **focal enhancement, diffusion restriction, perfusion 상승, 비조영 CT 고음영, 대사 이상** 부위",
  "절제 예정이어도 영상상 가장 high-grade로 예상되는 부위를 **따로 표시해 제출**(spatially annotated sampling) → 병리·분자 평가",
  "비조영·침윤성 종양에서도 **국소 ADC 저하·rCBV 상승**이 공존할 수 있음(molecular GBM에서 강조된 원칙 — IDH-mut AST에도 적용 가능 [보충])",
  "생검에서 grade 4 조직 소견이 빠질 수 있음 → 생검 조직만으로 grade를 확정하지 말고 영상과 대조",
]));
C.push(SRC("Kurokawa, Jpn J Radiol 2026 (PMC13526116) Introduction, Molecular GBM"));

C.push(H2("6-5. 판독 체크리스트"));
C.push(...BOX("[제안] IDH-mutant astrocytoma 판독 시", [
  "T2-FLAIR mismatch(완전/부분) 여부 — 있으면 AST 시사, **없거나 사라지면** 고위험 분자 변화 가능성",
  "조영증강: 유무, **경계(poorly defined margin)**, 새로 생긴 증강",
  "DWI/ADC: 최저 ADC 부위 위치",
  "DSC/DCE/ASL perfusion: rCBV 최고 부위 위치",
  "Midline crossing, ependymal extension",
  "**조직 grade와 영상이 불일치**(예: 조직 grade 2–3인데 ADC↓·rCBV↑)하면 보고서·MDT에서 **CNV(PDGFRA, CDKN2A/B) 및 NGS 확인 요청**",
  "Biopsy/절제 계획 시 aggressive focus 위치를 명시",
], GREEN));

// ---------- 7. Treatment ----------
C.push(H1("7. 치료와 임상적 의미 — 왜 grade가 중요한가"));
C.push(H2("7-1. CATNON (1p/19q non-codeleted anaplastic glioma)"));
C.push(...B([
  "Phase 3, 751명 무작위 배정: RT 단독(59.4 Gy/33 fx) / RT + **concurrent** TMZ / RT + **adjuvant** TMZ 12 cycles / 둘 다",
  "2nd interim (2021)",
  "  Concurrent TMZ: median OS 66.9 vs 60.4개월, HR 0.97 → **futility**",
  "  Adjuvant TMZ: 82.3 vs 46.9개월, **HR 0.64**",
  "Final (2026, 추적 10.9년)",
  "  ITT: adjuvant HR 0.65 / concurrent HR 0.91(유의하지 않음)",
  "  **IDH-mut(444명)**: adjuvant TMZ median OS **12.5년 vs 6.0년(HR 0.54)** / concurrent 9.7 vs 7.2년(HR 0.81, 유의하지 않음)",
  "  IDH-wt: TMZ 이득 없음",
  "  PDGFRA·CDK4 amp, CDKN2A HD, CNV load는 예후 불량과 연관되지만 **TMZ 효과 predictive 아님**",
  "결론: grade 3급 공격적 IDH-mut astrocytoma → **RT 후 adjuvant TMZ 12 cycles** (concurrent 없이)",
]));
C.push(SRC("van den Bent, Lancet Oncol 2021;22:813 (10.1016/S1470-2045(21)00090-5); van den Bent, Lancet Oncol 2026;27:45 (10.1016/S1470-2045(25)00614-X) — PubMed abstracts"));

C.push(H2("7-2. INDIGO (vorasidenib)"));
C.push(...B([
  "대상: 수술 외 치료를 받지 않은 **잔존/재발 grade 2 IDH-mutant glioma**, 331명",
  "PFS **27.7 vs 11.1개월** (HR 0.39), 다음 치료까지 기간 HR 0.26",
  "Grade 3 이상 ALT 상승 9.6%",
  "→ **Grade 2에 한정된 근거**. Grade 3/4로 upgrade되면 이 전략의 대상에서 벗어남",
  "Vorasidenib은 non-enhancing, 임상적으로 안정된 잔존 grade 2 종양에서 RT·항암을 미루는 \"active delay\" 전략으로 자리매김 [보충]",
]));
C.push(SRC("Mellinghoff, NEJM 2023;389:589 (10.1056/NEJMoa2304194); Wu, Curr Treat Options Oncol 2026 (abstract)"));

C.push(H2("7-3. Grade별 치료 개요"));
C.push(...T(["Grade", "치료 개요", "근거"], [
  ["2", "최대 안전 절제 → 관찰(GTR, 저위험) / 잔존 non-enhancing → vorasidenib 고려 / 고위험 → RT + 항암", "INDIGO; Wu 2026 [보충]"],
  ["3", "RT → **adjuvant TMZ 12 cycles**", "CATNON"],
  ["**4**", "고등급 치료(RT + TMZ) — 세부 regimen(concurrent 포함 여부)은 가이드라인·기관 프로토콜 [원문 확인 권장]", "EANO 2021 (초록만 확인)"],
], [0.9, 5.6, 2.7], { highlight: (r) => r[0] === "**4**" }));
C.push(...B([
  "Grade 상향의 의미: 치료 강도·시점, 임상시험 적격성(grade별 eligibility), 예후 상담",
  "Virata 2026 저자들: 확장된 분자 기준은 \"systemic underestimation of true tumor grade\"를 막고 고위험 환자에서 **치료 escalation**을 고려하게 함",
]));
C.push(SRC("Weller, Nat Rev Clin Oncol 2021;18:170 (10.1038/s41571-020-00447-z, abstract); CATNON; INDIGO; Virata 2026 Discussion"));

// ---------- 8. Case ----------
C.push(H1("8. 증례 적용 — 63/F, Astrocytoma, IDH-mutant"));
C.push(...T(["단계", "내용"], [
  ["① 초기 병리", "Astrocytoma, IDH-mutant, **CNS WHO grade 3** (조직 소견 기반)"],
  ["② 영상", "**ADC·rCBV에서만** aggressiveness 시사 (조직 grade보다 공격적인 소견)"],
  ["③ 피드백", "영상의학과 → 병리과 feedback"],
  ["④ 분자 재평가", "**PDGFRA amplification** 확인"],
  ["⑤ 통합 진단", "cIMPACT-NOW Update 12에 따라 **grade 4로 재판독**"],
  ["⑥ 치료", "**CCRT 시행** (grade 3 기준이었다면 다른 치료 전략 고려 대상)"],
], [2.0, 7.2]));
C.push(...B([
  "발표 포인트",
  "  나이: IDH-mut AST의 평균 진단 나이는 30대 후반(Virata 2026: 37.8세, ≥55세 8.6%) → 63세는 드문 고령. 고령 IDH-mut AST에서는 mismatch가 grade 2에만 보였다는 보고(Alfattal 2026)",
  "  조영증강이 뚜렷하지 않아도 **ADC·rCBV가 aggressive**하면 upgrade 가능성을 판독에 적어야 함",
  "  영상 판독 → 분자 검사 → grade 변경 → 치료 변경으로 이어진 **영상-병리 피드백의 모범 사례**",
  "[원문 확인 권장] 실제 환자의 조영증강 양상, 최저 ADC 값, rCBV 수치, PDGFRA 검출 방법(FISH/NGS/methylation)과 CDKN2A/B 상태를 차트에서 확인해 슬라이드에 넣을 것",
]));
C.push(SRC("증례 정보는 사용자 제공. Virata 2026 Table 1; Alfattal 2026"));

// ---------- 9. Q&A ----------
C.push(H1("9. 예상 질문과 답변"));
C.push(...T(["질문", "답변 요지"], [
  ["Update 12는 WHO 공식 기준인가?", "아니다. cIMPACT-NOW는 'Not Officially WHO'. 다음 WHO(CNS6) 전까지의 권고이며, 기관마다 적용 여부가 다를 수 있다."],
  ["PDGFRA amp은 얼마나 흔한가?", "연구마다 다름: grade II/III FISH 18.8%(Yang 2019), alteration 6.5%(Virata 2026), 1,140명 중 grade 2/3→4 재분류 22례(Richardson 2026). 검출법·정의·코호트 차이 때문."],
  ["CDKN2A/B HD와 같은 무게인가?", "Update 12는 동일한 grade 4 criterion으로 취급. Richardson 2026에서 재분류 grade 4 군의 OS는 원래 grade 4와 차이 없음. Virata 2026은 intermediate로 분류해 이견이 있음."],
  ["영상으로 PDGFRA amp을 예측할 수 있나?", "확립된 소견 없음. 영상은 aggressive focus를 찾아 sampling과 분자 검사를 유도하는 역할."],
  ["PDGFRA amp이면 TMZ가 더 잘 듣나?", "아니다. CATNON final에서 prognostic이지만 predictive는 아님."],
  ["Grade 3에서 concurrent TMZ는?", "CATNON에서 IDH-mut에서도 concurrent TMZ의 유의한 OS 이득은 없었고(HR 0.81), adjuvant TMZ가 이득(HR 0.54)."],
  ["Mitosis ≥3/2.4 mm²의 한계는?", "생검 크기·sampling에 민감. IHC 논문도 cut-off가 낮아 생검에서 grading이 어렵다고 지적 → 분자 검사가 보완."],
  ["IDH-wt에서 PDGFRA는?", "Update 11: <25세에서 pedHGG H3/IDH-wt의 key molecular feature. IDH-mut AST의 grading marker(Update 12)와 맥락이 다르다."],
], [3.0, 6.2]));

// ---------- 10. Summary table ----------
C.push(H1("10. 최종 요약표"));
C.push(...T(["항목", "핵심", "출처"], [
  ["WHO CNS5 기준", "AST IDH-mut grade 2–4, CDKN2A/B HD → grade 4", "Louis 2021; Update 5"],
  ["Update 8", "Meningioma: 1p del + 22q del/NF2 → grade 2", "Sahm 2025"],
  ["Update 9", "Methylation classifier는 통합 진단의 한 도구", "Aldape 2025"],
  ["Update 10", "새 tumor type 정의(영상 포함 clinical phenotype 필수)", "Hawkins 2025"],
  ["Update 11", "<40세 EGFR amp·+7/−10 단독 GBM 금지, PDGFRA는 <25세 pedHGG marker", "Wesseling 2025"],
  ["**Update 12**", "**PDGFRA amp → grade 4**, mitosis ≥3/2.4 mm² → grade 3, grade 3 분자 marker", "Brat 2026"],
  ["Validation", "PDGFRA amp 재분류 grade 4 OS 51.4 vs grade 3 126.2개월 (HR 1.9)", "Richardson 2026"],
  ["실제 진료", "24% grade 변경, CNV(methylation) 기여 최대", "Tauziède-Espariat 2026"],
  ["검출", "FISH, NGS, methylation CNV, SNP/CGH, IHC screening(≥10% 강양성: 민감도 100%)", "Tauziède-Espariat 2026"],
  ["영상", "PDGFRA 특이 소견 없음. Mismatch 소실·poorly defined enhancing margin·ADC↓/rCBV↑ = 고위험 단서", "Ozono, Yang, Imai, Kurokawa"],
  ["치료", "G2 vorasidenib / G3 RT→adjuvant TMZ / G4 고등급 치료", "INDIGO, CATNON, EANO"],
], [1.8, 5.2, 2.2], { highlight: (r) => r[0] === "**Update 12**" }));

// ---------- References ----------
C.push(H1("참고문헌 (PubMed)"));
const refs = [
  "Brat DJ, et al. cIMPACT-NOW Update 12: Refining pathology-based risk stratification and grading for IDH-mutant gliomas. Neuro-Oncol 2026;28(9):2069–2089. doi:10.1093/neuonc/noag108",
  "Brat DJ, et al. cIMPACT-NOW update 5: recommended grading criteria and terminologies for IDH-mutant astrocytomas. Acta Neuropathol 2020;139:603–608. doi:10.1007/s00401-020-02127-9",
  "Sahm F, et al. cIMPACT-NOW update 8: meningioma. Neuro-Oncol 2025;27(2):319–330. doi:10.1093/neuonc/noae170",
  "Aldape K, et al. cIMPACT-NOW update 9: DNA methylation profiling. Neuro-Oncol Adv 2025;7(1):vdae228. doi:10.1093/noajnl/vdae228",
  "Hawkins C, et al. cIMPACT-NOW update 10: defining new types. Brain Pathol 2025;35(6):e70018. doi:10.1111/bpa.70018",
  "Wesseling P, et al. cIMPACT-NOW update 11: IDH- and H3-wildtype HGG, PF ependymal tumors. Brain Pathol 2025;36(1):e70035. doi:10.1111/bpa.70035",
  "Louis DN, et al. cIMPACT-NOW update 1. Acta Neuropathol 2018;135:481–484. doi:10.1007/s00401-018-1808-0",
  "Louis DN, et al. cIMPACT-NOW update 2. Acta Neuropathol 2018;135:639–642. doi:10.1007/s00401-018-1826-y",
  "Brat DJ, et al. cIMPACT-NOW update 3. Acta Neuropathol 2018;136:805–810. doi:10.1007/s00401-018-1913-0",
  "Ellison DW, et al. cIMPACT-NOW update 4. Acta Neuropathol 2019;137:683–687. doi:10.1007/s00401-019-01987-0",
  "Louis DN, et al. cIMPACT-NOW update 6. Brain Pathol 2020;30:844–856. doi:10.1111/bpa.12832",
  "Ellison DW, et al. cIMPACT-NOW update 7. Brain Pathol 2020;30:863–866. doi:10.1111/bpa.12866",
  "Richardson TE, et al. Validation of proposed cIMPACT-NOW update 12 molecular grading criteria for IDH-mutant astrocytoma. Acta Neuropathol 2026;152. doi:10.1007/s00401-026-03058-7",
  "Virata MC, et al. Expanding the molecular grading criteria in IDH-mutant astrocytoma. Neuro-Oncol 2026. doi:10.1093/neuonc/noag069",
  "Tauziède-Espariat A, et al. [Update 12 in routine practice]. Ann Pathol 2026. doi:10.1016/j.annpat.2026.09.001",
  "Tauziède-Espariat A, et al. Immunohistochemistry as a tool for identifying PDGFRA amplification in CNS tumors. Brain Pathol 2026:e70142. doi:10.1111/bpa.70142",
  "Yang RR, et al. IDH mutant lower grade astrocytomas can be stratified by CDKN2A, CDK4 and PDGFRA copy number alterations. Brain Pathol 2019. doi:10.1111/bpa.12801",
  "Tesileanu CMS, et al. Prognostic significance of genome-wide DNA methylation profiles within the CATNON trial. Neuro-Oncol 2021. doi:10.1093/neuonc/noab088",
  "Tesileanu CMS, et al. Molecular markers related to outcome in IDH-mutant astrocytomas grade 2–4: a systematic review. Eur J Cancer 2022. doi:10.1016/j.ejca.2022.08.016",
  "van den Bent MJ, et al. CATNON second interim analysis. Lancet Oncol 2021;22:813–823. doi:10.1016/S1470-2045(21)00090-5",
  "van den Bent MJ, et al. CATNON final and exploratory analyses. Lancet Oncol 2026;27:45–56. doi:10.1016/S1470-2045(25)00614-X",
  "Mellinghoff IK, et al. Vorasidenib in IDH1- or IDH2-mutant low-grade glioma (INDIGO). N Engl J Med 2023;389:589–601. doi:10.1056/NEJMoa2304194",
  "Weller M, et al. EANO guidelines on the diagnosis and treatment of diffuse gliomas of adulthood. Nat Rev Clin Oncol 2021;18:170–186. doi:10.1038/s41571-020-00447-z",
  "Kurokawa R, et al. Imaging diagnosis of adult-type diffuse gliomas in the era of WHO CNS5 and cIMPACT-NOW updates 8–11. Jpn J Radiol 2026;44:1534–1542. doi:10.1007/s11604-026-02013-6",
  "Ozono K, et al. T2-FLAIR mismatch sign in nonenhancing astrocytoma, IDH-mutant. J Neurooncol 2026. doi:10.1007/s11060-026-05519-1",
  "Alfattal R, et al. IDH-mutant astrocytomas in patients older than 55 years. Int J Surg Pathol 2026. doi:10.1177/10668969261489616",
  "Yang, et al. Clinical and VASARI features to predict CDKN2A homozygous deletion in IDH-mutant astrocytomas. AJNR 2025. doi:10.3174/ajnr.A8861",
  "Singh, et al. MRI-based prediction of DNA methylation grade in IDH-mutant astrocytomas. Neuroradiology 2025. doi:10.1007/s00234-025-03766-z",
  "Imai, et al. [11C]-methionine PET and CBV imaging in non-contrast-enhanced gliomas. J Clin Med 2025;14:6777. doi:10.3390/jcm14196777",
  "Wu, et al. Treatment options for adult grade 2 IDH-mutant diffuse gliomas in the vorasidenib era. Curr Treat Options Oncol 2026. doi:10.1007/s11864-026-01401-4",
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
