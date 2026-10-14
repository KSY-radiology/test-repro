// Disease review: cIMPACT-NOW updates after WHO CNS5, focus on update 12 (PDGFRA amp -> grade 4)
// Style matched to the user's black-background template (4:3, gold titles, white Calibri body).
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/16a09025-2f4e-4864-9733-fb6e0e14f271_61656db6-4f8d-4ea1-9b5f-fdb7dd74c88e/pptx/scripts/apply_theme.js");

const THEME = {
  name: "Neuro Black",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "000000", lt1: "FFFFFF", dk2: "1A1A1A", lt2: "BFBFBF",
    accent1: "FFC000", accent2: "E36C0A", accent3: "5B9BD5", accent4: "C00000",
    accent5: "A6A6A6", accent6: "70AD47", hlink: "FFC000", folHlink: "E36C0A",
  },
};
const GOLD = "FFC000", ORANGE = "E36C0A", WHITE = "FFFFFF", GRAY = "BFBFBF",
  DARK = "1A1A1A", DARK2 = "262626", BLUE = "5B9BD5", PINK = "D9A0C8";

const pres = new pptxgen();
pres.defineLayout({ name: "SCREEN_4x3", width: 10, height: 7.5 });
pres.layout = "SCREEN_4x3";
pres.theme = { headFontFace: "Calibri", bodyFontFace: "Calibri" };
pres.title = "cIMPACT-NOW update 12 – IDH-mutant astrocytoma grading";

pres.defineSlideMaster({
  title: "BLACK_TITLE_ONLY",
  background: { color: "000000" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.4, y: 0.25, w: 9.2, h: 0.6,
      fontSize: 26, bold: true, color: GOLD, align: "center", valign: "middle", fontFace: "Calibri" },
      text: "" } },
  ],
});

const T = (s, o = {}) => ({ text: s, options: o });
function cite(slide, text) {
  slide.addText(text, { x: 3.2, y: 7.0, w: 6.6, h: 0.4, fontSize: 9, color: GRAY, align: "right",
    valign: "bottom", italic: true, isTextBox: true, margin: 0, fontFace: "Calibri", objectName: "Citation" });
}
function bullets(items, base = {}) {
  // items: [text | {t, lvl, b, c}]
  return items.map((it, i) => {
    const o = typeof it === "string" ? { t: it } : it;
    return {
      text: o.t,
      options: {
        bullet: o.lvl ? { indent: 14 } : { indent: 14 },
        indentLevel: o.lvl || 0,
        bold: !!o.b, color: o.c || WHITE,
        fontSize: o.fs || base.fontSize || 14,
        paraSpaceAfter: base.paraSpaceAfter ?? 6,
        breakLine: i < items.length - 1,
      },
    };
  });
}
function box(slide, x, y, w, h, text, o = {}) {
  slide.addText(text, {
    shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x, y, w, h,
    fill: { color: o.fill || DARK2 }, line: { color: o.line || o.fill || DARK2, width: o.lw || 1 },
    fontSize: o.fs || 12, color: o.color || WHITE, bold: o.bold ?? false, align: o.align || "center",
    valign: "middle", margin: o.margin ?? 4, fontFace: "Calibri", isTextBox: true, objectName: o.name || "Box",
  });
}
// straight arrow from (x1,y1) to (x2,y2)
function arrow(slide, x1, y1, x2, y2, color = WHITE) {
  slide.addShape(pres.shapes.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001),
    flipH: x2 < x1, flipV: y2 < y1, line: { color, width: 1.5, endArrowType: "triangle" }, objectName: "Arrow",
  });
}
function lbl(slide, x, y, s, color = GRAY) {
  slide.addText(s, { x, y, w: 0.3, h: 0.25, fontSize: 11, bold: true, color, align: "center", margin: 0, isTextBox: true });
}

// ---------------- Slide 1: cIMPACT-NOW after WHO CNS5 ----------------
{
  const s = pres.addSlide({ masterName: "BLACK_TITLE_ONLY" });
  s.addText("cIMPACT-NOW after WHO CNS5 (2021)", { placeholder: "title" });

  s.addText([
    T("WHO CNS5 기준점  ", { bold: true, color: GOLD }),
    T("Astrocytoma, IDH-mutant = 하나의 type, CNS WHO grade 2–4 / CDKN2A/B homozygous deletion → 조직 소견과 무관하게 grade 4", { color: WHITE }),
  ], { x: 0.4, y: 1.0, w: 9.2, h: 0.62, fontSize: 13, fill: { color: DARK2 }, margin: 8, valign: "middle",
    isTextBox: true, objectName: "CNS5 baseline" });

  const H = (s) => ({ text: s, options: { bold: true, color: "000000", fill: { color: GOLD }, align: "center" } });
  const R = (cells, hi) => cells.map((c, i) => ({ text: c, options: {
    color: hi ? GOLD : WHITE, bold: hi && i === 0, fill: { color: hi ? "332600" : (i === 0 ? DARK : "000000") },
    align: i === 0 ? "center" : "left" } }));
  const rows = [
    [H("Update"), H("발표"), H("대상"), H("핵심 내용")],
    R(["8", "Neuro-Oncol 2025", "Meningioma grading", "Grade 1 형태라도 1p deletion + 22q del and/or NF2 → grade 2 제안, brain invasion 판정 표준화"]),
    R(["9", "Neuro-Oncol Adv 2025", "DNA methylation profiling", "Classifier는 integrated diagnosis의 한 도구 (단독 진단 X), layered report 작성 예시"]),
    R(["10", "Brain Pathol 2025", "New tumor type 정의", "Clinical phenotype + methylation cluster + driver + histology, ≥2 독립 보고 → 미충족 시 NEC"]),
    R(["11", "Brain Pathol 2025", "IDH/H3-wt HGG, PF EPN", "<40세: EGFR amp·+7/−10 단독으로 GBM 진단 금지, TERTp 단독 저등급 → grade 4 신중"]),
    R(["12", "Neuro-Oncol 2026", "IDH-mutant glioma grading", "PDGFRA amp → grade 4, mitosis ≥3/2.4 mm² → grade 3, grade 3 분자 marker 제시"], true),
  ];
  s.addTable(rows, { x: 0.4, y: 1.82, w: 9.2, colW: [0.8, 1.55, 1.95, 4.9], fontSize: 11.5, fontFace: "Calibri",
    border: { type: "solid", pt: 0.75, color: "404040" }, valign: "middle", rowH: [0.34, 0.68, 0.68, 0.68, 0.68, 0.68],
    margin: [3, 5, 3, 5], objectName: "Updates table" });

  s.addText([
    T("Update 1–7 (2018–2020)", { bold: true, color: GOLD }),
    T("은 WHO CNS5에 이미 반영됨 (예: update 3 molecular GBM, update 5 CDKN2A/B HD grading)", { color: WHITE }),
  ], { x: 0.4, y: 5.85, w: 9.2, h: 0.4, fontSize: 12, margin: 0, isTextBox: true, objectName: "Prior updates note" });

  cite(s, "Louis DN, Neuro-Oncol 2021;23:1231 · Sahm F, Neuro-Oncol 2025;27:319 · Aldape K, Neuro-Oncol Adv 2025;7:vdae228 · Hawkins C, Brain Pathol 2025;35:e70018 · Wesseling P, Brain Pathol 2026;36:e70035 · Brat DJ, Neuro-Oncol 2026;28:2069");
  s.addNotes(
`WHO CNS5 이후 나온 cIMPACT-NOW update를 짧게 정리.
- Update 8: 수막종 grading. 1p del + 22q del/NF2 → grade 2 제안.
- Update 9: methylation classifier 사용 원칙.
- Update 10: 새 tumor type을 인정하는 기준. 기준 미충족 시 NEC로 보고.
- Update 11: 40세 미만에서 EGFR amp나 +7/-10만으로 GBM 진단하지 말 것 (pediatric-type HGG 감별).
- 오늘 핵심인 Update 12: IDH-mutant glioma grading 개정.`);
}

// ---------------- Slide 2: Update 12 algorithm ----------------
{
  const s = pres.addSlide({ masterName: "BLACK_TITLE_ONLY" });
  s.addText("cIMPACT-NOW Update 12: Grading of IDH-mutant Gliomas", { placeholder: "title", fontSize: 24 });

  // headers
  box(s, 0.3, 1.0, 6.3, 0.36, "IDH-mutant astrocytoma", { fill: "1F3A4D", bold: true, fs: 13, name: "Astro header" });
  box(s, 6.85, 1.0, 2.85, 0.36, "Oligodendroglioma", { fill: "3A2A4D", bold: true, fs: 13, name: "Oligo header" });

  // astro decision boxes
  box(s, 3.75, 1.55, 2.85, 0.95, [
    T("MVP, necrosis,", { breakLine: true }), T("CDKN2A/B HD", { breakLine: true }),
    T("and/or ", {}), T("PDGFRA amp", { bold: true, color: GOLD }),
  ], { fill: "2B3A2B", line: GOLD, lw: 1.5, fs: 12, name: "Grade 4 criteria" });
  box(s, 0.3, 1.72, 2.6, 0.6, "Mitotic activity ≥3 / 2.4 mm²", { fill: "2B3A2B", fs: 12, name: "Mitosis criterion" });
  arrow(s, 3.75, 2.02, 2.9, 2.02); lbl(s, 3.18, 1.75, "N");

  // outcomes
  const oy = 3.25, oh = 0.55;
  box(s, 0.15, oy, 1.75, oh, "Astrocytoma\nWHO grade 2", { fill: "F2C14E", color: "000000", bold: true, fs: 11.5, name: "G2" });
  box(s, 2.05, oy, 1.75, oh, "Astrocytoma\nWHO grade 3", { fill: "F2C14E", color: "000000", bold: true, fs: 11.5, name: "G3" });
  box(s, 4.3, oy, 1.75, oh, "Astrocytoma\nWHO grade 4", { fill: GOLD, color: "000000", bold: true, fs: 11.5, line: ORANGE, lw: 2.5, name: "G4" });
  arrow(s, 1.25, 2.32, 1.03, oy); lbl(s, 0.78, 2.6, "N");
  arrow(s, 1.95, 2.32, 2.93, oy); lbl(s, 2.55, 2.6, "Y");
  arrow(s, 5.17, 2.5, 5.17, oy); lbl(s, 5.22, 2.75, "Y");

  // grade-3 molecular box
  s.addText([
    T("Molecular alterations suggesting WHO grade 3", { bold: true, color: GOLD, breakLine: true }),
    T("CDK4 / CDK6 / CCND2 amp, CDKN2A mutation, RB1 HD or mutation", { bullet: { indent: 10 }, breakLine: true }),
    T("PIK3CA / PIK3R1 mutation", { bullet: { indent: 10 }, breakLine: true }),
    T("EGFR alteration, MYCN amp", { bullet: { indent: 10 }, breakLine: true }),
    T("DNA methylation: A_IDH_HG or G-CIMP-low", { bullet: { indent: 10 } }),
  ], { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x: 0.3, y: 4.0, w: 5.0, h: 1.15,
    fill: { color: "3D2614" }, line: { color: ORANGE, width: 1 }, fontSize: 11, color: WHITE, margin: 6,
    valign: "top", isTextBox: true, objectName: "Grade 3 molecular" });
  arrow(s, 2.93, 4.0, 2.93, oy + oh, ORANGE);

  // oligo
  box(s, 6.85, 1.55, 2.85, 0.95, "MVP, necrosis,\nand/or mitoses ≥6 / 2.4 mm²", { fill: "2B3A2B", fs: 12, name: "Oligo criteria" });
  box(s, 6.85, oy, 1.35, oh, "Oligo\ngrade 2", { fill: PINK, color: "000000", bold: true, fs: 11.5, name: "O2" });
  box(s, 8.35, oy, 1.35, oh, "Oligo\ngrade 3", { fill: PINK, color: "000000", bold: true, fs: 11.5, name: "O3" });
  arrow(s, 7.8, 2.5, 7.52, oy); lbl(s, 7.35, 2.75, "N");
  arrow(s, 8.75, 2.5, 9.02, oy); lbl(s, 9.05, 2.75, "Y");
  s.addText("Oligodendroglioma: grading 개선 근거 없음 (기존 기준 유지)", { x: 5.6, y: 4.0, w: 4.1, h: 0.5,
    fontSize: 11, color: GRAY, italic: true, margin: 0, isTextBox: true, align: "right" });

  // key changes
  s.addText(bullets([
    { t: "Grade 4 criterion에 PDGFRA amplification 추가 (CDKN2A/B HD와 동급)", c: GOLD, b: true },
    "Grade 2/3 경계: mitosis ≥3 / 2.4 mm² 수치 기준 제안, grade 3 분자 marker 명시",
    "Histology grade 2/3여도 NGS · methylation CNV 결과로 upgrade 가능",
  ], { fontSize: 14, paraSpaceAfter: 4 }), { x: 0.3, y: 5.5, w: 9.4, h: 1.25, valign: "top", margin: 0,
    isTextBox: true, objectName: "Key changes" });

  cite(s, "Adapted from Brat DJ et al. cIMPACT-NOW update 12. Neuro-Oncol 2026;28(9):2069-2089. doi:10.1093/neuonc/noag108");
  s.addNotes(
`Update 12 grading algorithm (Brat et al. 2026).
1) MVP, necrosis, CDKN2A/B HD, 그리고 새로 PDGFRA amplification 중 하나라도 있으면 grade 4.
2) 없으면 mitotic count ≥3/2.4 mm²이면 grade 3, 아니면 grade 2.
3) 주황 박스: grade 3를 시사하는 분자 변화 (CDK4/6, CCND2 amp, RB1, PIK3CA/PIK3R1, EGFR, MYCN, methylation A_IDH_HG/G-CIMP-low).
4) Oligodendroglioma는 grading을 바꿀 근거가 없었다고 결론.
※ 그림은 논문 Fig.를 기반으로 다시 그림.`);
}

// ---------------- Slide 3: PDGFRA amplification ----------------
{
  const s = pres.addSlide({ masterName: "BLACK_TITLE_ONLY" });
  s.addText("PDGFRA Amplification → WHO Grade 4", { placeholder: "title" });

  // stat column
  const stat = (y, big, small, src) => {
    s.addText(big, { x: 0.3, y, w: 3.2, h: 0.62, fontSize: 30, bold: true, color: GOLD, margin: 0, isTextBox: true, objectName: "Stat" });
    s.addText(small, { x: 0.3, y: y + 0.62, w: 3.2, h: 0.62, fontSize: 11.5, color: WHITE, margin: 0, valign: "top", isTextBox: true });
    s.addText(src, { x: 0.3, y: y + 1.22, w: 3.2, h: 0.25, fontSize: 9, color: GRAY, italic: true, margin: 0, isTextBox: true });
  };
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.15, y: 1.05, w: 3.5, h: 5.7, rectRadius: 0.1,
    fill: { color: DARK }, line: { color: DARK }, objectName: "Stat panel" });
  stat(1.2, "51 vs 126 mo", "Median OS: PDGFRA amp으로 grade 4 재분류 vs grade 3 유지 (HR 1.9, p=0.028)", "Richardson TE, Acta Neuropathol 2026 (n=1,140)");
  stat(3.05, "18.8%", "Grade 2/3 IDH-mut astrocytoma의 PDGFRA amp (FISH). CDKN2A HD · CDK4 amp와 mutually exclusive", "Yang RR, Brain Pathol 2019 (n=160)");
  stat(4.9, "24%", "실제 진료에서 molecular 결과로 grade 변경 (grade 3 → 4: 3례)", "Tauziède-Espariat A, Ann Pathol 2026 (12/50)");

  s.addText(bullets([
    { t: "Update 12의 결론", b: true, c: GOLD },
    { t: "PDGFRA amp = highly aggressive behavior → CNS WHO grade 4", lvl: 1 },
    { t: "근거", b: true, c: GOLD },
    { t: "Yang 2019: PDGFRA amp → 가장 나쁜 risk group (mean OS 40.5 vs 71.5 mo)", lvl: 1 },
    { t: "CATNON (IDH-mut anaplastic astrocytoma 432명): 독립 예후인자", lvl: 1 },
    { t: "Virata 2026: 위 분자 변화가 있으면 grade 2/3의 median OS 67–82 mo vs 135–141 mo", lvl: 1 },
    { t: "검출 방법", b: true, c: GOLD },
    { t: "NGS CNV, methylation array CNV, FISH, SNP/CGH array", lvl: 1 },
    { t: "PDGFRA IHC: amplification screening 도구로 제안", lvl: 1 },
    { t: "Pitfall", b: true, c: GOLD },
    { t: "형태학적 grade 2/3 → CNV 검사 없으면 놓침", lvl: 1 },
  ], { fontSize: 13.5, paraSpaceAfter: 3 }), { x: 3.95, y: 1.05, w: 5.75, h: 5.75, valign: "top", margin: 0,
    isTextBox: true, objectName: "PDGFRA evidence" });

  cite(s, "Brat DJ, Neuro-Oncol 2026 · Richardson TE, Acta Neuropathol 2026;152 · Yang RR, Brain Pathol 2019 · Tesileanu CMS, Neuro-Oncol 2021 · Virata MC, Neuro-Oncol 2026 · Tauziède-Espariat A, Brain Pathol 2026;e70142");
  s.addNotes(
`PDGFRA amplification이 grade 4 criterion이 된 근거.
- Validation (Richardson 2026, 1,140명): PDGFRA amp으로 grade 4가 된 군의 median OS 51.4개월, grade 3 유지군은 126.2개월 (HR 1.9). Grade 2→4 7례, 3→4 15례.
- Yang 2019 (160명, FISH): PDGFRA amp 18.8%, CDKN2A HD 15.0%, CDK4 amp 18.1%로 서로 겹치지 않음. 3 risk group의 mean OS 40.5/62.9/71.5개월.
- CATNON (Tesileanu 2021): IDH-mut anaplastic astrocytoma 432명에서 PDGFRA amp는 독립 예후인자.
- 실제 진료 (Tauziède-Espariat 2026): 50례 중 12례 (24%)에서 grade 변경, grade 3→4가 3례.
- 검출: NGS, methylation array CNV, FISH 등. IHC를 screening에 쓰자는 보고 (Brain Pathol 2026).
※ CATNON final analysis: PDGFRA amp는 예후 인자이지만 TMZ 효과의 predictive marker는 아님.`);
}

// ---------------- Slide 4: Imaging ----------------
{
  const s = pres.addSlide({ masterName: "BLACK_TITLE_ONLY" });
  s.addText("Imaging Findings", { placeholder: "title" });

  const ph = (x, y, w, h, label) => {
    s.addText(label, { shape: pres.shapes.RECTANGLE, x, y, w, h, fill: { color: DARK },
      line: { color: "595959", width: 1, dashType: "dash" }, fontSize: 12, color: GRAY, align: "center",
      valign: "middle", isTextBox: true, objectName: "Image placeholder" });
  };
  ph(0.3, 1.05, 2.15, 2.1, "T2 / FLAIR");
  ph(2.55, 1.05, 2.15, 2.1, "CE T1");
  ph(0.3, 3.25, 2.15, 2.1, "ADC");
  ph(2.55, 3.25, 2.15, 2.1, "rCBV");
  s.addText([
    T("63/F, Astrocytoma, IDH-mutant", { bold: true, color: GOLD, breakLine: true }),
    T("ADC · rCBV에서만 aggressive feature", { color: WHITE }),
  ], { x: 0.3, y: 5.45, w: 4.4, h: 0.65, fontSize: 13, margin: 0, isTextBox: true, objectName: "Case caption" });

  s.addText(bullets([
    { t: "T2-FLAIR mismatch: IDH-mut astrocytoma에 특이적, 민감도 낮음", b: true },
    { t: "주로 non-enhancing 저등급 (≥55세 cohort: G2 25%, G3/4 0%)", lvl: 1 },
    { t: "Enhancement ≠ grade", b: true },
    { t: "Non-enhancing high grade 존재, grade 2도 enhancement 가능", lvl: 1 },
    { t: "ADC ↓ / rCBV ↑ focus", b: true, c: GOLD },
    { t: "Non-enhancing glioma에서 rCBV로 grade 4 구분 AUC 0.72", lvl: 1 },
    { t: "Biopsy target + 병리에 molecular work-up 요청 근거", lvl: 1 },
    { t: "PDGFRA amp 특이 영상 소견: 확립된 근거 없음", b: true },
    { t: "→ 영상-병리 불일치 시 피드백이 핵심", lvl: 1 },
  ], { fontSize: 13.5, paraSpaceAfter: 4 }), { x: 4.95, y: 1.05, w: 4.75, h: 5.6, valign: "top", margin: 0,
    isTextBox: true, objectName: "Imaging bullets" });

  cite(s, "Alfattal, Int J Surg Pathol 2026 · Ozono, J Neurooncol 2026 · Imai, J Clin Med 2025;14:6777 · Yang, AJNR 2025 · Kurokawa, Jpn J Radiol 2026;44:1534");
  s.addNotes(
`(증례 영상을 placeholder 자리에 넣어 주세요.)
- T2-FLAIR mismatch: 특이도는 높지만 민감도는 낮음. ≥55세 cohort에서 grade 2의 25%에서만 보였고 grade 3/4에서는 없었음 (Alfattal 2026). CDKN2A HD 증례가 mismatch-negative와 연관 (Ozono 2026).
- Enhancement 여부만으로 grade를 판단할 수 없음 (위 cohort에서 grade 2의 50%가 enhancement).
- Non-enhancing glioma에서 rCBV는 grade 4 구분 AUC 0.718 (Imai 2025).
- CDKN2A HD 예측: poorly defined enhancing margin (OR 13.19, Yang AJNR 2025).
- PDGFRA amp에 특이적인 영상 소견은 아직 확립된 근거가 없음. 이 증례처럼 ADC/CBV가 조직 grade보다 aggressive하면 병리에 feedback을 주는 것이 핵심.`);
}

// ---------------- Slide 5: Why it matters ----------------
{
  const s = pres.addSlide({ masterName: "BLACK_TITLE_ONLY" });
  s.addText("Why It Matters: Grade Changes Treatment", { placeholder: "title" });

  // case flow
  const steps = [
    ["Initial", "Astrocytoma, IDH-mut, grade 3 (histology)"],
    ["Radiology", "ADC · rCBV: aggressive → 병리에 피드백"],
    ["Molecular", "PDGFRA amplification 확인"],
    ["Integrated", "Grade 4로 재판독 (Update 12)"],
    ["Treatment", "CCRT 시행"],
  ];
  steps.forEach(([k, v], i) => {
    const y = 1.05 + i * 1.0;
    const last = i === steps.length - 1;
    s.addText(String(i + 1), { shape: pres.shapes.OVAL, x: 0.3, y: y + 0.05, w: 0.62, h: 0.62,
      fill: { color: last ? GOLD : (i === 2 || i === 3 ? ORANGE : "404040") }, line: { color: "000000" },
      fontSize: 18, bold: true, color: last ? "000000" : WHITE, align: "center", valign: "middle", margin: 0,
      isTextBox: true, objectName: "Step dot" });
    s.addText([T(k, { bold: true, color: last ? GOLD : (i === 2 || i === 3 ? ORANGE : GRAY), fontSize: 11, breakLine: true }),
      T(v, { color: last ? GOLD : WHITE, bold: last })], { x: 1.05, y: y - 0.02, w: 2.95, h: 0.78, fontSize: 13,
      valign: "middle", margin: 0, isTextBox: true, objectName: "Step text" });
    if (!last) arrow(s, 0.61, y + 0.67, 0.61, y + 1.05, GRAY);
  });

  // treatment table
  const H = (t) => ({ text: t, options: { bold: true, color: "000000", fill: { color: GOLD }, align: "center" } });
  const C = (t, hi, ctr) => ({ text: t, options: { color: hi ? GOLD : WHITE, bold: !!hi, align: ctr ? "center" : "left", fill: { color: hi ? "332600" : DARK } } });
  s.addTable([
    [H("Grade"), H("주요 치료 근거")],
    [C("2", false, true), C("관찰 (GTR) / 잔존 non-enhancing → vorasidenib (INDIGO)")],
    [C("3", false, true), C("RT → adjuvant TMZ 12 cycles (CATNON: concurrent TMZ 이득 없음)")],
    [C("4", true, true), C("RT + TMZ, 고등급 치료 (EANO) — vorasidenib 대상 아님", true)],
  ], { x: 4.2, y: 1.05, w: 5.5, colW: [0.8, 4.7], fontSize: 12, fontFace: "Calibri",
    border: { type: "solid", pt: 0.75, color: "404040" }, valign: "middle", rowH: [0.34, 0.66, 0.66, 0.66],
    margin: [3, 6, 3, 6], objectName: "Treatment table" });

  s.addText(bullets([
    { t: "Take-home", b: true, c: GOLD },
    { t: "Grade = 치료 강도 · trial 적격성 · 예후 상담을 결정", lvl: 1 },
    { t: "ADC/rCBV가 조직 grade보다 aggressive하면 PDGFRA · CDKN2A 등 CNV 확인 요청", lvl: 1 },
    { t: "영상 판독이 molecular upgrade의 trigger가 될 수 있음", lvl: 1 },
  ], { fontSize: 13.5, paraSpaceAfter: 4 }), { x: 4.2, y: 3.65, w: 5.5, h: 2.9, valign: "top", margin: 0,
    isTextBox: true, objectName: "Take-home" });

  cite(s, "van den Bent MJ, Lancet Oncol 2021;22:813 & 2026;27:45 · Mellinghoff IK, NEJM 2023;389:589 · Weller M, Nat Rev Clin Oncol 2021;18:170 · Brat DJ, Neuro-Oncol 2026");
  s.addNotes(
`이 증례: 처음에는 조직학적으로 grade 3. 영상에서 ADC/CBV상 aggressive해 병리에 feedback → PDGFRA amplification 확인 → cIMPACT-NOW 12에 따라 grade 4로 재판독 → CCRT 시행.
- Grade 3 IDH-mut astrocytoma: CATNON에서 adjuvant TMZ는 생존 이득 (IDH-mut에서 median OS 12.5 vs 6.0년, HR 0.54), concurrent TMZ는 유의한 이득 없음 (HR 0.81, NS).
- Grade 2: INDIGO에서 vorasidenib이 PFS 27.7 vs 11.1개월로 개선 (잔존 non-enhancing grade 2 대상). Grade 4로 바뀌면 대상이 아님.
- Grade 4: 고등급 치료 (RT + TMZ).
- 메시지: 영상의학과의 feedback이 molecular upgrade와 치료 변경으로 이어진 증례.`);
}

(async () => {
  const out = "/home/user/test-repro/deck/cIMPACT-NOW_update12_disease_review.pptx";
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("wrote", out);
})();
