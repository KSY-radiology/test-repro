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
    T("WHO CNS5 baseline  ", { bold: true, color: GOLD }),
    T("Astrocytoma, IDH-mutant = single tumor type, CNS WHO grade 2–4 / CDKN2A/B homozygous deletion → grade 4 regardless of histology", { color: WHITE }),
  ], { x: 0.4, y: 1.0, w: 9.2, h: 0.62, fontSize: 13, fill: { color: DARK2 }, margin: 8, valign: "middle",
    isTextBox: true, objectName: "CNS5 baseline" });

  const H = (s) => ({ text: s, options: { bold: true, color: "000000", fill: { color: GOLD }, align: "center" } });
  const R = (cells, hi) => cells.map((c, i) => ({ text: c, options: {
    color: hi ? GOLD : WHITE, bold: hi && i === 0, fill: { color: hi ? "332600" : (i === 0 ? DARK : "000000") },
    align: i === 0 ? "center" : "left" } }));
  const rows = [
    [H("Update"), H("Journal"), H("Topic"), H("Key recommendation")],
    R(["8", "Neuro-Oncol 2025", "Meningioma grading", "Grade 1 histology + 1p deletion with 22q del and/or NF2 variant → grade 2; standardized brain invasion"]),
    R(["9", "Neuro-Oncol Adv 2025", "DNA methylation profiling", "Classifier = one tool within integrated diagnosis (not stand-alone); layered report examples"]),
    R(["10", "Brain Pathol 2025", "Defining new tumor types", "Clinical phenotype + methylation cluster + driver + histology, ≥2 independent reports; otherwise NEC"]),
    R(["11", "Brain Pathol 2025", "IDH/H3-wt HGG, PF EPN", "Age <40: EGFR amp or +7/−10 alone insufficient for GBM; caution with TERTp-only low-grade tumors"]),
    R(["12", "Neuro-Oncol 2026", "IDH-mutant glioma grading", "PDGFRA amp → grade 4; mitoses ≥3/2.4 mm² → grade 3; molecular markers of grade 3"], true),
  ];
  s.addTable(rows, { x: 0.4, y: 1.82, w: 9.2, colW: [0.8, 1.55, 1.95, 4.9], fontSize: 11.5, fontFace: "Calibri",
    border: { type: "solid", pt: 0.75, color: "404040" }, valign: "middle", rowH: [0.34, 0.68, 0.68, 0.68, 0.68, 0.68],
    margin: [3, 5, 3, 5], objectName: "Updates table" });

  s.addText([
    T("Update 1–7 (2018–2020)", { bold: true, color: GOLD }),
    T(" were incorporated into WHO CNS5 (e.g., update 3: molecular GBM; update 5: CDKN2A/B HD grading)", { color: WHITE }),
  ], { x: 0.4, y: 5.85, w: 9.2, h: 0.4, fontSize: 12, margin: 0, isTextBox: true, objectName: "Prior updates note" });

  cite(s, "Louis DN, Neuro-Oncol 2021;23:1231 · Sahm F, Neuro-Oncol 2025;27:319 · Aldape K, Neuro-Oncol Adv 2025;7:vdae228 · Hawkins C, Brain Pathol 2025;35:e70018 · Wesseling P, Brain Pathol 2026;36:e70035 · Brat DJ, Neuro-Oncol 2026;28:2069");
  s.addNotes(
`cIMPACT-NOW updates published after WHO CNS5.
- Update 8: meningioma grading. 1p deletion with 22q deletion and/or NF2 variant upgrades a histologic grade 1 meningioma to grade 2.
- Update 9: principles for using the DNA methylation classifier.
- Update 10: criteria for accepting a new tumor type; tumors that do not meet them are reported as NEC.
- Update 11: in patients under 40, EGFR amplification or +7/-10 alone should not establish GBM (consider pediatric-type HGG).
- Update 12, today's focus: revised grading of IDH-mutant gliomas.`);
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
  s.addText("Oligodendroglioma: no new grading criteria", { x: 5.6, y: 4.0, w: 4.1, h: 0.5,
    fontSize: 11, color: GRAY, italic: true, margin: 0, isTextBox: true, align: "right" });

  // key changes
  s.addText(bullets([
    { t: "PDGFRA amplification added as a grade 4 criterion (same weight as CDKN2A/B HD)", c: GOLD, b: true },
    "Grade 2 vs 3: numeric mitotic threshold (≥3 / 2.4 mm²) and molecular markers of grade 3",
    "Histologic grade 2/3 tumors can be upgraded by NGS or methylation-array CNV",
  ], { fontSize: 14, paraSpaceAfter: 4 }), { x: 0.3, y: 5.5, w: 9.4, h: 1.25, valign: "top", margin: 0,
    isTextBox: true, objectName: "Key changes" });

  cite(s, "Figure: Park YW et al. Korean J Radiol 2026 (in press), Fig. 2 · Original criteria: Brat DJ et al. Neuro-Oncol 2026;28(9):2069-2089");
  s.addNotes(
`Update 12 grading algorithm (Brat et al. 2026).
1) Any of MVP, necrosis, CDKN2A/B HD or, newly, PDGFRA amplification → grade 4.
2) Otherwise, mitotic count ≥3 per 2.4 mm² → grade 3; if not, grade 2.
3) Orange box: molecular alterations that may suggest grade 3 (CDK4/6, CCND2 amp, RB1, PIK3CA/PIK3R1, EGFR, MYCN, methylation A_IDH_HG / G-CIMP-low).
4) Oligodendroglioma: the working group found no new evidence to change grading.
Figure redrawn from Park YW et al. Imaging of IDH-Mutant Gliomas in the Era of IDH Inhibitors: What Radiologists Need to Know. Korean J Radiol 2026 (in press), Fig. 2. Criteria from Brat et al. 2026.`);
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
  stat(1.2, "51 vs 126 mo", "Median OS: reclassified to grade 4 by PDGFRA amp vs remaining grade 3 (HR 1.9, p=0.028)", "Richardson TE, Acta Neuropathol 2026 (n=1,140)");
  stat(3.05, "18.8%", "PDGFRA amp in grade 2/3 IDH-mut astrocytoma (FISH); mutually exclusive with CDKN2A HD and CDK4 amp", "Yang RR, Brain Pathol 2019 (n=160)");
  stat(4.9, "24%", "Grade changed by molecular results in routine practice (grade 3 → 4: 3 cases)", "Tauziède-Espariat A, Ann Pathol 2026 (12/50)");

  s.addText(bullets([
    { t: "Update 12 conclusion", b: true, c: GOLD },
    { t: "PDGFRA amp = highly aggressive behavior → CNS WHO grade 4", lvl: 1 },
    { t: "Evidence", b: true, c: GOLD },
    { t: "Yang 2019: PDGFRA amp = highest-risk group (mean OS 40.5 vs 71.5 mo)", lvl: 1 },
    { t: "CATNON (432 IDH-mut anaplastic astrocytomas): independent prognostic factor", lvl: 1 },
    { t: "Virata 2026: grade 2/3 with these alterations, median OS 67–82 vs 135–141 mo", lvl: 1 },
    { t: "Detection", b: true, c: GOLD },
    { t: "NGS CNV, methylation array CNV, FISH, SNP/CGH array", lvl: 1 },
    { t: "PDGFRA IHC proposed as a screening tool", lvl: 1 },
    { t: "Pitfall", b: true, c: GOLD },
    { t: "Missed in histologic grade 2/3 tumors without CNV testing", lvl: 1 },
  ], { fontSize: 13.5, paraSpaceAfter: 3 }), { x: 3.95, y: 1.05, w: 5.75, h: 5.75, valign: "top", margin: 0,
    isTextBox: true, objectName: "PDGFRA evidence" });

  cite(s, "Brat DJ, Neuro-Oncol 2026 · Richardson TE, Acta Neuropathol 2026;152 · Yang RR, Brain Pathol 2019 · Tesileanu CMS, Neuro-Oncol 2021 · Virata MC, Neuro-Oncol 2026 · Tauziède-Espariat A, Brain Pathol 2026;e70142");
  s.addNotes(
`Why PDGFRA amplification became a grade 4 criterion.
- Validation (Richardson 2026, n=1,140): tumors reclassified to grade 4 by PDGFRA amp had median OS 51.4 months vs 126.2 months for tumors that stayed grade 3 (HR 1.9). 7 cases moved from grade 2 to 4, 15 from grade 3 to 4.
- Yang 2019 (n=160, FISH): PDGFRA amp 18.8%, CDKN2A HD 15.0%, CDK4 amp 18.1%, mutually exclusive. Mean OS of the three risk groups: 40.5 / 62.9 / 71.5 months.
- CATNON (Tesileanu 2021): PDGFRA amp was an independent prognostic factor in 432 IDH-mutant anaplastic astrocytomas.
- Routine practice (Tauziede-Espariat 2026): grade changed in 12 of 50 cases (24%), including 3 grade 3 → 4.
- Detection: NGS, methylation-array CNV, FISH; IHC proposed as a screening tool (Brain Pathol 2026).
Caveat: in the CATNON final analysis PDGFRA amp was prognostic but did not predict benefit from temozolomide.`);
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
    T("Aggressive features on ADC and rCBV only", { color: WHITE }),
  ], { x: 0.3, y: 5.45, w: 4.4, h: 0.65, fontSize: 13, margin: 0, isTextBox: true, objectName: "Case caption" });

  s.addText(bullets([
    { t: "T2-FLAIR mismatch: specific for IDH-mut astrocytoma, low sensitivity", b: true },
    { t: "Mostly non-enhancing lower grade (age ≥55 cohort: G2 25%, G3/4 0%)", lvl: 1 },
    { t: "Enhancement ≠ grade", b: true },
    { t: "High-grade tumors can be non-enhancing; grade 2 tumors can enhance", lvl: 1 },
    { t: "ADC ↓ / rCBV ↑ focus", b: true, c: GOLD },
    { t: "rCBV discriminates grade 4 in non-enhancing gliomas (AUC 0.72)", lvl: 1 },
    { t: "Guides biopsy targeting and requests for molecular work-up", lvl: 1 },
    { t: "No established imaging signature of PDGFRA amp", b: true },
    { t: "→ Feedback to pathology when imaging and histology disagree", lvl: 1 },
  ], { fontSize: 13.5, paraSpaceAfter: 4 }), { x: 4.95, y: 1.05, w: 4.75, h: 5.6, valign: "top", margin: 0,
    isTextBox: true, objectName: "Imaging bullets" });

  cite(s, "Alfattal, Int J Surg Pathol 2026 · Ozono, J Neurooncol 2026 · Imai, J Clin Med 2025;14:6777 · Yang, AJNR 2025 · Kurokawa, Jpn J Radiol 2026;44:1534");
  s.addNotes(
`(Insert the case images into the placeholders.)
- T2-FLAIR mismatch: high specificity, low sensitivity. In a cohort aged ≥55 it was seen in 25% of grade 2 and in no grade 3/4 tumors (Alfattal 2026). All CDKN2A HD cases were mismatch-negative (Ozono 2026).
- Enhancement alone does not determine grade (50% of grade 2 tumors enhanced in that cohort).
- In non-enhancing gliomas, rCBV discriminated grade 4 with AUC 0.718 (Imai 2025).
- CDKN2A HD was predicted by a poorly defined enhancing margin (OR 13.19, Yang AJNR 2025).
- No imaging feature specific to PDGFRA amplification has been established. When ADC/CBV look more aggressive than the histologic grade, as in this case, feed that back to pathology.`);
}

// ---------------- Slide 5: Why it matters ----------------
{
  const s = pres.addSlide({ masterName: "BLACK_TITLE_ONLY" });
  s.addText("Why It Matters: Grade Changes Treatment", { placeholder: "title" });

  // case flow
  const steps = [
    ["Initial", "Astrocytoma, IDH-mut, grade 3 (histology)"],
    ["Radiology", "Aggressive on ADC · rCBV → feedback to pathology"],
    ["Molecular", "PDGFRA amplification detected"],
    ["Integrated", "Regraded to grade 4 (Update 12)"],
    ["Treatment", "Concurrent chemoradiotherapy"],
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
    [H("Grade"), H("Treatment evidence")],
    [C("2", false, true), C("Observation after GTR / residual non-enhancing → vorasidenib (INDIGO; MFDS approval Jan 2026)")],
    [C("3", false, true), C("RT → adjuvant TMZ × 12 cycles (CATNON; no significant benefit of concurrent TMZ)")],
    [C("4", true, true), C("Grade 3 or GBM-type regimen (ASCO-SNO); not a vorasidenib candidate", true)],
  ], { x: 4.2, y: 1.05, w: 5.5, colW: [0.8, 4.7], fontSize: 12, fontFace: "Calibri",
    border: { type: "solid", pt: 0.75, color: "404040" }, valign: "middle", rowH: [0.34, 0.66, 0.66, 0.66],
    margin: [3, 6, 3, 6], objectName: "Treatment table" });

  s.addText(bullets([
    { t: "Take-home", b: true, c: GOLD },
    { t: "Grade drives treatment intensity, trial eligibility and counseling", lvl: 1 },
    { t: "If ADC/rCBV look worse than the histologic grade, ask for CNV (PDGFRA, CDKN2A/B)", lvl: 1 },
    { t: "The radiology report can trigger a molecular upgrade", lvl: 1 },
  ], { fontSize: 13.5, paraSpaceAfter: 4 }), { x: 4.2, y: 3.65, w: 5.5, h: 2.9, valign: "top", margin: 0,
    isTextBox: true, objectName: "Take-home" });

  cite(s, "van den Bent MJ, Lancet Oncol 2021;22:813 & 2026;27:45 · Mellinghoff IK, NEJM 2023;389:589 · Mohile NA, J Clin Oncol 2022;40:403 · Park YW, Korean J Radiol 2026 (in press)");
  s.addNotes(
`This case: histologic grade 3 at first. ADC and CBV looked aggressive, so we fed this back to pathology → PDGFRA amplification found → regraded to grade 4 under cIMPACT-NOW update 12 → concurrent chemoradiotherapy.
- Grade 3 IDH-mut astrocytoma: in CATNON, adjuvant TMZ improved survival (IDH-mut median OS 12.5 vs 6.0 years, HR 0.54); concurrent TMZ did not reach significance (HR 0.81).
- Grade 2: in INDIGO, vorasidenib improved PFS (27.7 vs 11.1 months) in residual non-enhancing grade 2 tumors. FDA approval Aug 2024, Korean MFDS Jan 2026 (Park YW, KJR 2026). A grade 4 tumor is not a candidate.
- Grade 4 IDH-mut: ASCO-SNO 2022 says these patients may follow either the grade 3 astrocytoma or the IDH-wildtype GBM recommendations.
- Message: radiology feedback led to a molecular upgrade and a change in treatment.`);
}

(async () => {
  const out = "/home/user/test-repro/deck/cIMPACT-NOW_update12_disease_review.pptx";
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("wrote", out);
})();
