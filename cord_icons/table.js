// Rebuilds the "Cord" row of the table with corrected axial icons (anterior = top)
const fs = require('fs');
const { chromium } = require('playwright');
const ax = (f) => fs.readFileSync(f, 'utf8').replace(/width="\d+"/, 'width="120"');

// Sagittal sketch: cord line (top = cranial) with orange lesion segments [y1,y2] in 0-130
const sag = (segs) => `<svg viewBox="0 0 30 140" width="26" height="122">
 <path d="M8,4 C18,30 22,60 16,90 C13,105 12,120 16,136" fill="none" stroke="#E9D9C4" stroke-width="3" stroke-linecap="round"/>
 ${segs.map(([a, b, x]) => `<line x1="${x}" y1="${a}" x2="${x}" y2="${b}" stroke="#F07A2A" stroke-width="5" stroke-linecap="round"/>`).join('')}</svg>`;

const cols = [
  { f: '1_MS_peripheral.svg', s: sag([[8, 11, 10], [22, 25, 14]]),
    t: `<span class="g">short segment,</span><br><b>focal, peripheral<br>(posterior and lateral)<br>cervical lesions,<br>multiple lesions</b>` },
  { f: '2_NMOSD_central_LETM.svg', s: sag([[4, 34, 12]]),
    t: `<b>LETM, central GM,<br>&gt; 1/2 of cross sectional area,<br>bright spotty lesion on T2WI,<br>single cervical lesion with<br>brainstem extension</b>` },
  { f: '3_lower_spine_conus.svg', s: sag([[118, 132, 14]]),
    t: `<b>lower spine<br>(thoracolumbar)<br>and conus medullaris</b>` },
  { f: '4_long_segment_confluent.svg', s: sag([[62, 100, 18]]),
    t: `<b>variable,</b><br><span class="g">long segment,</span><br><b>confluent and large,<br>thoracic lesion</b>` },
];

const html = `<html><head><style>
body{margin:0;background:#fff;font-family:Arial,Helvetica,sans-serif}
table{border-collapse:separate;border-spacing:3px;margin:12px}
td{background:#F7E4E6;text-align:center;vertical-align:middle;font-size:13px;line-height:1.55;color:#222;width:215px;padding:10px 6px}
td.h{width:110px;background:#E9CFD2;color:#8a6a6e;font-size:14px}
td.img{background:#fff;height:130px;padding:0}
.g{color:#999;font-weight:normal}
.pic{display:flex;align-items:center;justify-content:center;gap:6px}
</style></head><body><table>
<tr><td class="h" rowspan="2">Cord</td>${cols.map(c => `<td class="img"><div class="pic">${c.s}${ax(c.f)}</div></td>`).join('')}</tr>
<tr>${cols.map(c => `<td>${c.t}</td>`).join('')}</tr>
</table></body></html>`;

(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ deviceScaleFactor: 3 });
  await p.setContent(html);
  await p.locator('table').screenshot({ path: 'cord_table_fixed.png' });
  await b.close();
})();
