// Spinal cord axial icons (radiologic convention: anterior = top, patient right = image left)
const fs = require('fs');
const { chromium } = require('playwright');

const CORD = '#FBE3C8', CORD_EDGE = '#D9A57A', GM = '#E2C4AA', LES = '#F07A2A';

// Right half of the H-shaped gray matter (x >= 100); mirrored for the left half.
// Anterior horn: broad & blunt (top). Posterior horn: slender, extends to posterolateral sulcus (bottom).
const gmHalf = `M100,70 L110,70 C113,62 114,52 120,45 C127,38 142,39 148,47
 C153,54 152,64 146,70 C142,74 137,76 135,80
 C136,90 142,104 147,118 C149,124 151,130 150,134
 C147,136 143,134 141,130 C135,116 126,100 118,92
 C113,88 106,87 100,87 Z`;

const cord = `
 <ellipse cx="100" cy="82" rx="82" ry="64" fill="${CORD}" stroke="${CORD_EDGE}" stroke-width="2"/>
 <path d="${gmHalf}" fill="${GM}"/>
 <path d="${gmHalf}" fill="${GM}" transform="translate(200,0) scale(-1,1)"/>
 <circle cx="100" cy="78" r="2.2" fill="${CORD}" stroke="${CORD_EDGE}" stroke-width="1"/>
 <!-- anterior median fissure (deep, top) -->
 <path d="M100,17 L100,42" stroke="#FFFFFF" stroke-width="5" stroke-linecap="round"/>
 <path d="M96.5,18 L98,42 M103.5,18 L102,42" stroke="${CORD_EDGE}" stroke-width="1"/>
 <!-- posterior median sulcus/septum (shallow, bottom) -->
 <path d="M100,146 L100,92" stroke="${CORD_EDGE}" stroke-width="1.2"/>`;

const L = (s) => `<g fill="${LES}" fill-opacity="0.55">${s}</g>`;
const lesions = {
  '1_MS_peripheral': L(`
    <ellipse cx="100" cy="120" rx="15" ry="13"/>
    <path d="M178,68 C184,80 182,98 172,110 C160,104 152,94 153,84 C156,76 166,70 178,68 Z"/>`),
  '2_NMOSD_central_LETM': L(`
    <circle cx="80" cy="80" r="40"/><circle cx="120" cy="80" r="40"/><circle cx="100" cy="82" r="44"/>`),
  '3_lower_spine_conus': L(`
    <circle cx="82" cy="82" r="42"/><circle cx="118" cy="82" r="42"/>`),
  '4_long_segment_confluent': L(`
    <circle cx="82" cy="82" r="44"/><circle cx="118" cy="82" r="44"/><circle cx="100" cy="82" r="30"/>`),
};

const svg = (inner, labels) => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 ${labels ? -14 : 12} 200 ${labels ? 188 : 142}" width="${labels ? 400 : 1000}">
${cord}${inner}${labels ? `<text x="100" y="-2" font-size="12" text-anchor="middle" font-family="Arial" font-weight="bold">A</text>
<text x="100" y="170" font-size="12" text-anchor="middle" font-family="Arial" font-weight="bold">P</text>
<text x="4" y="86" font-size="11" font-family="Arial">R</text><text x="188" y="86" font-size="11" font-family="Arial">L</text>` : ''}</svg>`;

(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  const shoot = async (s, file, transparent) => {
    await p.setContent(`<html><body style="margin:0;background:${transparent ? 'transparent' : '#fff'}">${s}</body></html>`);
    await p.locator('svg').screenshot({ path: file, omitBackground: transparent });
  };
  for (const [k, v] of Object.entries(lesions)) {
    fs.writeFileSync(`${k}.svg`, svg(v, false));
    await shoot(svg(v, false), `${k}.png`, true);
  }
  const prev = `<div style="display:flex;gap:10px;padding:10px;font-family:Arial">` +
    Object.entries(lesions).map(([k, v]) => `<div style="text-align:center;font-size:13px">${svg(v, true).replace('width="400"', 'width="230"')}<br>${k}</div>`).join('') + `</div>`;
  await p.setContent(`<html><body style="margin:0;background:#fff">${prev}</body></html>`);
  await p.locator('div').first().screenshot({ path: 'preview.png' });
  await b.close();
})();
