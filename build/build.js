// Build the whole deck.  usage: SKILL_DIR=... NODE_PATH=... node build.js [part...]
const L = require('./lib');
const PARTS = ['p0_open', 'p1_intro', 'p2_keys', 'p3_hash', 'p4_ledger', 'p5_pow', 'p6_econ', 'p7_achieve', 'p8_flaws', 'p8b_energy_value', 'p9_legacy', 'p10_appendix'];
(async () => {
  await L.warmIcons(); L.defineLayouts();
  const only = process.argv.slice(2);
  for (const p of PARTS) {
    if (only.length && !only.includes(p)) continue;
    let mod; try { mod = require('./parts/' + p); } catch (e) { if (e.code === 'MODULE_NOT_FOUND' && e.message.includes('parts/' + p)) { console.log('skip', p); continue; } throw e; }
    await mod(L);
  }
  const out = process.env.OUT || 'out/deck.pptx';
  await L.finish(out);
  console.log('written', out, 'slides:', L.pres.slides.length);
  if (L.warnings.length) { console.log('WARNINGS (' + L.warnings.length + '):'); L.warnings.forEach((w) => console.log(' -', w)); }
})();
