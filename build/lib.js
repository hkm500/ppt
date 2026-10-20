/* Design system + component library for the revised Chapter-4 deck.
   - 4:3 (10 x 7.5 in) to stay consistent with the other chapters of the course.
   - Theme colours / fonts live in the theme; every colour below is a scheme colour.
   - All diagrams are native shapes, all charts native charts, so the teacher can edit everything. */
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const FA = require('react-icons/fa');

const FONT = 'Microsoft YaHei';
const THEME = {
  name: '数字货币-第四章-比特币',
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: '1F2430',      // ink
    lt1: 'FFFFFF',
    dk2: '2E3445',      // lighter ink (dark cards)
    lt2: 'F3F4F6',      // card background
    accent1: 'F7931A',  // bitcoin orange (fills; text on dark)
    accent2: 'B42318',  // crimson (warnings, errors, "wrong")
    accent3: '0F766E',  // teal (verified / correct)
    accent4: '475569',  // slate (secondary text)
    accent5: '94A3B8',  // light slate (lines, muted)
    accent6: 'B45309',  // burnt orange (emphasis text / big numbers on light)
    hlink: 'B45309',
    folHlink: '7C2D12',
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = 'LAYOUT_4x3';
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = '数字货币 第四章 比特币（2026修订版）';
pres.subject = '数字货币课程课件';
pres.author = '数字货币课程组';
const C = pres.SchemeColor;
const SH = pres.shapes;

const W = 10, H = 7.5, MX = 0.5;
const BODY_Y = 1.5, BODY_B = 6.7;           // body area
const warnings = [];
let CUR = { title: '' };

/* ------------------------------------------------------------------ icons */
const ICON_NAMES = [
  'FaBitcoin','FaKey','FaFingerprint','FaLink','FaCoins','FaServer','FaUsers','FaBolt','FaLock','FaLockOpen','FaCubes','FaBalanceScale',
  'FaChartLine','FaLeaf','FaGlobe','FaShieldAlt','FaExclamationTriangle','FaCheckCircle','FaTimesCircle','FaQuestionCircle','FaLightbulb',
  'FaBook','FaGavel','FaPizzaSlice','FaMicrochip','FaNetworkWired','FaSearch','FaHandshake','FaFileSignature','FaDatabase','FaClock',
  'FaFire','FaUniversity','FaWallet','FaCube','FaRandom','FaSitemap','FaIdCard','FaEnvelope','FaMoneyBillWave','FaIndustry','FaMask',
  'FaUserSecret','FaPlug','FaCodeBranch','FaProjectDiagram','FaPlay','FaStar','FaFlask','FaComments','FaPen','FaCalculator','FaTrophy',
  'FaRegClock','FaMapMarkerAlt','FaHourglassHalf','FaThumbsUp','FaThumbsDown','FaEye','FaMobileAlt','FaCreditCard','FaFileAlt','FaStamp',
];
const ICON_COLORS = ['FFFFFF', 'F7931A', 'B45309', '1F2430', '0F766E', 'B42318', '475569'];
const iconCache = {};
async function warmIcons() {
  for (const n of ICON_NAMES) for (const c of ICON_COLORS) {
    const Cmp = FA[n];
    if (!Cmp) throw new Error('react-icons/fa has no ' + n);
    const svg = RDS.renderToStaticMarkup(React.createElement(Cmp, { color: '#' + c, size: '256' }));
    const buf = await sharp(Buffer.from(svg)).resize(192, 192, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
    iconCache[n + c] = 'image/png;base64,' + buf.toString('base64');
  }
}
function ic(name, color) {
  const k = name + color;
  if (!iconCache[k]) throw new Error('icon not warmed: ' + k);
  return iconCache[k];
}

/* ------------------------------------------------------------ text helpers */
// markup:  **bold**   [[orange emphasis]]   ((red warning))   <<teal ok>>
function parseMarkup(s) {
  const out = [];
  const re = /(\*\*.+?\*\*|\[\[.+?\]\]|\(\(.+?\)\)|<<.+?>>)/gs;
  let last = 0, m;
  while ((m = re.exec(s))) {
    if (m.index > last) out.push({ text: s.slice(last, m.index), o: {} });
    const tok = m[0], inner = tok.slice(2, -2);
    if (tok.startsWith('**')) out.push({ text: inner, o: { bold: true } });
    else if (tok.startsWith('[[')) out.push({ text: inner, o: { bold: true, color: C.accent6 } });
    else if (tok.startsWith('((')) out.push({ text: inner, o: { bold: true, color: C.accent2 } });
    else out.push({ text: inner, o: { bold: true, color: C.accent3 } });
    last = m.index + tok.length;
  }
  if (last < s.length) out.push({ text: s.slice(last), o: {} });
  return out;
}
const strip = (s) => s.replace(/\*\*|\[\[|\]\]|\(\(|\)\)|<<|>>/g, '');

function cw(ch, pt) {
  const c = ch.codePointAt(0);
  if (c >= 0x2E80 || (c >= 0xFF00 && c <= 0xFFEF) || (c >= 0x2000 && c <= 0x206F && ch !== ' ')) return pt / 72;
  if (ch === ' ') return pt * 0.3 / 72;
  if (/[A-Z0-9@#%&]/.test(ch)) return pt * 0.64 / 72;
  if (/[iljtf.,:;'!|]/.test(ch)) return pt * 0.3 / 72;
  return pt * 0.54 / 72;
}
function nLines(text, widthIn, pt) {
  let n = 0;
  for (const para of strip(text).split('\n')) {
    let w = 0, l = 1;
    for (const ch of para) { const d = cw(ch, pt); if (w + d > widthIn) { l++; w = d; } else w += d; }
    n += l;
  }
  return n;
}
// height needed (inches) for a list of paragraphs at a given size
function measure(paras, widthIn, pt, opts = {}) {
  const lh = opts.lineSpacing || 1.32, gap = opts.gap ?? 0.07;
  let h = 0;
  paras.forEach((p, i) => {
    const t = typeof p === 'string' ? p : p.t;
    const sz = (typeof p === 'object' && p.size) || pt;
    const indent = (typeof p === 'object' && p.bullet) || opts.bullet ? 0.28 : 0;
    h += nLines(t, widthIn - indent, sz) * sz * lh / 72 + (i < paras.length - 1 ? gap : 0);
  });
  return h;
}
function warn(msg) { warnings.push(`[${CUR.title}] ${msg}`); }

function toRuns(paras, o = {}) {
  const runs = [];
  paras.forEach((p, i) => {
    const po = typeof p === 'string' ? { t: p } : p;
    const last = i === paras.length - 1;
    const parts = parseMarkup(po.t);
    parts.forEach((r, j) => {
      const opt = { ...r.o };
      if (po.size) opt.fontSize = po.size;
      if (po.color) opt.color = po.color;
      if (po.bold) opt.bold = true;
      if (po.italic) opt.italic = true;
      if (po.align) opt.align = po.align;
      if (po.bullet || o.bullet) opt.bullet = { indent: 20 };
      if (!last || po.breakLast) opt.breakLine = j === parts.length - 1;
      else opt.breakLine = false;
      if (j !== parts.length - 1) opt.breakLine = false;
      if (i < paras.length - 1) opt.paraSpaceAfter = po.gap ?? o.gap_pt ?? 5;
      runs.push({ text: r.text, options: opt });
    });
  });
  return runs;
}

// generic text box (with auto-overflow warning)
function T(slide, paras, o = {}) {
  if (typeof paras === 'string') paras = [paras];
  const size = o.size || 18;
  const w = o.w, h = o.h;
  if (h && w) {
    const need = measure(paras, w - (o.pad ?? 0) * 2, size, { bullet: o.bullet, lineSpacing: o.lh });
    if (need > h - (o.pad ?? 0) * 2 + 0.02) warn(`text overflow: needs ${need.toFixed(2)}in, box ${h.toFixed(2)}in :: ${strip(typeof paras[0] === 'string' ? paras[0] : paras[0].t).slice(0, 26)}`);
  }
  slide.addText(toRuns(paras, o), {
    x: o.x, y: o.y, w, h,
    fontSize: size, color: o.color || C.text1, bold: o.bold, italic: o.italic,
    align: o.align || 'left', valign: o.valign || 'top',
    margin: o.pad ?? 0, lineSpacingMultiple: o.lh ? o.lh * 0.92 : 1.05,
    fit: 'none', isTextBox: true, objectName: o.name,
    fontFace: o.mono ? 'Consolas' : undefined,
  });
}

/* ---------------------------------------------------------------- layouts */
function defineLayouts() {
  const footer = (dark) => ({
    text: { text: '数字货币 · 第四章 比特币', options: { x: MX, y: 7.08, w: 4, h: 0.28, fontSize: 10, color: dark ? C.accent5 : C.accent5, margin: 0, isTextBox: true } },
  });
  pres.defineSlideMaster({
    title: 'COVER', background: { color: HEX.dk1 },
    objects: [],
  });
  pres.defineSlideMaster({
    title: 'SECTION', background: { color: HEX.dk1 },
    objects: [],
  });
  pres.defineSlideMaster({
    title: 'CONTENT', background: { color: HEX.lt1 },
    objects: [
      footer(false),
      { placeholder: { options: { name: 'title', type: 'title', x: MX, y: 0.52, w: 9, h: 0.95, fontSize: 28, bold: true, color: HEX.dk1, align: 'left', valign: 'top', margin: 0 }, text: '' } },
    ],
    slideNumber: { x: 9.0, y: 7.08, w: 0.5, h: 0.28, fontSize: 10, color: HEX.accent5, align: 'right' },
  });
  pres.defineSlideMaster({
    title: 'CONTENT_DARK', background: { color: HEX.dk1 },
    objects: [
      footer(true),
      { placeholder: { options: { name: 'title', type: 'title', x: MX, y: 0.52, w: 9, h: 0.95, fontSize: 28, bold: true, color: HEX.lt1, align: 'left', valign: 'top', margin: 0 }, text: '' } },
    ],
    slideNumber: { x: 9.0, y: 7.08, w: 0.5, h: 0.28, fontSize: 10, color: HEX.accent5, align: 'right' },
  });
}

/* ------------------------------------------------------------- slide makers */
let currentSection = null;
function section(title) { pres.addSection({ title }); currentSection = title; }

// standard content slide: tag (small orange label), assertion title, notes, optional source
function S(o) {
  CUR = { title: o.title };
  const dark = !!o.dark;
  const s = pres.addSlide({ masterName: dark ? 'CONTENT_DARK' : 'CONTENT', sectionTitle: currentSection });
  s.addText(o.title, { placeholder: 'title' });
  if (nLines(o.title, 9, 28) > 2) warn('title > 2 lines');
  if (o.tag) s.addText(o.tag, { x: MX, y: 0.22, w: 6, h: 0.26, fontSize: 12, bold: true, color: dark ? C.accent1 : C.accent6, margin: 0, isTextBox: true, objectName: 'tag' });
  if (o.src) s.addText('来源：' + o.src, { x: MX, y: 6.76, w: 9, h: 0.26, fontSize: 10, color: dark ? C.accent5 : C.accent4, margin: 0, isTextBox: true, objectName: 'source', valign: 'top' });
  if (o.notes) s.addNotes(o.notes);
  s._dark = dark;
  return s;
}

/* ------------------------------------------------------------- components */
function rect(s, x, y, w, h, o = {}) {
  s.addShape(o.round === false ? SH.RECTANGLE : SH.ROUNDED_RECTANGLE, {
    x, y, w, h, fill: { color: o.fill || C.background2, transparency: o.tr || 0 },
    line: o.line ? { color: o.line, width: o.lw || 1 } : { color: o.fill || C.background2, width: 0, transparency: 100 },
    rectRadius: o.round === false ? undefined : (o.r ?? 0.08), objectName: o.name,
  });
}
function circle(s, x, y, d, fill, o = {}) {
  s.addShape(SH.OVAL, { x, y, w: d, h: d, fill: { color: fill, transparency: o.tr || 0 }, line: { color: fill, width: 0, transparency: 100 }, objectName: o.name });
}
function iconBadge(s, name, x, y, d, bg, fg) {
  circle(s, x, y, d, bg);
  const p = d * 0.22;
  s.addImage({ data: ic(name, fg), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, altText: name });
}
function numBadge(s, n, x, y, d, bg, fg, size) {
  circle(s, x, y, d, bg);
  s.addText(String(n), { x, y, w: d, h: d, align: 'center', valign: 'middle', fontSize: size || Math.round(d * 28), bold: true, color: fg || C.background1, margin: 0, isTextBox: true });
}
function chip(s, text, x, y, w, h, fill, color, size) {
  s.addShape(SH.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0, transparency: 100 }, rectRadius: h / 2 });
  s.addText(text, { x, y, w, h, align: 'center', valign: 'middle', fontSize: size || 12, bold: true, color: color || C.background1, margin: 0, isTextBox: true });
}

// bullets list
function bullets(s, items, o = {}) {
  const paras = items.map((t) => (typeof t === 'string' ? { t, bullet: true } : { bullet: true, ...t }));
  T(s, paras, { x: o.x ?? MX, y: o.y ?? BODY_Y, w: o.w ?? 9, h: o.h ?? 4.6, size: o.size || 18, gap_pt: o.gap || 9, color: o.color, lh: o.lh });
}

// grid of cards: [{icon, h, b, n, fill, hc}]
function cards(s, items, o = {}) {
  const cols = o.cols || items.length, rows = Math.ceil(items.length / cols);
  const x0 = o.x ?? MX, y0 = o.y ?? BODY_Y, W0 = o.w ?? 9, H0 = o.h ?? (BODY_B - BODY_Y), gap = o.gap ?? 0.2;
  const cwid = (W0 - gap * (cols - 1)) / cols, chei = (H0 - gap * (rows - 1)) / rows;
  const hs = o.hsize || 18, bs = o.size || 15;
  items.forEach((it, i) => {
    const cx = x0 + (i % cols) * (cwid + gap), cy = y0 + Math.floor(i / cols) * (chei + gap);
    const dark = o.dark;
    rect(s, cx, cy, cwid, chei, { fill: it.fill || (dark ? C.text2 : C.background2) });
    let ty = cy + 0.18;
    const px = 0.2;
    if (it.icon) { iconBadge(s, it.icon, cx + px, ty, 0.55, it.badge || C.accent1, it.fg || (it.badge ? 'FFFFFF' : '1F2430')); ty += 0.68; }
    else if (it.n !== undefined) { numBadge(s, it.n, cx + px, ty, 0.5, it.badge || C.accent1, C.text1, 16); ty += 0.62; }
    const hh = measure([it.h], cwid - 2 * px, hs) ;
    T(s, [{ t: it.h, bold: true }], { x: cx + px, y: ty, w: cwid - 2 * px, h: hh + 0.02, size: hs, color: dark ? C.background1 : C.text1 });
    ty += hh + 0.08;
    if (it.b) {
      const paras = Array.isArray(it.b) ? it.b : [it.b];
      const rem = cy + chei - ty - 0.12;
      T(s, paras, { x: cx + px, y: ty, w: cwid - 2 * px, h: rem, size: bs, color: dark ? C.background2 : C.accent4, gap_pt: 4 });
    }
  });
}

// big stat
function stat(s, big, label, o = {}) {
  const x = o.x, y = o.y, w = o.w ?? 2.8, h = o.h ?? 1.5;
  rect(s, x, y, w, h, { fill: o.fill || (o.dark ? C.text2 : C.background2) });
  T(s, [{ t: big, bold: true }], { x: x + 0.15, y: y + 0.12, w: w - 0.3, h: h * 0.5, size: o.bigSize || 34, color: o.color || (o.dark ? C.accent1 : C.accent6), align: 'center', valign: 'middle' });
  T(s, label, { x: x + 0.15, y: y + h * 0.58, w: w - 0.3, h: h * 0.38, size: o.size || 14, color: o.dark ? C.background2 : C.accent4, align: 'center', valign: 'top' });
}

// callout: tinted box with a leading icon
function callout(s, text, o = {}) {
  const kind = o.kind || 'key';
  const map = { key: ['FaLightbulb', C.accent1, 'F7931A'], warn: ['FaExclamationTriangle', C.accent2, 'B42318'], ok: ['FaCheckCircle', C.accent3, '0F766E'], ask: ['FaQuestionCircle', C.accent1, 'F7931A'] };
  const [icn, col] = map[kind];
  const x = o.x ?? MX, y = o.y, w = o.w ?? 9, h = o.h ?? 0.9;
  s.addShape(SH.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: col, transparency: 88 }, line: { color: col, width: 0, transparency: 100 }, rectRadius: 0.08 });
  const iconName = o.icon || icn;
  const fg = kind === 'warn' ? 'B42318' : kind === 'ok' ? '0F766E' : 'B45309';
  s.addImage({ data: ic(iconName, fg), x: x + 0.2, y: y + (h - 0.42) / 2, w: 0.42, h: 0.42, altText: kind });
  T(s, Array.isArray(text) ? text : [text], { x: x + 0.8, y: y + 0.08, w: w - 1.0, h: h - 0.16, size: o.size || 16, valign: 'middle', gap_pt: 3 });
}

// quote block
function quote(s, text, who, o = {}) {
  const x = o.x ?? MX, y = o.y ?? BODY_Y, w = o.w ?? 9, h = o.h ?? 3.4, dark = o.dark;
  rect(s, x, y, w, h, { fill: dark ? C.text2 : C.background2 });
  s.addText('“', { x: x + 0.25, y: y + 0.05, w: 0.8, h: 0.9, fontSize: 60, bold: true, color: C.accent1, margin: 0, isTextBox: true, valign: 'top' });
  T(s, text, { x: x + 0.9, y: y + 0.3, w: w - 1.3, h: h - (who ? 1.0 : 0.5), size: o.size || 18, color: dark ? C.background1 : C.text1, gap_pt: 6, lh: 1.35 });
  if (who) T(s, who, { x: x + 0.9, y: y + h - 0.6, w: w - 1.3, h: 0.4, size: 13, color: dark ? C.accent5 : C.accent4, valign: 'top' });
}

const ALT = {
  'email.png': '中本聪 2008 年 10 月 31 日发给密码学邮件列表的邮件截图', 'wp_abstract.png': '比特币白皮书首页与摘要', 'navajo.png': '二战中的纳瓦霍密语译电员',
  'turing.png': '艾伦·图灵肖像', 'xm_claim_bad.png': '漫画：小红声称收到小明付款', 'xm_claim_good.png': '漫画：小明声称付款给小红', 'sign_flow.png': '数字签名与验证流程示意',
  'bitcoin_card.png': '“比特币银行卡”示意图', 'mysql_md5.png': '软件下载页上的 MD5 校验值', 'merkle4.png': '四笔交易的默克尔树', 'merkle_tamper.png': '篡改一笔交易后默克尔根完全改变',
  'block_hdr.png': '区块头与交易列表示意', 'nonce_list.png': '不同 nonce 对应的哈希值列表', 'poly1.png': '漫画：小红出的方程题', 'coinbase.png': 'coinbase 交易示意图',
  'chart_conf_time.png': '比特币平均确认时间图表（blockchain.com）', 'chart_fee.png': '比特币每笔交易手续费图表（blockchain.com）', 'mtgox.png': 'Mt.Gox 债权人举牌抗议',
  'tweet_kanye.png': '2020 年被盗推特账号发布的比特币诈骗推文', 'asic.png': '比特币专用矿机（ASIC）', 'farm.png': '俄罗斯 Bratsk 的加密货币矿场', 'ndrc_notice.png': '国家发改委等《关于整治虚拟货币“挖矿”活动的通知》',
  'carney.png': '马克·卡尼肖像', 'carstens.png': '阿古斯丁·卡斯滕斯肖像', 's2f_chart.png': '存量流量比（S2F）模型价格图', 's2f_panel.png': 'S2F 模型当前数值面板（2025-10）', 'silkroad.png': '丝绸之路网站页面截图',
  'wp_2_transactions.png': '白皮书第 2 节 交易', 'wp_3_timestamp.png': '白皮书第 3 节 时间戳服务器', 'wp_4_pow.png': '白皮书第 4 节 工作量证明', 'wp_5_network.png': '白皮书第 5 节 网络',
  'wp_6_incentive.png': '白皮书第 6 节 激励', 'wp_7_disk.png': '白皮书第 7 节 回收磁盘空间', 'wp_8_spv.png': '白皮书第 8 节 简化支付验证',
};

// picture with contain-fit inside a box
const MAN = JSON.parse(fs.readFileSync(path.join(__dirname, 'assets', 'manifest.json'), 'utf8'));
function img(s, key, o = {}) {
  const m = MAN[key]; if (!m) throw new Error('no asset ' + key);
  const bx = o.x, by = o.y, bw = o.w, bh = o.h;
  let w = bw, h = bw * m.h / m.w;
  if (h > bh) { h = bh; w = bh * m.w / m.h; }
  const x = bx + (o.align === 'left' ? 0 : o.align === 'right' ? bw - w : (bw - w) / 2);
  const y = by + (o.valign === 'top' ? 0 : o.valign === 'bottom' ? bh - h : (bh - h) / 2);
  const p = path.join(__dirname, 'assets', key);
  if (o.border !== false) s.addShape(SH.RECTANGLE, { x: x - 0.02, y: y - 0.02, w: w + 0.04, h: h + 0.04, fill: { color: C.background1 }, line: { color: C.accent5, width: 0.75 } });
  s.addImage({ path: p, x, y, w, h, altText: o.alt || ALT[key] || key });
  if (o.caption) T(s, o.caption, { x: bx, y: y + h + 0.06, w: bw, h: 0.3, size: 11, color: C.accent4, align: 'center' });
  return { x, y, w, h };
}

// flow of boxes joined by arrows (horizontal)
function flow(s, steps, o = {}) {
  const n = steps.length, x0 = o.x ?? MX, y0 = o.y, W0 = o.w ?? 9, bh = o.h ?? 1.0, ag = o.arrow ?? 0.32;
  const bw = (W0 - ag * (n - 1)) / n;
  steps.forEach((st, i) => {
    const x = x0 + i * (bw + ag);
    const fill = st.fill || (o.dark ? C.text2 : C.background2);
    rect(s, x, y0, bw, bh, { fill });
    T(s, [{ t: st.h, bold: true, size: o.hsize || 15 }].concat(st.b ? [{ t: st.b, size: o.size || 12, color: o.dark ? C.accent5 : C.accent4 }] : []), { x: x + 0.08, y: y0 + 0.06, w: bw - 0.16, h: bh - 0.12, size: o.size || 12, align: 'center', valign: 'middle', color: o.dark ? C.background1 : C.text1, gap_pt: 2 });
    if (i < n - 1) s.addShape(SH.RIGHT_ARROW, { x: x + bw + 0.04, y: y0 + bh / 2 - 0.11, w: ag - 0.08, h: 0.22, fill: { color: C.accent1 }, line: { color: C.accent1, width: 0, transparency: 100 } });
  });
  return { bw };
}

// vertical timeline with date chips
function timelineV(s, items, o = {}) {
  const x0 = o.x ?? MX, y0 = o.y ?? BODY_Y, W0 = o.w ?? 9, H0 = o.h ?? (BODY_B - BODY_Y);
  const n = items.length, rh = H0 / n, cwid = o.chipW ?? 1.35;
  s.addShape(SH.LINE, { x: x0 + cwid + 0.22, y: y0 + rh / 2, w: 0, h: rh * (n - 1), line: { color: C.accent5, width: 1.5 } });
  items.forEach((it, i) => {
    const y = y0 + i * rh;
    const hot = it.hot;
    chip(s, it.when, x0, y + rh / 2 - 0.17, cwid, 0.34, hot ? C.accent1 : C.text1, hot ? C.text1 : C.background1, 12);
    circle(s, x0 + cwid + 0.14, y + rh / 2 - 0.08, 0.16, hot ? C.accent1 : C.accent4);
    const paras = [{ t: it.h, bold: true, size: o.hsize || 16 }].concat(it.b ? [{ t: it.b, size: o.size || 13, color: C.accent4 }] : []);
    T(s, paras, { x: x0 + cwid + 0.5, y: y + 0.02, w: W0 - cwid - 0.5, h: rh - 0.04, size: o.size || 13, valign: 'middle', gap_pt: 1 });
  });
}

// table
function table(s, head, rows, o = {}) {
  const x = o.x ?? MX, y = o.y ?? BODY_Y, w = o.w ?? 9;
  const fs_ = o.size || 14;
  const hdr = head.map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, fontSize: fs_, align: o.align || 'left', valign: 'middle' } }));
  const body = rows.map((r, ri) => r.map((c, ci) => {
    const isObj = typeof c === 'object' && c !== null && !Array.isArray(c);
    const txt = isObj ? c.t : c;
    const runs = parseMarkup(String(txt)).map((p) => ({ text: p.text, options: { ...p.o } }));
    return { text: runs, options: { fontSize: fs_, color: C.text1, fill: { color: ri % 2 ? C.background1 : C.background2 }, valign: 'middle', bold: ci === 0 && o.boldFirst !== false, align: (isObj && c.align) || (o.align || 'left') } };
  }));
  s.addTable([hdr, ...body], {
    x, y, w, colW: o.colW, rowH: (typeof o.rowH === 'number') ? [o.hdrH ?? 0.5].concat(rows.map(() => o.rowH)) : o.rowH, border: { type: 'solid', color: HEX.lt2, pt: 0.5 },
    margin: [0.05, 0.1, 0.05, 0.1], autoPage: false,
  });
}

/* ------------------------------------------------------------------ charts */
function chartBase(extra = {}) {
  return Object.assign({
    catAxisLabelColor: HEX.accent4, valAxisLabelColor: HEX.accent4,
    catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
    catAxisLabelFontFace: '+mn-lt', valAxisLabelFontFace: '+mn-lt',
    dataLabelColor: HEX.dk1, dataLabelFontSize: 12, dataLabelFontFace: '+mn-lt',
    valGridLine: { color: 'E5E7EB', size: 0.5 }, catGridLine: { style: 'none' },
    legendFontSize: 12, legendColor: HEX.accent4, legendFontFace: '+mn-lt',
    titleFontFace: '+mn-lt', titleColor: HEX.dk1, titleFontSize: 14,
  }, extra);
}




// quiz slide: qs = [{q, o (options/hint), kind}]
function quiz(o) {
  const s = S({ tag: o.tag, title: o.title, notes: o.notes });
  const n = o.qs.length, gap = 0.12, rh = Math.min(1.55, (BODY_B - BODY_Y - gap * (n - 1)) / n);
  o.qs.forEach((q, i) => {
    const y = BODY_Y + i * (rh + gap);
    rect(s, MX, y, 9, rh, { fill: C.background2 });
    numBadge(s, 'Q' + (i + 1), MX + 0.2, y + (rh - 0.62) / 2, 0.62, C.accent1, C.text1, 15);
    const paras = [{ t: q.q, bold: true, size: 17 }].concat(q.o ? [{ t: q.o, size: 14, color: C.accent4 }] : []);
    T(s, paras, { x: MX + 1.1, y: y + 0.08, w: 6.7, h: rh - 0.16, size: 15, valign: 'middle', gap_pt: 3 });
    chip(s, q.kind || '选择', MX + 8.2 - 0.2, y + rh / 2 - 0.2, 0.85, 0.4, C.text1, C.background1, 12);
  });
  return s;
}

/* ------------------------------------------------------------ diagram tools */
function arrow(s, x1, y1, x2, y2, o = {}) {
  const x = Math.min(x1, x2), y = Math.min(y1, y2), w = Math.abs(x2 - x1), h = Math.abs(y2 - y1);
  s.addShape(SH.LINE, {
    x, y, w, h, flipH: x1 > x2, flipV: y1 > y2,
    line: { color: o.color || C.accent4, width: o.width || 2, dashType: o.dash, endArrowType: o.noHead ? undefined : 'triangle', beginArrowType: o.both ? 'triangle' : undefined },
  });
}
function label(s, text, x, y, w, h, o = {}) {
  T(s, text, { x, y, w, h, size: o.size || 12, color: o.color || C.accent4, align: o.align || 'center', valign: o.valign || 'middle', bold: o.bold, mono: o.mono, gap_pt: 2 });
}
// nodes joined by labelled arrows. nodes: [{h,b,fill,color}], ops: ['label above arrow', ...]
function flowL(s, nodes, ops, o = {}) {
  const n = nodes.length, x0 = o.x ?? MX, y0 = o.y, W0 = o.w ?? 9, bh = o.h ?? 1.0, ag = o.ag ?? 1.5;
  const bw = (W0 - ag * (n - 1)) / n;
  nodes.forEach((nd, i) => {
    const x = x0 + i * (bw + ag);
    rect(s, x, y0, bw, bh, { fill: nd.fill || (o.dark ? C.text2 : C.background2) });
    T(s, [{ t: nd.h, bold: true, size: o.hsize || 15 }].concat(nd.b ? [{ t: nd.b, size: o.size || 12, color: nd.color || (o.dark ? C.accent5 : C.accent4) }] : []),
      { x: x + 0.08, y: y0 + 0.05, w: bw - 0.16, h: bh - 0.1, size: o.size || 12, align: 'center', valign: 'middle', color: nd.hcolor || (o.dark ? C.background1 : C.text1), gap_pt: 2 });
    if (i < n - 1) {
      arrow(s, x + bw + 0.05, y0 + bh / 2, x + bw + ag - 0.05, y0 + bh / 2, { color: C.accent1, width: 2.5 });
      if (ops && ops[i]) label(s, ops[i], x + bw + 0.02, y0 + bh / 2 - 0.62, ag - 0.04, 0.55, { size: o.opSize || 12, valign: 'bottom', color: o.dark ? C.accent5 : C.accent4 });
    }
  });
  return { bw };
}
// small mesh network of n nodes on a circle
function mesh(s, cx, cy, r, n, o = {}) {
  const pts = [];
  for (let i = 0; i < n; i++) { const a = -Math.PI / 2 + (2 * Math.PI * i) / n; pts.push([cx + r * Math.cos(a), cy + r * Math.sin(a)]); }
  for (let i = 0; i < n; i++) for (let j = i + 1; j < n; j++) arrow(s, pts[i][0], pts[i][1], pts[j][0], pts[j][1], { color: C.accent5, width: 1, noHead: true });
  const d = o.d || 0.62;
  pts.forEach(([x, y], i) => {
    const bad = o.bad && o.bad.includes(i);
    circle(s, x - d / 2, y - d / 2, d, bad ? C.accent2 : (o.fill || C.accent1));
    s.addImage({ data: ic(bad ? (o.badIcon || o.icon || 'FaDatabase') : (o.icon || 'FaDatabase'), bad ? 'FFFFFF' : '1F2430'), x: x - d * 0.28, y: y - d * 0.28, w: d * 0.56, h: d * 0.56, altText: 'node' });
    if (o.labels) label(s, o.labels[i], x - 0.7, y + d / 2 + 0.02, 1.4, 0.3, { size: 11 });
  });
  return pts;
}

/* -------------------------------------------------------- cover / section */
const REQS = ['① 识别交易双方', '② 便捷促成并验证交易', '③ 记录可验证、不可更改', '④ 防止同一单位货币重复使用'];
function cover(o) {
  CUR = { title: 'cover' };
  const s = pres.addSlide({ masterName: 'COVER', sectionTitle: currentSection });
  rect(s, MX, 0.55, 2.9, 1.85, { fill: C.background1, r: 0.1 });
  const m = MAN['logo.jpeg']; const lw = 2.6, lh = lw * m.h / m.w;
  s.addImage({ path: path.join(__dirname, 'assets', 'logo.jpeg'), x: MX + (2.9 - lw) / 2, y: 0.55 + (1.85 - lh) / 2, w: lw, h: lh, altText: '江西财经大学校徽' });
  s.addImage({ data: ic('FaBitcoin', 'F7931A'), x: 6.6, y: 0.7, w: 2.6, h: 2.6, altText: 'bitcoin' });
  s.addText(o.title, { x: MX, y: 3.25, w: 9, h: 1.0, fontSize: 54, bold: true, color: C.background1, margin: 0, isTextBox: true, valign: 'middle' });
  s.addText(o.subtitle, { x: MX, y: 4.35, w: 9, h: 0.7, fontSize: 24, color: C.accent1, margin: 0, isTextBox: true, valign: 'top' });
  s.addText(o.lines.map((t, i) => ({ text: t, options: { breakLine: i < o.lines.length - 1 } })), { x: MX, y: 5.45, w: 9, h: 1.2, fontSize: 14, color: C.accent5, margin: 0, isTextBox: true, valign: 'top', paraSpaceAfter: 4 });
  if (o.notes) s.addNotes(o.notes);
  return s;
}
function sectionSlide(o) {
  CUR = { title: 'section ' + o.n };
  const s = pres.addSlide({ masterName: 'SECTION', sectionTitle: currentSection });
  s.addText(o.n, { x: MX, y: 0.7, w: 4, h: 1.5, fontSize: 96, bold: true, color: C.accent1, margin: 0, isTextBox: true, valign: 'top' });
  if (o.icon) s.addImage({ data: ic(o.icon, 'F7931A'), x: 7.0, y: 0.9, w: 2.0, h: 2.0, altText: o.icon, transparency: 25 });
  s.addText(o.title, { x: MX, y: 2.55, w: 9, h: 1.3, fontSize: 40, bold: true, color: C.background1, margin: 0, isTextBox: true, valign: 'top' });
  if (nLines(o.title, 9, 40) > 2) warn('section title > 2 lines');
  s.addText(o.sub, { x: MX, y: 3.95, w: 8.6, h: 1.1, fontSize: 20, color: C.accent5, margin: 0, isTextBox: true, valign: 'top' });
  if (o.req) {
    s.addText('这一部分回答支付系统的哪个要求？', { x: MX, y: 5.2, w: 6, h: 0.3, fontSize: 13, color: C.accent5, margin: 0, isTextBox: true });
    REQS.forEach((t, i) => {
      const hot = o.req.includes(i), x = MX + (i % 2) * 4.55, y = 5.6 + Math.floor(i / 2) * 0.55;
      chip(s, t, x, y, 4.4, 0.42, hot ? C.accent1 : C.text2, hot ? C.text1 : C.accent5, 13);
    });
  }
  if (o.notes) s.addNotes(o.notes);
  return s;
}

/* ------------------------------------------------------------ finishing */
async function finish(file) {
  const { applyTheme } = require(process.env.SKILL_DIR + '/scripts/apply_theme.js');
  await pres.writeFile({ fileName: file });
  await applyTheme(file, THEME);
  // CJK: give the theme an East-Asian font so Chinese text never falls back to a serif face
  const JSZip = require(require.resolve('jszip', { paths: [require.resolve('pptxgenjs')] }));
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  let xml = await zip.file('ppt/theme/theme1.xml').async('string');
  xml = xml.replace(/<a:ea typeface="[^"]*"\s*\/>/g, `<a:ea typeface="${FONT}"/>`);
  zip.file('ppt/theme/theme1.xml', xml);
  fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

module.exports = {
  quiz, arrow, label, flowL, mesh, cover, sectionSlide, finish, REQS,
  pres, C, SH, THEME, HEX, W, H, MX, BODY_Y, BODY_B, FONT,
  warmIcons, ic, defineLayouts, section, S, T, rect, circle, iconBadge, numBadge, chip, bullets, cards, stat, callout, quote,
  img, flow, timelineV, table, chartBase, measure, nLines, warnings, MAN, parseMarkup, strip,
};
