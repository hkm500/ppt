const L = require('./lib'); const { pres, C, HEX } = L;
(async () => {
  await L.warmIcons(); L.defineLayouts();
  L.section('开场');
  L.cover({ title: '第四章 比特币', subtitle: '没有银行的支付系统，是怎样炼成的？', lines: ['数字货币课程 · 2026 年 10 月修订版', '基于 Eswar Prasad《The Future of Money》第四章，并据白皮书与最新资料增订'] });
  L.sectionSlide({ n: '01', title: '积木一：密码学与公私钥', sub: '没有银行，怎么证明“这笔钱是我的、是我愿意付的”？', icon: 'FaKey', req: [0, 1] });
  let s = L.S({ tag: '积木一 · 密码学', title: '数字签名：不可伪造、不可篡改、不可抵赖', notes: 'x', src: '白皮书 §2' });
  L.cards(s, [
    { icon: 'FaFileSignature', h: '不可伪造', b: '私钥只有签名人自己知道，别人造不出有效签名。' },
    { icon: 'FaLock', h: '不可篡改', b: '消息被改动一个字节，验证立刻失败。' },
    { icon: 'FaHandshake', h: '不可抵赖', b: '验证通过，就说明确实是私钥持有人签的。' },
  ], { y: 1.6, h: 2.9, cols: 3 });
  L.callout(s, '比特币**不加密任何东西**：账本是公开的，密码学只用来“签名”和“做指纹”。', { y: 4.8, h: 1.0 });
  s = L.S({ tag: '对比', title: '四种支付方式，凭什么证明你有权花这笔钱？' });
  L.table(s, ['', '凭什么证明', '被盗/泄露后', '谁来背书'], [
    ['纸币', '占有实物', '丢了就没了', '央行'],
    ['支票', '收款人签名', '可止付', '银行'],
    ['借记卡', '卡 + PIN 码', '可挂失', '银行'],
    ['比特币', '**私钥签名**', '((无法找回))', '协议与全网'],
  ], { colW: [1.5, 2.6, 2.5, 2.4], y: 1.7, rowH: 0.7 });
  s = L.S({ tag: '数据', title: '减半让比特币越来越稀缺：已发行比例逼近 100%', src: '协议规则；作者按 210,000 块/次计算' });
  s.addChart(pres.charts.BAR, [{ name: '累计已发行比例', labels: ['2012', '2016', '2020', '2024', '2028'], values: [50, 75, 87.5, 93.75, 96.875] }],
    L.chartBase({ x: 0.5, y: 1.6, w: 6.0, h: 4.8, barDir: 'col', chartColors: [HEX.accent1], showValue: true, dataLabelPosition: 'outEnd', valAxisMaxVal: 100, valAxisMinVal: 0, showLegend: false, dataLabelFormatCode: '0.0"%"', valAxisLabelFormatCode: '0"%"' }));
  L.stat(s, '93.75%', '2024 年减半后已发行', { x: 6.8, y: 1.7, w: 2.7, h: 1.6 });
  L.stat(s, '≈2140', '最后一枚比特币发行', { x: 6.8, y: 3.5, w: 2.7, h: 1.6 });
  s = L.S({ tag: '素材', title: '白皮书 §2：一枚电子币，就是一条数字签名链' });
  L.img(s, 'wp_2_transactions.png', { x: 0.5, y: 1.55, w: 5.4, h: 4.2, caption: '来源：Satoshi Nakamoto (2008), Bitcoin white paper, §2' });
  L.bullets(s, ['每个所有者用**私钥**在上一笔交易的哈希与下一位所有者的**公钥**上签名', '任何人用公钥即可**验证**所有权链条'], { x: 6.2, y: 1.7, w: 3.3, h: 4, size: 16 });
  s = L.S({ tag: '流程', title: '一笔交易的旅程' });
  L.flow(s, [{ h: '签名', b: '私钥签署' }, { h: '广播', b: 'P2P 网络' }, { h: '打包', b: '矿工入块' }, { h: '确认', b: '6 个区块' }], { y: 1.8, h: 1.2 });
  L.timelineV(s, [
    { when: '1982', h: 'Chaum 盲签名', b: '第一个匿名电子现金设想' },
    { when: '1997', h: 'Hashcash', b: '用 CPU 代价对抗垃圾邮件', hot: true },
    { when: '2008', h: '比特币白皮书', b: '10 月 31 日发布' },
  ], { y: 3.4, h: 3.0 });
  s = L.S({ dark: true, tag: '中本聪', title: '问题的根源是信任' });
  L.quote(s, ['传统货币的根本问题，在于它的运行完全依赖信任。我们必须信任中央银行不会滥发货币……'], '中本聪，P2P Foundation，2009-02-11', { dark: true, y: 1.7, h: 3.2 });
  await L.finish('out/sample.pptx');
  console.log('warnings:', L.warnings);
})();
