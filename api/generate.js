// api/generate.js
// EPP Quarterly Review - PPTX generator
// Receives JSON data from NetSuite, returns a base64 PPTX

const pptxgen = require('pptxgenjs');
const LOGO_B64 = require('../logo.js');

// Shared secret - must match the value in the NetSuite script
const API_TOKEN = process.env.EPP_API_TOKEN || 'kleenrite-epp-2026';

// ---- Theme -----------------------------------------------------------------
const C = {
  black:    '1A1A1A',
  red:      'CC0000',
  darkRed:  '990000',
  white:    'FFFFFF',
  off:      'F9F9F9',
  lgray:    'EEEEEE',
  mgray:    '888888',
  dgray:    '444444',
  redpale:  'FFF0F0',
  redmid:   'FFD5D5',
  blue:     '1F497D',
  bluepale: 'E8F4FF',
  blueline: '5B9BD5',
};
const FONT = 'Calibri';

const money = (n) => {
  const v = Number(n) || 0;
  const s = Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return (v < 0 ? '-$' : '$') + s;
};
// Same value with an explicit + when positive (used on the growth card)
const moneySigned = (n) => (Number(n) > 0 ? '+' : '') + money(n);
const num = (n) => Math.round(Number(n) || 0).toLocaleString('en-US');

// ---- Deck builder ----------------------------------------------------------
function buildDeck(d) {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9';
  pres.title = `${d.customerName} - EPP Quarterly Account Review`;

  const LW = 2.1, LH = LW * (168 / 740), LX = 10 - LW - 0.18, LY = 0.1;
  const logo = (s) => s.addImage({ data: 'image/png;base64,' + LOGO_B64, x: LX, y: LY, w: LW, h: LH });

  const hdr = (s, title, sub) => {
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.65, fill: { color: C.black }, line: { color: C.black } });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.65, w: 10, h: 0.032, fill: { color: C.red }, line: { color: C.red } });
    s.addText(title, { x: 0.3, y: 0, w: 7.3, h: 0.65, fontSize: 18, bold: true, color: C.white, fontFace: FONT, valign: 'middle' });
    if (sub) s.addText(sub, { x: 0.3, y: 0, w: 7.3, h: 0.65, fontSize: 8.5, color: 'AAAAAA', fontFace: FONT, valign: 'bottom', italic: true });
    logo(s);
  };

  const statCard = (s, x, y, w, h, val, lbl, sub, vc) => {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.white }, line: { color: C.lgray, pt: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.05, fill: { color: C.red }, line: { color: C.red } });
    s.addText(val, { x, y: y + 0.06, w, h: h * 0.55, fontSize: 19, bold: true, color: vc || C.black, fontFace: FONT, align: 'center', valign: 'middle' });
    s.addText(lbl, { x, y: y + h * 0.65, w, h: 0.18, fontSize: 7.5, bold: true, color: C.mgray, fontFace: FONT, align: 'center', charSpacing: 0.5 });
    if (sub) s.addText(sub, { x, y: y + h * 0.82, w, h: 0.14, fontSize: 7, color: 'BBBBBB', fontFace: FONT, align: 'center', italic: true });
  };

  // ===== SLIDE 1 - Title =====================================================
  {
    const s = pres.addSlide(); s.background = { color: C.black };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: C.red }, line: { color: C.red } });
    logo(s);
    s.addText('EPP QUARTERLY ACCOUNT REVIEW', { x: 0.38, y: 0.9, w: 8.5, h: 0.4, fontSize: 11, bold: true, color: C.red, fontFace: FONT, charSpacing: 5 });
    s.addText('Account Review &\nGrowth Opportunity', { x: 0.38, y: 1.32, w: 8.5, h: 1.65, fontSize: 38, bold: true, color: C.white, fontFace: FONT, lineSpacingMultiple: 1.12 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.38, y: 2.94, w: 5.4, h: 0.038, fill: { color: C.red }, line: { color: C.red } });
    s.addText(d.customerName, { x: 0.38, y: 3.02, w: 8.5, h: 0.44, fontSize: 18, bold: true, color: C.white, fontFace: FONT });
    s.addText(`Customer #${d.customerNumber}`, { x: 0.38, y: 3.46, w: 8.5, h: 0.3, fontSize: 13, color: C.mgray, fontFace: FONT });
    s.addText(`Period: ${d.dateFrom} - ${d.dateTo}  |  Prior Year: ${d.priorFrom} - ${d.priorTo}`, { x: 0.38, y: 3.78, w: 8.5, h: 0.28, fontSize: 12, color: C.mgray, fontFace: FONT });
    if (d.rollupLabel) {
      s.addShape(pres.shapes.RECTANGLE, { x: 0.38, y: 4.1, w: 3.5, h: 0.26,
        fill: { color: '2A2A2A' }, line: { color: C.red, pt: 0.75 } });
      s.addText(d.rollupLabel, { x: 0.38, y: 4.1, w: 3.5, h: 0.26, fontSize: 8.5,
        bold: true, color: C.red, fontFace: FONT, align: 'center', valign: 'middle' });
    }
    s.addText('Sales Trends  |  Rebate Results  |  Product Opportunity Analysis', { x: 0.38, y: 4.44, w: 8.5, h: 0.28, fontSize: 11, color: '666666', fontFace: FONT, italic: true });
    s.addText(`Generated ${d.generatedDate}  |  Interim Report - Final rebate paid at year end`, { x: 0.38, y: 5.18, w: 8.5, h: 0.22, fontSize: 8, color: '555555', fontFace: FONT, italic: true });
  }

  // ===== SLIDE 2 - Sales Trend ===============================================
  {
    const s = pres.addSlide(); s.background = { color: C.off };
    hdr(s, 'SALES TREND', `Current: ${d.dateFrom} - ${d.dateTo}   vs.   Prior Year: ${d.priorFrom} - ${d.priorTo}` +
      (d.rollupLabel ? '   |   ' + d.rollupLabel : ''));

    statCard(s, 0.28, 0.8, 2.15, 0.9, money(d.currTotal), 'CURRENT PERIOD', `${d.dateFrom}-${d.dateTo}`);
    statCard(s, 2.51, 0.8, 2.15, 0.9, money(d.priorTotal), 'PRIOR YEAR', `${d.priorFrom}-${d.priorTo}`);
    statCard(s, 4.74, 0.8, 2.15, 0.9, moneySigned(d.growth), 'NET GROWTH', 'vs prior year', C.red);
    statCard(s, 6.97, 0.8, 2.15, 0.9, d.growthPct + '%', 'GROWTH RATE', 'year over year', C.red);

    s.addChart(pres.charts.LINE,
      [
        { name: `Current (${d.dateFrom}-${d.dateTo})`, labels: d.weeklyLabels, values: d.currWeekly },
        { name: `Prior Year (${d.priorFrom}-${d.priorTo})`, labels: d.weeklyLabels, values: d.priorWeekly },
      ],
      {
        x: 0.28, y: 1.85, w: 9.44, h: 3.08,
        chartColors: ['CC0000', 'AAAAAA'],
        lineSize: 2.5, lineSmooth: true,
        showLegend: true, legendPos: 'b', legendFontSize: 9,
        showValue: false, showDot: true,
        chartArea: { fill: { color: C.white } },
        plotArea: { fill: { color: C.white } },
        catAxisLabelColor: C.dgray, valAxisLabelColor: C.dgray,
        valGridLine: { color: 'DDDDDD', size: 0.5 }, catGridLine: { style: 'none' },
        valAxisNumFmt: '$#,##0',
        catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
      });

    const ok = !!d.aovOk;
    const aovBg = ok ? 'E8F5E9' : 'FFFBF0';
    const aovBd = ok ? '81C784' : 'FFE082';
    const aovMsg = ok
      ? `Average Order Value: ${money(d.avgOrderValue)} across ${num(d.numOrders)} orders  |  Meets the $300 rebate program minimum  |  Program participation is in good standing.`
      : `Average Order Value: ${money(d.avgOrderValue)} across ${num(d.numOrders)} orders  |  Below the $300 rebate program minimum  |  Program participation could be at risk if average order value is not $300 or more.`;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: 5.22, w: 9.44, h: 0.32, fill: { color: aovBg }, line: { color: aovBd, pt: 0.75 } });
    s.addText(aovMsg, { x: 0.42, y: 5.22, w: 9.1, h: 0.32, fontSize: 8.5, color: '5A4000', fontFace: FONT, valign: 'middle' });
  }

  // ===== SLIDE 3 - Rebate ====================================================
  {
    const s = pres.addSlide(); s.background = { color: C.off };
    hdr(s, 'REBATE CALCULATIONS - CURRENT YEAR STATUS',
      `Interim Report  |  Period: ${d.dateFrom} - ${d.dateTo}  |  Final rebate calculated and paid at year end`);

    const PW = 4.55, GAP = 0.34, LM = 0.28;

    const tier = (xStart, hdrColor, title, desc, rows, totLabel, totVal) => {
      const TOP = 0.82;
      const colW = [1.90, 1.22, 0.58, 0.85];

      s.addShape(pres.shapes.RECTANGLE, { x: xStart, y: TOP, w: PW, h: 0.36, fill: { color: hdrColor }, line: { color: hdrColor } });
      s.addText(title, { x: xStart + 0.1, y: TOP, w: PW - 0.2, h: 0.36, fontSize: 12, bold: true, color: C.white, fontFace: FONT, valign: 'middle' });

      s.addShape(pres.shapes.RECTANGLE, { x: xStart, y: TOP + 0.36, w: PW, h: 0.27, fill: { color: 'F2F2F2' }, line: { color: 'E4E4E4', pt: 0.3 } });
      s.addText(desc, { x: xStart + 0.12, y: TOP + 0.36, w: PW - 0.24, h: 0.27, fontSize: 8, color: C.mgray, fontFace: FONT, valign: 'middle', italic: true });

      const tableRows = [[
        { text: '', options: { fill: { color: 'EBEBEB' } } },
        { text: 'Amount', options: { bold: true, fontSize: 8, color: C.dgray, align: 'right', fill: { color: 'EBEBEB' } } },
        { text: 'Rate', options: { bold: true, fontSize: 8, color: C.dgray, align: 'center', fill: { color: 'EBEBEB' } } },
        { text: 'Rebate', options: { bold: true, fontSize: 8, color: C.dgray, align: 'right', fill: { color: 'EBEBEB' } } },
      ]];
      rows.forEach((r, i) => {
        const bg = i % 2 === 0 ? 'F8F8F8' : 'FFFFFF';
        tableRows.push([
          { text: r[0], options: { fontSize: 9, color: C.black, align: 'left',   fill: { color: bg }, valign: 'middle' } },
          { text: r[1], options: { fontSize: 9, color: C.black, align: 'right',  fill: { color: bg }, valign: 'middle' } },
          { text: r[2], options: { fontSize: 9, color: C.mgray, align: 'center', fill: { color: bg }, valign: 'middle' } },
          { text: r[3], options: { fontSize: 9, color: C.black, align: 'right',  fill: { color: bg }, valign: 'middle', bold: r[3] !== '' } },
        ]);
      });

      s.addTable(tableRows, {
        x: xStart, y: TOP + 0.63, w: PW, colW, rowH: 0.46,
        border: { type: 'solid', color: 'EBEBEB', pt: 0.3 }, fontFace: FONT,
        margin: [0.02, 0.05, 0.02, 0.05],   // tight padding - stops Rate wrapping
        autoPage: false,
      });

      const tableBot = TOP + 0.63 + tableRows.length * 0.46;
      s.addShape(pres.shapes.LINE, { x: xStart + 0.1, y: tableBot + 0.08, w: PW - 0.2, h: 0, line: { color: C.lgray, pt: 0.75 } });
      const ty = tableBot + 0.2;
      s.addShape(pres.shapes.RECTANGLE, { x: xStart, y: ty, w: PW, h: 0.60, fill: { color: C.redpale }, line: { color: C.redmid, pt: 0.75 } });
      s.addText(totLabel, { x: xStart + 0.12, y: ty, w: 2.0, h: 0.60, fontSize: 11, bold: true, color: C.darkRed, fontFace: FONT, valign: 'middle' });
      s.addText(totVal, { x: xStart + 2.12, y: ty, w: PW - 2.24, h: 0.60, fontSize: 21, bold: true, color: C.darkRed, fontFace: FONT, valign: 'middle', align: 'right' });
    };

    tier(LM, C.black, 'Tier 1 - Revenue Rebate',
      '1% of total purchases during the program period',
      [
        ['Program Period Revenue', money(d.currTotal), 'x1%', money(d.revRebate)],
        ['Every dollar purchased earns 1%', '', '', ''],
      ],
      'Tier 1 Earned', money(d.revRebate));

    tier(LM + PW + GAP, C.red, 'Tier 2 - Growth Rebate',
      '1% of purchases above prior-year same period',
      [
        [`Current Period (${d.dateFrom}-${d.dateTo})`, money(d.currTotal), '', ''],
        [`Prior Year Same Period (${d.priorFrom}-${d.priorTo})`, money(d.priorTotal), '', ''],
        ['Net Growth', money(d.growth), d.growth > 0 ? 'x1%' : '', d.growth > 0 ? money(d.growthRebate) : 'No growth rebate'],
      ],
      'Tier 2 Earned', money(d.growthRebate));

    s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: 5.2, w: 9.44, h: 0.36, fill: { color: C.black }, line: { color: C.black } });
    s.addText('TOTAL REBATE EARNED TO DATE:', { x: 0.44, y: 5.2, w: 3.4, h: 0.36, fontSize: 10, bold: true, color: C.white, fontFace: FONT, valign: 'middle' });
    s.addText(money(d.totalRebate), { x: 3.84, y: 5.2, w: 2.2, h: 0.36, fontSize: 18, bold: true, color: C.red, fontFace: FONT, valign: 'middle' });
    s.addText(`Tier 1: ${money(d.revRebate)}   +   Tier 2: ${money(d.growthRebate)}   |   Interim - Final paid at year end`,
      { x: 6.04, y: 5.2, w: 3.58, h: 0.36, fontSize: 8.5, color: C.white, fontFace: FONT, valign: 'middle' });
  }


  // ===== SLIDE 4 - Free Shipping Benefit =====================================
  {
    const s = pres.addSlide(); s.background = { color: C.off };
    hdr(s, 'THE VALUE OF FREE SHIPPING',
      `Every shipment ships free under the EPP program  |  ${d.dateFrom} - ${d.dateTo}`);

    const shipPctOfRev = d.currTotal > 0 ? (d.shipFreight / d.currTotal * 100) : 0;
    const per100       = d.currTotal > 0 ? (d.shipFreight / d.currTotal * 100) : 0;
    const totalValue   = d.shipFreight + d.totalRebate;
    const shipMultiple = d.totalRebate > 0 ? (d.shipFreight / d.totalRebate) : 0;

    // Three headline cards
    [
      { v: money(d.shipFreight),   l: 'TOTAL SHIPPING ABSORBED', sub: 'paid by Kleen-Rite, not you', red: true },
      { v: num(d.shipShipments),   l: 'SHIPMENTS COVERED',       sub: 'every one shipped free' },
      { v: money(d.shipAvg),       l: 'AVERAGE PER SHIPMENT',    sub: 'typical freight cost' },
    ].forEach((c, i) => {
      const x = 0.28 + i * 3.18;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 0.82, w: 2.9, h: 1.0,
        fill: { color: C.white }, line: { color: C.lgray, pt: 1 } });
      s.addShape(pres.shapes.RECTANGLE, { x, y: 0.82, w: 2.9, h: 0.05,
        fill: { color: C.red }, line: { color: C.red } });
      s.addText(c.v, { x, y: 0.88, w: 2.9, h: 0.52, fontSize: c.red ? 22 : 20, bold: true,
        color: c.red ? C.red : C.black, fontFace: FONT, align: 'center', valign: 'middle' });
      s.addText(c.l, { x, y: 1.40, w: 2.9, h: 0.2, fontSize: 8, bold: true,
        color: C.mgray, fontFace: FONT, align: 'center', charSpacing: 0.5 });
      s.addText(c.sub, { x, y: 1.58, w: 2.9, h: 0.18, fontSize: 7.5,
        color: 'BBBBBB', fontFace: FONT, align: 'center', italic: true });
    });

    // ---- Left panel: how the figure is built
    const LX0 = 0.28, LW0 = 4.55, PY0 = 1.98;
    s.addShape(pres.shapes.RECTANGLE, { x: LX0, y: PY0, w: LW0, h: 2.86,
      fill: { color: C.white }, line: { color: C.lgray, pt: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x: LX0, y: PY0, w: LW0, h: 0.34,
      fill: { color: C.black }, line: { color: C.black } });
    s.addText('HOW THIS IS CALCULATED', { x: LX0 + 0.12, y: PY0, w: LW0 - 0.24, h: 0.34,
      fontSize: 10, bold: true, color: C.white, fontFace: FONT, valign: 'middle', charSpacing: 0.5 });

    const rows = [
      ['Shipments fulfilled in period', num(d.shipShipments)],
      ['Average freight per shipment',  money(d.shipAvg)],
      ['Total freight absorbed',        money(d.shipFreight)],
      ['Your purchases this period',    money(d.currTotal)],
      ['Shipping as % of purchases',    shipPctOfRev.toFixed(1) + '%'],
    ];
    let ry = PY0 + 0.34;
    rows.forEach((r, i) => {
      const last = i === rows.length - 1;
      const bg = last ? C.redpale : (i % 2 === 0 ? 'F8F8F8' : C.white);
      s.addShape(pres.shapes.RECTANGLE, { x: LX0, y: ry, w: LW0, h: 0.42,
        fill: { color: bg }, line: { color: 'EBEBEB', pt: 0.2 } });
      s.addText(r[0], { x: LX0 + 0.14, y: ry, w: 2.9, h: 0.42, fontSize: 9.5,
        bold: last, color: last ? C.darkRed : C.black, fontFace: FONT, valign: 'middle' });
      s.addText(r[1], { x: LX0 + 3.0, y: ry, w: LW0 - 3.14, h: 0.42, fontSize: 10,
        bold: true, color: last ? C.darkRed : C.black, fontFace: FONT,
        valign: 'middle', align: 'right' });
      ry += 0.42;
    });
    s.addText('Orders that split across multiple shipments are covered on every one - ' +
              'that is why shipments exceed order count.',
      { x: LX0 + 0.14, y: ry + 0.04, w: LW0 - 0.28, h: 0.34, fontSize: 8,
        color: C.mgray, fontFace: FONT, italic: true, wrap: true });

    // ---- Right panel: total program value
    const RX = 5.17, RW = 4.55;
    s.addShape(pres.shapes.RECTANGLE, { x: RX, y: PY0, w: RW, h: 2.86,
      fill: { color: C.white }, line: { color: C.lgray, pt: 1 } });
    s.addShape(pres.shapes.RECTANGLE, { x: RX, y: PY0, w: RW, h: 0.34,
      fill: { color: C.red }, line: { color: C.red } });
    s.addText('TOTAL PROGRAM VALUE DELIVERED', { x: RX + 0.12, y: PY0, w: RW - 0.24, h: 0.34,
      fontSize: 10, bold: true, color: C.white, fontFace: FONT, valign: 'middle', charSpacing: 0.5 });

    // Proportional bars
    const barMax = Math.max(d.shipFreight, d.totalRebate, 1);
    const BAR_X  = RX + 1.55, BAR_MAXW = 2.35;
    [
      { label: 'Free Shipping', val: d.shipFreight,  color: C.red },
      { label: 'Rebate Earned', val: d.totalRebate,  color: '888888' },
    ].forEach((b, i) => {
      const by = PY0 + 0.55 + i * 0.62;
      const w  = Math.max((b.val / barMax) * BAR_MAXW, 0.06);
      s.addText(b.label, { x: RX + 0.14, y: by, w: 1.35, h: 0.34, fontSize: 9.5,
        bold: true, color: C.black, fontFace: FONT, valign: 'middle' });
      s.addShape(pres.shapes.RECTANGLE, { x: BAR_X, y: by + 0.05, w: w, h: 0.24,
        fill: { color: b.color }, line: { color: b.color } });
      s.addText(money(b.val), { x: BAR_X, y: by, w: BAR_MAXW + 0.5, h: 0.34,
        fontSize: 9, bold: true, color: C.black, fontFace: FONT,
        valign: 'middle', align: 'right' });
    });

    // Total
    const tY = PY0 + 1.86;
    s.addShape(pres.shapes.LINE, { x: RX + 0.14, y: tY - 0.06, w: RW - 0.28, h: 0,
      line: { color: C.lgray, pt: 0.75 } });
    s.addShape(pres.shapes.RECTANGLE, { x: RX, y: tY, w: RW, h: 0.5,
      fill: { color: C.redpale }, line: { color: C.redmid, pt: 0.75 } });
    s.addText('Combined Value', { x: RX + 0.14, y: tY, w: 2.0, h: 0.5, fontSize: 11,
      bold: true, color: C.darkRed, fontFace: FONT, valign: 'middle' });
    s.addText(money(totalValue), { x: RX + 2.1, y: tY, w: RW - 2.24, h: 0.5,
      fontSize: 18, bold: true, color: C.darkRed, fontFace: FONT,
      valign: 'middle', align: 'right' });

    if (shipMultiple >= 1.5) {
      s.addText('Your free shipping benefit is ' + shipMultiple.toFixed(1) +
                'x larger than your rebate.',
        { x: RX + 0.14, y: tY + 0.54, w: RW - 0.28, h: 0.3, fontSize: 9,
          color: C.dgray, fontFace: FONT, italic: true });
    }

    // ---- Bottom highlight
    s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: 5.0, w: 9.44, h: 0.42,
      fill: { color: C.black }, line: { color: C.black } });
    s.addText('For every $100 you spend with Kleen-Rite, we absorb ' +
              money(per100) + ' in freight so you do not pay a cent for shipping.',
      { x: 0.44, y: 5.0, w: 9.12, h: 0.42, fontSize: 11, bold: true,
        color: C.white, fontFace: FONT, valign: 'middle' });

    s.addText('Freight figures are taken from Item Fulfillment records, so orders shipped ' +
              'in multiple boxes are counted on each shipment.',
      { x: 0.28, y: 5.46, w: 9.44, h: 0.2, fontSize: 7.5, color: C.mgray,
        fontFace: FONT, italic: true });
  }

  // ===== SLIDE 5 - Product Opportunity =======================================
  {
    const s = pres.addSlide(); s.background = { color: C.off };
    hdr(s, 'PRODUCT OPPORTUNITY COMPARISON',
      `Sample Group Top 100 Items by Quantity vs. ${d.customerName}  |  ${d.dateFrom} - ${d.dateTo}`);

    const blocks = [
      { val: String(d.oppCount),  numSz: 44, label: 'OPPORTUNITY ITEMS', sub: 'Sample Group buys; you do NOT', bg: C.redpale, acc: C.red, tc: C.darkRed },
      { val: String(d.bothCount), numSz: 44, label: 'BOTH PURCHASE', sub: 'Items in common', bg: C.lgray, acc: C.dgray, tc: C.dgray },
      { val: String(d.custOnlyCount), numSz: 38, label: `${d.shortName.toUpperCase()} ONLY`,
        sub: d.custOnlyCount === 0 ? '(None this period)' : 'Items you buy; Sample Group does not',
        bg: C.bluepale, acc: C.blueline, tc: C.blue },
    ];
    const BW = 3.1, BH = 1.02, BY = 0.82;
    blocks.forEach((b, i) => {
      const bx = 0.28 + i * (BW + 0.06);
      s.addShape(pres.shapes.RECTANGLE, { x: bx, y: BY, w: BW, h: BH, fill: { color: b.bg }, line: { color: b.acc, pt: 1 } });
      s.addShape(pres.shapes.RECTANGLE, { x: bx, y: BY, w: 0.09, h: BH, fill: { color: b.acc }, line: { color: b.acc } });
      s.addText(b.val, { x: bx + 0.1, y: BY + 0.06, w: 1.25, h: 0.72, fontSize: b.numSz, bold: true, color: b.tc, fontFace: FONT, valign: 'middle', align: 'center' });
      s.addText(b.label, { x: bx + 1.38, y: BY + 0.1, w: 1.6, h: 0.28, fontSize: 8.5, bold: true, color: b.tc, fontFace: FONT });
      s.addText(b.sub, { x: bx + 1.38, y: BY + 0.38, w: 1.6, h: 0.38, fontSize: 7.5, color: C.mgray, fontFace: FONT, italic: true });
    });

    s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: 1.97, w: 9.44, h: 0.25, fill: { color: C.black }, line: { color: C.black } });
    s.addText(`TOP OPPORTUNITY ITEMS - Sample Group Buys; ${d.shortName} Does NOT  (showing top ${d.topOppRows.length} of ${d.oppCount})`,
      { x: 0.35, y: 1.97, w: 9.3, h: 0.25, fontSize: 8.5, bold: true, color: C.white, fontFace: FONT, valign: 'middle' });

    const COLS = [
      { l: 'Rank', x: 0.28, w: 0.62 },
      { l: 'Item', x: 0.90, w: 1.45 },
      { l: 'Description', x: 2.35, w: 3.62 },
      { l: 'Cat', x: 5.97, w: 0.58 },
      { l: 'Sample Grp Qty', x: 6.55, w: 1.55 },
      { l: 'Your Qty', x: 8.10, w: 1.62 },
    ];
    s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: 2.22, w: 9.44, h: 0.28, fill: { color: C.dgray }, line: { color: C.dgray } });
    COLS.forEach(c => s.addText(c.l, { x: c.x, y: 2.22, w: c.w, h: 0.28, fontSize: 8.5, bold: true, color: C.white, fontFace: FONT, valign: 'middle', align: 'center' }));

    d.topOppRows.forEach((row, i) => {
      const ry = 2.50 + i * 0.29;
      const bg = i % 2 === 0 ? C.redpale : C.white;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: ry, w: 9.44, h: 0.29, fill: { color: bg }, line: { color: 'F0D8D8', pt: 0.2 } });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: ry, w: 0.055, h: 0.29, fill: { color: C.red }, line: { color: C.red } });
      s.addText(String(row.rank),      { x: COLS[0].x, y: ry, w: COLS[0].w, h: 0.29, fontSize: 8.5, color: C.mgray, fontFace: FONT, valign: 'middle', align: 'center' });
      s.addText(row.itemid,            { x: COLS[1].x, y: ry, w: COLS[1].w, h: 0.29, fontSize: 8.5, bold: true, color: C.red, fontFace: FONT, valign: 'middle' });
      s.addText(row.displayname || '', { x: COLS[2].x, y: ry, w: COLS[2].w, h: 0.29, fontSize: 8.5, color: C.black, fontFace: FONT, valign: 'middle' });
      s.addText(row.category,          { x: COLS[3].x, y: ry, w: COLS[3].w, h: 0.29, fontSize: 8.5, color: C.mgray, fontFace: FONT, valign: 'middle', align: 'center' });
      s.addText(num(row.groupQty),     { x: COLS[4].x, y: ry, w: COLS[4].w, h: 0.29, fontSize: 8.5, bold: true, color: C.darkRed, fontFace: FONT, valign: 'middle', align: 'right' });
      s.addText(row.custQty > 0 ? num(row.custQty) : '-', { x: COLS[5].x, y: ry, w: COLS[5].w, h: 0.29, fontSize: 8.5, color: 'BBBBBB', fontFace: FONT, valign: 'middle', align: 'center' });
    });

    const bothY = 2.50 + d.topOppRows.length * 0.29 + 0.07;
    const bothStr = (d.topBothRows || []).slice(0, 5)
      .map(r => `${r.itemid}: ${num(r.groupQty)}/${num(r.custQty)}`).join('   |   ');
    s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: bothY, w: 9.44, h: 0.22, fill: { color: 'F0F0F0' }, line: { color: C.lgray, pt: 0.5 } });
    s.addText('BOTH BUY (Sample Group Qty / Your Qty):', { x: 0.35, y: bothY, w: 2.9, h: 0.22, fontSize: 8, bold: true, color: C.dgray, fontFace: FONT, valign: 'middle' });
    s.addText(bothStr || 'None', { x: 3.2, y: bothY, w: 6.48, h: 0.22, fontSize: 8, color: C.dgray, fontFace: FONT, valign: 'middle' });

    s.addText('Category: W = Wash Bay   P = Prep Station   V = Vacuum Island  |  See attached spreadsheet for complete sample group comparison',
      { x: 0.28, y: 5.35, w: 9.44, h: 0.2, fontSize: 7.5, color: C.mgray, fontFace: FONT, italic: true });
  }

  // ===== SLIDE 6 - Progress & Opportunities ==================================
  {
    const s = pres.addSlide(); s.background = { color: C.off };
    hdr(s, 'YOUR PROGRESS & OPPORTUNITIES', '');

    s.addText('YOUR PROGRESS THIS YEAR', { x: 0.28, y: 0.82, w: 4.55, h: 0.26, fontSize: 9, bold: true, color: C.black, fontFace: FONT, charSpacing: 1 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: 1.07, w: 4.55, h: 0.038, fill: { color: C.red }, line: { color: C.red } });

    const findings = [
      { head: d.growth >= 0 ? 'Strong Revenue Growth' : 'Revenue Below Prior Year',
        body: `Your purchases of ${money(d.currTotal)} in this period compare to ${money(d.priorTotal)} in the prior year same period - a ${d.growthPct}% change. This reflects your ongoing operational investment and relationship with Kleen-Rite.` },
      { head: 'Rebate Program Fully Active',
        body: `You have earned ${money(d.totalRebate)} in rebates to date - ${money(d.revRebate)} (Tier 1 revenue) and ${money(d.growthRebate)} (Tier 2 growth). Your average order value is ${money(d.avgOrderValue)}. Final rebate is paid at year end.` },
      { head: 'Free Shipping Is Your Largest Benefit',
        body: `Kleen-Rite absorbed ${money(d.shipFreight)} in freight across ${num(d.shipShipments)} shipments this period - ` +
              (d.currTotal > 0 ? (d.shipFreight / d.currTotal * 100).toFixed(1) + '% of everything you purchased' : 'every shipment covered') +
              `. Combined with your rebate, that is ${money(d.shipFreight + d.totalRebate)} of value delivered.` },
      { head: 'Sample Group Comparison Available',
        body: `The sample group purchases ${d.oppCount} of the top 100 items in high volume not yet in your program. You also purchase ${d.custOnlyCount} items the sample group does not - showing a broad and specialized product mix.` },
    ];
    findings.forEach((f, i) => {
      const fy = 1.16 + i * 1.0;
      s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: fy, w: 4.55, h: 0.9, fill: { color: C.white }, line: { color: C.lgray, pt: 0.75 } });
      s.addShape(pres.shapes.RECTANGLE, { x: 0.28, y: fy, w: 0.08, h: 0.9, fill: { color: C.red }, line: { color: C.red } });
      s.addText(f.head, { x: 0.44, y: fy + 0.07, w: 4.3, h: 0.25, fontSize: 10.5, bold: true, color: C.black, fontFace: FONT });
      s.addText(f.body, { x: 0.44, y: fy + 0.31, w: 4.3, h: 0.53, fontSize: 9, color: C.dgray, fontFace: FONT, wrap: true });
    });

    s.addText('WHERE YOU COULD BE GAINING', { x: 5.17, y: 0.82, w: 4.55, h: 0.26, fontSize: 9, bold: true, color: C.black, fontFace: FONT, charSpacing: 1 });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.17, y: 1.07, w: 4.55, h: 0.038, fill: { color: C.red }, line: { color: C.red } });

    const actions = [
      { num: '01', head: 'Review Your High-Volume Consumables',
        body: 'The top opportunity items are high-volume consumables your peers replenish regularly. A review of your current stock levels could help prevent unexpected downtime and maintain wash performance.' },
      { num: '02', head: 'Add Prep Station Consumables',
        body: 'Prep station items are in consistent high-volume use across the sample group. Stocking these reduces emergency orders and keeps your prep operation running at full capacity.' },
      { num: '03', head: 'Evaluate Your Hose & Tubing Program',
        body: 'Hose and tubing items appear prominently in sample group purchases. A regular replacement schedule protects equipment, maintains wash performance, and avoids unplanned downtime.' },
      { num: '04', head: 'Keep Building - Your Rebate Grows With You',
        body: `Every purchase above your ${money(d.priorTotal)} prior-year baseline earns an additional 1% growth rebate. Your rebate is calculated on your full-year results and paid next year.` },
    ];
    actions.forEach((a, i) => {
      const ay = 1.16 + i * 1.0;
      s.addShape(pres.shapes.RECTANGLE, { x: 5.17, y: ay, w: 4.55, h: 0.9, fill: { color: C.white }, line: { color: C.lgray, pt: 0.75 } });
      s.addShape(pres.shapes.RECTANGLE, { x: 5.17, y: ay, w: 0.36, h: 0.9, fill: { color: C.redpale }, line: { color: C.redpale } });
      s.addText(a.num, { x: 5.17, y: ay, w: 0.36, h: 0.9, fontSize: 15, bold: true, color: C.red, fontFace: FONT, align: 'center', valign: 'middle' });
      s.addText(a.head, { x: 5.6, y: ay + 0.07, w: 4.0, h: 0.25, fontSize: 10.5, bold: true, color: C.black, fontFace: FONT });
      s.addText(a.body, { x: 5.6, y: ay + 0.31, w: 4.0, h: 0.53, fontSize: 9, color: C.dgray, fontFace: FONT, wrap: true });
    });
  }

  // ===== SLIDE 7 - Closing ===================================================
  {
    const s = pres.addSlide(); s.background = { color: C.black };
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: C.red }, line: { color: C.red } });
    logo(s);
    s.addText('THE OPPORTUNITY AHEAD', { x: 0.38, y: 0.55, w: 8.5, h: 0.36, fontSize: 11, bold: true, color: C.red, fontFace: FONT, charSpacing: 5 });
    s.addText(`${d.shortName} is a Valued Partner.`, { x: 0.38, y: 0.92, w: 8.5, h: 0.9, fontSize: 34, bold: true, color: C.white, fontFace: FONT });
    s.addText("Let's Keep Growing Together.", { x: 0.38, y: 1.84, w: 8.5, h: 0.42, fontSize: 20, color: 'CCCCCC', fontFace: FONT });

    s.addText('PROGRAM HIGHLIGHTS', { x: 0.38, y: 2.36, w: 8.5, h: 0.24, fontSize: 8.5, bold: true, color: '666666', fontFace: FONT, charSpacing: 2 });
    [
      { val: money(d.currTotal), label: 'Current Period Revenue' },
      { val: money(d.totalRebate), label: 'Rebate Earned to Date' },
      { val: d.growthPct + '%', label: 'Growth vs Prior Year' },
    ].forEach((cs, i) => {
      const x = 0.38 + i * 2.75;
      s.addShape(pres.shapes.RECTANGLE, { x, y: 2.62, w: 2.45, h: 0.86, fill: { color: '222222' }, line: { color: C.red, pt: 0.75 } });
      s.addText(cs.val, { x, y: 2.65, w: 2.45, h: 0.5, fontSize: 18, bold: true, color: C.red, fontFace: FONT, align: 'center', valign: 'bottom' });
      s.addText(cs.label, { x, y: 3.15, w: 2.45, h: 0.28, fontSize: 8, color: 'AAAAAA', fontFace: FONT, align: 'center', valign: 'top' });
    });

    // Second row: two boxes, widened to span the same width as the row above.
    // Labels sit on the boxes themselves, so no section header here.
    [
      [String(d.oppCount),   'Missed Items Opportunity'],
      [money(d.shipFreight), 'Freight Savings'],
    ].forEach((b, i) => {
      const w = 3.82;
      const x = 0.38 + i * (w + 0.31);
      s.addShape(pres.shapes.RECTANGLE, { x, y: 3.78, w: w, h: 0.9,
        fill: { color: '222222' }, line: { color: C.red, pt: 0.75 } });
      s.addText(b[0], { x, y: 3.82, w: w, h: 0.52, fontSize: 22, bold: true,
        color: C.red, fontFace: FONT, align: 'center', valign: 'bottom' });
      s.addText(b[1], { x, y: 4.34, w: w, h: 0.3, fontSize: 8.5,
        color: 'AAAAAA', fontFace: FONT, align: 'center', valign: 'top' });
    });

    s.addText(`Data: NetSuite  |  Period: ${d.dateFrom}-${d.dateTo}  |  Generated ${d.generatedDate}  |  Interim Report - Final rebate paid at year end`,
      { x: 0.38, y: 4.88, w: 9.0, h: 0.22, fontSize: 7.5, color: '444444', fontFace: FONT, italic: true });
  }

  return pres;
}


// ---- Excel workbook builder ------------------------------------------------
// Full item comparison, formatted to match the Kleen-Rite red/black theme.
async function buildWorkbook(d) {
  const ExcelJS = require('exceljs');
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Kleen-Rite EPP Quarterly Review';
  wb.created = new Date();

  const ws = wb.addWorksheet('Item Comparison', {
    views: [{ state: 'frozen', ySplit: 3 }]
  });

  const W = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const THIN = { style: 'thin', color: { argb: 'FFCCCCCC' } };
  const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };

  ws.columns = [
    { key: 'rank', width: 8 },
    { key: 'item', width: 20 },
    { key: 'desc', width: 46 },
    { key: 'cat',  width: 10 },
    { key: 'gqty', width: 18 },
    { key: 'cqty', width: 16 },
    { key: 'stat', width: 20 },
  ];

  // Title
  ws.mergeCells('A1:G1');
  const t = ws.getCell('A1');
  t.value = `${d.customerName} (${d.customerNumber})  -  Item Comparison vs. Sample Group  |  ${d.dateFrom} - ${d.dateTo}`;
  t.font = { name: 'Calibri', bold: true, size: 12, color: { argb: 'FFFFFFFF' } };
  t.fill = W('FF1A1A1A');
  t.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 24;

  // Subtitle
  ws.mergeCells('A2:G2');
  const st = ws.getCell('A2');
  st.value = 'Source: Customer Group Item Purchases  |  Excludes DJ items  |  Category: W = Wash Bay, P = Prep Station, V = Vacuum Island';
  st.font = { name: 'Calibri', italic: true, size: 9, color: { argb: 'FF666666' } };
  st.fill = W('FFF5F5F5');
  st.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(2).height = 16;

  ws.getRow(3).height = 6;

  let r = 4;

  const sectionHeader = (label, argb) => {
    ws.mergeCells(`A${r}:G${r}`);
    const c = ws.getCell(`A${r}`);
    c.value = '  ' + label;
    c.font = { name: 'Calibri', bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    c.fill = W(argb);
    c.alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getRow(r).height = 18;
    r++;
  };

  const columnHeader = (argb) => {
    const labels = ['Rank', 'Item', 'Description', 'Category', 'Sample Group Qty', 'Customer Qty', 'Status'];
    const row = ws.getRow(r);
    labels.forEach((lbl, i) => {
      const c = row.getCell(i + 1);
      c.value = lbl;
      c.font = { name: 'Calibri', bold: true, size: 9, color: { argb: 'FFFFFFFF' } };
      c.fill = W(argb);
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      c.border = BORDER;
    });
    row.height = 20;
    r++;
  };

  const dataRow = (vals, bgArgb, itemArgb, itemBold) => {
    const row = ws.getRow(r);
    vals.forEach((v, i) => {
      const c = row.getCell(i + 1);
      c.value = v;
      c.font = {
        name: 'Calibri', size: 10,
        bold: i === 1 ? !!itemBold : false,
        color: { argb: i === 1 && itemArgb ? itemArgb : 'FF000000' }
      };
      c.fill = W(bgArgb);
      c.border = BORDER;
      c.alignment = { horizontal: i === 2 ? 'left' : 'center', vertical: 'middle' };
      if (i === 4 || i === 5) c.numFmt = '#,##0';
    });
    row.height = 16;
    r++;
  };

  // --- Opportunity
  const opp = d.allOppRows || [];
  sectionHeader(`OPPORTUNITY  -  Sample Group buys these ${opp.length} items; ${d.shortName} does NOT`, 'FFCC0000');
  columnHeader('FF444444');
  opp.forEach((row, i) => {
    dataRow(
      [row.rank, row.itemid, row.displayname || '', row.category, Math.round(row.groupQty), null, 'OPPORTUNITY'],
      i % 2 === 0 ? 'FFFFF0F0' : 'FFFFFFFF',
      'FFCC0000', true
    );
  });

  r++;

  // --- Both buy
  const both = d.allBothRows || [];
  sectionHeader(`BOTH PURCHASE  -  ${both.length} items ${d.shortName} and the Sample Group both buy`, 'FF444444');
  columnHeader('FF666666');
  both.forEach((row, i) => {
    dataRow(
      [row.rank, row.itemid, row.displayname || '', row.category, Math.round(row.groupQty), Math.round(row.custQty), 'BOTH BUY'],
      i % 2 === 0 ? 'FFF2F2F2' : 'FFFFFFFF',
      'FF444444', false
    );
  });

  r++;

  // --- Customer only
  const only = d.allCustOnly || [];
  sectionHeader(`${d.shortName.toUpperCase()} ONLY  -  ${only.length} items purchased that are not in the Sample Group top 100`, 'FF1F497D');
  columnHeader('FF33557A');
  only.forEach((row, i) => {
    dataRow(
      [null, row.itemid, row.displayname || '', row.category, null, Math.round(row.custQty), 'CUSTOMER ONLY'],
      i % 2 === 0 ? 'FFF0F6FC' : 'FFFFFFFF',
      'FF1F497D', false
    );
  });

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf).toString('base64');
}

// ---- Handler ---------------------------------------------------------------
module.exports = async (req, res) => {
  // CORS / preflight
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.status(200).end();

  if (req.method === 'GET') {
    return res.status(200).json({ status: 'ok', service: 'EPP Quarterly Review PPTX generator' });
  }
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  // Auth
  const auth = req.headers.authorization || '';
  const token = auth.replace(/^Bearer\s+/i, '').trim();
  if (token !== API_TOKEN) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  try {
    const d = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;

    // Defaults so a missing field never crashes the build
    const data = Object.assign({
      customerName: 'Customer',
      shortName: 'Customer',
      customerNumber: '',
      dateFrom: '', dateTo: '', priorFrom: '', priorTo: '',
      generatedDate: new Date().toLocaleDateString('en-US'),
      currTotal: 0, priorTotal: 0, growth: 0, growthPct: 0,
      revRebate: 0, growthRebate: 0, totalRebate: 0,
      numOrders: 0, avgOrderValue: 0, aovOk: false,
      weeklyLabels: [], currWeekly: [], priorWeekly: [],
      oppCount: 0, bothCount: 0, custOnlyCount: 0,
      topOppRows: [], topBothRows: [],
      allOppRows: [], allBothRows: [], allCustOnly: [],
      shipShipments: 0, shipFreight: 0, shipAvg: 0,
      rollupLabel: '',
    }, d || {});

    const pres = buildDeck(data);
    const b64 = await pres.write({ outputType: 'base64' });

    // Excel is best-effort: never let a workbook problem block the PowerPoint
    let xlsxB64 = null;
    let xlsxError = null;
    try {
      xlsxB64 = await buildWorkbook(data);
    } catch (xe) {
      xlsxError = xe.message;
    }

    return res.status(200).json({
      success: true,
      pptx: b64,
      xlsx: xlsxB64,
      xlsxError: xlsxError
    });
  } catch (err) {
    return res.status(500).json({ error: err.message, stack: err.stack });
  }
};
