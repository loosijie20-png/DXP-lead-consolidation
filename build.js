const pptxgen = require('pptxgenjs');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa6');

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = 'Lead Consolidation MVP';

const W = 13.333, H = 7.5, M = 0.6, CW = W - 2 * M;
const C = {
  PETROL: '0B3C49', TEAL: '1F7A8C', TINT: 'E6F0F2', TINT2: 'F5F9FA',
  AMBER: 'E09F3E', AMBER_D: '8A5A12', AMBER_T: 'FBF1E1',
  RED: 'B23A48', RED_T: 'F7E3E5', INK: '1B2631', MUTED: '5A6B73',
  LINE: 'C9D6DA', WHITE: 'FFFFFF', PALE: 'CFE3E8', GRAY_T: 'F1F3F4',
};
const TF = 'Cambria', BF = 'Calibri';

// ---------- helpers ----------
const iconCache = {};
async function iconData(name, color) {
  const key = name + color;
  if (iconCache[key]) return iconCache[key];
  const Comp = fa[name];
  if (!Comp) throw new Error('Missing icon ' + name);
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { size: 256, color: '#' + color }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  iconCache[key] = 'image/png;base64,' + buf.toString('base64');
  return iconCache[key];
}
const shadow = () => ({ type: 'outer', blur: 8, offset: 2, angle: 90, color: '000000', opacity: 0.10 });

function txt(s, text, o) {
  s.addText(text, Object.assign({ fontFace: BF, fontSize: 14, color: C.INK, margin: 0, valign: 'top', isTextBox: true }, o));
}
function box(s, o, rounded = true) {
  s.addShape(rounded ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE,
    Object.assign(rounded ? { rectRadius: 0.08 } : {}, { line: { type: 'none' } }, o));
}
function circ(s, x, y, d, fill, label, o = {}) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: 'none' } });
  if (label !== undefined) {
    txt(s, String(label), { x, y, w: d, h: d, align: 'center', valign: 'middle', bold: true,
      color: o.color || C.WHITE, fontSize: o.fontSize || 14, fontFace: o.fontFace || BF });
  }
}
async function iconCirc(s, name, x, y, d, fill, iconColor = C.WHITE) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: 'none' } });
  const k = d * 0.5;
  s.addImage({ data: await iconData(name, iconColor), x: x + (d - k) / 2, y: y + (d - k) / 2, w: k, h: k });
}
async function icon(s, name, x, y, size, color) {
  s.addImage({ data: await iconData(name, color), x, y, w: size, h: size });
}
function vline(s, x, y, h, color, width = 1.5, dash) {
  s.addShape(pres.shapes.LINE, { x, y, w: 0, h, line: Object.assign({ color, width }, dash ? { dashType: dash } : {}) });
}
function hline(s, x, y, w, color, width = 1.5) {
  s.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color, width } });
}
function title(s, t, kicker) {
  if (kicker) txt(s, kicker, { x: M, y: 0.3, w: CW, h: 0.3, fontSize: 12, bold: true, color: C.TEAL, charSpacing: 2 });
  txt(s, t, { x: M, y: 0.58, w: CW, h: 0.72, fontFace: TF, fontSize: 36, bold: true, color: C.PETROL, valign: 'middle' });
}
function footer(s, n) {
  txt(s, 'Lead Consolidation MVP', { x: M, y: 7.08, w: 5, h: 0.28, fontSize: 10, color: C.MUTED });
  txt(s, String(n), { x: W - M - 1, y: 7.08, w: 1, h: 0.28, fontSize: 10, color: C.MUTED, align: 'right' });
}
function intro(s, t, y = 1.38) {
  txt(s, t, { x: M, y, w: CW, h: 0.42, fontSize: 16, italic: true, color: C.MUTED, valign: 'middle' });
}
function header(s, t, x, y, w, o = {}) {
  txt(s, t, Object.assign({ x, y, w, h: 0.4, fontSize: 18, bold: true, color: C.PETROL, valign: 'middle' }, o));
}
function bullets(items, o = {}) {
  return items.map((t, i) => ({
    text: t,
    options: Object.assign({ bullet: { indent: 14 }, breakLine: i < items.length - 1 }, o),
  }));
}
function table(s, head, rows, o) {
  const hdr = head.map(h => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.PETROL }, fontSize: o.hfs || o.fs } }));
  const body = rows.map((r, i) => r.map((c, j) => {
    const base = Object.assign({ fill: { color: i % 2 ? C.WHITE : C.TINT2 } }, (o.colOpts && o.colOpts[j]) || {});
    if (c && typeof c === 'object' && !Array.isArray(c) && c.text !== undefined) {
      return { text: c.text, options: Object.assign(base, c.options || {}) };
    }
    return { text: c, options: base };
  }));
  s.addTable([hdr, ...body], {
    x: o.x, y: o.y, w: o.w, colW: o.colW, fontFace: BF, fontSize: o.fs, color: C.INK, valign: 'middle',
    border: { type: 'solid', pt: 0.75, color: C.LINE }, margin: [0.06, 0.1, 0.06, 0.1], rowH: [o.hRowH || Math.min(o.rowH, 0.45), ...rows.map(() => o.rowH)], autoPage: false,
  });
}
function labelDesc(label, desc, lfs, dfs, lcolor = C.PETROL, dcolor = C.INK) {
  return [
    { text: label, options: { bold: true, fontSize: lfs, color: lcolor, breakLine: true } },
    { text: desc, options: { fontSize: dfs, color: dcolor } },
  ];
}

// ---------- slides ----------
async function build() {
  let n = 0;

  // 1. Title
  {
    const s = pres.addSlide(); n++;
    s.background = { color: C.PETROL };
    txt(s, 'CONCEPT PROPOSAL', { x: 0.8, y: 1.7, w: 8, h: 0.4, fontSize: 16, bold: true, color: C.AMBER, charSpacing: 6 });
    txt(s, 'Lead Consolidation MVP', { x: 0.8, y: 2.15, w: 11.5, h: 1.15, fontFace: TF, fontSize: 54, bold: true, color: C.WHITE, valign: 'middle' });
    txt(s, 'A focused, low-code workflow for capturing, validating, assigning, and tracking leads end-to-end',
      { x: 0.8, y: 3.4, w: 9.5, h: 1.0, fontSize: 22, color: C.PALE });
    const cy = 4.8, d = 0.3, step = 0.75;
    hline(s, 0.8 + d / 2, cy + d / 2, 7 * step, '5FB3C2', 2);
    for (let i = 0; i < 8; i++) {
      s.addShape(pres.shapes.OVAL, { x: 0.8 + i * step, y: cy, w: d, h: d, fill: { color: i === 7 ? C.AMBER : C.TEAL }, line: { color: '5FB3C2', width: 1.5 } });
    }
    txt(s, 'Chin Yung Xuan   ·   Loo Si Jie   ·   Teo Jing Ying', { x: 0.8, y: 5.7, w: 11, h: 0.45, fontSize: 18, color: C.WHITE, valign: 'middle' });
  }

  // 2. Problem & Pain Points
  {
    const s = pres.addSlide(); n++;
    title(s, 'Problem & Pain Points');
    box(s, { x: M, y: 1.45, w: CW, h: 0.75, fill: { color: C.PETROL } });
    txt(s, [
      { text: 'The Problem   ', options: { bold: true, color: C.AMBER } },
      { text: 'Leads arrive through multiple channels. No single controlled workflow captures, validates, assigns, tracks, or monitors them.', options: { color: C.WHITE } },
    ], { x: M + 0.25, y: 1.45, w: CW - 0.5, h: 0.75, fontSize: 15, valign: 'middle' });

    header(s, 'Current Workflow', M, 2.45, 5.1);
    const steps = [
      ['Leads arrive', 'Web forms, email, phone calls, walk-ins, referrals'],
      ['Manual capture', 'Staff copy lead details into spreadsheets by hand'],
      ['Scattered records', 'Data sits in inboxes, spreadsheets, and personal notes'],
      ['Unclear follow-up', 'No set owner, no deadline, no tracking'],
    ];
    const sy = 2.98, sstep = 1.0, d = 0.5;
    vline(s, M + d / 2, sy + d / 2, 3 * sstep, C.LINE, 2);
    steps.forEach(([l, t], i) => {
      circ(s, M, sy + i * sstep, d, C.TEAL, i + 1, { fontSize: 15 });
      txt(s, labelDesc(l, t, 15, 13), { x: M + 0.7, y: sy + i * sstep - 0.02, w: 4.4, h: 0.85 });
    });

    const rx = 6.2, rw = W - M - rx;
    header(s, 'Key Pain Points', rx, 2.45, rw);
    const pains = [
      ['FaKeyboard', 'Manual consolidation', 'Slow and error-prone'],
      ['FaTriangleExclamation', 'Inconsistent data', 'Phone, email, or company often missing'],
      ['FaClone', 'Duplicates', 'Two records, two follow-ups, one confused customer'],
      ['FaUserSlash', 'Unclear ownership', 'Leads sit untouched for days'],
      ['FaHourglassEnd', 'SLA breaches', 'Overdue leads found too late'],
      ['FaEyeSlash', 'Hidden exceptions', 'Incomplete, duplicate, and unassigned records not surfaced'],
      ['FaChartColumn', 'Limited visibility', 'No central dashboard or audit history'],
    ];
    const gw = 0.15, cw = (rw - gw) / 2, ch = 0.88, gh = 0.14, py = 2.98;
    for (let i = 0; i < pains.length; i++) {
      const [ic, l, t] = pains[i];
      const col = i % 2, row = Math.floor(i / 2);
      const x = rx + col * (cw + gw), y = py + row * (ch + gh);
      const w = i === 6 ? rw : cw;
      box(s, { x, y, w, h: ch, fill: { color: C.WHITE }, line: { color: C.LINE, width: 0.75 } });
      await iconCirc(s, ic, x + 0.15, y + (ch - 0.46) / 2, 0.46, C.AMBER);
      txt(s, labelDesc(l, t, 13, 12), { x: x + 0.75, y: y + 0.1, w: w - 0.9, h: ch - 0.15 });
    }
    footer(s, n);
  }

  // 3. Proposed Solution Overview
  {
    const s = pres.addSlide(); n++;
    title(s, 'Proposed Solution Overview');
    intro(s, 'One eight-stage pipeline. Every lead passes through every stage.');
    const stages = [
      ['Intake', 'Lead arrives via Microsoft Forms, shared email inbox, or Power Apps manual entry'],
      ['Register', 'Saved to the central Lead Register with source and timestamp'],
      ['Validate', 'Values standardised; mandatory fields checked'],
      ['Deduplicate', 'Compared with existing leads; duplicates flagged'],
      ['Assign', 'Routed to an owner by region, product, or round-robin'],
      ['Notify', 'Owner alerted; SLA timer starts; reminders sent'],
      ['Track', 'Status, ageing, and SLA state monitored; overdue leads escalated'],
      ['Report', 'Power BI dashboard and full audit history'],
    ];
    const k = 8, gap = 0.17, cw = (CW - gap * (k - 1)) / k, cy = 2.0, d = 0.56, cardY = 2.75, cardH = 2.2;
    hline(s, M + cw / 2, cy + d / 2, (k - 1) * (cw + gap), C.TEAL, 2.5);
    stages.forEach(([nm, ds], i) => {
      const x = M + i * (cw + gap);
      circ(s, x + (cw - d) / 2, cy, d, C.TEAL, i + 1, { fontSize: 16 });
      box(s, { x, y: cardY, w: cw, h: cardH, fill: { color: C.TINT } });
      txt(s, nm, { x: x + 0.1, y: cardY + 0.14, w: cw - 0.2, h: 0.35, fontSize: 13.5, bold: true, color: C.PETROL });
      txt(s, ds, { x: x + 0.1, y: cardY + 0.55, w: cw - 0.2, h: cardH - 0.65, fontSize: 12.5 });
    });
    const exY = cardY + cardH + 0.45;
    [2, 3, 4, 6].forEach(i => {
      const x = M + i * (cw + gap) + cw / 2;
      vline(s, x, cardY + cardH, 0.45, C.AMBER, 1.5, 'dash');
    });
    box(s, { x: M, y: exY, w: CW, h: 1.1, fill: { color: C.AMBER_T }, line: { color: C.AMBER, width: 1 } });
    txt(s, [
      { text: 'Exception Branch', options: { bold: true, fontSize: 15, color: C.AMBER_D } },
      { text: '   below stages 3, 4, 5, and 7', options: { fontSize: 12, color: C.MUTED, breakLine: true } },
      { text: 'Incomplete records · Suspected duplicates · Unassigned leads · Breached SLAs   →   ', options: { fontSize: 14 } },
      { text: 'Exception Queue', options: { fontSize: 14, bold: true, color: C.AMBER_D } },
      { text: '   →   human review', options: { fontSize: 14 } },
    ], { x: M + 0.25, y: exY + 0.12, w: CW - 0.5, h: 0.9, valign: 'middle' });
    footer(s, n);
  }

  // 4. Scope of the MVP
  {
    const s = pres.addSlide(); n++;
    title(s, 'Scope of the MVP');
    const y = 1.5, h = 5.4, lw = 7.7, rx = M + lw + 0.3, rw = W - M - rx;
    box(s, { x: M, y, w: lw, h, fill: { color: C.TINT2 }, line: { color: C.LINE, width: 0.75 } });
    box(s, { x: rx, y, w: rw, h, fill: { color: C.GRAY_T }, line: { color: C.LINE, width: 0.75 } });
    await iconCirc(s, 'FaCheck', M + 0.3, y + 0.25, 0.48, C.TEAL);
    header(s, 'In Scope', M + 0.95, y + 0.29, 4);
    await iconCirc(s, 'FaXmark', rx + 0.3, y + 0.25, 0.48, C.MUTED);
    header(s, 'Out of Scope', rx + 0.95, y + 0.29, 3);
    const ins = [
      'Lead capture from 3 approved intake channels',
      'Central Lead Register with agreed mandatory and optional fields',
      'Data standardisation and validation, with a review queue for incomplete records',
      'Rule-based duplicate detection, duplicate flags, and human review',
      'Lead assignment to approved owners by agreed rules',
      'Automated notifications: new assignments, approaching SLAs, overdue actions',
      'Tracking of status, ageing, SLA state, and exceptions',
      'Power BI management dashboard',
      'Audit history: who did what, and when',
      'Test evidence, user and admin documentation, handover session',
    ];
    for (let i = 0; i < ins.length; i++) {
      const iy = y + 1.0 + i * 0.43;
      await icon(s, 'FaCheck', M + 0.42, iy + 0.06, 0.2, C.TEAL);
      txt(s, ins[i], { x: M + 0.8, y: iy, w: lw - 1.0, h: 0.34, fontSize: 13, valign: 'middle' });
    }
    const outs = [
      'Two-way CRM integration, such as Salesforce (one-way export possible)',
      'Marketing automation, campaign management, and ML-based lead scoring',
      'Customer-facing portals',
      'Mobile app builds beyond the Power Apps mobile experience',
      'Migration of large volumes of historical leads',
    ];
    for (let i = 0; i < outs.length; i++) {
      const iy = y + 1.0 + i * 0.82;
      await icon(s, 'FaXmark', rx + 0.42, iy + 0.06, 0.2, C.MUTED);
      txt(s, outs[i], { x: rx + 0.8, y: iy, w: rw - 1.0, h: 0.7, fontSize: 13 });
    }
    footer(s, n);
  }

  // 5. Three Approved Intake Methods
  {
    const s = pres.addSlide(); n++;
    title(s, 'Three Approved Intake Methods');
    const ch = [
      ['FaRegFileLines', 'Channel 1', 'Digital Intake', 'Microsoft Forms', [
        'Public form on the company website or shared as a link',
        'Mobile-friendly; works on every device',
        'Built-in checks: required fields and email format',
        'Each submission triggers a Power Automate flow immediately',
        'Included in the existing Microsoft 365 licence']],
      ['FaEnvelope', 'Channel 2', 'Passive Intake', 'Shared Email Inbox', [
        'Monitored mailbox, for example leads@company.com',
        'Power Automate creates a lead record for every new email',
        'Sender name, sender email, subject, and body parsed into the register',
        'Attachments saved to SharePoint and linked to the lead',
        'Fields that cannot be parsed are flagged "Needs Enrichment"']],
      ['FaMobileScreenButton', 'Channel 3', 'Staff Intake', 'Power Apps Manual Entry', [
        'For phone calls, walk-ins, trade shows, and referrals',
        'Mobile and desktop form with the same fields as the register',
        'Mandatory fields enforced at the point of entry',
        'Source, channel, and captured-by filled automatically',
        'Offline capture at events; syncs when connectivity returns']],
    ];
    const gap = 0.3, cw = (CW - 2 * gap) / 3, y = 1.5, h = 4.5;
    for (let i = 0; i < 3; i++) {
      const [ic, lab, sub, nm, bl] = ch[i];
      const x = M + i * (cw + gap);
      box(s, { x, y, w: cw, h, fill: { color: C.WHITE }, line: { color: C.LINE, width: 0.75 }, shadow: shadow() });
      await iconCirc(s, ic, x + 0.3, y + 0.3, 0.7, C.TEAL);
      txt(s, [
        { text: lab.toUpperCase(), options: { fontSize: 11, bold: true, color: C.MUTED, charSpacing: 1, breakLine: true } },
        { text: sub, options: { fontSize: 15, bold: true, color: C.TEAL } },
      ], { x: x + 1.15, y: y + 0.33, w: cw - 1.3, h: 0.65, valign: 'middle' });
      txt(s, nm, { x: x + 0.3, y: y + 1.15, w: cw - 0.6, h: 0.45, fontSize: 19, bold: true, color: C.PETROL, valign: 'middle' });
      txt(s, bullets(bl), { x: x + 0.3, y: y + 1.72, w: cw - 0.55, h: h - 1.85, fontSize: 13, paraSpaceAfter: 5 });
    }
    box(s, { x: M, y: 6.25, w: CW, h: 0.6, fill: { color: C.PETROL } });
    txt(s, 'All three channels feed the same Lead Register. The downstream workflow is identical.',
      { x: M + 0.25, y: 6.25, w: CW - 0.5, h: 0.6, fontSize: 15, color: C.WHITE, valign: 'middle' });
    footer(s, n);
  }

  // 6. The Central Lead Register
  {
    const s = pres.addSlide(); n++;
    title(s, 'The Central Lead Register');
    const lw = 3.35;
    box(s, { x: M, y: 1.5, w: lw, h: 1.55, fill: { color: C.PETROL } });
    txt(s, labelDesc('Single Source of Truth', 'Every channel writes to it. Every rule, notification, and dashboard reads from it.', 16, 13, C.AMBER, C.WHITE),
      { x: M + 0.2, y: 1.62, w: lw - 0.4, h: 1.35, valign: 'middle', paraSpaceAfter: 4 });
    await iconCirc(s, 'FaDatabase', M, 3.3, 0.45, C.TEAL);
    txt(s, 'Where It Lives: SharePoint List', { x: M + 0.55, y: 3.3, w: lw - 0.55, h: 0.45, fontSize: 14, bold: true, color: C.PETROL, valign: 'middle' });
    txt(s, bullets([
      'Covered by existing Microsoft 365 licences',
      'Native version history, permissions, and audit',
      'Integrates with Power Automate, Power Apps, and Power BI',
      'No separate database administration',
      'Upgrade path to Dataverse as volume grows',
    ]), { x: M, y: 3.9, w: lw, h: 2.4, fontSize: 12.5, paraSpaceAfter: 4 });
    txt(s, 'Final field list confirmed with the client in the Discover phase.',
      { x: M, y: 6.35, w: lw, h: 0.5, fontSize: 11, italic: true, color: C.MUTED });
    const tx = 4.25, tw = W - M - tx;
    const rows = [
      ['Lead ID', 'Auto', 'Unique system identifier'],
      ['Full Name', 'Text (mandatory)', 'Contact name'],
      ['Email', 'Email (mandatory*)', '*Mandatory if Phone not provided'],
      ['Phone', 'Text (mandatory*)', '*Mandatory if Email not provided'],
      ['Company', 'Text (optional)', 'Organisation name'],
      ['Source', 'Choice (mandatory)', 'Form, Email, Manual'],
      ['Channel Detail', 'Text (optional)', 'Which form, inbox, or event'],
      ['Product Interest', 'Choice (optional)', 'Drives assignment rules'],
      ['Region', 'Choice (optional)', 'Drives assignment rules'],
      ['Received Date', 'DateTime (auto)', 'When the lead arrived'],
      ['Assigned Owner', 'Person (auto)', 'Set by the assignment engine'],
      ['Assignment Date', 'DateTime (auto)', 'When ownership was set'],
      ['SLA Due', 'DateTime (auto)', 'Assignment Date + SLA hours'],
      ['Status', 'Choice (auto)', 'New, Assigned, In Progress, Contacted, Qualified, Closed – Won, Closed – Lost'],
      ['SLA State', 'Choice (auto)', 'On Track, Approaching, Breached'],
      ['Duplicate Flag', 'Choice (auto)', 'None, Suspected, Confirmed'],
      ['Duplicate Of', 'Lookup (auto)', 'Link to the master lead'],
      ['Notes', 'Multi-line (optional)', 'Owner comments and interaction history'],
    ];
    table(s, ['Field', 'Type', 'Notes'], rows, {
      x: tx, y: 1.5, w: tw, colW: [1.5, 1.6, tw - 3.1], fs: 10, hfs: 11, rowH: 0.27,
      colOpts: [{ bold: true, color: C.PETROL }, {}, {}],
    });
    footer(s, n);
  }

  // 7. Data Standardisation
  {
    const s = pres.addSlide(); n++;
    title(s, 'Data Standardisation', 'BUSINESS RULES  ·  1 OF 5');
    intro(s, 'Runs first. Converts every lead to one common format before any other rule.');
    const lw = 4.3;
    header(s, 'How It Works', M, 2.05, lw);
    const items = [
      'One shared step for all three intake channels; same rules in the Power App',
      'Raw values kept exactly as captured, for audit',
      'Standardised values stored in indexed match keys: Email Key, Phone Key, Name Key, Company Key',
    ];
    items.forEach((t, i) => {
      const y = 2.6 + i * 1.3;
      circ(s, M, y, 0.46, C.TEAL, i + 1);
      txt(s, t, { x: M + 0.65, y: y - 0.02, w: lw - 0.65, h: 1.1, fontSize: 14 });
    });
    const tx = 5.3, tw = W - M - tx;
    table(s, ['Field', 'Rule'], [
      ['Email', 'Remove extra spaces; convert to lowercase'],
      ['Phone', { text: [{ text: 'Remove spaces, dashes, brackets, dots; convert to +60 format', options: { breakLine: true } }, { text: 'Example: 012-345 6789 → +60123456789', options: { color: C.MUTED } }] }],
      ['Full Name', 'Remove extra spaces; keep original case'],
      ['Company', 'Match key: lowercase; remove punctuation and legal suffixes (Sdn Bhd, Berhad, PLT)'],
      ['Region, Product Interest', 'Fixed choice lists only; no free text'],
      ['Dates and times', 'Stored in UTC; displayed in Malaysia time (UTC+8)'],
    ], { x: tx, y: 2.1, w: tw, colW: [2.3, tw - 2.3], fs: 13, hfs: 13, rowH: 0.6, colOpts: [{ bold: true, color: C.PETROL }, {}] });
    footer(s, n);
  }

  // 8. Validation
  {
    const s = pres.addSlide(); n++;
    title(s, 'Validation', 'BUSINESS RULES  ·  2 OF 5');
    intro(s, 'Checks that a lead is complete enough to work. Forgiving by design: no lead is rejected outright.');
    const lw = 6.6;
    header(s, 'Rules', M, 2.05, lw);
    const rules = [
      'Full Name must be present',
      'At least one of Email or Phone must be present',
      'Email, if present, must be a valid email format',
      'Phone, if present, must be a valid phone format (+60 followed by 8 to 10 digits)',
      'Source must be an approved value',
    ];
    for (let i = 0; i < rules.length; i++) {
      const y = 2.6 + i * 0.84;
      box(s, { x: M, y, w: lw, h: 0.68, fill: { color: C.TINT2 }, line: { color: C.LINE, width: 0.75 } });
      await iconCirc(s, 'FaCheck', M + 0.15, y + 0.13, 0.42, C.TEAL);
      txt(s, rules[i], { x: M + 0.75, y, w: lw - 0.9, h: 0.68, fontSize: 14, valign: 'middle' });
    }
    const rx = 7.7, rw = W - M - rx;
    box(s, { x: rx, y: 2.05, w: rw, h: 4.85, fill: { color: C.AMBER_T }, line: { color: C.AMBER, width: 1 } });
    txt(s, 'If a Rule Fails', { x: rx + 0.3, y: 2.25, w: rw - 0.6, h: 0.4, fontSize: 18, bold: true, color: C.AMBER_D, valign: 'middle' });
    const fails = [
      'Lead is still saved to the register; nothing is lost',
      'Lead is flagged "Incomplete"',
      'Lead goes to the Exception Queue for a staff member to enrich or reject',
    ];
    vline(s, rx + 0.3 + 0.25, 3.05, 2.5, C.AMBER, 2);
    fails.forEach((t, i) => {
      const y = 2.8 + i * 1.25;
      circ(s, rx + 0.3, y, 0.5, C.AMBER, i + 1, { fontSize: 15 });
      txt(s, t, { x: rx + 1.0, y: y + 0.02, w: rw - 1.3, h: 1.0, fontSize: 14 });
    });
    footer(s, n);
  }

  // 9. Duplicate Detection
  {
    const s = pres.addSlide(); n++;
    title(s, 'Duplicate Detection', 'BUSINESS RULES  ·  3 OF 5');
    intro(s, 'Rules compare the standardised match keys. The strongest match wins.');
    const conf = { text: 'Confirmed duplicate', options: { bold: true, color: C.TEAL, fill: { color: C.TINT } } };
    const susp = { text: 'Suspected duplicate', options: { bold: true, color: C.AMBER_D, fill: { color: C.AMBER_T } } };
    const pr = t => ({ text: t, options: { bold: true, align: 'center', color: C.PETROL } });
    table(s, ['Priority', 'Rule', 'Outcome'], [
      [pr('1'), 'Exact email match, last 90 days', Object.assign({}, conf)],
      [pr('2'), 'Exact phone match, last 90 days', Object.assign({}, conf, { options: Object.assign({}, conf.options) })],
      [pr('3'), 'Same full name + same company, last 30 days', Object.assign({}, susp)],
      [pr('4'), 'Similar name (fuzzy) + same email domain, last 30 days', Object.assign({}, susp, { options: Object.assign({}, susp.options) })],
    ], { x: M, y: 2.0, w: CW, colW: [1.3, CW - 1.3 - 3.4, 3.4], fs: 14, hfs: 14, rowH: 0.48 });
    header(s, 'Handling', M, 4.6, 4);
    const hs = [
      ['FaLink', C.TEAL, 'Confirmed', 'Linked to the master record via Duplicate Of; not assigned to an owner'],
      ['FaClone', C.AMBER, 'Suspected', 'Sent to the Exception Queue for a human to accept or reject'],
      ['FaShieldHalved', C.PETROL, 'Nothing deleted', 'All evidence kept for audit'],
    ];
    const gap = 0.3, cw = (CW - 2 * gap) / 3;
    for (let i = 0; i < 3; i++) {
      const [ic, col, l, t] = hs[i];
      const x = M + i * (cw + gap), y = 5.1;
      box(s, { x, y, w: cw, h: 1.75, fill: { color: C.WHITE }, line: { color: C.LINE, width: 0.75 }, shadow: shadow() });
      await iconCirc(s, ic, x + 0.25, y + 0.25, 0.55, col);
      txt(s, l, { x: x + 0.95, y: y + 0.25, w: cw - 1.1, h: 0.55, fontSize: 16, bold: true, color: C.PETROL, valign: 'middle' });
      txt(s, t, { x: x + 0.25, y: y + 0.95, w: cw - 0.5, h: 0.7, fontSize: 13 });
    }
    footer(s, n);
  }

  // 10. Lead Assignment
  {
    const s = pres.addSlide(); n++;
    title(s, 'Lead Assignment', 'BUSINESS RULES  ·  4 OF 5');
    intro(s, 'Rules stored in a SharePoint "Assignment Rules" list. Business users update them without code.');
    const rules = [
      ['FaMapLocationDot', 'By Region', 'For example, North → Owner A, South → Owner B'],
      ['FaBoxOpen', 'By Product Interest', 'Product leads → product specialist'],
      ['FaArrowsRotate', 'Round-robin fallback', 'No rule match → next owner in the rotation list'],
      ['FaUserClock', 'Out-of-office', 'Owner on leave (per "Owner Availability" list) → next eligible owner'],
    ];
    const gap = 0.25, cw = (CW - 3 * gap) / 4, y = 2.1, h = 2.95;
    for (let i = 0; i < 4; i++) {
      const [ic, l, t] = rules[i];
      const x = M + i * (cw + gap);
      box(s, { x, y, w: cw, h, fill: { color: C.TINT } });
      await iconCirc(s, ic, x + 0.3, y + 0.3, 0.75, C.TEAL);
      txt(s, l, { x: x + 0.3, y: y + 1.25, w: cw - 0.6, h: 0.45, fontSize: 17, bold: true, color: C.PETROL, valign: 'middle' });
      txt(s, t, { x: x + 0.3, y: y + 1.8, w: cw - 0.6, h: 1.05, fontSize: 15 });
    }
    box(s, { x: M, y: 5.4, w: CW, h: 1.2, fill: { color: C.PETROL } });
    await icon(s, 'FaClipboardCheck', M + 0.35, 5.4 + 0.35, 0.5, C.AMBER);
    txt(s, labelDesc('Traceability', 'Every assignment is logged with the rule that triggered it.', 16, 15, C.AMBER, C.WHITE),
      { x: M + 1.15, y: 5.4, w: CW - 1.4, h: 1.2, valign: 'middle' });
    footer(s, n);
  }

  // 11. SLA Calculation
  {
    const s = pres.addSlide(); n++;
    title(s, 'SLA Calculation', 'BUSINESS RULES  ·  5 OF 5');
    box(s, { x: M, y: 1.5, w: CW, h: 0.9, fill: { color: C.PETROL } });
    txt(s, [
      { text: 'SLA Due', options: { bold: true, color: C.AMBER } },
      { text: '  =  Assignment Date  +  agreed SLA hours  ', options: { color: C.WHITE } },
      { text: '(working hours only)', options: { italic: true, color: C.PALE } },
    ], { x: M + 0.3, y: 1.5, w: CW - 0.6, h: 0.9, fontSize: 22, valign: 'middle' });
    const lw = 3.7;
    txt(s, '48', { x: M, y: 2.65, w: lw, h: 1.15, fontFace: TF, fontSize: 72, bold: true, color: C.TEAL, valign: 'middle' });
    txt(s, 'Working hours: default SLA', { x: M, y: 3.85, w: lw, h: 0.4, fontSize: 15, bold: true });
    txt(s, 'Tunable per source or product in a later phase', { x: M, y: 4.3, w: lw, h: 0.7, fontSize: 13, color: C.MUTED });
    txt(s, 'SLA window', { x: M, y: 5.45, w: lw, h: 0.3, fontSize: 11, bold: true, color: C.MUTED });
    const bx = M, by = 5.85, bh = 0.32;
    box(s, { x: bx, y: by, w: 2.25, h: bh, fill: { color: C.TEAL } }, false);
    box(s, { x: bx + 2.25, y: by, w: 0.75, h: bh, fill: { color: C.AMBER } }, false);
    box(s, { x: bx + 3.0, y: by, w: 0.6, h: bh, fill: { color: C.RED } }, false);
    vline(s, bx + 3.0, by - 0.12, bh + 0.24, C.PETROL, 2);
    txt(s, 'Assigned', { x: bx, y: by + 0.42, w: 1.2, h: 0.28, fontSize: 10.5, color: C.MUTED });
    txt(s, 'SLA Due', { x: bx + 2.5, y: by + 0.42, w: 1.0, h: 0.28, fontSize: 10.5, color: C.MUTED, align: 'center' });

    const rx = 4.8, rw = W - M - rx;
    header(s, 'SLA States', rx, 2.65, rw);
    const st = [
      ['On Track', C.TEAL, C.TINT, 'Before SLA Due; more than 25% of the SLA window remains'],
      ['Approaching', C.AMBER, C.AMBER_T, '25% or less of the SLA window remains'],
      ['Breached', C.RED, C.RED_T, 'Past SLA Due; lead not yet Contacted or Closed'],
    ];
    st.forEach(([l, col, tint, t], i) => {
      const y = 3.2 + i * 1.25;
      box(s, { x: rx, y, w: rw, h: 1.05, fill: { color: tint } });
      box(s, { x: rx, y, w: 2.2, h: 1.05, fill: { color: col } });
      txt(s, l, { x: rx, y, w: 2.2, h: 1.05, fontSize: 17, bold: true, color: C.WHITE, align: 'center', valign: 'middle' });
      txt(s, t, { x: rx + 2.5, y, w: rw - 2.8, h: 1.05, fontSize: 15, valign: 'middle' });
    });
    footer(s, n);
  }

  // 12. Notifications
  {
    const s = pres.addSlide(); n++;
    title(s, 'Notifications, Reminders & SLA Tracking');
    intro(s, 'All notifications sent by Power Automate via email and Microsoft Teams.');
    table(s, ['Trigger', 'Sent To', 'Content'], [
      ['New lead assigned', 'Assigned Owner', 'Lead summary, SLA Due date, one-click link'],
      ['Approaching SLA (25% left)', 'Assigned Owner', 'Reminder with SLA countdown'],
      ['SLA breached', 'Owner + Team Manager', 'Escalation with lead details and time overdue'],
      ['Duplicate suspected', 'Exception Reviewer', 'Link to both records for side-by-side comparison'],
      ['Incomplete record', 'Exception Reviewer', 'Missing fields and link to enrich'],
      ['Unassigned after 1 hour', 'Team Manager', 'No assignment rule matched'],
    ], { x: M, y: 1.95, w: CW, colW: [3.1, 2.7, CW - 5.8], fs: 13, hfs: 13, rowH: 0.42, colOpts: [{ bold: true, color: C.PETROL }, {}, {}] });
    header(s, 'Controls', M, 5.15, 4, { fontSize: 16 });
    const ctl = [
      ['FaLink', 'Deep link', 'Every message opens the lead in one click'],
      ['FaMoon', 'Quiet hours', 'No non-urgent messages outside working hours'],
      ['FaGaugeHigh', 'Rate-limiting', 'A burst of leads does not create a burst of messages'],
    ];
    const gap = 0.3, cw = (CW - 2 * gap) / 3;
    for (let i = 0; i < 3; i++) {
      const [ic, l, t] = ctl[i];
      const x = M + i * (cw + gap), y = 5.6;
      box(s, { x, y, w: cw, h: 1.25, fill: { color: C.TINT } });
      await iconCirc(s, ic, x + 0.2, y + 0.33, 0.58, C.TEAL);
      txt(s, labelDesc(l, t, 15, 12.5), { x: x + 0.95, y: y + 0.12, w: cw - 1.1, h: 1.0, valign: 'middle' });
    }
    footer(s, n);
  }

  // 13. Exception Management
  {
    const s = pres.addSlide(); n++;
    title(s, 'Exception Management');
    intro(s, 'Leads that need human judgement get one dedicated home: the Exception Queue in the Power App.');
    header(s, 'Four Exception Categories', M, 2.05, 7.4);
    const cats = [
      ['Incomplete records', 'Missing or invalid mandatory fields'],
      ['Suspected duplicates', 'Awaiting confirmation or rejection'],
      ['Unassigned leads', 'No assignment rule matched'],
      ['Breached SLAs', 'Overdue and awaiting action'],
    ];
    const lw = 7.4, gap = 0.3, cw = (lw - gap) / 2, ch = 2.0;
    cats.forEach(([l, t], i) => {
      const x = M + (i % 2) * (cw + gap), y = 2.6 + Math.floor(i / 2) * (ch + 0.25);
      box(s, { x, y, w: cw, h: ch, fill: { color: C.AMBER_T } });
      circ(s, x + 0.3, y + 0.3, 0.55, C.AMBER, i + 1, { fontSize: 16 });
      txt(s, l, { x: x + 0.3, y: y + 1.0, w: cw - 0.6, h: 0.4, fontSize: 17, bold: true, color: C.PETROL, valign: 'middle' });
      txt(s, t, { x: x + 0.3, y: y + 1.42, w: cw - 0.6, h: 0.45, fontSize: 14 });
    });
    const rx = M + lw + 0.5, rw = W - M - rx;
    header(s, 'Reviewer Actions', rx, 2.05, rw);
    ['Enrich', 'Reassign', 'Merge', 'Mark Not-Duplicate', 'Close as Invalid'].forEach((a, i) => {
      const y = 2.6 + i * 0.56;
      box(s, { x: rx, y, w: rw, h: 0.44, fill: { color: C.TINT }, rectRadius: 0.22 });
      txt(s, a, { x: rx + 0.25, y, w: rw - 0.5, h: 0.44, fontSize: 14, bold: true, color: C.TEAL, valign: 'middle' });
    });
    header(s, 'Controls', rx, 5.5, rw);
    txt(s, bullets([
      'Every action is written to the audit history',
      'Queue size and exception age are measured on the dashboard',
    ]), { x: rx, y: 5.95, w: rw, h: 0.95, fontSize: 13, paraSpaceAfter: 4 });
    footer(s, n);
  }

  // 14. Management Dashboard
  {
    const s = pres.addSlide(); n++;
    title(s, 'Management Dashboard', 'DASHBOARDS & AUDIT HISTORY  ·  1 OF 2');
    intro(s, 'One Power BI report on the SharePoint register. Refreshed on a schedule or on demand.');
    const fx = M, fy = 1.95, fw = 8.45, fh = 4.95;
    box(s, { x: fx, y: fy, w: fw, h: fh, fill: { color: C.GRAY_T }, line: { color: C.LINE, width: 0.75 } });
    txt(s, 'Views', { x: fx + 0.2, y: fy + 0.1, w: 3, h: 0.35, fontSize: 14, bold: true, color: C.PETROL, valign: 'middle' });
    const kpis = ['Total open leads', 'New this week', 'SLA compliance %', 'Average time to first contact'];
    const iw = fw - 0.4, kg = 0.15, kw = (iw - 3 * kg) / 4;
    kpis.forEach((k, i) => {
      const x = fx + 0.2 + i * (kw + kg), y = fy + 0.5;
      box(s, { x, y, w: kw, h: 0.95, fill: { color: C.WHITE }, line: { color: C.LINE, width: 0.75 } });
      txt(s, [
        { text: 'KPI TILE', options: { fontSize: 9, bold: true, color: C.TEAL, charSpacing: 1, breakLine: true } },
        { text: k, options: { fontSize: 13, bold: true, color: C.PETROL } },
      ], { x: x + 0.15, y: y + 0.08, w: kw - 0.3, h: 0.8, valign: 'middle' });
    });
    const panels = [
      ['Leads by Source', 'Volume per intake channel'],
      ['Leads by Owner', 'Open leads per owner'],
      ['Ageing buckets', '0–2, 3–5, 6–10, 11+ days'],
      ['SLA status', 'On Track vs Approaching vs Breached'],
      ['Funnel', 'New → Assigned → Contacted → Qualified → Closed'],
      ['Exceptions', 'Queue size and average time to resolution'],
    ];
    const pw = (iw - 2 * kg) / 3, ph = 1.52;
    panels.forEach(([l, t], i) => {
      const x = fx + 0.2 + (i % 3) * (pw + kg), y = fy + 1.6 + Math.floor(i / 3) * (ph + kg);
      box(s, { x, y, w: pw, h: ph, fill: { color: C.WHITE }, line: { color: C.LINE, width: 0.75 } });
      txt(s, labelDesc(l, t, 14, 12.5), { x: x + 0.18, y: y + 0.15, w: pw - 0.36, h: ph - 0.3 });
    });
    const rx = fx + fw + 0.3, rw = W - M - rx;
    box(s, { x: rx, y: fy, w: rw, h: fh, fill: { color: C.PETROL } });
    txt(s, 'Access', { x: rx + 0.3, y: fy + 0.25, w: rw - 0.6, h: 0.45, fontSize: 18, bold: true, color: C.AMBER, valign: 'middle' });
    await icon(s, 'FaListCheck', rx + 0.3, fy + 1.0, 0.4, C.PALE);
    txt(s, labelDesc('Filters', 'Date range, source, region, owner', 15, 13, C.WHITE, C.PALE), { x: rx + 0.3, y: fy + 1.5, w: rw - 0.6, h: 1.1 });
    await icon(s, 'FaShieldHalved', rx + 0.3, fy + 2.75, 0.4, C.PALE);
    txt(s, labelDesc('Row-level security', 'Owners see their own leads; managers see all leads', 15, 13, C.WHITE, C.PALE), { x: rx + 0.3, y: fy + 3.25, w: rw - 0.6, h: 1.4 });
    footer(s, n);
  }

  // 15. Audit History
  {
    const s = pres.addSlide(); n++;
    title(s, 'Audit History', 'DASHBOARDS & AUDIT HISTORY  ·  2 OF 2');
    const lw = 4.9;
    header(s, 'Two Layers', M, 1.5, lw);
    const layers = [
      ['FaCodeBranch', 'SharePoint version history', 'Raw field changes'],
      ['FaFileLines', 'Audit Log list', 'Key events in plain English'],
    ];
    for (let i = 0; i < 2; i++) {
      const [ic, l, t] = layers[i];
      const y = 2.0 + i * 1.5;
      box(s, { x: M, y, w: lw, h: 1.3, fill: { color: C.TINT } });
      await iconCirc(s, ic, M + 0.25, y + 0.33, 0.64, C.TEAL);
      txt(s, labelDesc(l, t, 16, 14), { x: M + 1.1, y: y + 0.15, w: lw - 1.3, h: 1.0, valign: 'middle' });
    }
    box(s, { x: M, y: 5.2, w: lw, h: 1.7, fill: { color: C.PETROL } });
    txt(s, labelDesc('Outcome', 'Every lead is fully traceable. This supports operational review and compliance.', 16, 15, C.AMBER, C.WHITE),
      { x: M + 0.3, y: 5.2, w: lw - 0.6, h: 1.7, valign: 'middle', paraSpaceAfter: 4 });
    const rx = 6.1, rw = W - M - rx;
    header(s, 'Events Logged', rx, 1.5, rw);
    const ev = [
      ['Lead created', 'source, channel, captured by'],
      ['Validation', 'outcome'],
      ['Duplicate check', 'rule triggered, matched lead ID'],
      ['Assignment', 'rule applied, owner assigned'],
      ['Status change', 'from, to, by whom'],
      ['Notification sent', 'recipient, type, time'],
      ['SLA state change', 'On Track → Approaching → Breached'],
      ['Exception resolution', 'action taken, by whom'],
    ];
    const ey = 2.05, es = 0.6;
    vline(s, rx + 0.11, ey + 0.2, (ev.length - 1) * es, C.LINE, 2);
    ev.forEach(([l, t], i) => {
      const y = ey + i * es;
      s.addShape(pres.shapes.OVAL, { x: rx, y: y + 0.09, w: 0.22, h: 0.22, fill: { color: C.TEAL }, line: { color: C.WHITE, width: 1.5 } });
      txt(s, [
        { text: l + ':  ', options: { bold: true, color: C.PETROL } },
        { text: t, options: { color: C.INK } },
      ], { x: rx + 0.45, y, w: rw - 0.45, h: 0.4, fontSize: 14.5, valign: 'middle' });
    });
    footer(s, n);
  }

  // 16. Technology Stack
  {
    const s = pres.addSlide(); n++;
    title(s, 'Technology Stack');
    txt(s, 'Pipeline and Tools', { x: M, y: 1.38, w: CW, h: 0.4, fontSize: 16, bold: true, color: C.PETROL, valign: 'middle' });
    const st = t => ({ text: t, options: { bold: true, color: C.WHITE, fill: { color: C.TEAL } } });
    table(s, ['Stage', 'Tool', 'Role'], [
      [st('1. Intake'), 'Microsoft Forms · Power Automate (Outlook connector) · Power Apps (canvas)', 'Web form · shared inbox · staff manual entry'],
      [st('2. Register'), 'SharePoint Lists', 'Lead Register and configuration lists'],
      [st('3. Validate'), 'Power Automate · Power Apps', 'Standardisation and validation rules; checks at the point of entry'],
      [st('4. Deduplicate'), 'Power Automate · SharePoint Lists', 'Match rules against indexed match keys'],
      [st('5. Assign'), 'Power Automate · SharePoint Lists', 'Assignment Rules and Owner Availability lists'],
      [st('6. Notify'), 'Power Automate · Outlook · Microsoft Teams', 'Assignment alerts, reminders, escalations'],
      [st('7. Track'), 'Power Automate · Power Apps', 'Scheduled SLA timers and escalations; status updates and Exception Queue'],
      [st('8. Report'), 'Power BI · SharePoint Lists', 'Management dashboard; Audit Log list and version history'],
    ], { x: M, y: 1.85, w: CW, colW: [1.9, 4.9, CW - 6.8], fs: 12, hfs: 12.5, rowH: 0.41 });
    const fy = 6.0, fh = 0.9;
    box(s, { x: M, y: fy, w: CW, h: fh, fill: { color: C.PETROL } });
    txt(s, [
      { text: 'Foundation', options: { bold: true, fontSize: 15, color: C.AMBER, breakLine: true } },
      { text: 'All stages', options: { fontSize: 11.5, color: C.PALE } },
    ], { x: M + 0.25, y: fy, w: 1.6, h: fh, valign: 'middle' });
    const fnd = [
      ['Microsoft Entra ID', 'Single sign-on, role-based access, row-level security'],
      ['SharePoint document library', 'User guide, admin guide, test evidence'],
      ['Environment', "The client's existing Microsoft 365 tenant"],
    ];
    const fx0 = M + 1.95, fw = (CW - 1.95 - 0.2) / 3;
    fnd.forEach(([l, t], i) => {
      txt(s, labelDesc(l, t, 13, 11.5, C.WHITE, C.PALE), { x: fx0 + i * fw, y: fy + 0.08, w: fw - 0.2, h: fh - 0.16, valign: 'middle' });
    });
    footer(s, n);
  }

  // 17. Success Metrics
  {
    const s = pres.addSlide(); n++;
    title(s, 'Success Metrics');
    table(s, ['Metric', 'Definition', 'Target*'], [
      ['Manual consolidation effort', 'Staff hours per week to collect, copy, and reconcile leads', 'At least 70% reduction from baseline'],
      ['Time to assignment', 'Median time from lead arrival to owner assignment; % assigned within 15 minutes', 'Median under 15 minutes for complete leads'],
      ['First response and SLA compliance', 'Median working hours from assignment to first contact; % contacted before SLA Due', '90% or more contacted within SLA'],
      ['Duplicate detection quality', '% of assigned leads later found to be duplicates; % of suspected duplicates confirmed by reviewers', 'Under 2% missed; confirm rate tracked for rule tuning'],
      ['Exception handling', '% of leads entering the Exception Queue, by category; median time to resolve', 'Median resolution under 1 working day'],
      ['Data completeness at intake', '% of leads that pass validation on arrival, by channel', 'Forms and Power App near 100%; email tracked separately'],
      ['Capture coverage', '% of leads recorded through the three approved channels, from a spot check of staff contacts', '95% or more'],
    ], { x: M, y: 1.4, w: CW, colW: [2.9, CW - 2.9 - 3.6, 3.6], fs: 11.5, hfs: 12.5, rowH: 0.5, hRowH: 0.4, colOpts: [{ bold: true, color: C.PETROL }, {}, {}] });
    header(s, 'Measurement Plan', M, 5.55, 5, { fontSize: 15, h: 0.35 });
    const mp = [
      ['Baseline', 'Discover phase (Week 1), from recent leads in inboxes and spreadsheets and a short staff time log'],
      ['Follow-up', '4 weeks after go-live, from the Power BI dashboard and the Audit Log'],
    ];
    const gap = 0.3, cw = (CW - gap) / 2;
    mp.forEach(([l, t], i) => {
      const x = M + i * (cw + gap);
      box(s, { x, y: 5.95, w: cw, h: 0.72, fill: { color: C.TINT } });
      txt(s, [
        { text: l + '   ', options: { bold: true, color: C.TEAL } },
        { text: t, options: {} },
      ], { x: x + 0.2, y: 5.95, w: cw - 0.4, h: 0.72, fontSize: 12.5, valign: 'middle' });
    });
    txt(s, '*Example targets. Final targets agreed with the client after the Discover baseline.',
      { x: M, y: 6.75, w: CW, h: 0.3, fontSize: 10.5, italic: true, color: C.MUTED });
    footer(s, n);
  }

  // 18. Delivery Approach
  {
    const s = pres.addSlide(); n++;
    title(s, 'Delivery Approach');
    intro(s, 'Five phases. 10 weeks total. Client feedback at every phase.');
    const ph = [
      ['Discover', 'Workshops to confirm intake channels, mandatory fields, assignment rules, SLA definition, and owners', 'Signed-off requirements pack'],
      ['Design', 'Data model, workflow diagrams, dashboard wireframes, test plan', 'Approved design pack'],
      ['Build', 'Configure SharePoint lists, build the Power App, wire up Power Automate flows, build the Power BI report', 'Weekly demos'],
      ['Test', 'Unit tests, end-to-end scenarios, UAT with real business users, fixes', 'UAT sign-off'],
      ['Deploy & Handover', 'Move to production, load reference data, run training, hand over documentation', 'Live system and handover pack'],
    ];
    const k = 5, gap = 0.2, cw = (CW - gap * (k - 1)) / k, cy = 2.0, d = 0.55, cardY = 2.8, cardH = 4.05;
    hline(s, M + cw / 2, cy + d / 2, (k - 1) * (cw + gap), C.TEAL, 2.5);
    ph.forEach(([nm, wh, out], i) => {
      const x = M + i * (cw + gap);
      circ(s, x + (cw - d) / 2, cy, d, C.TEAL, i + 1, { fontSize: 16 });
      box(s, { x, y: cardY, w: cw, h: cardH, fill: { color: C.TINT } });
      txt(s, nm, { x: x + 0.18, y: cardY + 0.18, w: cw - 0.36, h: 0.45, fontSize: 16, bold: true, color: C.PETROL, valign: 'middle' });
      txt(s, wh, { x: x + 0.18, y: cardY + 0.75, w: cw - 0.36, h: 2.0, fontSize: 13.5 });
      box(s, { x: x + 0.12, y: cardY + cardH - 1.07, w: cw - 0.24, h: 0.95, fill: { color: C.WHITE } });
      txt(s, [
        { text: 'OUTPUT', options: { fontSize: 9.5, bold: true, color: C.TEAL, charSpacing: 1, breakLine: true } },
        { text: out, options: { fontSize: 12.5, bold: true, color: C.PETROL } },
      ], { x: x + 0.24, y: cardY + cardH - 1.02, w: cw - 0.48, h: 0.85, valign: 'middle' });
    });
    footer(s, n);
  }

  // 19. Testing & Quality Assurance
  {
    const s = pres.addSlide(); n++;
    title(s, 'Testing & Quality Assurance');
    intro(s, 'Three layers catch small bugs and workflow-level issues before go-live.');
    const lw = 4.6;
    const cardHead = (x, y, num, t, w) => {
      circ(s, x + 0.25, y + 0.22, 0.46, C.TEAL, num, { fontSize: 15 });
      txt(s, t, { x: x + 0.85, y: y + 0.22, w: w - 1.0, h: 0.46, fontSize: 17, bold: true, color: C.PETROL, valign: 'middle' });
    };
    box(s, { x: M, y: 1.95, w: lw, h: 2.4, fill: { color: C.TINT2 }, line: { color: C.LINE, width: 0.75 } });
    cardHead(M, 1.95, 1, 'Unit Testing', lw);
    txt(s, bullets([
      'Each Power Automate flow tested alone with representative inputs: valid lead, incomplete lead, duplicate lead, boundary SLA times',
      'Each Power App screen tested for field validation and error handling',
    ]), { x: M + 0.25, y: 2.8, w: lw - 0.5, h: 1.45, fontSize: 12.5, paraSpaceAfter: 5 });
    box(s, { x: M, y: 4.5, w: lw, h: 2.4, fill: { color: C.TINT2 }, line: { color: C.LINE, width: 0.75 } });
    cardHead(M, 4.5, 3, 'User Acceptance Testing (UAT)', lw);
    txt(s, bullets([
      'Real business users test against a UAT script and in a free-form session',
      'Their feedback drives the final round of fixes',
      'All results logged as test evidence for handover',
    ]), { x: M + 0.25, y: 5.35, w: lw - 0.5, h: 1.45, fontSize: 12.5, paraSpaceAfter: 5 });
    const rx = M + lw + 0.3, rw = W - M - rx;
    box(s, { x: rx, y: 1.95, w: rw, h: 4.95, fill: { color: C.TINT2 }, line: { color: C.LINE, width: 0.75 } });
    cardHead(rx, 1.95, 2, 'End-to-End Testing', rw);
    txt(s, 'Scripted scenarios move a lead through the full pipeline. The audit log is checked at each stage.',
      { x: rx + 0.25, y: 2.8, w: rw - 0.5, h: 0.6, fontSize: 13 });
    table(s, ['Scenario', 'Test'], [
      ['Happy path', 'Valid form → assigned → contacted → qualified → closed within SLA'],
      ['Duplicate path', 'Form submission with an email that matches an existing lead'],
      ['Exception path', 'Email with a missing sender name'],
      ['SLA breach path', 'Assigned lead with no follow-up past SLA Due'],
      ['Unassigned path', 'Lead with no matching assignment rule'],
    ], { x: rx + 0.25, y: 3.5, w: rw - 0.5, colW: [1.9, rw - 0.5 - 1.9], fs: 12.5, hfs: 12.5, rowH: 0.52, colOpts: [{ bold: true, color: C.PETROL }, {}] });
    footer(s, n);
  }

  // 20. Documentation & Handover
  {
    const s = pres.addSlide(); n++;
    title(s, 'Documentation & Handover');
    box(s, { x: M, y: 1.45, w: CW, h: 0.7, fill: { color: C.PETROL } });
    txt(s, [
      { text: 'Goal   ', options: { bold: true, color: C.AMBER } },
      { text: 'The client can run, adjust, and extend the solution without us.', options: { color: C.WHITE } },
    ], { x: M + 0.25, y: 1.45, w: CW - 0.5, h: 0.7, fontSize: 17, valign: 'middle' });
    const lw = 8.2;
    header(s, 'Handover Pack', M, 2.35, lw);
    const pack = [
      ['FaBook', 'User guide', 'Screenshots and one-page quick-reference cards'],
      ['FaScrewdriverWrench', 'Admin guide', 'Assignment rules, SLA settings, owner lists, Exception Queue'],
      ['FaSitemap', 'Solution architecture document', 'One-page overview plus detail on each flow'],
      ['FaTableList', 'Data dictionary', 'Every field, its type, and its validation rule'],
      ['FaClipboardCheck', 'Test evidence pack', 'Test cases, screenshots, UAT sign-off'],
      ['FaClockRotateLeft', 'Change log', 'Decisions made during the build'],
    ];
    const gap = 0.2, cw = (lw - 2 * gap) / 3, ch = 1.95;
    for (let i = 0; i < pack.length; i++) {
      const [ic, l, t] = pack[i];
      const x = M + (i % 3) * (cw + gap), y = 2.85 + Math.floor(i / 3) * (ch + 0.15);
      box(s, { x, y, w: cw, h: ch, fill: { color: C.WHITE }, line: { color: C.LINE, width: 0.75 }, shadow: shadow() });
      await iconCirc(s, ic, x + 0.2, y + 0.22, 0.48, C.TEAL);
      txt(s, l, { x: x + 0.8, y: y + 0.08, w: cw - 0.95, h: 0.75, fontSize: 14, bold: true, color: C.PETROL, valign: 'middle' });
      txt(s, t, { x: x + 0.2, y: y + 0.95, w: cw - 0.4, h: ch - 1.05, fontSize: 12.5 });
    }
    const rx = M + lw + 0.3, rw = W - M - rx;
    box(s, { x: rx, y: 2.35, w: rw, h: 2.5, fill: { color: C.TINT } });
    await iconCirc(s, 'FaChalkboardUser', rx + 0.25, 2.55, 0.55, C.TEAL);
    txt(s, 'Training', { x: rx + 0.95, y: 2.55, w: rw - 1.1, h: 0.55, fontSize: 17, bold: true, color: C.PETROL, valign: 'middle' });
    txt(s, bullets([
      'Two live sessions: one for end users, one for administrators',
      'Recordings kept in SharePoint',
    ]), { x: rx + 0.25, y: 3.3, w: rw - 0.5, h: 1.45, fontSize: 13, paraSpaceAfter: 5 });
    box(s, { x: rx, y: 5.0, w: rw, h: 1.9, fill: { color: C.AMBER_T }, line: { color: C.AMBER, width: 1 } });
    txt(s, labelDesc('Handover Standard', 'Any competent Power Platform maker at the client can continue the solution.', 16, 13.5, C.AMBER_D, C.INK),
      { x: rx + 0.25, y: 5.12, w: rw - 0.5, h: 1.66, valign: 'middle', paraSpaceAfter: 4 });
    footer(s, n);
  }

  // 21. Team Roles
  {
    const s = pres.addSlide(); n++;
    title(s, 'Team Roles');
    const roles = [
      ['A', 'Business Analyst', 'Client-Facing Lead', [
        'Discovery workshops and requirements pack',
        'Field dictionary and logical data model',
        'Power BI dashboard and row-level security',
        'UAT sessions and test evidence pack',
        'Handover pack and end-user training',
        'Client communication, backlog, and scope control']],
      ['B', 'Tech Lead', null, [
        'Solution architecture and environment setup',
        'SharePoint lists, permissions, and Audit Log',
        'Power Apps canvas app and Exception Queue',
        'Assignment engine, SLA timers, and escalations',
        'Flow-failure alerts and production deployment',
        'Admin guide and admin training']],
      ['C', 'Integration & QA Lead', null, [
        'Microsoft Forms and email intake flows',
        'Standardisation, validation, and duplicate detection',
        'Duplicate rule tuning',
        'Notification content and delivery',
        'Unit and end-to-end testing',
        'Peer review of Tech Lead flows']],
    ];
    const gap = 0.3, cw = (CW - 2 * gap) / 3, y = 1.5, h = 5.4;
    roles.forEach(([L, nm, sub, bl], i) => {
      const x = M + i * (cw + gap);
      box(s, { x, y, w: cw, h, fill: { color: C.WHITE }, line: { color: C.LINE, width: 0.75 }, shadow: shadow() });
      circ(s, x + 0.3, y + 0.3, 0.75, C.TEAL, L, { fontSize: 28, fontFace: TF });
      txt(s, [
        { text: 'ROLE ' + L, options: { fontSize: 11, bold: true, color: C.MUTED, charSpacing: 1, breakLine: true } },
        { text: nm, options: { fontSize: 17, bold: true, color: C.PETROL } },
      ], { x: x + 1.2, y: y + 0.28, w: cw - 1.35, h: 0.8, valign: 'middle' });
      if (sub) txt(s, sub, { x: x + 0.3, y: y + 1.2, w: cw - 0.6, h: 0.35, fontSize: 13, bold: true, italic: true, color: C.TEAL, valign: 'middle' });
      hline(s, x + 0.3, y + 1.65, cw - 0.6, C.LINE, 0.75);
      txt(s, bullets(bl), { x: x + 0.3, y: y + 1.85, w: cw - 0.55, h: h - 2.0, fontSize: 14.5, paraSpaceAfter: 10 });
    });
    footer(s, n);
  }

  // 22. Roles by Phase
  {
    const s = pres.addSlide(); n++;
    title(s, 'Roles by Phase');
    const lead = t => ({ text: [
      { text: 'Lead: ', options: { bold: true, color: C.AMBER_D } },
      { text: t, options: {} },
    ], options: { fill: { color: C.AMBER_T } } });
    const ph = t => ({ text: t, options: { bold: true, color: C.PETROL, fill: { color: C.TINT } } });
    table(s, ['Phase', 'Business Analyst', 'Tech Lead', 'Integration & QA Lead'], [
      [ph('1. Discover'), lead('workshops, requirements pack'), 'Licence and connector check', 'Intake channel details'],
      [ph('2. Design'), 'Field dictionary, data model, dashboard wireframes', lead('architecture, list design, environments'), 'Test plan, duplicate rules'],
      [ph('3. Build'), 'Power BI report, configuration lists, UAT scripts', lead('Power App, assignment engine, SLA timers'), 'Intake flows, data quality flows, notifications'],
      [ph('4. Test'), 'UAT sessions, test evidence pack', 'Defect fixes', lead('unit and end-to-end tests, rule tuning')],
      [ph('5. Deploy & Handover'), lead('end-user training, handover pack'), 'Production deployment, admin training', 'Final test evidence, peer review'],
    ], { x: M, y: 1.55, w: CW, colW: [2.35, (CW - 2.35) / 3, (CW - 2.35) / 3, (CW - 2.35) / 3], fs: 13.5, hfs: 14, rowH: 0.95 });
    footer(s, n);
  }

  // 23. Key Assumptions
  {
    const s = pres.addSlide(); n++;
    title(s, 'Key Assumptions', 'RISKS & ASSUMPTIONS  ·  1 OF 2');
    const as = [
      "The client's Microsoft 365 licences include Power Apps, Power Automate, and Power BI at the required level",
      "A sandbox or development environment is available in the client's tenant",
      'The client nominates a business owner and up to two subject-matter experts for workshops and UAT',
      'Historical lead migration, if needed, is a separate small project after go-live',
    ];
    const gap = 0.3, cw = (CW - gap) / 2, ch = 2.45;
    as.forEach((t, i) => {
      const x = M + (i % 2) * (cw + gap), y = 1.6 + Math.floor(i / 2) * (ch + gap);
      box(s, { x, y, w: cw, h: ch, fill: { color: C.TINT2 }, line: { color: C.LINE, width: 0.75 } });
      txt(s, '0' + (i + 1), { x: x + 0.35, y, w: 1.3, h: ch, fontFace: TF, fontSize: 48, bold: true, color: C.TEAL, valign: 'middle' });
      txt(s, t, { x: x + 1.75, y, w: cw - 2.05, h: ch, fontSize: 17, valign: 'middle' });
    });
    footer(s, n);
  }

  // 24. Risks & Mitigations
  {
    const s = pres.addSlide(); n++;
    title(s, 'Risks & Mitigations', 'RISKS & ASSUMPTIONS  ·  2 OF 2');
    const med = { text: 'Medium', options: { bold: true, color: C.AMBER_D, fill: { color: C.AMBER_T }, align: 'center' } };
    const low = { text: 'Low', options: { bold: true, color: C.TEAL, fill: { color: C.TINT }, align: 'center' } };
    const cp = o => ({ text: o.text, options: Object.assign({}, o.options) });
    table(s, ['Risk', 'Likelihood', 'Mitigation'], [
      ['Scope creep during build', cp(med), 'Freeze scope after Design; new requests go to a phase 2 backlog'],
      ['Assignment rules more complex than expected', cp(med), 'Rules kept in a configuration list; business users adjust them without code'],
      ['Adoption resistance from staff used to email', cp(med), 'Involve owners in UAT; short video walkthroughs; simple UI'],
      ['Licensing gap (missing premium connectors)', cp(low), 'Confirm licence coverage in Discover; prefer standard connectors'],
      ['SharePoint list volume grows beyond comfortable limits', cp(low), 'Index fields from the start; migration path to Dataverse'],
      ['Duplicate false positives frustrate users', cp(med), 'Suspected duplicates go to review, not auto-merge; tune rules during UAT'],
    ], { x: M, y: 1.55, w: CW, colW: [4.3, 1.6, CW - 5.9], fs: 14, hfs: 14, rowH: 0.78, colOpts: [{ bold: true, color: C.PETROL }, {}, {}] });
    footer(s, n);
  }

  await pres.writeFile({ fileName: 'Lead_Consolidation_MVP.pptx' });
  console.log('slides:', n);
}

build().catch(e => { console.error(e); process.exit(1); });
