const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TabStopType, TabStopPosition, PositionalTab, PositionalTabAlignment,
  PositionalTabRelativeTo, PositionalTabLeader, ExternalHyperlink,
  TableOfContents, Bookmark
} = require('docx');
const fs = require('fs');

// ════════════════════════════════════════════════════════════════════
// MASTER COLOR TOKENS
// ════════════════════════════════════════════════════════════════════
const C = {
  // Primary Palette
  NAVY:        "0A1628",   // WIBT Sovereign Navy — primary brand authority color
  GOLD:        "C9A84C",   // WIBT Sovereign Gold — accent, highlights, dividers
  SILVER:      "8E9BB5",   // WIBT Silver — secondary text, metadata, captions
  // Neutrals
  WHITE:       "FFFFFF",
  OFF_WHITE:   "F7F8FA",
  LIGHT_GRAY:  "E8ECF2",
  MID_GRAY:    "D0D6E4",
  DARK_GRAY:   "4A5568",
  BLACK:       "0D0D0D",
  // Entity Accent Colors
  WDT_BLUE:    "1A56DB",   // Wiz-Dev Tech — electric tech blue
  MPRS_GREEN:  "0B6E4F",   // MPRS Capital — wealth green
  MR_CRIMSON:  "8B1A1A",   // Massey & Rosupo — legal crimson
  WAC_TEAL:    "0E7490",   // Wisdom Analytics — data teal
  WLE_EARTH:   "6B4226",   // Wisdom Legacy Estates — earth brown
  WIBT_NAVY:   "0A1628",   // WIBT parent — same as NAVY
  // Semantic Colors
  SUCCESS:     "0B6E4F",
  WARNING:     "B45309",
  ERROR:       "8B1A1A",
  INFO:        "1A56DB",
  // Surface variants
  SURFACE_1:   "F7F8FA",
  SURFACE_2:   "E8ECF2",
  SURFACE_3:   "D0D6E4",
};

// ════════════════════════════════════════════════════════════════════
// BORDER HELPERS
// ════════════════════════════════════════════════════════════════════
const b = {
  none:    () => ({ style: BorderStyle.NONE,   size: 0, color: "FFFFFF" }),
  hair:    (c=C.MID_GRAY)  => ({ style: BorderStyle.SINGLE, size: 1, color: c }),
  thin:    (c=C.SILVER)    => ({ style: BorderStyle.SINGLE, size: 2, color: c }),
  mid:     (c=C.NAVY)      => ({ style: BorderStyle.SINGLE, size: 4, color: c }),
  thick:   (c=C.GOLD)      => ({ style: BorderStyle.SINGLE, size: 8, color: c }),
  double:  (c=C.NAVY)      => ({ style: BorderStyle.DOUBLE, size: 4, color: c }),
};

function allBorders(style) {
  return { top: style, bottom: style, left: style, right: style };
}
function sideBorders(opts={}) {
  return {
    top:    opts.top    || b.none(),
    bottom: opts.bottom || b.none(),
    left:   opts.left   || b.none(),
    right:  opts.right  || b.none(),
  };
}

// ════════════════════════════════════════════════════════════════════
// TYPOGRAPHY HELPERS
// ════════════════════════════════════════════════════════════════════
// Font Stack:
//   Display/Headings:  Georgia (serif authority — evokes legal/institutional weight)
//   Body:              Calibri (clean, readable, universally available)
//   Mono/Code:         Courier New (technical precision)
//   Labels/Caps:       Georgia Bold + letter spacing simulation via character spacing

function run(text, opts={}) {
  return new TextRun({
    text,
    font:         opts.font     || "Calibri",
    size:         opts.size     || 22,
    bold:         opts.bold     || false,
    italics:      opts.italics  || false,
    color:        opts.color    || C.BLACK,
    underline:    opts.underline,
    characterSpacing: opts.tracking || 0,
    highlight:    opts.highlight,
  });
}

function P(children, opts={}) {
  const c = typeof children === 'string' ? [run(children, opts)] : children;
  return new Paragraph({
    children: c,
    alignment:    opts.align     || AlignmentType.LEFT,
    spacing: {
      before: opts.before  !== undefined ? opts.before : 0,
      after:  opts.after   !== undefined ? opts.after  : 120,
      line:   opts.line    || undefined,
    },
    indent:       opts.indent    ? { left: opts.indent, hanging: opts.hanging || 0 } : undefined,
    numbering:    opts.numbering,
    border:       opts.border,
    shading:      opts.shading,
    pageBreakBefore: opts.pageBreak || false,
  });
}

// Heading constructors
function H1(text, bookmark) {
  const children = bookmark
    ? [new Bookmark({ id: bookmark, children: [run(text, { font:"Georgia", size:40, bold:true, color:C.WHITE })] })]
    : [run(text, { font:"Georgia", size:40, bold:true, color:C.WHITE })];
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children,
    shading:   { fill: C.NAVY, type: ShadingType.CLEAR },
    spacing:   { before: 480, after: 240 },
    alignment: AlignmentType.LEFT,
  });
}

function H2(text, bookmark) {
  const children = bookmark
    ? [new Bookmark({ id: bookmark, children: [run(text, { font:"Georgia", size:30, bold:true, color:C.NAVY })] })]
    : [run(text, { font:"Georgia", size:30, bold:true, color:C.NAVY })];
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children,
    spacing: { before: 400, after: 160 },
    border:  { bottom: b.mid(C.GOLD) },
  });
}

function H3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [run(text, { font:"Georgia", size:26, bold:true, color:C.NAVY })],
    spacing: { before: 280, after: 100 },
  });
}

function H4(text) {
  return new Paragraph({
    children: [run(text, { font:"Georgia", size:22, bold:true, color:C.GOLD })],
    spacing:  { before: 200, after: 80 },
  });
}

function BODY(text, opts={}) {
  return P([run(text, { font:"Calibri", size:22, color: opts.color || C.BLACK, bold: opts.bold||false, italics: opts.italics||false })],
    { after: 140, ...opts });
}

function CAPTION(text) {
  return P([run(text, { font:"Calibri", size:18, color:C.SILVER, italics:true })], { after:200, align:AlignmentType.CENTER });
}

function LABEL(text, color=C.NAVY) {
  return P([run(text.toUpperCase(), { font:"Georgia", size:18, bold:true, color, tracking:40 })], { after:60 });
}

function CODE(text) {
  return P([run(text, { font:"Courier New", size:20, color:"1A56DB" })],
    { after:80, shading:{ fill:C.LIGHT_GRAY, type:ShadingType.CLEAR },
      border: sideBorders({ left: b.thick(C.WDT_BLUE) }),
      indent: 360 });
}

function MONO(text, color=C.DARK_GRAY) {
  return run(text, { font:"Courier New", size:20, color });
}

function NOTE(text) {
  return new Paragraph({
    children: [
      run("NOTE  ", { font:"Georgia", size:20, bold:true, color:C.GOLD }),
      run(text, { font:"Calibri", size:20, italics:true, color:C.DARK_GRAY }),
    ],
    spacing: { before:80, after:80 },
    indent:  { left: 360 },
    border:  sideBorders({ left: b.thick(C.GOLD) }),
    shading: { fill:"FEFBF0", type: ShadingType.CLEAR },
  });
}

function WARNING_BOX(text) {
  return new Paragraph({
    children: [
      run("⚠  STRICT RULE  ", { font:"Georgia", size:20, bold:true, color:C.MR_CRIMSON }),
      run(text, { font:"Calibri", size:20, color:C.DARK_GRAY }),
    ],
    spacing: { before:100, after:100 },
    indent:  { left: 360 },
    border:  sideBorders({ left: b.thick(C.MR_CRIMSON) }),
    shading: { fill:"FFF5F5", type: ShadingType.CLEAR },
  });
}

function DIVIDER(color=C.GOLD) {
  return P([run("", {})], {
    before: 160, after: 160,
    border: { bottom: { style:BorderStyle.SINGLE, size:6, color } }
  });
}

function SPACER(n=1) {
  return Array.from({length:n}, ()=> P("", { after:0, before:0 }));
}

function PAGE_BREAK() {
  return new Paragraph({ children:[new PageBreak()] });
}

function bullet(text, level=0, font="Calibri") {
  return new Paragraph({
    children: [run(text, { font, size:22 })],
    numbering: { reference:"bullets", level },
    spacing: { before:40, after:80 },
  });
}

function numItem(text) {
  return new Paragraph({
    children: [run(text, { font:"Calibri", size:22 })],
    numbering: { reference:"numbers", level:0 },
    spacing: { before:40, after:80 },
  });
}

// ════════════════════════════════════════════════════════════════════
// TABLE HELPERS
// ════════════════════════════════════════════════════════════════════
function TH(text, w, color=C.NAVY) {
  return new TableCell({
    children: [new Paragraph({
      children:[run(text.toUpperCase(), { font:"Georgia", size:18, bold:true, color:C.WHITE, tracking:20 })],
      spacing:{ before:80, after:80 }, alignment:AlignmentType.LEFT,
    })],
    width: { size:w, type:WidthType.DXA },
    shading: { fill:color, type:ShadingType.CLEAR },
    margins: { top:100, bottom:100, left:160, right:160 },
    borders: allBorders(b.none()),
  });
}

function TD(text, w, shade=false, opts={}) {
  return new TableCell({
    children: [new Paragraph({
      children:[run(text, { font:"Calibri", size:20, color: opts.color||C.BLACK, bold:opts.bold||false })],
      spacing:{ before:80, after:80 },
    })],
    width: { size:w, type:WidthType.DXA },
    shading: { fill: shade ? C.OFF_WHITE : C.WHITE, type:ShadingType.CLEAR },
    margins: { top:100, bottom:100, left:160, right:160 },
    borders: { top:b.hair(), bottom:b.hair(), left:b.none(), right:b.none() },
    verticalAlign: VerticalAlign.CENTER,
  });
}

function TD_ACCENT(text, w, accentColor=C.GOLD) {
  return new TableCell({
    children: [new Paragraph({
      children:[run(text, { font:"Calibri", size:20, bold:true, color:C.NAVY })],
      spacing:{ before:80, after:80 },
    })],
    width: { size:w, type:WidthType.DXA },
    shading: { fill: C.LIGHT_GRAY, type:ShadingType.CLEAR },
    margins: { top:100, bottom:100, left:160, right:160 },
    borders: { top:b.hair(), bottom:b.hair(), left:b.thick(accentColor), right:b.none() },
    verticalAlign: VerticalAlign.CENTER,
  });
}

function TD_SWATCH(hexColor, w) {
  return new TableCell({
    children: [new Paragraph({
      children:[MONO("#" + hexColor, C.WHITE)],
      spacing:{ before:80, after:80 }, alignment:AlignmentType.CENTER,
    })],
    width: { size:w, type:WidthType.DXA },
    shading: { fill: hexColor, type:ShadingType.CLEAR },
    margins: { top:100, bottom:100, left:100, right:100 },
    borders: allBorders(b.none()),
    verticalAlign: VerticalAlign.CENTER,
  });
}

function makeTable(headers, rows, colWidths, headerColor=C.NAVY) {
  const total = colWidths.reduce((a,b)=>a+b,0);
  return new Table({
    width: { size:total, type:WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        children: headers.map((h,i)=>TH(h, colWidths[i], headerColor)),
        tableHeader: true,
      }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((cell,ci) => {
          if (typeof cell === 'object' && cell._type === 'swatch') return TD_SWATCH(cell.hex, colWidths[ci]);
          if (typeof cell === 'object' && cell._type === 'accent') return TD_ACCENT(cell.text, colWidths[ci], cell.color);
          return TD(cell, colWidths[ci], ri%2===1);
        }),
      }))
    ],
  });
}

// Color swatch helper
const swatch = (hex) => ({ _type:'swatch', hex });
const accent = (text, color=C.GOLD) => ({ _type:'accent', text, color });

// ════════════════════════════════════════════════════════════════════
// SECTION BUILDERS
// ════════════════════════════════════════════════════════════════════

// ── COVER ────────────────────────────────────────────────────────────
function makeCover() {
  return [
    // Top rule
    P([], { before:0, after:0, border:{ bottom:{ style:BorderStyle.SINGLE, size:24, color:C.GOLD } } }),
    ...SPACER(3),
    new Paragraph({
      children:[ run("WIBT ECOSYSTEM", { font:"Georgia", size:72, bold:true, color:C.NAVY, tracking:80 }) ],
      alignment: AlignmentType.LEFT, spacing:{ before:0, after:80 },
    }),
    new Paragraph({
      children:[ run("Brand & Theme Standards", { font:"Georgia", size:44, bold:false, color:C.GOLD }) ],
      alignment: AlignmentType.LEFT, spacing:{ before:0, after:60 },
    }),
    new Paragraph({
      children:[ run("Development Reference Document  |  Strict Compliance Required", { font:"Calibri", size:22, italics:true, color:C.SILVER }) ],
      alignment: AlignmentType.LEFT, spacing:{ before:0, after:600 },
    }),
    // Metadata block
    makeTable(
      ["Property","Value"],
      [
        ["Document ID",        "WIBT-THEME-20260512-001"],
        ["Version",            "1.0 — Authoritative"],
        ["Status",             "ACTIVE — All new development must comply"],
        ["Owner",              "Wiz-Dev Tech LLC (WDT) — approved by WIBT"],
        ["Scope",              "CipherNex platform, WIBT Treasury Ledger, masseyrosupo.com, wizdevtech.com, wisdomignited.com, MPRS Capital, WAC, WLE"],
        ["Enforcement",        "Non-compliant UIs must be remediated before merge to main"],
        ["Capacity Vault ID",  "Assigned on blockchain mint"],
        ["Last Updated",       "May 12, 2026"],
      ],
      [2800, 6560]
    ),
    ...SPACER(3),
    new Paragraph({
      children:[ run("Wisdom Ignited Business Trust  ·  EIN 92-6269777  ·  CipherNex Chain ID 777287", { font:"Calibri", size:18, color:C.SILVER }) ],
      alignment: AlignmentType.CENTER, spacing:{ before:0, after:60 },
    }),
    P([], { before:0, after:0, border:{ bottom:{ style:BorderStyle.SINGLE, size:24, color:C.NAVY } } }),
    PAGE_BREAK(),
  ];
}

// ── TOC ──────────────────────────────────────────────────────────────
function makeTOC() {
  return [
    H1("TABLE OF CONTENTS"),
    new TableOfContents("Table of Contents", { hyperlink:true, headingStyleRange:"1-3" }),
    PAGE_BREAK(),
  ];
}

// ── 1. PURPOSE & SCOPE ───────────────────────────────────────────────
function makeSection1() {
  return [
    H1("1.  PURPOSE & SCOPE", "sec1"),
    BODY("This document is the single authoritative source for all visual, typographic, color, spacing, and component standards across the WIBT ecosystem. Every developer, designer, and contractor working on any WIBT-affiliated interface must comply with these specifications without deviation."),
    BODY("This is not a style suggestion. It is a strict compliance standard. Interfaces that deviate from these specifications introduce brand fragmentation that undermines the institutional authority the ecosystem depends on. WIBT's credibility as a sovereign financial and governance platform rests on its visual presentation being as rigorous as its legal framework."),
    DIVIDER(),

    H2("1.1  Governing Principle"),
    BODY("The WIBT visual language embodies one overarching principle:"),
    new Paragraph({
      children:[ run("Sovereign Authority Expressed Through Restraint.", { font:"Georgia", size:28, bold:true, color:C.NAVY }) ],
      spacing:{ before:120, after:120 }, alignment:AlignmentType.CENTER,
    }),
    BODY("Every design decision — color, type, spacing, layout — should communicate institutional gravitas. Nothing should feel casual, consumer-grade, or speculative. The ecosystem deals in bonds, sovereign settlement, and legal instruments. The interface must look and feel like it does."),

    H2("1.2  Scope of Application"),
    makeTable(
      ["Surface","Applies To","Compliance Level"],
      [
        ["CipherNex Admin Portal","admin.wisdomignited.com (Port 3005)","STRICT — zero deviation"],
        ["WIBT Treasury Ledger","45-service modular platform (wisdomignited.com)","STRICT — zero deviation"],
        ["Massey & Rosupo Platform","masseyrosupo.com","STRICT — zero deviation"],
        ["WizDevTech Public Site","wizdevtech.com","STRICT — zero deviation"],
        ["Entity Onboarding","entity-register.html + onboarding wizard","STRICT — zero deviation"],
        ["DocumentService UI","Port 3004 interfaces","STRICT — zero deviation"],
        ["API Documentation","Developer-facing docs","STANDARD — entity colors permitted"],
        ["Internal Tools","Plane, Mattermost, local tooling","ADVISORY — follow where practical"],
        ["Entity ICO Pages","WDT, MPRS, WAC token offering pages","STRICT — entity color variants apply"],
        ["Print / PDF Output","Court packages, governance docs","STRICT — print palette rules apply"],
      ],
      [2600, 3800, 2960]
    ),
    WARNING_BOX("Do not apply consumer-grade design patterns (purple gradients, rounded pill buttons, pastel backgrounds, drop shadows on everything) to any WIBT surface. These patterns signal speculation and distrust — the opposite of what this ecosystem must communicate."),
    PAGE_BREAK(),
  ];
}

// ── 2. COLOR SYSTEM ──────────────────────────────────────────────────
function makeSection2() {
  return [
    H1("2.  COLOR SYSTEM", "sec2"),
    BODY("The WIBT color system is built on a sovereign palette: deep navy authority, gold precision, and silver restraint. The palette has no soft edges, no pastels, and no gradients in primary UI surfaces. All colors are specified as CSS custom properties (variables) and hex codes. No color may be used that is not on this list without a documented exception approved by WDT."),

    H2("2.1  Primary Brand Palette"),
    makeTable(
      ["Token","Swatch","Hex","RGB","Usage"],
      [
        ["--color-navy",    swatch(C.NAVY),   "#0A1628", "10, 22, 40",   "Page backgrounds, section headers, primary authority surfaces"],
        ["--color-gold",    swatch(C.GOLD),   "#C9A84C", "201, 168, 76", "Accents, dividers, icon strokes, QR frame, CTA borders"],
        ["--color-silver",  swatch(C.SILVER), "#8E9BB5", "142, 155, 181","Secondary text, metadata, captions, placeholder text"],
        ["--color-white",   swatch(C.WHITE),  "#FFFFFF", "255, 255, 255","Body text on dark, card backgrounds on light surfaces"],
        ["--color-off-white",swatch(C.OFF_WHITE),"#F7F8FA","247, 248, 250","Page background on light-mode surfaces"],
      ],
      [1600, 720, 1000, 1400, 4640]
    ),
    CAPTION("Table 2.1 — Primary palette. These five tokens are mandatory in every interface."),

    H2("2.2  Neutral Palette"),
    makeTable(
      ["Token","Swatch","Hex","Usage"],
      [
        ["--color-light-gray",swatch(C.LIGHT_GRAY),"#E8ECF2","Card borders, table row stripes, input borders"],
        ["--color-mid-gray",  swatch(C.MID_GRAY),  "#D0D6E4","Dividers, disabled states, skeleton loaders"],
        ["--color-dark-gray", swatch(C.DARK_GRAY),  "#4A5568","Body text on light backgrounds"],
        ["--color-black",     swatch(C.BLACK),      "#0D0D0D","High-emphasis text, data values, legal instrument text"],
      ],
      [1800, 720, 1000, 5840]
    ),

    H2("2.3  Entity Accent Colors"),
    BODY("Each subsidiary entity has a single designated accent color. This color is used for entity-specific elements only — entity badges, entity-scoped navigation tabs, entity ICO pages, and entity headers within the multi-entity platform. Entity accent colors must never replace or compete with the primary navy/gold palette."),
    makeTable(
      ["Entity","Token","Swatch","Hex","Permitted Use"],
      [
        ["Wisdom Ignited Business Trust (WIBT)", "--entity-wibt",  swatch(C.WIBT_NAVY),  "#0A1628","Parent brand — uses navy directly"],
        ["Wiz-Dev Tech LLC (WDT)",               "--entity-wdt",   swatch(C.WDT_BLUE),   "#1A56DB","WDT badges, API docs, tech service headers"],
        ["MPRS Capital",                          "--entity-mprs",  swatch(C.MPRS_GREEN), "#0B6E4F","Settlement indicators, banking UI elements, USD coin badge"],
        ["Massey & Rosupo Co.",                   "--entity-mr",    swatch(C.MR_CRIMSON), "#8B1A1A","Legal matter status, enforcement alerts, arbitration pipeline"],
        ["Wisdom Analytics LLC (WAC)",            "--entity-wac",   swatch(C.WAC_TEAL),   "#0E7490","Charts, analytics dashboards, data outputs"],
        ["Wisdom Legacy Estates (WLE)",           "--entity-wle",   swatch(C.WLE_EARTH),  "#6B4226","Asset records, real property cards, estate instruments"],
      ],
      [2600, 1400, 720, 1000, 3640]
    ),
    NOTE("Entity accent colors are identity markers, not decoration. Apply them only to elements that explicitly belong to that entity's scope. A transaction record in the MPRS settlement module uses --entity-mprs green; the platform chrome around it does not."),

    H2("2.4  Semantic Colors"),
    makeTable(
      ["State","Token","Hex","Swatch","Usage"],
      [
        ["Success / Confirmed", "--color-success", "#0B6E4F", swatch(C.SUCCESS), "Blockchain mint confirmed, settlement cleared, notice delivered"],
        ["Warning / Pending",   "--color-warning", "#B45309", swatch(C.WARNING), "Awaiting signature, pending cure period, 90-day window open"],
        ["Error / Violation",   "--color-error",   "#8B1A1A", swatch(C.ERROR),   "Authentication failure, CI Matrix rejection, enforcement triggered"],
        ["Info / Processing",   "--color-info",    "#1A56DB", swatch(C.INFO),    "API call in progress, document hashing, XRPL submission pending"],
      ],
      [2000, 1400, 1000, 720, 4240]
    ),
    WARNING_BOX("NEVER use green (#00FF00 variants), red (#FF0000 variants), or yellow (#FFFF00 variants) as semantic colors. They are cheap and consumer-grade. Use the semantic tokens above exclusively."),

    H2("2.5  CSS Custom Properties — Required Declaration"),
    BODY("Every frontend surface in the WIBT ecosystem must declare the following CSS custom properties in :root. This is mandatory. Hard-coding color hex values in component styles is prohibited."),
    CODE(":root {"),
    CODE("  /* Primary */"),
    CODE("  --color-navy:       #0A1628;"),
    CODE("  --color-gold:       #C9A84C;"),
    CODE("  --color-silver:     #8E9BB5;"),
    CODE("  --color-white:      #FFFFFF;"),
    CODE("  --color-off-white:  #F7F8FA;"),
    CODE(""),
    CODE("  /* Neutrals */"),
    CODE("  --color-light-gray: #E8ECF2;"),
    CODE("  --color-mid-gray:   #D0D6E4;"),
    CODE("  --color-dark-gray:  #4A5568;"),
    CODE("  --color-black:      #0D0D0D;"),
    CODE(""),
    CODE("  /* Entity Accents */"),
    CODE("  --entity-wibt:  #0A1628;"),
    CODE("  --entity-wdt:   #1A56DB;"),
    CODE("  --entity-mprs:  #0B6E4F;"),
    CODE("  --entity-mr:    #8B1A1A;"),
    CODE("  --entity-wac:   #0E7490;"),
    CODE("  --entity-wle:   #6B4226;"),
    CODE(""),
    CODE("  /* Semantic */"),
    CODE("  --color-success: #0B6E4F;"),
    CODE("  --color-warning: #B45309;"),
    CODE("  --color-error:   #8B1A1A;"),
    CODE("  --color-info:    #1A56DB;"),
    CODE("}"),

    H2("2.6  Dark Mode vs. Light Mode"),
    makeTable(
      ["Surface","Background","Primary Text","Secondary Text","Borders"],
      [
        ["Dark mode (default for all admin/platform)","--color-navy","--color-white","--color-silver","--color-gold"],
        ["Light mode (documents, reports, public pages)","--color-off-white","--color-black","--color-dark-gray","--color-mid-gray"],
        ["Cards on dark","#0F1E36 (navy +10% light)","--color-white","--color-silver","--color-gold at 40% opacity"],
        ["Cards on light","--color-white","--color-black","--color-dark-gray","--color-light-gray"],
        ["Tables on dark","alternate: navy / #0F1E36","--color-white","--color-silver","--color-gold"],
        ["Tables on light","alternate: white / --color-off-white","--color-black","--color-dark-gray","--color-light-gray"],
      ],
      [2400, 1800, 1600, 1600, 1960]
    ),
    WARNING_BOX("Dark mode is the default for all authenticated platform surfaces. Light mode is for document outputs, public marketing pages, and print. Never invert this — a light-mode admin dashboard communicates consumer-grade, not sovereign infrastructure."),
    PAGE_BREAK(),
  ];
}

// ── 3. TYPOGRAPHY ─────────────────────────────────────────────────────
function makeSection3() {
  return [
    H1("3.  TYPOGRAPHY", "sec3"),
    BODY("Typography is the primary carrier of institutional authority in the WIBT ecosystem. The font system is built on a three-family stack: Georgia for display and headings, Calibri for body and UI text, and Courier New for all technical/code content. This stack is universally available without web font dependencies, ensuring rendering consistency across all environments."),
    WARNING_BOX("Inter, Roboto, Arial, and system-ui are prohibited on all WIBT surfaces. They signal generic consumer software. Georgia communicates legal weight and institutional permanence. Do not substitute."),

    H2("3.1  Font Stack"),
    makeTable(
      ["Role","Font","Fallback","Usage"],
      [
        ["Display / Headings", "Georgia",     "'Times New Roman', serif",  "All H1–H4 elements, section titles, entity names, document headers"],
        ["Body / UI Text",     "Calibri",     "Candara, 'Segoe UI', sans-serif","Body copy, labels, form fields, table data, nav items, tooltips"],
        ["Technical / Code",   "Courier New", "'Lucida Console', monospace", "All hash values, addresses, API endpoints, code blocks, transaction IDs"],
      ],
      [1800, 1600, 2800, 3160]
    ),

    H2("3.2  Type Scale"),
    BODY("All sizes are expressed in CSS rem units, with 1rem = 16px base. Pixel equivalents provided for reference. Use the scale tokens — do not interpolate between sizes."),
    makeTable(
      ["Token","Rem","Px","CSS Font","Weight","Tracking","Usage"],
      [
        ["--text-display",  "3.0rem",  "48px", "Georgia",  "700", "0.02em",  "Page titles, hero headings (H1 on landing)"],
        ["--text-h1",       "2.5rem",  "40px", "Georgia",  "700", "0.01em",  "Section headings (H1 in platform)"],
        ["--text-h2",       "1.875rem","30px", "Georgia",  "700", "0",       "Sub-section headings (H2)"],
        ["--text-h3",       "1.5rem",  "24px", "Georgia",  "700", "0",       "Component headings (H3)"],
        ["--text-h4",       "1.25rem", "20px", "Georgia",  "700", "0.03em",  "Card titles, label headings (H4)"],
        ["--text-body-lg",  "1.125rem","18px", "Calibri",  "400", "0",       "Lead paragraphs, important body text"],
        ["--text-body",     "1rem",    "16px", "Calibri",  "400", "0",       "Standard body, form fields, table data"],
        ["--text-body-sm",  "0.875rem","14px", "Calibri",  "400", "0",       "Captions, help text, secondary labels"],
        ["--text-label",    "0.75rem", "12px", "Georgia",  "700", "0.08em",  "ALL CAPS labels, nav items, badge text"],
        ["--text-mono",     "0.875rem","14px", "Courier New","400","0",      "Hash values, wallet addresses, API endpoints"],
      ],
      [1600, 800, 600, 1200, 800, 1000, 3360]
    ),

    H2("3.3  Type Hierarchy Rules"),
    H4("Heading Rules"),
    bullet("H1: one per page section — authority title, always navy background, white text in platform"),
    bullet("H2: section subdivision — navy text, gold bottom border"),
    bullet("H3: component/card title — navy text, no border"),
    bullet("H4: label-weight titles — gold text, often ALL CAPS with tracking"),
    bullet("Headings never use color outside the navy/gold/white triplet"),
    bullet("Heading line-height: 1.2 for H1/H2, 1.3 for H3/H4"),

    H4("Body Text Rules"),
    bullet("Body text on dark backgrounds: --color-white, Calibri, 16px, line-height 1.6"),
    bullet("Body text on light backgrounds: --color-black, Calibri, 16px, line-height 1.6"),
    bullet("Never justify body text — left-aligned only"),
    bullet("Maximum line length: 75 characters (680px at 16px). Do not allow text to span full page width"),
    bullet("Paragraph spacing: 1em bottom margin between paragraphs, no top margin"),

    H4("Label Rules"),
    bullet("Labels always ALL CAPS, Georgia, 12px, letter-spacing 0.08em"),
    bullet("Labels use --color-silver on dark, --color-dark-gray on light"),
    bullet("Form field labels: above the field, 4px gap, never inside (no placeholder-as-label)"),

    H4("Technical / Mono Rules"),
    bullet("All wallet addresses: Courier New, 14px, --entity-wdt blue on light, --color-silver on dark"),
    bullet("All hash values (SHA-256, transaction IDs): Courier New, break-all, with copy-to-clipboard affordance"),
    bullet("All API endpoints: Courier New, preceded by METHOD badge (GET/POST) in entity accent color"),
    bullet("Chain IDs, port numbers, block heights: Courier New inline within Calibri body"),

    H2("3.4  CSS Typography Tokens"),
    CODE(":root {"),
    CODE("  --font-display:  'Georgia', 'Times New Roman', serif;"),
    CODE("  --font-body:     'Calibri', 'Candara', sans-serif;"),
    CODE("  --font-mono:     'Courier New', 'Lucida Console', monospace;"),
    CODE(""),
    CODE("  --text-display:  3rem;"),
    CODE("  --text-h1:       2.5rem;"),
    CODE("  --text-h2:       1.875rem;"),
    CODE("  --text-h3:       1.5rem;"),
    CODE("  --text-h4:       1.25rem;"),
    CODE("  --text-body-lg:  1.125rem;"),
    CODE("  --text-body:     1rem;"),
    CODE("  --text-body-sm:  0.875rem;"),
    CODE("  --text-label:    0.75rem;"),
    CODE("  --text-mono:     0.875rem;"),
    CODE(""),
    CODE("  --leading-tight:   1.2;"),
    CODE("  --leading-normal:  1.6;"),
    CODE("  --leading-relaxed: 1.8;"),
    CODE(""),
    CODE("  --tracking-tight:  -0.01em;"),
    CODE("  --tracking-normal:  0;"),
    CODE("  --tracking-wide:    0.03em;"),
    CODE("  --tracking-widest:  0.08em;"),
    CODE("}"),
    PAGE_BREAK(),
  ];
}

// ── 4. SPACING & LAYOUT ───────────────────────────────────────────────
function makeSection4() {
  return [
    H1("4.  SPACING & LAYOUT", "sec4"),
    BODY("The WIBT layout system uses an 8px base grid. All spacing values are multiples of 8px. No half-steps, no odd values, no component-specific magic numbers. Consistent spatial rhythm communicates precision and control — qualities that are essential for a platform handling sovereign instruments and legal documents."),

    H2("4.1  Spacing Scale"),
    makeTable(
      ["Token","Value","Px","Usage"],
      [
        ["--space-1",  "0.25rem", "4px",   "Icon padding, tight inline gaps"],
        ["--space-2",  "0.5rem",  "8px",   "Base unit — minimum padding for any element"],
        ["--space-3",  "0.75rem", "12px",  "Input internal padding (vertical)"],
        ["--space-4",  "1rem",    "16px",  "Input internal padding (horizontal), card internal padding"],
        ["--space-5",  "1.25rem", "20px",  "Component internal padding — standard"],
        ["--space-6",  "1.5rem",  "24px",  "Section internal padding — standard"],
        ["--space-8",  "2rem",    "32px",  "Component gap — between cards in a grid"],
        ["--space-10", "2.5rem",  "40px",  "Section gap — between major sections"],
        ["--space-12", "3rem",    "48px",  "Page section padding — top/bottom"],
        ["--space-16", "4rem",    "64px",  "Hero / masthead padding"],
        ["--space-20", "5rem",    "80px",  "Maximum section separation"],
      ],
      [1600, 1000, 600, 6160]
    ),
    CODE(":root {"),
    CODE("  --space-1: 0.25rem;  --space-2: 0.5rem;   --space-3: 0.75rem;"),
    CODE("  --space-4: 1rem;     --space-5: 1.25rem;  --space-6: 1.5rem;"),
    CODE("  --space-8: 2rem;     --space-10: 2.5rem;  --space-12: 3rem;"),
    CODE("  --space-16: 4rem;    --space-20: 5rem;"),
    CODE("}"),

    H2("4.2  Layout Grid"),
    makeTable(
      ["Breakpoint","Min Width","Columns","Gutter","Max Content Width"],
      [
        ["Mobile",  "320px",   "4",  "16px", "100% - 32px padding"],
        ["Tablet",  "768px",   "8",  "24px", "720px"],
        ["Desktop", "1024px",  "12", "32px", "1200px"],
        ["Wide",    "1440px",  "12", "32px", "1320px"],
        ["Admin",   "1280px+", "12", "24px", "1440px (admin portal preferred)"],
      ],
      [1200, 1200, 1000, 1000, 4960]
    ),
    bullet("Sidebar navigation width: 240px (collapsed: 64px) — fixed, navy background"),
    bullet("Main content area: full grid minus sidebar and 32px padding"),
    bullet("Modal max-width: 640px (standard), 860px (wide/document view)"),
    bullet("Drawer width: 480px from right edge"),
    bullet("Top navigation height: 64px — fixed, navy background, gold bottom border 1px"),

    H2("4.3  Border Radius"),
    BODY("The WIBT UI is angular. It does not use heavy border radius. Rounded elements signal consumer-grade softness. WIBT infrastructure is precise and institutional."),
    makeTable(
      ["Token","Value","Usage"],
      [
        ["--radius-none",  "0px",   "All primary surfaces: cards, panels, modals, navigation"],
        ["--radius-sm",    "2px",   "Badges, status pills, tags only"],
        ["--radius-md",    "4px",   "Input fields, small buttons only"],
        ["--radius-full",  "9999px","Circular avatar only — no other use"],
      ],
      [1600, 1000, 6760]
    ),
    WARNING_BOX("Do not use border-radius: 8px, 12px, 16px, or 24px on any card, panel, or major container. These values produce the consumer-grade 'app-like' aesthetic that is incompatible with the WIBT brand."),

    H2("4.4  Elevation & Shadow"),
    BODY("Shadows are used sparingly and with precision. The WIBT platform communicates depth through color contrast and border weight, not drop shadows."),
    makeTable(
      ["Token","CSS Value","Usage"],
      [
        ["--shadow-none",    "none","Default — all cards and panels at rest"],
        ["--shadow-border",  "inset 0 0 0 1px var(--color-gold)","Active state on interactive cards (NOT a shadow — a gold inset border)"],
        ["--shadow-sm",      "0 1px 3px rgba(10,22,40,0.12), 0 1px 2px rgba(10,22,40,0.08)","Dropdowns, tooltips only"],
        ["--shadow-md",      "0 4px 6px rgba(10,22,40,0.10), 0 2px 4px rgba(10,22,40,0.08)","Modals, drawers only"],
        ["--shadow-document","0 0 0 1px var(--color-mid-gray)","Document preview cards — flat outline shadow"],
      ],
      [1800, 3400, 4160]
    ),
    PAGE_BREAK(),
  ];
}

// ── 5. COMPONENT STANDARDS ────────────────────────────────────────────
function makeSection5() {
  return [
    H1("5.  COMPONENT STANDARDS", "sec5"),
    BODY("Each component family below specifies exact appearance, states, and CSS variable usage. Deviations require a written exception filed with WDT and approved by WIBT before implementation."),

    H2("5.1  Buttons"),
    makeTable(
      ["Variant","Background","Text","Border","Hover","Use Case"],
      [
        ["Primary",   "var(--color-navy)","var(--color-white)","2px solid var(--color-gold)","background: #1A2F52 (navy +15%)","Main CTA — Submit, Confirm, Mint, Sign"],
        ["Secondary", "transparent","var(--color-navy)","2px solid var(--color-navy)","background: var(--color-light-gray)","Alternative action — Cancel, Back, Export"],
        ["Gold",      "var(--color-gold)","var(--color-navy)","none","opacity: 0.9","High-visibility CTA on dark backgrounds"],
        ["Danger",    "transparent","var(--color-error)","2px solid var(--color-error)","background: rgba(139,26,26,0.08)","Destructive actions — Burn, Revoke, Freeze"],
        ["Ghost",     "transparent","var(--color-silver)","1px solid var(--color-mid-gray)","border-color: var(--color-silver)","Low-emphasis — view, details, more"],
        ["Disabled",  "var(--color-light-gray)","var(--color-silver)","none","cursor: not-allowed — no hover change","All disabled states"],
      ],
      [1200, 2000, 1600, 2000, 2000, 2560]
    ),
    bullet("Button height: 40px (standard), 32px (compact), 48px (large primary CTA)"),
    bullet("Button padding: 0 24px (standard), 0 16px (compact)"),
    bullet("Button font: Calibri, 14px, font-weight 600, letter-spacing 0.02em"),
    bullet("Button border-radius: 4px (--radius-md) — do not use 0 or larger values"),
    bullet("Button text: sentence case for standard buttons; ALL CAPS reserved for primary CTA on hero sections only"),
    bullet("Icons in buttons: 16px, left of text, 8px gap — never icon-only without tooltip"),

    H2("5.2  Form Inputs"),
    makeTable(
      ["State","Border","Background","Label","Placeholder"],
      [
        ["Default",   "1px solid var(--color-mid-gray)","var(--color-white)","var(--color-dark-gray) — above field","var(--color-silver)"],
        ["Focus",     "1px solid var(--color-navy) + 2px navy outline","var(--color-white)","var(--color-navy)","hidden"],
        ["Filled",    "1px solid var(--color-navy)","var(--color-white)","var(--color-navy)","N/A — has value"],
        ["Error",     "1px solid var(--color-error)","#FFF5F5","var(--color-error)","Replaced by error message below"],
        ["Disabled",  "1px solid var(--color-light-gray)","var(--color-off-white)","var(--color-silver)","var(--color-light-gray)"],
        ["Read-only", "1px solid var(--color-light-gray)","var(--color-off-white) + lock icon","var(--color-dark-gray)","N/A"],
      ],
      [1200, 2400, 1800, 1800, 2160]
    ),
    bullet("Input height: 40px — no exceptions"),
    bullet("Input font: Calibri 16px — ensures legibility at platform scale"),
    bullet("Input padding: 0 16px — horizontal; vertically centered text"),
    bullet("Wallet address inputs: Courier New 14px, with inline ENS/UD resolution indicator"),
    bullet("Hash inputs: Courier New 14px, with 64-character validation indicator"),
    bullet("All monetary inputs: right-aligned text, currency indicator on left"),

    H2("5.3  Tables"),
    makeTable(
      ["Element","Specification","Notes"],
      [
        ["Header row",   "Background: var(--color-navy); text: var(--color-white); font: Georgia 12px ALL CAPS, tracking 0.08em","No rounded corners on any header cell"],
        ["Data rows",    "Alternating: var(--color-white) / var(--color-off-white); text: var(--color-black) Calibri 14px","Row height minimum 48px for accessibility"],
        ["Hover row",    "Background: rgba(10,22,40,0.04) — subtle navy tint","Do not use gold highlight on hover"],
        ["Active/selected row","Left border: 3px solid var(--color-gold); background: rgba(201,168,76,0.06)","Gold left accent = selected state"],
        ["Borders",      "Only horizontal 1px var(--color-light-gray) between rows — no vertical cell borders","No table outline border"],
        ["Numeric columns","Right-aligned; Courier New 14px for precision values (amounts, ratios, hashes)","All reserve ratios: 4 decimal places"],
        ["Status column","Badge with semantic color background at 10% opacity, text at 100%","Status: Confirmed, Pending, Error, Processing"],
        ["Pagination",   "Georgia font; Previous/Next as text links, not icons; current page gold-underlined","Do not use icon-only pagination arrows"],
      ],
      [1600, 4400, 3360]
    ),

    H2("5.4  Cards"),
    makeTable(
      ["Card Type","Border","Padding","Header","Notes"],
      [
        ["Standard",      "1px solid var(--color-light-gray)","var(--space-6) all sides","H4 label, entity badge if applicable","Background: white (light) / #0F1E36 (dark)"],
        ["Authority",     "Top: 3px solid var(--color-gold); others: 1px mid-gray","var(--space-6)","H3, gold top accent bar","For governance documents, Capacity Vault entries"],
        ["Entity-scoped", "Top: 3px solid var(--entity-*)","var(--space-6)","H4 label + entity badge","Entity accent color top border only"],
        ["Alert / Notice","Left: 4px solid (semantic color)","var(--space-4) left, var(--space-6) others","Semantic icon + label","Used for Notice Registry items, enforcement alerts"],
        ["Transaction",   "1px solid var(--color-light-gray); left: 2px solid semantic","var(--space-5)","Status badge + timestamp","XRPL hash always visible, monospace"],
        ["Document",      "1px solid var(--color-mid-gray)","var(--space-5)","Document ID (monospace) + doc type badge","SHA-256 hash truncated (first 8 + ...); QR icon right"],
      ],
      [1600, 2400, 1600, 2200, 3560]
    ),
    WARNING_BOX("No drop shadows on cards. No border-radius above 4px on cards. Cards use border weight and color to communicate hierarchy, not elevation or softness."),

    H2("5.5  Navigation"),
    H4("Sidebar Navigation (Admin Portal, Treasury Ledger)"),
    bullet("Width: 240px — background: var(--color-navy)"),
    bullet("Logo zone: 64px height — WIBT wordmark left-aligned, gold treatment"),
    bullet("Nav items: Calibri 14px, color: var(--color-silver); active: var(--color-white) + left border 3px var(--color-gold)"),
    bullet("Section labels: Georgia 11px ALL CAPS, tracking 0.1em, var(--color-silver) at 60% opacity — no hover state"),
    bullet("Hover state: background rgba(255,255,255,0.06) — barely perceptible tint"),
    bullet("Entity switcher: bottom of sidebar, entity name + badge color indicator"),
    bullet("Collapsed state (64px): icon only, tooltip on hover, gold dot on active icon"),

    H4("Top Navigation (Public surfaces)"),
    bullet("Height: 64px — background: var(--color-navy) — gold bottom border 1px"),
    bullet("Logo left, nav items center, CTA button right"),
    bullet("Nav link font: Georgia 14px, var(--color-silver); hover: var(--color-white)"),
    bullet("Active page: var(--color-white), gold bottom underline 2px"),
    PAGE_BREAK(),
  ];
}

// ── 6. ICONOGRAPHY & VISUAL MARKS ────────────────────────────────────
function makeSection6() {
  return [
    H1("6.  ICONOGRAPHY & VISUAL MARKS", "sec6"),

    H2("6.1  Icon System"),
    makeTable(
      ["Rule","Specification"],
      [
        ["Icon library",    "Lucide React (open source) — preferred for all platform UI. Phosphor Icons (thin weight) as secondary source."],
        ["Icon size",       "16px in text / nav items; 20px in cards; 24px in headers; 32px in empty states; 48px in hero / onboarding flows"],
        ["Icon color",      "Inherits from text context. Never apply gradient fills or multiple colors to a single icon."],
        ["Icon weight",     "Stroke weight 1.5px at 24px size; 1px at 16px — do not use filled variants in UI. Filled = status only (e.g., filled circle for confirmed state)"],
        ["Icon + label",    "Always 8px gap between icon and label text. Icon left of text, vertically centered."],
        ["Prohibited",      "3D icons, skeuomorphic icons, emoji used as UI icons, icons with drop shadows, animated icons except for loading states"],
      ],
      [1800, 7560]
    ),

    H2("6.2  Status Indicators"),
    makeTable(
      ["Status","Icon","Color","Animation"],
      [
        ["Confirmed / Minted",    "Check circle (filled)",   "var(--color-success)",  "None — static"],
        ["Pending / Processing",  "Clock or spinner",        "var(--color-warning)",  "Spinner: 1.2s linear infinite rotation"],
        ["Error / Rejected",      "X circle (filled)",       "var(--color-error)",    "None — static"],
        ["Info / Broadcasting",   "Info circle",             "var(--color-info)",     "None — static"],
        ["Locked / Frozen",       "Lock (closed)",           "var(--color-silver)",   "None — static"],
        ["On-chain / Minted",     "Link or chain icon",      "var(--color-gold)",     "None — use gold to signal blockchain presence"],
        ["Arbitration Active",    "Scales / gavel icon",     "var(--entity-mr)",      "None — red signals active enforcement"],
        ["QR Verified",           "QR code icon",            "var(--color-success)",  "None — appears only when hash verified"],
      ],
      [2000, 2200, 2000, 3160]
    ),

    H2("6.3  The WIBT Wordmark & Identity Marks"),
    makeTable(
      ["Mark","Usage","Prohibited Use"],
      [
        ["WIBT Wordmark (horizontal)","Primary placement — left-aligned in nav headers, document headers, official communications","Stretching, recoloring outside approved palette, placing on busy backgrounds"],
        ["CipherNex Logotype","CipherNex platform surfaces, Chain ID references, blockchain explorer contexts","Do not use CipherNex mark to represent WIBT parent brand"],
        ["Entity Badges","Small rectangular badge — entity abbreviation (WDT, MPRS, M&R, WAC, WLE) in entity accent color on navy background","Circular badges, badges without abbreviated text, badges using non-entity colors"],
        ["Chain ID Badge","0xbdc47 or 777287 in Courier New monospace badge — gold border on navy","Do not display Chain ID in any other format"],
        ["Verified Seal","QR + hash + checkmark composite mark — appears on finalized governance documents only","Do not display Verified Seal on documents that have not completed the Document Authentication Protocol"],
      ],
      [2200, 3800, 3360]
    ),
    PAGE_BREAK(),
  ];
}

// ── 7. DATA DISPLAY ──────────────────────────────────────────────────
function makeSection7() {
  return [
    H1("7.  DATA DISPLAY STANDARDS", "sec7"),
    BODY("The WIBT platform handles financial instruments, legal records, and blockchain data. Data display is not decorative — it is evidentiary. Formatting standards exist to ensure data is unambiguous, auditable, and consistent across every surface."),

    H2("7.1  Financial Data"),
    makeTable(
      ["Data Type","Format","Font","Example"],
      [
        ["CIPR amounts",    "Comma-separated, 4 decimal places","Courier New 14px right-aligned","1,000,000.0000 CIPR"],
        ["USD amounts",     "USD prefix, 2 decimal places, comma-separated","Courier New 14px right-aligned","USD 45,000.00"],
        ["Reserve ratio",   "Always 4 decimal places, with ratio label","Courier New 14px center-aligned","1.0000 (RESERVE)"],
        ["Bond face value", "USD prefix, 0 decimal places for round, 2 for cents","Courier New 14px right-aligned","USD 100,000"],
        ["Percentages",     "2 decimal places, % suffix, no space","Calibri 14px right-aligned","62.50%"],
        ["Timestamps",      "ISO 8601: YYYY-MM-DD HH:MM:SS UTC","Calibri 14px","2026-05-12 14:23:41 UTC"],
      ],
      [2000, 2400, 2000, 2960]
    ),
    WARNING_BOX("Never display the reserve ratio with fewer than 4 decimal places. 1.0000 is materially different from 1.0 — the former is an audit-precise statement; the latter is ambiguous. Any interface displaying '1.0' for the reserve ratio is in violation."),

    H2("7.2  Blockchain Data"),
    makeTable(
      ["Data Type","Display Format","Truncation Rule","Copy Affordance"],
      [
        ["Wallet addresses (EVM)",  "0x + 40 hex chars, Courier New","First 6 + ... + last 4 in compact view","Always — clipboard icon right"],
        ["XRPL addresses",          "r + 25-34 chars, Courier New","First 8 + ... + last 6 in compact view","Always"],
        ["Transaction hashes (ETH)","0x + 64 hex chars, Courier New","First 8 + ... + last 8 in compact view","Always"],
        ["XRPL tx hashes",          "64 hex chars uppercase, Courier New","First 8 + ... + last 8 in compact view","Always"],
        ["SHA-256 doc hashes",       "64 hex chars lowercase, Courier New","First 8 + ... + last 8; full on hover","Always + QR icon"],
        ["Block numbers",            "Integer, comma-separated at 1000s, Courier New","Never truncate","Optional"],
        ["CipherNex Chain ID",       "777287 (0xbdc47) — both forms displayed","Never truncate","Optional"],
        ["Mint transaction IDs",     "Full hash, Courier New, with Verified badge if hash matches doc","Never truncate","Always"],
      ],
      [2400, 2200, 2000, 2760]
    ),

    H2("7.3  Legal & Document Data"),
    makeTable(
      ["Data Type","Format","Notes"],
      [
        ["Document ID",        "WIBT-[CAT]-[YYYYMMDD]-[SEQ] — Courier New","Category codes: THEME, GOV, FID, ARB, COM, BEN, ENF, OPS"],
        ["Capacity Vault ID",  "CV-[YYYY]-[SEQ6] — e.g., CV-2026-000001 — Courier New","Sequential, never reused or deleted"],
        ["UCC references",     "§ symbol required — e.g., UCC § 3-311, 12 U.S.C. § 411","Georgia italics for inline citations"],
        ["Case references",    "Standard legal citation format — name in italics, no bold","Do not abbreviate case names in legal instruments"],
        ["Notice status",      "Badge: DRAFT / DISPATCHED / DELIVERED / ACQUIESCED / MINTED","Entity-mr crimson for notices in active enforcement"],
        ["Award status",       "PENDING / SIGNED / FINAL / SETTLED — with date","90-day window countdown displayed in amber when < 14 days"],
        ["Form 1041 lines",    "Schedule letter + line number — e.g., Sch B, Line 2b","Always cite alongside journal entries in GAAPCLAW outputs"],
      ],
      [2000, 3200, 4160]
    ),
    PAGE_BREAK(),
  ];
}

// ── 8. MOTION & ANIMATION ────────────────────────────────────────────
function makeSection8() {
  return [
    H1("8.  MOTION & ANIMATION", "sec8"),
    BODY("Motion is used with restraint. The WIBT platform communicates precision and institutional authority — not delight mechanics. Every animation must have a functional justification. If the animation cannot be explained in terms of what information it conveys to the user, it should not exist."),

    H2("8.1  Permitted Animations"),
    makeTable(
      ["Animation","Duration","Easing","Trigger","Functional Purpose"],
      [
        ["Page section reveal",   "300ms","ease-out","Route change","Confirms navigation occurred; prevents disorientation"],
        ["Modal open/close",      "200ms","ease-in-out","User action","Confirms modal state change; spatial context"],
        ["Sidebar collapse/expand","250ms","ease-in-out","User toggle","Confirms layout state change"],
        ["Loading spinner",       "1200ms linear infinite","—","API call in progress","Communicates system is working"],
        ["Toast notification",    "200ms in, 3000ms hold, 200ms out","ease","System event","Confirms action completed without interrupting flow"],
        ["Table row highlight",   "150ms","ease","Filter/search match","Communicates which rows match criteria"],
        ["Hash verification tick","400ms","ease-out","Hash confirmed","Communicates on-chain verification completed"],
        ["Progress bar fill",     "Tied to actual progress","linear","Multi-step process","Communicates workflow advancement"],
      ],
      [2000, 1200, 1200, 1800, 3160]
    ),

    H2("8.2  Prohibited Animations"),
    bullet("Parallax scrolling on any platform surface — disorienting in a data-dense admin UI"),
    bullet("Hover animations on table rows that shift position or size (background color change only)"),
    bullet("Animated backgrounds, particle effects, or canvas animations on any primary surface"),
    bullet("Page transition animations beyond opacity fade — no slides, zooms, or flips"),
    bullet("Auto-playing animations that the user cannot pause (beyond loading spinners)"),
    bullet("Animated chart data — charts render to final state immediately; no 'bar growing' animations"),
    bullet("Skeleton loaders with 'shimmer' animation — use static skeleton shapes (opacity 0.5) instead"),
    WARNING_BOX("The WIBT admin portal is not a consumer app. Delight animations that feel at home in a food delivery or social media app are brand-degrading in a sovereign financial platform. When in doubt, do not animate."),

    H2("8.3  Transition Tokens"),
    CODE(":root {"),
    CODE("  --duration-fast:   150ms;"),
    CODE("  --duration-normal: 250ms;"),
    CODE("  --duration-slow:   400ms;"),
    CODE(""),
    CODE("  --ease-default:    ease-in-out;"),
    CODE("  --ease-in:         cubic-bezier(0.4, 0, 1, 1);"),
    CODE("  --ease-out:        cubic-bezier(0, 0, 0.2, 1);"),
    CODE("  --ease-spring:     cubic-bezier(0.34, 1.56, 0.64, 1);  /* modals only */"),
    CODE("}"),
    PAGE_BREAK(),
  ];
}

// ── 9. ENTITY-SPECIFIC VARIANTS ────────────────────────────────────────
function makeSection9() {
  return [
    H1("9.  ENTITY-SPECIFIC VARIANTS", "sec9"),
    BODY("Each subsidiary entity has a defined visual variant built on top of the base WIBT system. These variants apply exclusively to entity-scoped surfaces — entity dashboards, entity ICO pages, entity document headers. They do not override or replace the base system; they extend it."),

    H2("9.1  Variant Specification"),
    makeTable(
      ["Entity","Primary Accent","H1 Header","Card Top Border","Badge Style","Chart/Data Color"],
      [
        ["WIBT (Parent)",        "#0A1628 Navy",     "Navy bg / white text / gold divider","3px gold","WIBT  navy/gold",              "var(--color-gold)"],
        ["WDT — Tech",           "#1A56DB Blue",     "Navy bg / white text / blue divider","3px #1A56DB","WDT  navy/blue",           "var(--entity-wdt)"],
        ["MPRS — Banking",       "#0B6E4F Green",    "Navy bg / white text / green divider","3px #0B6E4F","MPRS  navy/green",         "var(--entity-mprs)"],
        ["Massey & Rosupo",      "#8B1A1A Crimson",  "Navy bg / white text / crimson divider","3px #8B1A1A","M&R   navy/crimson",    "var(--entity-mr)"],
        ["WAC — Analytics",      "#0E7490 Teal",     "Navy bg / white text / teal divider","3px #0E7490","WAC   navy/teal",           "var(--entity-wac)"],
        ["WLE — Estates",        "#6B4226 Earth",    "Navy bg / white text / earth divider","3px #6B4226","WLE   navy/earth",         "var(--entity-wle)"],
      ],
      [2000, 1400, 2600, 2000, 1600, 1760]
    ),

    H2("9.2  Entity Variant CSS Override Pattern"),
    BODY("Apply entity variants via a data attribute on the root element of the entity-scoped section, never by overriding global CSS variables:"),
    CODE("[data-entity='wdt'] {"),
    CODE("  --entity-accent: var(--entity-wdt);"),
    CODE("  --entity-accent-subtle: rgba(26, 86, 219, 0.08);"),
    CODE("}"),
    CODE("[data-entity='mprs'] {"),
    CODE("  --entity-accent: var(--entity-mprs);"),
    CODE("  --entity-accent-subtle: rgba(11, 110, 79, 0.08);"),
    CODE("}"),
    CODE("[data-entity='mr'] {"),
    CODE("  --entity-accent: var(--entity-mr);"),
    CODE("  --entity-accent-subtle: rgba(139, 26, 26, 0.08);"),
    CODE("}"),
    BODY("Then all entity-scoped components reference --entity-accent and --entity-accent-subtle rather than the entity's specific hex value. This allows the system to scale to new entities without component-level code changes."),

    H2("9.3  Massey & Rosupo — Legal UI Specific Rules"),
    BODY("The Massey & Rosupo platform has additional typographic rules due to its legal document output function:"),
    bullet("All legal document body text: Georgia 12pt (not Calibri) — matches court document conventions"),
    bullet("All citation text: Georgia 11pt, italics for case names, regular weight for statutory citations"),
    bullet("Notice status badges: always use the Massey & Rosupo crimson (--entity-mr) — never green for any notice still in active procedure"),
    bullet("Arbitration pipeline stage indicator: horizontal stepper component with stage names in ALL CAPS Georgia 11px"),
    bullet("Schedule A fee amounts: right-aligned, Courier New, dollar sign, 2 decimal places — bold at total line"),
    bullet("90-day countdown: always visible on all active award records; amber (--color-warning) when < 14 days; crimson when < 3 days"),

    H2("9.4  MPRS Capital — Banking UI Specific Rules"),
    bullet("Reserve ratio must always be visible in the platform header when in MPRS scope — not buried in a sub-page"),
    bullet("CIPR amounts always 4 decimal places with CIPR suffix in Courier New"),
    bullet("USD amounts always with USD prefix — never $, never 'dollars', always 'USD'"),
    bullet("Bond corpus total: displayed as a running total, not as a static label — shows growth with new bond entries"),
    bullet("Settlement rail indicator: always show whether a transaction is on CipherNex chain vs. XRPL vs. Polygon"),
    PAGE_BREAK(),
  ];
}

// ── 10. PRINT & DOCUMENT OUTPUT ─────────────────────────────────────────
function makeSection10() {
  return [
    H1("10.  PRINT & DOCUMENT OUTPUT", "sec10"),
    BODY("Every document output from the WIBT ecosystem — governance documents, court packages, notice letters, arbitration awards, commercial invoices — must follow the print specification below. These documents are evidentiary instruments. Their appearance must be consistent, professional, and unmistakably institutional."),

    H2("10.1  Print Color Palette"),
    makeTable(
      ["Token","Print Hex","Usage in Print"],
      [
        ["Print Navy",  "#0A1628","Section headers only — full navy fill. Use sparingly — expensive to print."],
        ["Print Gold",  "#C9A84C","Dividers, border accents, document ID badges"],
        ["Print Black", "#0D0D0D","All body text — never pure #000000 (too harsh on paper)"],
        ["Print Gray",  "#6B7280","Metadata, captions, page numbers, secondary labels"],
        ["Print White", "#FFFFFF","All backgrounds — no colored backgrounds except navy header rows"],
        ["Print Tint",  "#F7F8FA","Alternating table rows, aside boxes — very light, prints as near-white"],
      ],
      [1600, 1600, 6160]
    ),
    WARNING_BOX("Do not use entity accent colors in print documents unless the printer is confirmed to support 4-color CMYK. Default to navy/gold/black/gray for all print output."),

    H2("10.2  Document Typography (Print)"),
    makeTable(
      ["Element","Font","Size","Weight","Notes"],
      [
        ["Document title",    "Georgia","18pt","Bold","ALL CAPS, centered or left-aligned by document type"],
        ["Entity header",     "Georgia","11pt","Bold","Entity name + EIN + date — right-aligned opposite document title"],
        ["Section heading",   "Georgia","13pt","Bold","Left-aligned; gold underline rule below"],
        ["Body text",         "Calibri","11pt","Regular","Left-aligned; 1.45 line height; first-line indent 0.25in on body paragraphs"],
        ["Table headers",     "Georgia","9pt","Bold","ALL CAPS; navy fill; white text"],
        ["Table body",        "Calibri","9pt","Regular","Left-aligned; alternating row tint"],
        ["Legal citations",   "Georgia","10pt","Italic (cases) / Regular (statutes)","Standard legal citation format"],
        ["Hash values",       "Courier New","8pt","Regular","Full SHA-256 — never truncate in print documents"],
        ["Document ID",       "Courier New","9pt","Regular","Upper-right footer on every page"],
        ["Page numbers",      "Calibri","9pt","Regular","Lower-center footer: Page N of M"],
      ],
      [2000, 1400, 800, 1200, 3960]
    ),

    H2("10.3  Document Layout Specification"),
    makeTable(
      ["Property","Value"],
      [
        ["Page size",        "US Letter (8.5\" × 11\") — default for all domestic documents"],
        ["Margins",          "Top: 1.0\" | Bottom: 1.0\" | Left: 1.25\" | Right: 1.0\""],
        ["Header zone",      "0.5\" from top — entity name left, document ID right, gold rule below"],
        ["Footer zone",      "0.5\" from bottom — Capacity Vault ID left, page number center, date right"],
        ["Content width",    "6.25\" (8.5\" - 1.25\" - 1.0\")"],
        ["Body text width",  "6.25\" — do not add further column constraints on single-column documents"],
        ["Two-column (allowed)","Left: 4.25\" | Right: 1.75\" — for definition/term layouts only"],
        ["Table full-width", "6.25\" — always span full content width"],
        ["Section spacing",  "18pt before H2; 12pt after H2; 6pt after body paragraphs"],
        ["QR code placement","Lower-right of document cover or first page — 1\" × 1\" minimum"],
        ["SHA-256 hash",     "Small, Courier New, below QR code — 8pt; links to verification URL"],
      ],
      [2800, 6560]
    ),
    PAGE_BREAK(),
  ];
}

// ── 11. ACCESSIBILITY ────────────────────────────────────────────────
function makeSection11() {
  return [
    H1("11.  ACCESSIBILITY STANDARDS", "sec11"),
    BODY("WIBT platform surfaces must meet WCAG 2.1 AA compliance as a minimum. Accessibility is not an optional enhancement — it is a compliance requirement for any platform handling financial instruments and legal processes."),

    H2("11.1  Color Contrast Requirements"),
    makeTable(
      ["Pair","Ratio","WCAG Level","Pass"],
      [
        ["Navy text (#0A1628) on White (#FFFFFF)",  "18.1:1", "AAA", "PASS"],
        ["Gold (#C9A84C) on Navy (#0A1628)",        "4.8:1",  "AA",  "PASS"],
        ["White (#FFFFFF) on Navy (#0A1628)",        "18.1:1", "AAA", "PASS"],
        ["Dark Gray (#4A5568) on White (#FFFFFF)",   "6.4:1",  "AA",  "PASS"],
        ["Silver (#8E9BB5) on Navy (#0A1628)",       "4.6:1",  "AA",  "PASS — body text only, 16px+"],
        ["WDT Blue (#1A56DB) on White (#FFFFFF)",   "4.7:1",  "AA",  "PASS — large text / UI only"],
        ["Crimson (#8B1A1A) on White (#FFFFFF)",    "7.2:1",  "AA",  "PASS"],
        ["Silver (#8E9BB5) on White (#FFFFFF)",      "2.9:1",  "FAIL","FAIL — do not use silver text on white"],
      ],
      [3800, 1200, 1400, 2960]
    ),
    WARNING_BOX("Silver (#8E9BB5) on white fails contrast. Use Dark Gray (#4A5568) for body text on light backgrounds. Silver is permitted only on navy backgrounds and for metadata/caption roles where it remains above 4.5:1."),

    H2("11.2  Additional Accessibility Rules"),
    bullet("All interactive elements: minimum 44px × 44px touch target"),
    bullet("All form inputs: associated <label> elements — no placeholder-only labeling"),
    bullet("All icons: aria-label when icon-only; aria-hidden when decorative with adjacent text"),
    bullet("All data tables: <th scope> attributes; summary attribute for complex tables"),
    bullet("Focus indicators: 2px solid var(--color-gold) outline, 2px offset — never remove focus outline"),
    bullet("Color must never be the sole means of conveying information — status always includes icon + text label + color"),
    bullet("All modals: focus trap, escape key closes, focus returns to trigger element on close"),
    bullet("All hash values: role='text' with aria-label describing what the hash represents"),
    PAGE_BREAK(),
  ];
}

// ── 12. IMPLEMENTATION CHECKLIST ─────────────────────────────────────
function makeSection12() {
  return [
    H1("12.  DEVELOPER IMPLEMENTATION CHECKLIST", "sec12"),
    BODY("This checklist must be completed before any new interface or component is submitted for review. It is not a suggestion — it is a merge gate requirement. Non-compliant submissions will be returned for remediation."),

    H2("12.1  New Surface Checklist"),
    makeTable(
      ["#","Check","How to Verify"],
      [
        ["C-01","CSS custom properties declared in :root following Section 2.5 exactly","Inspect :root in DevTools; confirm all 20 tokens present"],
        ["C-02","No hardcoded hex colors in component styles","grep for # in CSS/SCSS; zero results expected (except inside :root token declarations)"],
        ["C-03","Font stack uses Georgia / Calibri / Courier New only","Inspect computed font-family on heading, body, and mono elements"],
        ["C-04","No Inter, Roboto, Arial, or system-ui applied anywhere","grep for these font names in CSS; zero results expected"],
        ["C-05","All type sizes use --text-* tokens","grep for px/rem font-size not referencing a var(); zero non-token sizes expected"],
        ["C-06","All spacing uses --space-* tokens","Inspect margin/padding properties; no magic numbers"],
        ["C-07","Border radius: 0 or --radius-sm (2px) or --radius-md (4px) only","grep for border-radius values > 4px; zero expected"],
        ["C-08","Dark mode for authenticated surfaces, light mode for public/print","Confirm background color of authenticated wrapper is --color-navy or equivalent"],
        ["C-09","Entity accent applied via data-entity attribute, not inline styles","Inspect entity-scoped elements for data-entity attribute"],
        ["C-10","No drop shadows on cards or panels","grep for box-shadow on card/panel selectors; zero expected except --shadow-sm/md on overlays"],
        ["C-11","Reserve ratio displayed at 4 decimal places where present","Manual check on any surface displaying reserve ratio"],
        ["C-12","All wallet addresses and hashes in Courier New","Inspect computed font-family on address/hash elements"],
        ["C-13","Copy-to-clipboard present on all hash/address displays","Manual functional test"],
        ["C-14","Focus indicators visible at 2px gold outline","Tab through interface; confirm all focus rings visible"],
        ["C-15","Color contrast passes WCAG AA for all text/background pairs","Run axe DevTools or similar; zero contrast failures"],
        ["C-16","No animations beyond permitted list in Section 8","Audit all CSS animations/transitions against Section 8.1 permitted list"],
        ["C-17","Document Authentication Protocol referenced in any document-output UI","Confirm SHA-256 hash display + QR code + Capacity Vault ID field present"],
        ["C-18","Entity badges use abbreviation + entity accent color on navy only","Visual inspection of all badge instances"],
      ],
      [600, 3600, 5160]
    ),

    H2("12.2  Exception Process"),
    BODY("If a design requirement cannot be met with these standards (e.g., a third-party embed that cannot be styled), the following exception process applies:"),
    numItem("Developer documents the specific deviation, the reason it is necessary, and the surface it affects"),
    numItem("Exception request submitted to WDT lead with the completed checklist noting the failed item"),
    numItem("WDT lead reviews within 48 hours and either approves with conditions or requires remediation"),
    numItem("Approved exceptions logged in the Capacity Vault under document category WIBT-THEME-EXC"),
    numItem("Exception approval does not set a precedent — each exception is evaluated independently"),
    PAGE_BREAK(),
  ];
}

// ── 13. VERSION CONTROL ──────────────────────────────────────────────
function makeSection13() {
  return [
    H1("13.  DOCUMENT CONTROL & VERSION HISTORY", "sec13"),

    H2("13.1  Amendment Procedure"),
    BODY("This document may only be amended by WDT with written approval from the WIBT Trustee. Any amendment follows the Document Authentication Protocol: new version SHA-256 hashed, minted to CipherNex chain, Capacity Vault entry updated, old version annotated as superseded."),
    bullet("Minor updates (typo corrections, clarifications that do not change standards): WDT authority, Version X.Y increment"),
    bullet("Standard updates (new component standards, new entity variants): WDT drafts, Trustee approves, Version X.0 increment"),
    bullet("Major revisions (palette change, font change, layout system overhaul): Full ecosystem impact review required, Version N.0 increment"),
    WARNING_BOX("A color token change is a major revision. Changing --color-gold or --color-navy requires testing across all 45 Treasury Ledger services, all entity surfaces, and all print outputs before the change is ratified. Do not treat palette changes as minor updates."),

    H2("13.2  Version History"),
    makeTable(
      ["Version","Date","Author","Summary","Capacity Vault ID"],
      [
        ["1.0","2026-05-12","WDT / WIBT","Initial authoritative release — covers all six entities, full component system, print specification, accessibility standards","Assigned on mint"],
      ],
      [800, 1200, 1600, 4600, 1800]
    ),

    H2("13.3  Related Documents"),
    makeTable(
      ["Document","ID","Relationship"],
      [
        ["Architecture State v2",         "WIBT_Architecture_State_v2.md",         "Technical infrastructure this theme document governs the UI layer of"],
        ["Trust Indenture & Governance",  "WIBT-GOV-20260512-001",                 "Governance framework — Document Authentication Protocol defined therein"],
        ["CIPR Whitepaper v7",            "CIPR_Whitepaper_2026_Final.docx",        "Product specification — data display standards derive from CIPR requirements"],
        ["Business Plan v1",              "WIBT_Business_Plan_v1.docx",             "Member-facing platform description — informs public surface standards"],
        ["Massey Rosupo FAA Analysis",    "MRSP_Arbitration_Analysis.docx",         "Legal UI requirements — arbitration pipeline, Schedule A display"],
        ["Value Stream Mapping",          "WIBT_Value_Stream_Mapping.docx",          "Member experience context — onboarding flow, abstraction layer requirements"],
      ],
      [2400, 3000, 3960]
    ),
  ];
}

// ── EXECUTION PAGE ────────────────────────────────────────────────────
function makeExecution() {
  return [
    PAGE_BREAK(),
    H1("EXECUTION & AUTHENTICATION"),
    BODY("This Theme Template Document is executed under the governing authority of Wisdom Ignited Business Trust and constitutes the authoritative visual and design standard for all WIBT ecosystem surfaces. All development teams, contractors, and designers are bound by these specifications upon distribution of this document."),
    ...SPACER(2),
    makeTable(
      ["Field","Entry"],
      [
        ["Document ID",               "WIBT-THEME-20260512-001"],
        ["Authorized by (WIBT Trustee)","_______________________________________________"],
        ["WDT Technical Lead",         "_______________________________________________"],
        ["Date of Execution",          "_______________________________________________"],
        ["SHA-256 Document Hash",      "_______________________________________________"],
        ["CipherNex Mint Transaction", "_______________________________________________"],
        ["Capacity Vault Record ID",   "_______________________________________________"],
        ["Verification Endpoint",      "api.wisdomignited.com/verify/{hash}"],
        ["QR Code",                    "Affixed to physical copy — links to verification endpoint"],
        ["Next Scheduled Review",      "2026-11-12 (6-month cycle)"],
      ],
      [3000, 6360]
    ),
    ...SPACER(2),
    new Paragraph({
      children:[run(
        "This document is a private internal governance instrument of Wisdom Ignited Business Trust. Distribution is restricted to WDT development staff, contracted designers, and approved subsidiary entity administrators. External distribution requires Trustee written authorization.",
        { font:"Calibri", size:18, italics:true, color:C.SILVER }
      )],
      alignment: AlignmentType.CENTER, spacing:{ before:0, after:60 },
    }),
  ];
}

// ════════════════════════════════════════════════════════════════════
// ASSEMBLE DOCUMENT
// ════════════════════════════════════════════════════════════════════
const allChildren = [
  ...makeCover(),
  ...makeSection1(),
  ...makeSection2(),
  ...makeSection3(),
  ...makeSection4(),
  ...makeSection5(),
  ...makeSection6(),
  ...makeSection7(),
  ...makeSection8(),
  ...makeSection9(),
  ...makeSection10(),
  ...makeSection11(),
  ...makeSection12(),
  ...makeSection13(),
  ...makeExecution(),
];

const doc = new Document({
  styles: {
    default: {
      document: { run: { font:"Calibri", size:22 } }
    },
    paragraphStyles: [
      { id:"Heading1", name:"Heading 1", basedOn:"Normal", next:"Normal", quickFormat:true,
        run:{ size:40, bold:true, font:"Georgia", color:C.WHITE },
        paragraph:{ spacing:{ before:480, after:240 }, outlineLevel:0,
          shading:{ fill:C.NAVY, type:ShadingType.CLEAR } } },
      { id:"Heading2", name:"Heading 2", basedOn:"Normal", next:"Normal", quickFormat:true,
        run:{ size:30, bold:true, font:"Georgia", color:C.NAVY },
        paragraph:{ spacing:{ before:400, after:160 }, outlineLevel:1 } },
      { id:"Heading3", name:"Heading 3", basedOn:"Normal", next:"Normal", quickFormat:true,
        run:{ size:24, bold:true, font:"Georgia", color:C.NAVY },
        paragraph:{ spacing:{ before:280, after:100 }, outlineLevel:2 } },
    ]
  },
  numbering: {
    config: [
      { reference:"bullets", levels:[
          { level:0, format:LevelFormat.BULLET, text:"\u2022", alignment:AlignmentType.LEFT,
            style:{ paragraph:{ indent:{ left:720, hanging:360 } } } },
          { level:1, format:LevelFormat.BULLET, text:"\u25E6", alignment:AlignmentType.LEFT,
            style:{ paragraph:{ indent:{ left:1080, hanging:360 } } } }
        ] },
      { reference:"numbers", levels:[
          { level:0, format:LevelFormat.DECIMAL, text:"%1.", alignment:AlignmentType.LEFT,
            style:{ paragraph:{ indent:{ left:720, hanging:360 } } } }
        ] },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width:12240, height:15840 },
        margin: { top:1440, right:1440, bottom:1440, left:1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          children:[run("WIBT ECOSYSTEM  ·  Brand & Theme Standards  ·  WIBT-THEME-20260512-001  ·  v1.0",
            { font:"Georgia", size:16, color:C.SILVER })],
          border:{ bottom:{ style:BorderStyle.SINGLE, size:4, color:C.GOLD } },
          spacing:{ after:80 },
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          children:[
            run("STRICT COMPLIANCE REQUIRED  ·  ", { font:"Georgia", size:16, color:C.SILVER }),
            new TextRun({ children:[PageNumber.CURRENT], font:"Georgia", size:16, color:C.SILVER }),
          ],
          border:{ top:{ style:BorderStyle.SINGLE, size:4, color:C.GOLD } },
          spacing:{ before:80 }, alignment:AlignmentType.CENTER,
        })]
      })
    },
    children: allChildren,
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/WIBT_Brand_Theme_Standards_v1.docx', buf);
  console.log('Done');
}).catch(err=>{ console.error(err); process.exit(1); });