/**
 * OYA Monthly Report Builder
 * Reads data/YYYY-MM.json files → generates reports/YYYY-MM.pptx + index.html
 *
 * Run: node build.js
 * Env: BUILD_MONTH=2026-02 (optional, rebuilds specific month only)
 */

const pptxgen  = require("pptxgenjs");
const fs       = require("fs");
const path     = require("path");

// ── ENSURE OUTPUT DIR ─────────────────────────────────────────────────────
if (!fs.existsSync("reports")) fs.mkdirSync("reports");

// ── LOAD ALL DATA FILES ───────────────────────────────────────────────────
const buildMonth = process.env.BUILD_MONTH || null;
const dataFiles  = fs.readdirSync("data")
  .filter(f => f.match(/^\d{4}-\d{2}\.json$/))
  .sort()
  .reverse(); // latest first

if (dataFiles.length === 0) {
  console.log("No data files found in data/ — nothing to build.");
  process.exit(0);
}

const allMonths = dataFiles.map(f => ({
  monthStr: f.replace(".json", ""),
  data: JSON.parse(fs.readFileSync(path.join("data", f), "utf8")),
}));

// ── DESIGN SYSTEM ─────────────────────────────────────────────────────────
const C = {
  brown:      "5D300F",
  brownMid:   "7A4520",
  brownLight: "C4956A",
  cream:      "FFF9D9",
  creamDark:  "F0E0B8",
  teal:       "1A7A6E",
  red:        "C0392B",
  amber:      "C47A1A",
  slate:      "4A5568",
  slateLight: "718096",
  white:      "FFFFFF",
  offWhite:   "FAF7F2",
  lineGray:   "D6CFC4",
  darkGray:   "2D2A26",
};

const FONT  = "Century Gothic";
const FONT2 = "Calibri";
const W = 13.3;
const H = 7.5;
const CC       = ["TZ", "KE", "UG", "SL", "NG"];
const CC_NAMES = { TZ: "Tanzania", KE: "Kenya", UG: "Uganda", SL: "Sierra Leone", NG: "Nigeria" };

// ── FORMATTERS ────────────────────────────────────────────────────────────
const fmtUsd   = v => v == null ? "—" : "$" + Number(v).toLocaleString("en-US", { maximumFractionDigits: 0 });
const fmtNum   = v => v == null ? "—" : Number(v).toLocaleString("en-US", { maximumFractionDigits: 0 });
const fmtPct   = v => v == null ? "—" : (Number(v) * 100).toFixed(1) + "%";
const fmtBps   = (curr, prior) => {
  if (curr == null || prior == null) return "—";
  const bps = Math.round((curr - prior) * 10000);
  return (bps >= 0 ? "+" : "") + bps + " bps";
};
const fmtChg   = (curr, prior) => {
  if (curr == null || prior == null || prior === 0) return "—";
  const pct = ((curr - prior) / prior) * 100;
  return (pct >= 0 ? "+" : "") + pct.toFixed(1) + "%";
};
const fmtMonth = s => {
  const [y, m] = s.split("-");
  return new Date(+y, +m - 1, 1).toLocaleString("en-US", { month: "long", year: "numeric" });
};

// ── SLIDE HELPERS ─────────────────────────────────────────────────────────
function slideHeader(slide, title, reportMonth) {
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.55, fill: { color: C.brown }, line: { color: C.brown } });
  slide.addText(title, { x: 0.35, y: 0, w: W - 0.7, h: 0.55, fontSize: 22, bold: true, color: C.cream, fontFace: FONT, valign: "middle", margin: 0 });
  slide.addShape("rect", { x: 0, y: 0.55, w: W, h: 0.04, fill: { color: C.brownLight }, line: { color: C.brownLight } });
  slide.addShape("rect", { x: 0, y: H - 0.32, w: W, h: 0.32, fill: { color: C.brown }, line: { color: C.brown } });
  slide.addText(`OYA MICROCREDIT  |  ${reportMonth}`, { x: 0.3, y: H - 0.32, w: 6, h: 0.32, fontSize: 8, color: C.creamDark, fontFace: FONT, valign: "middle", margin: 0 });
  slide.addText("CONFIDENTIAL", { x: W - 2.5, y: H - 0.32, w: 2.2, h: 0.32, fontSize: 8, color: C.creamDark, fontFace: FONT, align: "right", valign: "middle", margin: 0 });
}

function sectionLabel(slide, x, y, w, text) {
  slide.addShape("rect", { x, y, w, h: 0.32, fill: { color: C.cream }, line: { color: C.creamDark } });
  slide.addShape("rect", { x, y, w: 0.06, h: 0.32, fill: { color: C.brown }, line: { color: C.brown } });
  slide.addText(text, { x: x + 0.12, y, w: w - 0.12, h: 0.32, fontSize: 12, bold: true, color: C.brown, fontFace: FONT, valign: "middle", margin: 0 });
}

function kpiBox(slide, x, y, w, h, label, value) {
  slide.addShape("rect", { x, y, w, h, fill: { color: C.white }, line: { color: C.lineGray, pt: 0.75 }, shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.08 } });
  slide.addShape("rect", { x, y, w, h: 0.05, fill: { color: C.brown }, line: { color: C.brown } });
  slide.addText(String(value), { x, y: y + 0.1, w, h: h * 0.55, fontSize: 26, bold: true, color: C.brown, fontFace: FONT, align: "center", valign: "middle" });
  slide.addText(label, { x: x + 0.05, y: y + h * 0.6, w: w - 0.1, h: h * 0.4, fontSize: 12, color: C.slateLight, fontFace: FONT2, align: "center", valign: "middle", wrap: true });
}

function commentaryBox(slide, x, y, w, h, text) {
  slide.addShape("rect", { x, y, w, h, fill: { color: C.offWhite }, line: { color: C.lineGray, pt: 0.5 } });
  slide.addShape("rect", { x, y, w: 0.06, h, fill: { color: C.brown }, line: { color: C.brown } });
  slide.addText("KEY HIGHLIGHTS", { x: x + 0.12, y: y + 0.06, w: w - 0.2, h: 0.25, fontSize: 12, bold: true, color: C.brown, fontFace: FONT, charSpacing: 1 });
  slide.addText(text || "—", { x: x + 0.12, y: y + 0.35, w: w - 0.2, h: h - 0.45, fontSize: 12, color: C.darkGray, fontFace: FONT2, valign: "top", wrap: true });
}

function countryTable(slide, x, y, w, rows, colWidths, headers) {
  const tblData = [];
  tblData.push(headers.map(h => ({
    text: h,
    options: { fill: { color: C.brown }, color: C.cream, bold: true, fontSize: 12, fontFace: FONT, align: "center", valign: "middle" }
  })));
  rows.forEach((row, i) => {
    const isTotal = row[0]?.text === "GROUP TOTAL";
    tblData.push(row.map((cell, j) => ({
      text: String(cell.text ?? cell),
      options: {
        fill: { color: isTotal ? C.cream : (i % 2 === 0 ? C.white : C.cream) },
        color: cell.color ?? (isTotal ? C.brown : C.darkGray),
        bold: cell.bold ?? (j === 0 || isTotal),
        fontSize: 12,
        fontFace: j === 0 ? FONT : FONT2,
        align: j === 0 ? "left" : "center",
        valign: "middle",
      }
    })));
  });
  slide.addTable(tblData, { x, y, w, colW: colWidths, rowH: 0.42, border: { pt: 0.5, color: C.lineGray } });
}

function barChart(slide, x, y, w, h, labels, values, color, isPercent) {
  slide.addChart("bar", [{ name: "", labels, values }], {
    x, y, w, h,
    barDir: "col",
    chartColors: Array(labels.length).fill(color),
    chartArea: { fill: { color: C.white }, roundedCorners: false },
    plotArea: { fill: { color: C.white } },
    catAxisLabelColor: C.darkGray,
    catAxisLabelFontSize: 12,
    catAxisLabelFontBold: true,
    valAxisHidden: true,
    valGridLine: { style: "none" },
    catGridLine: { style: "none" },
    showValue: true,
    dataLabelPosition: "outEnd",
    dataLabelColor: C.darkGray,
    dataLabelFontSize: 12,
    dataLabelFontBold: true,
    dataLabelFormatCode: isPercent ? "0.0%" : "#,##0",
    showLegend: false,
    showTitle: false,
  });
}

function changeRow(slide, values, colors, y, colW, startX) {
  values.forEach((val, i) => {
    slide.addText(val, {
      x: startX + i * colW, y, w: colW, h: 0.38,
      fontSize: 12, bold: true, color: colors[i] || C.slate, fontFace: FONT2, align: "center",
    });
  });
}

// ── AUTO-GENERATE COMMENTARY FROM DATA ───────────────────────────────────
function buildCommentary(d) {
  const best  = (key, higher=true) => CC.reduce((b, cc) => { const v = d[cc]?.[key]; return (v != null && (higher ? v > (d[b]?.[key]??-Infinity) : v < (d[b]?.[key]??Infinity))) ? cc : b; }, CC[0]);
  const pctCh = (c, p) => (c && p) ? ((c - p) / p * 100).toFixed(1) : null;

  const totalApps  = CC.reduce((s, cc) => s + (d[cc]?.total_apps || 0), 0);
  const priorApps  = CC.reduce((s, cc) => s + (d[cc]?.prior_total_apps || 0), 0);
  const appsChg    = pctCh(totalApps, priorApps);
  const topAppsCC  = best("apps_per_team");
  const appsLines  = [
    `Group total: ${fmtNum(totalApps)} applications${appsChg ? ` (${appsChg >= 0 ? "+" : ""}${appsChg}% vs prior month)` : ""}.`,
    `${CC_NAMES[topAppsCC]} had the highest avg applications per team (${fmtNum(d[topAppsCC]?.apps_per_team)}).`,
  ];

  const totalApp   = CC.reduce((s, cc) => s + (d[cc]?.total_approvals || 0), 0);
  const topEffCC   = best("approval_rate");
  const effLines   = [
    `Group total approvals: ${fmtNum(totalApp)}.`,
    `${CC_NAMES[topEffCC]} had the highest approval rate (${fmtPct(d[topEffCC]?.approval_rate)}).`,
  ];

  const totalDisb  = CC.reduce((s, cc) => s + (d[cc]?.total_disb_usd || 0), 0);
  const priorDisb  = CC.reduce((s, cc) => s + (d[cc]?.prior_total_disb_usd || 0), 0);
  const disbChg    = pctCh(totalDisb, priorDisb);
  const topDisbCC  = best("total_disb_usd");
  const disbLines  = [
    `Total disbursed: ${fmtUsd(totalDisb)}${disbChg ? ` (${disbChg >= 0 ? "+" : ""}${disbChg}% vs prior month)` : ""}.`,
    `${CC_NAMES[topDisbCC]} had the highest disbursements (${fmtUsd(d[topDisbCC]?.total_disb_usd)}).`,
  ];

  const totalBook  = CC.reduce((s, cc) => s + (d[cc]?.loan_book_usd || 0), 0);
  const priorBook  = CC.reduce((s, cc) => s + (d[cc]?.prior_loan_book_usd || 0), 0);
  const bookChg    = pctCh(totalBook, priorBook);
  const topBookCC  = best("loan_book_usd");
  const bookLines  = [
    `Total loan book: ${fmtUsd(totalBook)}${bookChg ? ` (${bookChg >= 0 ? "+" : ""}${bookChg}% vs prior month)` : ""}.`,
    `${CC_NAMES[topBookCC]} has the largest loan book (${fmtUsd(d[topBookCC]?.loan_book_usd)}).`,
  ];

  const totOI   = CC.reduce((s, cc) => s + (d[cc]?.operating_income_usd || 0), 0);
  const totRev  = CC.reduce((s, cc) => s + (d[cc]?.revenue_usd || 0), 0);
  const grpMarg = totRev ? (totOI / totRev * 100).toFixed(1) + "%" : "N/M";
  const profCC  = CC.filter(cc => (d[cc]?.operating_income_usd || 0) > 0);
  const lossCC  = CC.filter(cc => (d[cc]?.operating_income_usd || 0) < 0);
  const plLines = [
    `Group revenue: ${fmtUsd(totRev)}, operating income: ${fmtUsd(totOI)} (margin: ${grpMarg}).`,
    profCC.length ? `Profitable: ${profCC.map(cc => CC_NAMES[cc]).join(", ")}.` : "",
    lossCC.length ? `Operating loss: ${lossCC.map(cc => CC_NAMES[cc]).join(", ")}.` : "",
  ].filter(Boolean);

  return {
    apps:          appsLines.join("\n"),
    efficiency:    effLines.join("\n"),
    disbursements: disbLines.join("\n"),
    loan_book:     bookLines.join("\n"),
    financials:    plLines.join("\n"),
  };
}

// ── BUILD ONE MONTH'S PPTX ────────────────────────────────────────────────
async function buildPptx(monthStr, data) {
  const prs = new pptxgen();
  prs.layout  = "LAYOUT_WIDE";
  prs.author  = "Oya Microcredit";
  prs.title   = `Oya Group Monthly Review — ${fmtMonth(monthStr)}`;

  const RM = fmtMonth(monthStr); // "February 2026"
  const d  = data; // shorthand

  // Auto-generate commentary for any missing keys
  const _auto = buildCommentary(d);
  if (!d.commentary || typeof d.commentary !== "object") d.commentary = {};
  for (const k of Object.keys(_auto)) {
    if (!d.commentary[k]) d.commentary[k] = _auto[k];
  }

  // Helper: get change color
  const chgColor = (curr, prior, invert) => {
    if (!curr || !prior) return C.slateLight;
    const up = curr > prior;
    return (up !== !!invert) ? C.red : C.teal;
  };

  // ── SLIDE 1: COVER ──────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    s.background = { color: C.brown };
    s.addShape("rect", { x: 0, y: 0, w: 0.45, h: H, fill: { color: C.cream }, line: { color: C.cream } });
    s.addShape("rect", { x: 0.45, y: 0, w: 0.06, h: H, fill: { color: C.brownLight }, line: { color: C.brownLight } });
    s.addShape("ellipse", { x: 7.5, y: -1.5, w: 7, h: 7, fill: { color: C.cream, transparency: 92 }, line: { color: C.cream, transparency: 92 } });
    s.addText("OYA", { x: 0.8, y: 1.2, w: 5, h: 1.4, fontSize: 72, bold: true, color: C.cream, fontFace: FONT, charSpacing: 8 });
    s.addText("MICROCREDIT", { x: 0.8, y: 2.5, w: 7, h: 0.6, fontSize: 22, color: C.brownLight, fontFace: FONT, charSpacing: 6 });
    s.addShape("rect", { x: 0.8, y: 3.3, w: 6, h: 0.04, fill: { color: C.brownLight }, line: { color: C.brownLight } });
    s.addText("MONTHLY PERFORMANCE REVIEW", { x: 0.8, y: 3.5, w: 8, h: 0.5, fontSize: 16, bold: true, color: C.cream, fontFace: FONT, charSpacing: 3 });
    s.addText(RM.toUpperCase(), { x: 0.8, y: 4.1, w: 6, h: 0.5, fontSize: 20, bold: true, color: C.brownLight, fontFace: FONT });
    s.addText(`Prepared: ${d.preparedDate || new Date().toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" })}`,
      { x: 0.8, y: 4.7, w: 6, h: 0.35, fontSize: 12, color: C.creamDark, fontFace: FONT2 });
    s.addText("Tanzania  ·  Kenya  ·  Uganda  ·  Sierra Leone  ·  Nigeria",
      { x: 2, y: H - 0.8, w: W - 2.5, h: 0.4, fontSize: 12, color: C.creamDark, fontFace: FONT2, align: "right" });
  }

  // ── SLIDE 2: LOAN APPLICATIONS ──────────────────────────────────────────
  {
    const s = prs.addSlide();
    slideHeader(s, "Loan Applications", RM);
    sectionLabel(s, 0, 0.59, 8.5, `Average Number of Applications Per Team – ${RM}`);

    const appsVals  = CC.map(cc => d[cc]?.apps_per_team || 0);
    const appsChgs  = CC.map(cc => fmtChg(d[cc]?.apps_per_team, d[cc]?.prior_apps_per_team));
    const appsColors = CC.map(cc => chgColor(d[cc]?.apps_per_team, d[cc]?.prior_apps_per_team, false));

    barChart(s, 0.15, 0.92, 8.3, 5.1, CC.map(c => CC_NAMES[c]), appsVals, C.slate, false);
    sectionLabel(s, 0, 6.1, 8.5, "% Change from Prior Month");
    changeRow(s, appsChgs, appsColors, 6.48, 8.3 / 5, 0.15);

    const totalApps  = CC.reduce((s, cc) => s + (d[cc]?.total_apps || 0), 0);
    const totalTeams = CC.reduce((s, cc) => s + (d[cc]?.num_teams || 0), 0);
    const groupAppsTeam = totalTeams ? Math.round(totalApps / totalTeams) : 0;
    const priorTotalApps  = CC.reduce((s, cc) => s + (d[cc]?.prior_total_apps || 0), 0);
    const priorTotalTeams = CC.reduce((s, cc) => s + (d[cc]?.prior_num_teams || 0), 0);
    const priorGroupAppsTeam = priorTotalTeams
      ? Math.round(priorTotalApps / priorTotalTeams)
      : (() => { const t = priorTotalApps;
                 const w = CC.reduce((s,cc)=>s+(d[cc]?.prior_apps_per_team||0)*(d[cc]?.prior_total_apps||0),0);
                 return t ? Math.round(w/t) : 0; })();

    kpiBox(s, 8.7, 0.75, 4.45, 1.1, "Total Applications", fmtNum(totalApps));
    kpiBox(s, 8.7, 1.95, 4.45, 1.1, "Group Avg per Team", fmtNum(groupAppsTeam));
    kpiBox(s, 8.7, 3.15, 4.45, 1.1, "Change vs Prior Month", fmtChg(groupAppsTeam, priorGroupAppsTeam));
    commentaryBox(s, 8.7, 4.35, 4.45, 2.82, d.commentary?.apps);
  }

  // ── SLIDE 3: ASSESSMENT EFFICIENCY ──────────────────────────────────────
  {
    const s = prs.addSlide();
    slideHeader(s, "Assessment Efficiency", RM);
    sectionLabel(s, 0, 0.59, 8.5, `Loan Approval Rate (%) Per Country – ${RM}`);

    const effVals   = CC.map(cc => d[cc]?.approval_rate || 0);
    const effBps    = CC.map(cc => fmtBps(d[cc]?.approval_rate, d[cc]?.prior_approval_rate));
    const effColors = CC.map(cc => chgColor(d[cc]?.approval_rate, d[cc]?.prior_approval_rate, false));

    barChart(s, 0.15, 0.92, 8.3, 5.1, CC.map(c => CC_NAMES[c]), effVals, C.teal, true);
    sectionLabel(s, 0, 6.1, 8.5, "Change from Prior Month (bps)");
    changeRow(s, effBps, effColors, 6.48, 8.3 / 5, 0.15);

    const totalApprovals = CC.reduce((s, cc) => s + (d[cc]?.total_approvals || 0), 0);
    const totalApps      = CC.reduce((s, cc) => s + (d[cc]?.total_apps || 0), 0);
    const groupEff       = totalApps ? totalApprovals / totalApps : 0;
    const priorTotalApprovals = CC.reduce((s, cc) => s + (d[cc]?.prior_total_approvals || 0), 0);
    const priorTotalAppsEff   = CC.reduce((s, cc) => s + (d[cc]?.prior_total_apps || 0), 0);
    const priorGroupEff = priorTotalApprovals && priorTotalAppsEff
      ? priorTotalApprovals / priorTotalAppsEff
      : (() => { const t = priorTotalAppsEff;
                 const w = CC.reduce((s,cc)=>s+(d[cc]?.prior_approval_rate||0)*(d[cc]?.prior_total_apps||0),0);
                 return t ? w/t : 0; })();

    kpiBox(s, 8.7, 0.75, 4.45, 1.1, "Total Approvals", fmtNum(totalApprovals));
    kpiBox(s, 8.7, 1.95, 4.45, 1.1, "Group Approval Rate", fmtPct(groupEff));
    kpiBox(s, 8.7, 3.15, 4.45, 1.1, "Change from Prior Month", fmtBps(groupEff, priorGroupEff));
    commentaryBox(s, 8.7, 4.35, 4.45, 2.82, d.commentary?.efficiency);
  }

  // ── SLIDE 4: DISBURSEMENTS ───────────────────────────────────────────────
  {
    const s = prs.addSlide();
    slideHeader(s, "Disbursement Trend", RM);
    sectionLabel(s, 0, 0.59, 8.5, `Total Disbursements Per Country (USD) – ${RM}`);

    const disbVals   = CC.map(cc => d[cc]?.total_disb_usd || 0);
    const disbChgs   = CC.map(cc => fmtChg(d[cc]?.total_disb_usd, d[cc]?.prior_total_disb_usd));
    const disbColors = CC.map(cc => chgColor(d[cc]?.total_disb_usd, d[cc]?.prior_total_disb_usd, false));

    barChart(s, 0.15, 0.92, 8.3, 5.1, CC.map(c => CC_NAMES[c]), disbVals, C.slate, false);
    sectionLabel(s, 0, 6.1, 8.5, "% Change from Prior Month");
    changeRow(s, disbChgs, disbColors, 6.48, 8.3 / 5, 0.15);

    const totalDisb      = CC.reduce((s, cc) => s + (d[cc]?.total_disb_usd || 0), 0);
    const priorTotalDisb = CC.reduce((s, cc) => s + (d[cc]?.prior_total_disb_usd || 0), 0);
    const totalTeamsDisb = CC.reduce((s, cc) => s + (d[cc]?.num_teams || 0), 0);
    const priorTotalTeamsDisb = CC.reduce((s, cc) => s + (d[cc]?.prior_num_teams || 0), 0);
    const grpDisbPerTeam      = totalTeamsDisb      ? totalDisb / totalTeamsDisb : 0;
    const priorGrpDisbPerTeam = priorTotalTeamsDisb ? priorTotalDisb / priorTotalTeamsDisb : 0;

    kpiBox(s, 8.7, 0.75, 4.45, 1.15, "Total Disbursed (USD)", fmtUsd(totalDisb));
    kpiBox(s, 8.7, 2.0,  4.45, 1.15, "Group Disb per Team", fmtUsd(grpDisbPerTeam));
    kpiBox(s, 8.7, 3.25, 4.45, 0.8, "Change vs Prior Month", fmtChg(grpDisbPerTeam, priorGrpDisbPerTeam));

    // Teams per country grid
    sectionLabel(s, 8.7, 4.15, 4.45, "Number of Teams");
    CC.forEach((cc, i) => {
      const col = i % 3, row = Math.floor(i / 3);
      const x = 8.7 + col * 1.48, y = 4.52 + row * 0.72;
      s.addShape("rect", { x, y, w: 1.35, h: 0.62, fill: { color: C.cream }, line: { color: C.lineGray, pt: 0.5 } });
      s.addText(CC_NAMES[cc], { x, y, w: 1.35, h: 0.28, fontSize: 12, color: C.slateLight, fontFace: FONT2, align: "center", valign: "bottom" });
      s.addText(fmtNum(d[cc]?.num_teams), { x, y: y + 0.28, w: 1.35, h: 0.34, fontSize: 12, bold: true, color: C.brown, fontFace: FONT, align: "center", valign: "middle" });
    });

    commentaryBox(s, 8.7, 6.1, 4.45, 1.07, d.commentary?.disbursements);
  }

  // ── SLIDE 5: AVG LOAN SIZE & DISB PER TEAM ──────────────────────────────
  {
    const s = prs.addSlide();
    slideHeader(s, "Disbursement Stats", RM);
    sectionLabel(s, 0, 0.59, 6.55, `Average Loan Size (USD) – ${RM}`);
    sectionLabel(s, 6.75, 0.59, 6.4, `Avg Disbursements Per Team – ${RM}`);

    const loanVals   = CC.map(cc => d[cc]?.avg_loan_size_usd || 0);
    const loanChgs   = CC.map(cc => fmtChg(d[cc]?.avg_loan_size_usd, d[cc]?.prior_avg_loan_size_usd));
    const loanColors = CC.map(cc => chgColor(d[cc]?.avg_loan_size_usd, d[cc]?.prior_avg_loan_size_usd, false));

    const disbTeamVals   = CC.map(cc => d[cc]?.disb_per_team || 0);
    const disbTeamChgs   = CC.map(cc => fmtChg(d[cc]?.disb_per_team, d[cc]?.prior_disb_per_team));
    const disbTeamColors = CC.map(cc => chgColor(d[cc]?.disb_per_team, d[cc]?.prior_disb_per_team, false));

    barChart(s, 0.15, 0.92, 6.35, 5.1, CC.map(c => CC_NAMES[c]), loanVals, C.slate, false);
    barChart(s, 6.75, 0.92, 6.35, 5.1, CC.map(c => CC_NAMES[c]), disbTeamVals, C.teal, false);

    sectionLabel(s, 0, 6.1, 6.55, "% Change from Prior Month");
    changeRow(s, loanChgs, loanColors, 6.48, 6.35 / 5, 0.15);

    sectionLabel(s, 6.75, 6.1, 6.4, "% Change from Prior Month");
    changeRow(s, disbTeamChgs, disbTeamColors, 6.48, 6.35 / 5, 6.75);
  }

  // ── SLIDE 6: LOAN BOOK ───────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    slideHeader(s, "Loan Book Size", RM);
    sectionLabel(s, 0, 0.59, 8.5, `Loan Book – P+I Outstanding (≤60 Days Overdue) – ${RM}`);

    const bookVals   = CC.map(cc => d[cc]?.loan_book_usd || 0);
    const bookChgs   = CC.map(cc => fmtChg(d[cc]?.loan_book_usd, d[cc]?.prior_loan_book_usd));
    const bookColors = CC.map(cc => chgColor(d[cc]?.loan_book_usd, d[cc]?.prior_loan_book_usd, false));

    barChart(s, 0.15, 0.92, 8.3, 5.1, CC.map(c => CC_NAMES[c]), bookVals, C.slate, false);
    sectionLabel(s, 0, 6.1, 8.5, "% Change from Prior Month");
    changeRow(s, bookChgs, bookColors, 6.48, 8.3 / 5, 0.15);

    const totalBook      = CC.reduce((s, cc) => s + (d[cc]?.loan_book_usd || 0), 0);
    const priorTotalBook = CC.reduce((s, cc) => s + (d[cc]?.prior_loan_book_usd || 0), 0);

    kpiBox(s, 8.7, 0.75, 4.45, 1.1, "Total Loan Book (USD)", fmtUsd(totalBook));
    kpiBox(s, 8.7, 1.95, 4.45, 1.1, "Change vs Prior Month", fmtChg(totalBook, priorTotalBook));
    commentaryBox(s, 8.7, 3.15, 4.45, 4.02, d.commentary?.loan_book);
  }

  // ── SLIDE 7: CHRONIC RATES ───────────────────────────────────────────────
  {
    const s = prs.addSlide();
    slideHeader(s, "Default Rates", RM);
    sectionLabel(s, 0, 0.59, W, `Year to Date Chronic Rates (P+I) – ${RM}`);

    const chronicRows = CC.map(cc => {
      const curr  = d[cc]?.chronic_rate;
      const prior = d[cc]?.prior_chronic_rate;
      const bps   = fmtBps(curr, prior);
      const bpsColor = (curr != null && prior != null)
        ? (curr > prior ? C.red : C.teal) : C.slateLight;
      return [
        { text: CC_NAMES[cc] },
        { text: fmtPct(prior) },
        { text: fmtPct(curr) },
        { text: bps, color: bpsColor },
      ];
    });

    // Group total — weighted average by loan book
    const totChronicBook      = CC.reduce((s, cc) => s + (d[cc]?.loan_book_usd || 0), 0);
    const totPriorChronicBook = CC.reduce((s, cc) => s + (d[cc]?.prior_loan_book_usd || 0), 0);
    const grpChronic      = totChronicBook      ? CC.reduce((s, cc) => s + (d[cc]?.chronic_rate || 0) * (d[cc]?.loan_book_usd || 0), 0) / totChronicBook : null;
    const grpPriorChronic = totPriorChronicBook ? CC.reduce((s, cc) => s + (d[cc]?.prior_chronic_rate || 0) * (d[cc]?.prior_loan_book_usd || 0), 0) / totPriorChronicBook : null;
    const grpChronicBps = fmtBps(grpChronic, grpPriorChronic);
    const grpChronicBpsColor = (grpChronic != null && grpPriorChronic != null) ? (grpChronic > grpPriorChronic ? C.red : C.teal) : C.slateLight;
    chronicRows.push([
      { text: "GROUP TOTAL", bold: true },
      { text: fmtPct(grpPriorChronic), bold: true },
      { text: fmtPct(grpChronic), bold: true },
      { text: grpChronicBps, bold: true, color: grpChronicBpsColor },
    ]);

    countryTable(s, 0.15, 0.95, 13.0, chronicRows,
      [3.5, 3.0, 3.0, 3.5],
      ["Country", "Chronic Rate – Prior Month", "Chronic Rate – Month Under Review", "Change (bps)"]
    );
  }

  // ── SLIDE 8: PAR 30 ──────────────────────────────────────────────────────
  {
    const s = prs.addSlide();
    slideHeader(s, "Default Rates", RM);
    sectionLabel(s, 0, 0.59, W, `PAR 30 – ${RM}`);

    const par30Rows = CC.map(cc => {
      const curr  = d[cc]?.par30_rate;
      const prior = d[cc]?.prior_par30_rate;
      const bps   = fmtBps(curr, prior);
      const bpsColor = (curr != null && prior != null)
        ? (curr > prior ? C.red : C.teal) : C.slateLight;
      return [
        { text: CC_NAMES[cc] },
        { text: fmtPct(prior) },
        { text: fmtPct(curr) },
        { text: bps, color: bpsColor },
      ];
    });

    // Group total — weighted average by loan book
    const totParBook      = CC.reduce((s, cc) => s + (d[cc]?.loan_book_usd || 0), 0);
    const totPriorParBook = CC.reduce((s, cc) => s + (d[cc]?.prior_loan_book_usd || 0), 0);
    const grpPar      = totParBook      ? CC.reduce((s, cc) => s + (d[cc]?.par30_rate || 0) * (d[cc]?.loan_book_usd || 0), 0) / totParBook : null;
    const grpPriorPar = totPriorParBook ? CC.reduce((s, cc) => s + (d[cc]?.prior_par30_rate || 0) * (d[cc]?.prior_loan_book_usd || 0), 0) / totPriorParBook : null;
    const grpParBps = fmtBps(grpPar, grpPriorPar);
    const grpParBpsColor = (grpPar != null && grpPriorPar != null) ? (grpPar > grpPriorPar ? C.red : C.teal) : C.slateLight;
    par30Rows.push([
      { text: "GROUP TOTAL", bold: true },
      { text: fmtPct(grpPriorPar), bold: true },
      { text: fmtPct(grpPar), bold: true },
      { text: grpParBps, bold: true, color: grpParBpsColor },
    ]);

    countryTable(s, 0.15, 0.95, 13.0, par30Rows,
      [3.5, 3.0, 3.0, 3.5],
      ["Country", "PAR 30 – Prior Month", "PAR 30 – Month Under Review", "Change (bps)"]
    );
  }

  // ── SLIDE 9: FINANCIAL SUMMARY ───────────────────────────────────────────
  {
    const s = prs.addSlide();
    slideHeader(s, "Summary Financial Results", RM);
    sectionLabel(s, 0, 0.59, W, `Income & Expenses (USD) – ${RM}`);

    const headers = ["Country", "Revenue", "Provision", "Opex", "Operating Income", "Op. Margin", "Prior Month OI"];
    const colW    = [2.4, 1.8, 1.8, 1.8, 1.8, 1.35, 2.05];

    const plRows = CC.map(cc => {
      const cd = d[cc] || {};
      const margin = cd.revenue_usd ? cd.operating_income_usd / cd.revenue_usd : null;
      const marginStr = margin != null ? (margin * 100).toFixed(1) + "%" : "N/M";
      const marginNeg = margin != null && margin < 0;
      const priorOINeg = (cd.prior_operating_income_usd || 0) < 0;
      return [
        { text: CC_NAMES[cc] },
        { text: fmtUsd(cd.revenue_usd) },
        { text: fmtUsd(cd.provision_usd) },
        { text: fmtUsd(cd.opex_usd) },
        { text: fmtUsd(cd.operating_income_usd), color: cd.operating_income_usd < 0 ? C.red : C.darkGray },
        { text: marginStr, color: marginNeg ? C.red : C.darkGray },
        { text: fmtUsd(cd.prior_operating_income_usd), color: priorOINeg ? C.red : C.darkGray },
      ];
    });

    // Group totals
    const totRev  = CC.reduce((s, cc) => s + (d[cc]?.revenue_usd || 0), 0);
    const totProv = CC.reduce((s, cc) => s + (d[cc]?.provision_usd || 0), 0);
    const totOpex = CC.reduce((s, cc) => s + (d[cc]?.opex_usd || 0), 0);
    const totOI   = CC.reduce((s, cc) => s + (d[cc]?.operating_income_usd || 0), 0);
    const totPrOI = CC.reduce((s, cc) => s + (d[cc]?.prior_operating_income_usd || 0), 0);
    const grpMargNum = totRev ? totOI / totRev : null;
    const grpMarg = grpMargNum != null ? (grpMargNum * 100).toFixed(1) + "%" : "N/M";

    plRows.push([
      { text: "GROUP TOTAL", bold: true },
      { text: fmtUsd(totRev), bold: true },
      { text: fmtUsd(totProv), bold: true },
      { text: fmtUsd(totOpex), bold: true },
      { text: fmtUsd(totOI), bold: true, color: totOI < 0 ? C.red : C.brown },
      { text: grpMarg, bold: true, color: (grpMargNum != null && grpMargNum < 0) ? C.red : C.brown },
      { text: fmtUsd(totPrOI), bold: true, color: totPrOI < 0 ? C.red : C.brown },
    ]);

    const tblData = [];
    tblData.push(headers.map(h => ({
      text: h,
      options: { fill: { color: C.brown }, color: C.cream, bold: true, fontSize: 12, fontFace: FONT, align: "center", valign: "middle" }
    })));
    plRows.forEach((row, i) => {
      const isTotal = i === plRows.length - 1;
      tblData.push(row.map((cell, j) => ({
        text: String(cell.text ?? cell),
        options: {
          fill: { color: isTotal ? C.cream : (i % 2 === 0 ? C.white : C.offWhite) },
          color: cell.color ?? (isTotal ? C.brown : C.darkGray),
          bold: cell.bold ?? (j === 0 || isTotal),
          fontSize: isTotal ? 12 : 12,
          fontFace: j === 0 ? FONT : FONT2,
          align: j === 0 ? "left" : "center",
          valign: "middle",
        }
      })));
    });

    s.addTable(tblData, { x: 0.15, y: 0.95, w: 13.0, colW, rowH: 0.5, border: { pt: 0.5, color: C.lineGray } });
    s.addText("Revenue = Interest Income + Processing Fees  |  Opex includes Interest Expense  |  Figures in USD",
      { x: 0.15, y: 4.78, w: 13.0, h: 0.28, fontSize: 12, color: C.slateLight, fontFace: FONT2, italic: true, align: "center" });

    kpiBox(s, 0.15, 5.12, 4.2, 1.88, "Group Revenue", fmtUsd(totRev));
    kpiBox(s, 4.55, 5.12, 4.2, 1.88, "Group Operating Income", fmtUsd(totOI));
    kpiBox(s, 8.95, 5.12, 4.2, 1.88, "Group Margin", grpMarg);
  }

  // ── SLIDES 10–12: OPEX BREAKDOWN ────────────────────────────────────────
  const opexSlides = [
    { title: "Operating Expenses – Staff Cost",          key: "staff_cost_usd",    priorKey: "prior_staff_cost_usd",   color: C.slate },
    { title: "Operating Expenses – Fuel Cost",           key: "fuel_cost_usd",     priorKey: "prior_fuel_cost_usd",    color: C.amber },
    { title: "Operating Expenses – Vehicle Maintenance", key: "vehicle_cost_usd",  priorKey: "prior_vehicle_cost_usd", color: C.red   },
  ];

  for (const { title, key, priorKey, color } of opexSlides) {
    const s = prs.addSlide();
    slideHeader(s, title, RM);
    sectionLabel(s, 0, 0.59, W, `Average ${title.split("–")[1].trim()} Per Team by Country (USD) – ${RM}`);

    const teams = CC.map(cc => d[cc]?.num_teams || 1);
    const vals  = CC.map((cc, i) => teams[i] ? Math.round(Math.abs(d[cc]?.[key] || 0) / teams[i]) : 0);
    const priorVals = CC.map((cc, i) => teams[i] ? Math.round(Math.abs(d[cc]?.[priorKey] || 0) / teams[i]) : 0);
    const chgs  = CC.map((cc, i) => fmtChg(vals[i], priorVals[i]));
    const chgColors = CC.map((cc, i) => chgColor(vals[i], priorVals[i], true)); // inverted: up = bad

    barChart(s, 0.15, 0.92, 13.0, 5.2, CC.map(c => CC_NAMES[c]), vals, color, false);
    sectionLabel(s, 0, 6.17, W, "% Change from Prior Month");
    changeRow(s, chgs, chgColors, 6.5, 13.0 / 5, 0.15);
  }

  // ── WRITE FILE ───────────────────────────────────────────────────────────
  const outPath = path.join("reports", `${monthStr}.pptx`);
  await prs.writeFile({ fileName: outPath });
  console.log(`  ✓ ${outPath}`);
  return outPath;
}

// ── BUILD HTML REPORT ─────────────────────────────────────────────────────
function buildHtml(allMonths) {
  const CC_NAMES_FULL = { TZ: "Tanzania", KE: "Kenya", UG: "Uganda", SL: "Sierra Leone", NG: "Nigeria" };
  const fmtP = v => v == null ? "—" : (Number(v) * 100).toFixed(1) + "%";
  const fmtU = v => v == null ? "—" : "$" + Number(v).toLocaleString("en-US", { maximumFractionDigits: 0 });
  const fmtN = v => v == null ? "—" : Number(v).toLocaleString("en-US", { maximumFractionDigits: 0 });
  const fmtB = (c, p) => { if (c == null || p == null || p === 0) return "—"; const b = Math.round((c - p) * 10000); return (b >= 0 ? "+" : "") + b + " bps"; };
  const fmtC = (c, p) => { if (c == null || p == null || p === 0) return "—"; const x = ((c - p) / p * 100); return (x >= 0 ? "+" : "") + x.toFixed(1) + "%"; };
  // pos = green (improvement), neg = red (deterioration). inv=true means lower is better (PAR, chronic)
  const bpsClass = (c, p, inv) => { if (c == null || p == null || p === 0) return ""; const better = inv ? (c < p) : (c > p); return better ? "pos" : "neg"; };
  const chgClass = (c, p, inv) => { if (!c || !p) return ""; const better = inv ? (c < p) : (c > p); return better ? "pos" : "neg"; };
  // Arrow indicator: ↑ or ↓, colored by whether the direction is good or bad
  const fmtArrow = (c, p, inv) => { if (c == null || p == null || p === 0 || c === p) return "—"; return c > p ? "↑" : "↓"; };

  const monthSections = allMonths.map(({ monthStr, data: d }) => {
    const label = fmtMonth(monthStr);
    const totRev  = CC.reduce((s, cc) => s + (d[cc]?.revenue_usd || 0), 0);
    const totOI   = CC.reduce((s, cc) => s + (d[cc]?.operating_income_usd || 0), 0);
    const totBook = CC.reduce((s, cc) => s + (d[cc]?.loan_book_usd || 0), 0);
    const totDisb = CC.reduce((s, cc) => s + (d[cc]?.total_disb_usd || 0), 0);

    // Country rows for each table
    const appRows = CC.map(cc => `
      <tr>
        <td class="country">${CC_NAMES_FULL[cc]}</td>
        <td>${fmtN(d[cc]?.apps_per_team)}</td>
        <td class="${chgClass(d[cc]?.apps_per_team, d[cc]?.prior_apps_per_team, false)}">${fmtC(d[cc]?.apps_per_team, d[cc]?.prior_apps_per_team)}</td>
        <td>${fmtN(d[cc]?.total_apps)}</td>
        <td>${fmtN(d[cc]?.total_approvals)}</td>
        <td>${fmtP(d[cc]?.approval_rate)}</td>
        <td class="${bpsClass(d[cc]?.approval_rate, d[cc]?.prior_approval_rate, false)}">${fmtArrow(d[cc]?.approval_rate, d[cc]?.prior_approval_rate, false)}</td>
      </tr>`).join("");

    const disbRows = CC.map(cc => `
      <tr>
        <td class="country">${CC_NAMES_FULL[cc]}</td>
        <td>${fmtU(d[cc]?.total_disb_usd)}</td>
        <td class="${chgClass(d[cc]?.total_disb_usd, d[cc]?.prior_total_disb_usd, false)}">${fmtC(d[cc]?.total_disb_usd, d[cc]?.prior_total_disb_usd)}</td>
        <td>${fmtU(d[cc]?.avg_loan_size_usd)}</td>
        <td>${fmtN(d[cc]?.disb_per_team)}</td>
        <td>${fmtU(d[cc]?.loan_book_usd)}</td>
        <td class="${chgClass(d[cc]?.loan_book_usd, d[cc]?.prior_loan_book_usd, false)}">${fmtC(d[cc]?.loan_book_usd, d[cc]?.prior_loan_book_usd)}</td>
      </tr>`).join("");

    const defRows = CC.map(cc => `
      <tr>
        <td class="country">${CC_NAMES_FULL[cc]}</td>
        <td>${fmtP(d[cc]?.chronic_rate)}</td>
        <td class="${bpsClass(d[cc]?.chronic_rate, d[cc]?.prior_chronic_rate, true)}">${fmtArrow(d[cc]?.chronic_rate, d[cc]?.prior_chronic_rate, true)}</td>
        <td>${fmtP(d[cc]?.par30_rate)}</td>
        <td class="${bpsClass(d[cc]?.par30_rate, d[cc]?.prior_par30_rate, true)}">${fmtArrow(d[cc]?.par30_rate, d[cc]?.prior_par30_rate, true)}</td>
      </tr>`).join("");

    const plRows = CC.map(cc => {
      const cd = d[cc] || {};
      const margin = cd.revenue_usd ? (cd.operating_income_usd / cd.revenue_usd * 100).toFixed(1) + "%" : "N/M";
      const oiClass = cd.operating_income_usd < 0 ? "neg" : "";
      return `
      <tr>
        <td class="country">${CC_NAMES_FULL[cc]}</td>
        <td>${fmtU(cd.revenue_usd)}</td>
        <td>${fmtU(cd.provision_usd)}</td>
        <td>${fmtU(cd.opex_usd)}</td>
        <td class="${oiClass}">${fmtU(cd.operating_income_usd)}</td>
        <td>${margin}</td>
        <td>${fmtU(cd.prior_operating_income_usd)}</td>
      </tr>`;
    }).join("");
    const totPrOI = CC.reduce((s, cc) => s + (d[cc]?.prior_operating_income_usd || 0), 0);
    const totProv = CC.reduce((s, cc) => s + (d[cc]?.provision_usd || 0), 0);
    const totOpex = CC.reduce((s, cc) => s + (d[cc]?.opex_usd || 0), 0);
    const grpMarg = totRev ? (totOI / totRev * 100).toFixed(1) + "%" : "N/M";

    // Group totals for app/disb/def tables
    const grpTotalApps = CC.reduce((s, cc) => s + (d[cc]?.total_apps || 0), 0);
    const grpTotalTeams = CC.reduce((s, cc) => s + (d[cc]?.num_teams || 0), 0);
    const grpAppsTeam = grpTotalTeams ? Math.round(grpTotalApps / grpTotalTeams) : 0;
    const grpPriorTotalApps = CC.reduce((s, cc) => s + (d[cc]?.prior_total_apps || 0), 0);
    const grpPriorTeams = CC.reduce((s, cc) => s + (d[cc]?.prior_num_teams || 0), 0);
    // Fallback: weighted avg of prior_apps_per_team by prior_total_apps if prior_num_teams missing
    const grpPriorAppsTeam = grpPriorTeams
      ? Math.round(grpPriorTotalApps / grpPriorTeams)
      : (() => { const t = CC.reduce((s,cc)=>s+(d[cc]?.prior_total_apps||0),0);
                 const w = CC.reduce((s,cc)=>s+(d[cc]?.prior_apps_per_team||0)*(d[cc]?.prior_total_apps||0),0);
                 return t ? Math.round(w/t) : 0; })();
    const grpTotalApprovals = CC.reduce((s, cc) => s + (d[cc]?.total_approvals || 0), 0);
    const grpPriorApprovals = CC.reduce((s, cc) => s + (d[cc]?.prior_total_approvals || 0), 0);
    const grpAppRate = grpTotalApps ? grpTotalApprovals / grpTotalApps : 0;
    // Fallback: weighted avg of prior_approval_rate by prior_total_apps if prior_total_approvals missing
    const grpPriorAppRate = grpPriorApprovals && grpPriorTotalApps
      ? grpPriorApprovals / grpPriorTotalApps
      : (() => { const t = CC.reduce((s,cc)=>s+(d[cc]?.prior_total_apps||0),0);
                 const w = CC.reduce((s,cc)=>s+(d[cc]?.prior_approval_rate||0)*(d[cc]?.prior_total_apps||0),0);
                 return t ? w/t : 0; })();

    const grpTotalLoans = CC.reduce((s, cc) => s + (d[cc]?.num_loans || 0), 0);
    const grpAvgLoanSize = grpTotalLoans ? grpTotalDisb / grpTotalLoans : 0;
    const grpPriorDisb = CC.reduce((s, cc) => s + (d[cc]?.prior_total_disb_usd || 0), 0);
    const grpDisbPerTeam = grpTotalTeams ? grpTotalDisb / grpTotalTeams : 0;
    const grpPriorDisbPerTeam = grpPriorTeams ? grpPriorDisb / grpPriorTeams : 0;

    // Weighted avg PAR/Chronic
    const grpBookWt = CC.reduce((s, cc) => s + (d[cc]?.loan_book_usd || 0), 0);
    const grpPriorBookWt = CC.reduce((s, cc) => s + (d[cc]?.prior_loan_book_usd || 0), 0);
    const grpChronic = grpBookWt ? CC.reduce((s, cc) => s + (d[cc]?.chronic_rate || 0) * (d[cc]?.loan_book_usd || 0), 0) / grpBookWt : 0;
    const grpPriorChronic = grpPriorBookWt ? CC.reduce((s, cc) => s + (d[cc]?.prior_chronic_rate || 0) * (d[cc]?.prior_loan_book_usd || 0), 0) / grpPriorBookWt : 0;
    const grpPar = grpBookWt ? CC.reduce((s, cc) => s + (d[cc]?.par30_rate || 0) * (d[cc]?.loan_book_usd || 0), 0) / grpBookWt : 0;
    const grpPriorPar = grpPriorBookWt ? CC.reduce((s, cc) => s + (d[cc]?.prior_par30_rate || 0) * (d[cc]?.prior_loan_book_usd || 0), 0) / grpPriorBookWt : 0;

    const appTotalRow = `<tr class="total-row">
      <td>Group Total</td>
      <td>${fmtN(grpAppsTeam)}</td>
      <td class="${chgClass(grpAppsTeam, grpPriorAppsTeam, false)}">${fmtC(grpAppsTeam, grpPriorAppsTeam)}</td>
      <td>${fmtN(grpTotalApps)}</td>
      <td>${fmtN(grpTotalApprovals)}</td>
      <td>${fmtP(grpAppRate)}</td>
      <td class="${bpsClass(grpAppRate, grpPriorAppRate, false)}">${fmtArrow(grpAppRate, grpPriorAppRate, false)}</td>
    </tr>`;

    const disbTotalRow = `<tr class="total-row">
      <td>Group Total</td>
      <td>${fmtU(grpTotalDisb)}</td>
      <td class="${chgClass(grpTotalDisb, grpPriorDisb, false)}">${fmtC(grpTotalDisb, grpPriorDisb)}</td>
      <td>${grpAvgLoanSize ? fmtU(grpAvgLoanSize) : "—"}</td>
      <td>${fmtU(grpDisbPerTeam)}</td>
      <td>${fmtU(totBook)}</td>
      <td class="${chgClass(totBook, grpPriorBookWt, false)}">${fmtC(totBook, grpPriorBookWt)}</td>
    </tr>`;

    const defTotalRow = `<tr class="total-row">
      <td>Group Total</td>
      <td>${fmtP(grpChronic)}</td>
      <td class="${bpsClass(grpChronic, grpPriorChronic, true)}">${fmtArrow(grpChronic, grpPriorChronic, true)}</td>
      <td>${fmtP(grpPar)}</td>
      <td class="${bpsClass(grpPar, grpPriorPar, true)}">${fmtArrow(grpPar, grpPriorPar, true)}</td>
    </tr>`;

    return `
    <div class="month-page" id="page-${monthStr}">
      <div class="month-header">
        <div class="month-title-wrap">
          <span class="month-dot"></span>
          <h2 class="month-title">${label}</h2>
        </div>
        <div class="dl-btns">
          <a class="dl-btn" href="reports/${monthStr}.pptx" download>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Download PPTX
          </a>
          <a class="dl-btn pdf" href="reports/${monthStr}.pdf" download>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Download PDF
          </a>
        </div>
      </div>

      <div class="kpi-strip">
        <div class="kpi-card"><div class="kpi-val">${fmtU(totDisb)}</div><div class="kpi-lbl">Total Disbursed</div></div>
        <div class="kpi-card"><div class="kpi-val">${fmtU(totBook)}</div><div class="kpi-lbl">Total Loan Book</div></div>
        <div class="kpi-card"><div class="kpi-val">${fmtU(totRev)}</div><div class="kpi-lbl">Group Revenue</div></div>
        <div class="kpi-card"><div class="kpi-val ${totOI < 0 ? "neg" : ""}">${fmtU(totOI)}</div><div class="kpi-lbl">Group Operating Income</div></div>
        <div class="kpi-card"><div class="kpi-val">${grpMarg}</div><div class="kpi-lbl">Group Margin</div></div>
      </div>

      <div class="table-block">
        <div class="tbl-title">Loan Applications & Assessment Efficiency</div>
        <div class="tbl-wrap"><table>
          <thead><tr><th>Country</th><th>Apps / Team</th><th>vs Prior Month</th><th>Applications</th><th>Approvals</th><th>Approval Rate</th><th>Trend</th></tr></thead>
          <tbody>${appRows}${appTotalRow}</tbody>
        </table></div>
      </div>

      <div class="table-block">
        <div class="tbl-title">Disbursements & Loan Book</div>
        <div class="tbl-wrap"><table>
          <thead><tr><th>Country</th><th>Total Disbursed</th><th>vs Prior</th><th>Avg Loan Size</th><th>Disb / Team</th><th>Loan Book</th><th>vs Prior</th></tr></thead>
          <tbody>${disbRows}${disbTotalRow}</tbody>
        </table></div>
      </div>

      <div class="table-block">
        <div class="tbl-title">Default Rates</div>
        <div class="tbl-wrap"><table>
          <thead><tr><th>Country</th><th>Chronic Rate</th><th>Trend</th><th>PAR 30</th><th>Trend</th></tr></thead>
          <tbody>${defRows}${defTotalRow}</tbody>
        </table></div>
      </div>

      <div class="table-block">
        <div class="tbl-title">Financial Summary (USD)</div>
        <div class="tbl-wrap"><table>
          <thead><tr><th>Country</th><th>Revenue</th><th>Provision</th><th>Opex</th><th>Operating Income</th><th>Margin</th><th>Prior Month OI</th></tr></thead>
          <tbody>
            ${plRows}
            <tr class="total-row">
              <td>Group Total</td>
              <td>${fmtU(totRev)}</td>
              <td>${fmtU(totProv)}</td>
              <td>${fmtU(totOpex)}</td>
              <td class="${totOI < 0 ? "neg" : ""}">${fmtU(totOI)}</td>
              <td class="${(totRev && totOI/totRev < 0) ? "neg" : ""}">${grpMarg}</td>
              <td class="${totPrOI < 0 ? "neg" : ""}">${fmtU(totPrOI)}</td>
            </tr>
          </tbody>
        </table></div>
      </div>
    </div>`;
  }).join("\n");

  const latestMonth = allMonths[0]?.monthStr || "";

  const navItems = allMonths.map(({ monthStr }) =>
    `<button class="nav-month" data-month="${monthStr}" onclick="showMonth('${monthStr}')">${fmtMonth(monthStr)}</button>`
  ).join("");

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>OYA Microcredit — Monthly Performance Reports</title>
  <link rel="preconnect" href="https://fonts.googleapis.com"/>
  <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet"/>
  <style>
    :root {
      --brown:       #5D300F;
      --brown-mid:   #7A4520;
      --brown-lt:    #C4956A;
      --cream:       #FFF9D9;
      --cream-dk:    #F0E0B8;
      --bg:          #FAF7F2;
      --surface:     #FFFFFF;
      --border:      #E4D9C8;
      --text:        #2D2A26;
      --text-muted:  #7A7060;
      --pos:         #1E8A3E;
      --neg:         #C0392B;
      --sidebar-w:   220px;
      --header-h:    56px;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'DM Sans', sans-serif;
      background: var(--bg);
      color: var(--text);
      font-size: 14px;
      line-height: 1.6;
      overflow-x: hidden;
    }

    /* ── TOP HEADER ── */
    .site-header {
      background: var(--brown);
      height: var(--header-h);
      display: flex;
      align-items: center;
      padding: 0 1.25rem;
      position: fixed;
      top: 0; left: 0; right: 0;
      z-index: 200;
      box-shadow: 0 2px 12px rgba(93,48,15,0.3);
      gap: 0.75rem;
    }
    .hamburger {
      display: none;
      background: none;
      border: none;
      cursor: pointer;
      padding: 6px;
      color: var(--cream);
      flex-shrink: 0;
    }
    .hamburger svg { display: block; }
    .logo {
      font-family: 'Playfair Display', serif;
      font-size: 1.35rem;
      color: var(--cream);
      letter-spacing: 0.04em;
      flex: 1;
    }
    .logo span { color: var(--brown-lt); font-weight: 400; font-size: 0.82rem; margin-left: 0.5rem; font-family: 'DM Sans', sans-serif; }
    .header-tag {
      font-size: 0.68rem;
      color: var(--cream-dk);
      letter-spacing: 0.1em;
      text-transform: uppercase;
      display: none;
    }

    /* ── LAYOUT ── */
    .app-body {
      display: flex;
      padding-top: var(--header-h);
      min-height: 100vh;
    }

    /* ── SIDEBAR ── */
    .sidebar {
      width: var(--sidebar-w);
      background: var(--surface);
      border-right: 1px solid var(--border);
      position: fixed;
      top: var(--header-h);
      left: 0;
      bottom: 0;
      overflow-y: auto;
      z-index: 150;
      display: flex;
      flex-direction: column;
      transition: transform 0.25s ease;
    }
    .sidebar-label {
      padding: 1rem 1.1rem 0.5rem;
      font-size: 0.65rem;
      font-weight: 700;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.1em;
    }
    .nav-month {
      display: block;
      width: 100%;
      text-align: left;
      background: none;
      border: none;
      cursor: pointer;
      padding: 0.65rem 1.1rem;
      font-family: 'DM Sans', sans-serif;
      font-size: 0.88rem;
      font-weight: 500;
      color: var(--text-muted);
      border-left: 3px solid transparent;
      transition: all 0.15s;
      white-space: nowrap;
    }
    .nav-month:hover { color: var(--brown); background: var(--cream); border-left-color: var(--brown-lt); }
    .nav-month.active { color: var(--brown); background: var(--cream); border-left-color: var(--brown); font-weight: 700; }
    .sidebar-footer {
      margin-top: auto;
      padding: 1rem 1.1rem;
      font-size: 0.68rem;
      color: var(--text-muted);
      border-top: 1px solid var(--border);
      line-height: 1.5;
    }

    /* ── OVERLAY (mobile) ── */
    .sidebar-overlay {
      display: none;
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,0.4);
      z-index: 140;
    }
    .sidebar-overlay.open { display: block; }

    /* ── MAIN CONTENT ── */
    .main-content {
      flex: 1;
      margin-left: var(--sidebar-w);
      min-width: 0;
      padding: 2rem 2rem 4rem;
      max-width: calc(1200px + var(--sidebar-w));
    }

    /* ── MONTH PAGE ── */
    .month-page { display: none; animation: fadeUp 0.3s ease both; }
    .month-page.active { display: block; }
    @keyframes fadeUp { from { opacity: 0; transform: translateY(12px); } to { opacity: 1; transform: translateY(0); } }

    .month-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 0.75rem;
      margin-bottom: 1.5rem;
      padding-bottom: 1rem;
      border-bottom: 2px solid var(--brown);
    }
    .month-title-wrap { display: flex; align-items: center; gap: 0.75rem; }
    .month-dot { width: 10px; height: 10px; border-radius: 50%; background: var(--brown); flex-shrink: 0; }
    .month-title { font-family: 'Playfair Display', serif; font-size: 1.5rem; color: var(--brown); font-weight: 700; }

    .dl-btns { display: flex; gap: 0.5rem; flex-wrap: wrap; }
    .dl-btn {
      display: inline-flex; align-items: center; gap: 0.4rem;
      background: var(--brown); color: var(--cream);
      text-decoration: none; padding: 0.45rem 1rem;
      border-radius: 4px; font-size: 0.8rem; font-weight: 600;
      letter-spacing: 0.02em; transition: background 0.2s, transform 0.15s;
      white-space: nowrap;
    }
    .dl-btn:hover { background: var(--brown-mid); transform: translateY(-1px); }
    .dl-btn.pdf { background: var(--brown-mid); }
    .dl-btn.pdf:hover { background: var(--brown); }

    /* ── KPI STRIP ── */
    .kpi-strip {
      display: grid;
      grid-template-columns: repeat(5, 1fr);
      gap: 0.7rem;
      margin-bottom: 1.5rem;
    }
    .kpi-card {
      background: var(--surface); border: 1px solid var(--border);
      border-top: 3px solid var(--brown); padding: 0.9rem; border-radius: 3px;
    }
    .kpi-val { font-family: 'Playfair Display', serif; font-size: 1.2rem; font-weight: 700; color: var(--brown); line-height: 1.2; }
    .kpi-val.neg { color: var(--neg); }
    .kpi-lbl { font-size: 0.68rem; color: var(--text-muted); margin-top: 0.25rem; font-weight: 500; text-transform: uppercase; letter-spacing: 0.05em; }

    /* ── TABLES ── */
    .table-block { background: var(--surface); border: 1px solid var(--border); border-radius: 3px; margin-bottom: 1.1rem; overflow: hidden; }
    .tbl-title { background: var(--cream); border-bottom: 1px solid var(--border); border-left: 4px solid var(--brown); padding: 0.55rem 1rem; font-size: 0.75rem; font-weight: 700; color: var(--brown); text-transform: uppercase; letter-spacing: 0.06em; }
    .tbl-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
    table { width: 100%; border-collapse: collapse; min-width: 480px; }
    thead tr { background: var(--brown); }
    thead th { padding: 0.55rem 0.8rem; text-align: center; font-size: 0.72rem; font-weight: 600; color: var(--cream); letter-spacing: 0.04em; white-space: nowrap; }
    thead th:first-child { text-align: left; }
    tbody tr:nth-child(even) { background: var(--cream); }
    tbody tr:hover { background: var(--cream-dk); transition: background 0.15s; }
    tbody td { padding: 0.5rem 0.8rem; text-align: center; font-size: 0.8rem; color: var(--text); border-bottom: 1px solid var(--border); }
    tbody td:first-child, td.country { text-align: left; font-weight: 600; color: var(--brown); font-family: 'DM Sans', sans-serif; }
    .total-row td { background: var(--cream-dk) !important; font-weight: 700; color: var(--brown); }
    .pos { color: var(--pos); font-weight: 600; }
    .neg { color: var(--neg); font-weight: 600; }
    td.pos, td.neg { font-size: 1.15em; text-align: center; }

    /* ── EMPTY STATE ── */
    .empty-state { text-align: center; padding: 4rem 2rem; color: var(--text-muted); }
    .empty-state h3 { font-family: 'Playfair Display', serif; color: var(--brown); margin-bottom: 0.5rem; font-size: 1.2rem; }

    /* ── MOBILE ── */
    @media (max-width: 768px) {
      :root { --sidebar-w: 0px; }
      .hamburger { display: flex; }
      .header-tag { display: none; }
      .sidebar { transform: translateX(-220px); width: 220px; }
      .sidebar.open { transform: translateX(0); }
      .main-content { margin-left: 0; padding: 1.25rem 1rem 3rem; }
      .kpi-strip { grid-template-columns: repeat(2, 1fr); }
      .kpi-strip .kpi-card:last-child:nth-child(odd) { grid-column: span 2; }
      .month-header { flex-direction: column; align-items: flex-start; }
      .month-title { font-size: 1.25rem; }
    }
    @media (min-width: 769px) {
      .header-tag { display: block; }
      .sidebar { transform: none !important; }
    }
    @media (min-width: 900px) {
      .kpi-strip { grid-template-columns: repeat(5, 1fr); }
    }
  </style>
</head>
<body>
  <header class="site-header">
    <button class="hamburger" onclick="toggleSidebar()" aria-label="Toggle menu">
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
        <line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/>
      </svg>
    </button>
    <div class="logo">OYA <span>Microcredit</span></div>
    <div class="header-tag">Monthly Performance Reports</div>
  </header>

  <div class="sidebar-overlay" id="overlay" onclick="toggleSidebar()"></div>

  <div class="app-body">
    <aside class="sidebar" id="sidebar">
      <div class="sidebar-label">Reports</div>
      ${navItems}
      <div class="sidebar-footer">OYA Microcredit<br/>Confidential</div>
    </aside>

    <main class="main-content">
      ${monthSections}
      <div class="month-page active" id="page-empty" style="display:none"></div>
    </main>
  </div>

  <script>
    const LATEST = "${latestMonth}";

    function showMonth(monthStr) {
      document.querySelectorAll('.month-page').forEach(p => p.classList.remove('active'));
      document.querySelectorAll('.nav-month').forEach(b => b.classList.remove('active'));
      const page = document.getElementById('page-' + monthStr);
      if (page) { page.classList.add('active'); }
      const btn = document.querySelector('[data-month="' + monthStr + '"]');
      if (btn) { btn.classList.add('active'); }
      window.history.replaceState(null, '', '#' + monthStr);
      // Close sidebar on mobile
      if (window.innerWidth <= 768) { closeSidebar(); }
    }

    function toggleSidebar() {
      const sb = document.getElementById('sidebar');
      const ov = document.getElementById('overlay');
      sb.classList.toggle('open');
      ov.classList.toggle('open');
    }
    function closeSidebar() {
      document.getElementById('sidebar').classList.remove('open');
      document.getElementById('overlay').classList.remove('open');
    }

    // Show correct month on load
    const hash = window.location.hash.replace('#', '');
    const startMonth = (hash && document.getElementById('page-' + hash)) ? hash : LATEST;
    if (startMonth) showMonth(startMonth);
  </script>
</body>
</html>`;
}

// ── MAIN ──────────────────────────────────────────────────────────────────
(async () => {
  console.log(`Building ${allMonths.length} month(s)...`);

  const toBuild = buildMonth
    ? allMonths.filter(m => m.monthStr === buildMonth)
    : allMonths;

  for (const { monthStr, data } of toBuild) {
    console.log(`Building ${monthStr}...`);
    try {
      await buildPptx(monthStr, data);
    } catch (e) {
      console.error(`  ✗ PPTX failed for ${monthStr}: ${e.message}`);
    }
  }

  console.log("Building index.html...");
  fs.writeFileSync("index.html", buildHtml(allMonths), "utf8");
  console.log("  ✓ index.html");
  console.log("Done.");
})();
