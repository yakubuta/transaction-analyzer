import { useState, useCallback, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, PieChart, Pie, Cell
} from "recharts";

// ── helpers ────────────────────────────────────────────────────────────────
const fmt = (n) =>
  new Intl.NumberFormat("en-NG", { style: "currency", currency: "NGN", maximumFractionDigits: 0 }).format(n);

const fmtShort = (n) => {
  if (Math.abs(n) >= 1_000_000) return `₦${(n / 1_000_000).toFixed(1)}M`;
  if (Math.abs(n) >= 1_000) return `₦${(n / 1_000).toFixed(0)}K`;
  return `₦${n}`;
};

function detectColumns(headers) {
  // Strip empty / unnamed headers that Excel sometimes generates
  const clean = headers.filter((h) => h && String(h).trim() !== "" && !String(h).startsWith("__EMPTY"));
  const h = clean.map((x) => String(x).toLowerCase().trim());

  const find = (...terms) => {
    for (const t of terms) {
      const idx = h.findIndex((x) => x.includes(t));
      if (idx !== -1) return clean[idx];
    }
    return null;
  };

  return {
    date:        find("date", "time", "posted", "value date", "trans date"),
    credit:      find("settlement credit", "credit", "deposit", "amount in", " cr"),
    debit:       find("settlement debit", "debit", "withdrawal", "amount out", " dr"),
    amount:      find("transaction amount", "amount", "value", "sum", "trans amt"),
    description: find("narration", "description", "details", "remark", "memo", "transaction ref", "ref"),
    balance:     find("balance after", "balance before", "balance", "bal"),
    type:        find("transaction type", "type", "tran type"),
    status:      find("transaction status", "status"),
  };
}

function parseAmount(val) {
  if (val === null || val === undefined || val === "") return 0;
  const s = String(val).replace(/[₦,\s]/g, "").replace(/[()]/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function parseDate(val) {
  if (!val) return null;
  const s = String(val).trim();
  // Try various formats
  const d = new Date(s);
  if (!isNaN(d)) return d;
  // DD/MM/YYYY
  const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) {
    const year = m[3].length === 2 ? "20" + m[3] : m[3];
    return new Date(`${year}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`);
  }
  return null;
}

function analyzeTransactions(rows, cols) {
  const txns = rows
    .map((r, i) => {
      const date = parseDate(r[cols.date]);
      let credit = 0, debit = 0;
      if (cols.credit && cols.debit) {
        credit = parseAmount(r[cols.credit]);
        debit = parseAmount(r[cols.debit]);
      } else if (cols.amount) {
        const amt = parseAmount(r[cols.amount]);
        if (amt >= 0) credit = amt; else debit = Math.abs(amt);
      }
      return {
        id: i,
        date,
        dateStr: r[cols.date] || "",
        credit,
        debit,
        amount: credit || debit,
        isCredit: credit > 0,
        description: r[cols.description] || "",
        balance: parseAmount(r[cols.balance]),
        raw: r,
      };
    })
    .filter((t) => t.date && (t.credit > 0 || t.debit > 0));

  if (!txns.length) return null;

  txns.sort((a, b) => a.date - b.date);

  const totalCredit = txns.reduce((s, t) => s + t.credit, 0);
  const totalDebit = txns.reduce((s, t) => s + t.debit, 0);
  const netFlow = totalCredit - totalDebit;
  const avgCredit = totalCredit / (txns.filter((t) => t.credit > 0).length || 1);
  const avgDebit = totalDebit / (txns.filter((t) => t.debit > 0).length || 1);
  const maxBalance = Math.max(...txns.map((t) => t.balance).filter(Boolean));
  const minBalance = Math.min(...txns.map((t) => t.balance).filter(Boolean));
  const turnoverRatio = maxBalance > 0 ? totalCredit / maxBalance : 0;

  // ── Flag: high-velocity (>5 txns in 24h) ─────────────────────────────
  const velocityFlags = [];
  for (let i = 0; i < txns.length; i++) {
    const window = txns.filter(
      (t) => Math.abs(t.date - txns[i].date) <= 86400000
    );
    if (window.length >= 5 && !velocityFlags.find((f) => f.id === txns[i].id)) {
      velocityFlags.push({ ...txns[i], flag: "HIGH_VELOCITY", detail: `${window.length} txns within 24h` });
    }
  }

  // ── Flag: round-number transfers (multiples of 50k+) ─────────────────
  const roundFlags = txns.filter(
    (t) => t.amount >= 50000 && t.amount % 50000 === 0
  ).map((t) => ({ ...t, flag: "ROUND_NUMBER", detail: `${fmt(t.amount)} — exact round figure` }));

  // ── Flag: same-day round-trip (credit ≈ debit ±5%) ───────────────────
  const roundTrips = [];
  const dayMap = {};
  txns.forEach((t) => {
    const key = t.date.toISOString().split("T")[0];
    if (!dayMap[key]) dayMap[key] = { credits: [], debits: [] };
    if (t.credit > 0) dayMap[key].credits.push(t);
    else dayMap[key].debits.push(t);
  });
  Object.entries(dayMap).forEach(([day, { credits, debits }]) => {
    credits.forEach((c) => {
      debits.forEach((d) => {
        const diff = Math.abs(c.credit - d.debit) / c.credit;
        if (diff < 0.05 && c.credit > 10000) {
          roundTrips.push({
            ...c,
            flag: "ROUND_TRIP",
            detail: `In: ${fmt(c.credit)} | Out: ${fmt(d.debit)} same day (${day})`,
          });
        }
      });
    });
  });

  // ── Flag: pass-through (turnover ratio > 10x) ────────────────────────
  const passThroughFlag = turnoverRatio > 10;

  // ── Flag: late-night txns (23:00–04:00) ──────────────────────────────
  const lateNight = txns.filter((t) => {
    const h = t.date.getHours();
    return h >= 23 || h <= 4;
  }).map((t) => ({ ...t, flag: "LATE_NIGHT", detail: `Transaction at ${t.date.toTimeString().slice(0,5)}` }));

  // ── Flag: single large transaction (>5x avg) ─────────────────────────
  const largeFlags = txns.filter(
    (t) => t.amount > 5 * (t.isCredit ? avgCredit : avgDebit) && t.amount > 100000
  ).map((t) => ({ ...t, flag: "LARGE_OUTLIER", detail: `${fmt(t.amount)} — ${(t.amount / (t.isCredit ? avgCredit : avgDebit)).toFixed(1)}x avg` }));

  const allFlags = [
    ...velocityFlags,
    ...roundFlags,
    ...roundTrips,
    ...lateNight,
    ...largeFlags,
  ];

  // deduplicate by id
  const seen = new Set();
  const flaggedTxns = allFlags.filter((f) => {
    if (seen.has(`${f.id}-${f.flag}`)) return false;
    seen.add(`${f.id}-${f.flag}`);
    return true;
  });

  // ── Risk score ────────────────────────────────────────────────────────
  let riskScore = 0;
  if (velocityFlags.length > 0) riskScore += 25;
  if (roundTrips.length > 0) riskScore += 30;
  if (passThroughFlag) riskScore += 25;
  if (roundFlags.length > 3) riskScore += 10;
  if (largeFlags.length > 0) riskScore += 10;
  riskScore = Math.min(riskScore, 100);

  // ── Chart data ────────────────────────────────────────────────────────
  const monthMap = {};
  txns.forEach((t) => {
    const key = `${t.date.getFullYear()}-${String(t.date.getMonth() + 1).padStart(2,"0")}`;
    if (!monthMap[key]) monthMap[key] = { month: key, credit: 0, debit: 0 };
    monthMap[key].credit += t.credit;
    monthMap[key].debit += t.debit;
  });
  const monthlyData = Object.values(monthMap).sort((a, b) => a.month.localeCompare(b.month));

  const flagBreakdown = [
    { name: "High Velocity", value: velocityFlags.length, color: "#ef4444" },
    { name: "Round Numbers", value: roundFlags.length, color: "#f97316" },
    { name: "Round Trips", value: roundTrips.length, color: "#a855f7" },
    { name: "Late Night", value: lateNight.length, color: "#3b82f6" },
    { name: "Large Outliers", value: largeFlags.length, color: "#eab308" },
  ].filter((x) => x.value > 0);

  return {
    txns,
    totalCredit,
    totalDebit,
    netFlow,
    avgCredit,
    avgDebit,
    maxBalance,
    minBalance,
    turnoverRatio,
    passThroughFlag,
    flaggedTxns,
    velocityFlags,
    roundFlags,
    roundTrips,
    lateNight,
    largeFlags,
    riskScore,
    monthlyData,
    flagBreakdown,
    period: {
      start: txns[0].date,
      end: txns[txns.length - 1].date,
      days: Math.round((txns[txns.length - 1].date - txns[0].date) / 86400000) + 1,
    },
  };
}

const FLAG_COLORS = {
  HIGH_VELOCITY: { bg: "bg-red-900/40", border: "border-red-500/50", text: "text-red-400", label: "HIGH VELOCITY" },
  ROUND_NUMBER: { bg: "bg-orange-900/40", border: "border-orange-500/50", text: "text-orange-400", label: "ROUND NUMBER" },
  ROUND_TRIP: { bg: "bg-purple-900/40", border: "border-purple-500/50", text: "text-purple-400", label: "ROUND TRIP" },
  LATE_NIGHT: { bg: "bg-blue-900/40", border: "border-blue-500/50", text: "text-blue-400", label: "LATE NIGHT" },
  LARGE_OUTLIER: { bg: "bg-yellow-900/40", border: "border-yellow-500/50", text: "text-yellow-400", label: "LARGE OUTLIER" },
};

const RISK_COLORS = ["#22c55e","#22c55e","#eab308","#eab308","#ef4444"];
function riskLevel(score) {
  if (score < 20) return { label: "LOW", color: "#22c55e" };
  if (score < 50) return { label: "MEDIUM", color: "#eab308" };
  if (score < 75) return { label: "HIGH", color: "#f97316" };
  return { label: "CRITICAL", color: "#ef4444" };
}

// ── Local rule-based brief generator ──────────────────────────────────────
function generateBrief(analysis, onChunk) {
  return new Promise((resolve) => {
    const ref = "TPA-" + Date.now().toString().slice(-6);
    const date = new Date().toDateString().toUpperCase();
    const risk = riskLevel(analysis.riskScore);
    const period = `${analysis.period.start.toDateString()} to ${analysis.period.end.toDateString()} (${analysis.period.days} days)`;

    // BLUF
    const accountType = analysis.passThroughFlag
      ? "a high-turnover pass-through account with minimal balance retention"
      : analysis.turnoverRatio > 5
      ? "an active transactional account with elevated flow characteristics"
      : "a standard transactional account";

    const bluf = `This account exhibits characteristics consistent with ${accountType}. ` +
      `Over the assessment period of ${period}, a total of ${analysis.txns.length} transactions were recorded, ` +
      `with aggregate credits of ${fmt(analysis.totalCredit)} and aggregate debits of ${fmt(analysis.totalDebit)}. ` +
      `The composite risk score of ${analysis.riskScore}/100 places this account at ${risk.label} risk, ` +
      `with ${analysis.flaggedTxns.length} transactions flagged for further review.`;

    // Key judgments
    const j1 = analysis.passThroughFlag
      ? `The account demonstrates pass-through behavior with a turnover ratio of ${analysis.turnoverRatio.toFixed(1)}x, indicating funds are cycled through without meaningful accumulation — a pattern consistent with money mule activity, POS aggregation, or financial layering.`
      : `The account turnover ratio of ${analysis.turnoverRatio.toFixed(1)}x is within expected range for its transaction volume, suggesting standard business or personal use without clear pass-through indicators.`;

    const j2 = analysis.roundTrips.length > 0
      ? `${analysis.roundTrips.length} same-day round-trip transaction(s) were identified where inbound credits were matched within 5% by outbound debits on the same calendar date — a pattern associated with structuring, layering, or fund cycling in AML typologies.`
      : analysis.velocityFlags.length > 0
      ? `Elevated transaction velocity was detected on ${analysis.velocityFlags.length} occasion(s), with 5 or more transactions occurring within a 24-hour window — a pattern warranting review for structuring or coordinated fund movement.`
      : `No high-frequency clustering or round-trip patterns were detected during the assessment period. Transaction flow appears consistent with organic activity.`;

    const j3 = `Of the ${analysis.txns.length} total transactions reviewed, ${analysis.flaggedTxns.length} (${((analysis.flaggedTxns.length / analysis.txns.length) * 100).toFixed(1)}%) triggered one or more AML indicators. ` +
      `The concentration of flags — ${[
        analysis.velocityFlags.length > 0 ? `${analysis.velocityFlags.length} high-velocity` : "",
        analysis.roundTrips.length > 0 ? `${analysis.roundTrips.length} round-trip` : "",
        analysis.roundFlags.length > 0 ? `${analysis.roundFlags.length} round-number` : "",
        analysis.lateNight.length > 0 ? `${analysis.lateNight.length} late-night` : "",
        analysis.largeFlags.length > 0 ? `${analysis.largeFlags.length} large-outlier` : "",
      ].filter(Boolean).join(", ") || "none"} — ${analysis.riskScore >= 50 ? "collectively elevates concern and warrants escalation" : "does not individually or collectively meet the threshold for escalation at this time"}.`;

    // Indicators of concern
    let indicators = "";
    if (analysis.velocityFlags.length > 0) {
      indicators += `HIGH VELOCITY TRANSACTIONS: ${analysis.velocityFlags.length} instance(s) of 5 or more transactions within a 24-hour window were detected. Rapid sequential transactions are a recognised indicator of structuring — the deliberate breaking down of large sums into smaller transactions to avoid reporting thresholds. This pattern is also consistent with POS terminal abuse and unauthorised bulk transfers.\n\n`;
    }
    if (analysis.roundTrips.length > 0) {
      indicators += `ROUND-TRIP FLOWS: ${analysis.roundTrips.length} same-day round-trip(s) were identified. Funds received were offset by near-equivalent outflows on the same date. This is a classic layering technique used in money laundering and terrorism financing to obscure the origin and destination of funds while maintaining the appearance of legitimate activity.\n\n`;
    }
    if (analysis.roundFlags.length > 0) {
      indicators += `ROUND-NUMBER TRANSFERS: ${analysis.roundFlags.length} transaction(s) involving large round-figure amounts (multiples of ₦50,000 or above) were recorded. Round-number transactions, particularly at high frequency, are associated with structured payments, informal value transfer systems (hawala), and pre-arranged fund disbursements.\n\n`;
    }
    if (analysis.lateNight.length > 0) {
      indicators += `LATE-NIGHT ACTIVITY: ${analysis.lateNight.length} transaction(s) were executed between the hours of 23:00 and 04:00. While not individually conclusive, after-hours financial activity — particularly when combined with other indicators — may suggest evasion of real-time monitoring, use of automated transfer scripts, or coordination across time zones.\n\n`;
    }
    if (analysis.largeFlags.length > 0) {
      indicators += `LARGE OUTLIER TRANSACTIONS: ${analysis.largeFlags.length} transaction(s) exceeded 5 times the account average and surpassed ₦100,000. Sudden high-value transactions inconsistent with established account behaviour are a primary indicator for suspicious activity reporting under NFIU guidelines and FATF Recommendation 20.\n\n`;
    }
    if (!indicators) {
      indicators = "No significant individual indicators were detected during this assessment period. Risk score is driven by aggregate flow patterns rather than discrete transaction-level anomalies.\n\n";
    }

    // Account behavior profile
    const profile1 = analysis.passThroughFlag
      ? `Based on the volumetric analysis, this account exhibits behaviour most consistent with a cash-handling or payment aggregation operation — such as a POS terminal business, mobile money agent, or informal remittance node. The high turnover ratio of ${analysis.turnoverRatio.toFixed(1)}x relative to peak balance indicates that the account functions primarily as a conduit rather than a store of value.`
      : `The account demonstrates transaction patterns consistent with an active individual or small business account. Credit and debit flows show ${analysis.netFlow >= 0 ? "a net positive flow suggesting income exceeds expenditure" : "a net negative flow suggesting expenditure exceeds income"} over the assessment period, with an average credit of ${fmt(analysis.avgCredit)} and average debit of ${fmt(analysis.avgDebit)}.`;

    const profile2 = `The ${analysis.period.days}-day assessment window captured ${analysis.txns.length} transactions, averaging ${(analysis.txns.length / analysis.period.days).toFixed(1)} transactions per day. ` +
      (analysis.monthlyData.length > 1
        ? `Monthly flow analysis reveals variation across the period, suggesting activity is not uniformly distributed and may correspond to payment cycles, event-driven transfers, or seasonal business patterns.`
        : `The assessment window is limited to a single month, which constrains longitudinal trend analysis. Extended historical data would significantly strengthen behavioral profiling.`);

    // Recommended actions
    const actions = [
      analysis.riskScore >= 75
        ? "ESCALATE to the Nigerian Financial Intelligence Unit (NFIU) via a Suspicious Transaction Report (STR) in accordance with the Money Laundering (Prevention and Prohibition) Act 2022."
        : analysis.riskScore >= 50
        ? "FLAG account for enhanced due diligence (EDD) review. Request additional KYC documentation from the account holder including source of funds declaration."
        : "MONITOR account activity over next 90 days and reassess risk score with updated transaction data.",
      "OBTAIN full 12-month transaction history to establish baseline behaviour and identify longitudinal patterns not visible in the current assessment window.",
      analysis.passThroughFlag
        ? "VERIFY business registration and operational legitimacy of account holder. Confirm whether a POS or mobile money agent license is held with the Central Bank of Nigeria (CBN)."
        : "CROSS-REFERENCE flagged counterparties against NFIU watchlists, Interpol notices, and OFAC/UN sanctions databases.",
    ].filter(Boolean);

    const analystNote = `This assessment is based exclusively on quantitative transaction pattern analysis using open-source financial data. It does not constitute a definitive determination of criminal activity. The indicators identified herein are probabilistic in nature and require corroboration through additional intelligence streams including KYC records, counterparty analysis, beneficial ownership verification, and law enforcement liaison. Classification and dissemination of this brief should be handled in accordance with applicable data protection and intelligence sharing protocols.`;

    const brief = `CLASSIFICATION: RESTRICTED — FINANCIAL INTELLIGENCE BRIEF
REFERENCE: ${ref}
DATE: ${date}
PREPARED BY: Transaction Pattern Analysis System v1.0

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

BOTTOM LINE UP FRONT (BLUF):
${bluf}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

KEY JUDGMENTS:
1. ${j1}

2. ${j2}

3. ${j3}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

INDICATORS OF CONCERN:
${indicators}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ACCOUNT BEHAVIOR PROFILE:
${profile1}

${profile2}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

RECOMMENDED ACTIONS:
${actions.map((a, i) => `${i + 1}. ${a}`).join("\n")}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ANALYST NOTE:
${analystNote}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
END OF BRIEF — ${ref} | ANALYST: YAKUBU TANIMU | NAUB CRIMINOLOGY & SECURITY STUDIES`;

    // Simulate streaming by revealing text chunk by chunk
    let i = 0;
    const chunkSize = 8;
    const interval = setInterval(() => {
      i += chunkSize;
      onChunk(brief.slice(0, i));
      if (i >= brief.length) {
        clearInterval(interval);
        onChunk(brief);
        resolve(brief);
      }
    }, 10);
  });
}

// ── Main component ─────────────────────────────────────────────────────────
export default function App() {
  const [stage, setStage] = useState("upload"); // upload | preview | analysis | brief
  const [csvData, setCsvData] = useState(null);
  const [cols, setCols] = useState(null);
  const [analysis, setAnalysis] = useState(null);
  const [brief, setBrief] = useState("");
  const [briefLoading, setBriefLoading] = useState(false);
  const [error, setError] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const briefRef = useRef(null);

  const handleFile = useCallback((file) => {
    setError("");
    const name = file.name.toLowerCase();

    const processRows = (rows, fields) => {
      if (!rows.length) { setError("No data found in file."); return; }
      // Filter out __EMPTY keys SheetJS generates for blank Excel columns
      const cleanFields = fields.filter(f => f && !String(f).startsWith("__EMPTY") && String(f).trim() !== "");
      const detected = detectColumns(cleanFields);
      if (!detected.date) {
        setError(`Could not detect a date column. Columns found: ${cleanFields.slice(0,8).join(" | ")}`);
        return;
      }
      setCsvData({ data: rows, meta: { fields: cleanFields } });
      setCols(detected);
      setStage("preview");
    };

    if (name.endsWith(".csv")) {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (result) => processRows(result.data, result.meta.fields),
        error: () => setError("Failed to parse CSV."),
      });
    } else if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
          let bestSheet = wb.SheetNames[0];
          let bestCount = 0;
          wb.SheetNames.forEach((sname) => {
            const ws = wb.Sheets[sname];
            const ref = ws["!ref"];
            if (ref) {
              const range = XLSX.utils.decode_range(ref);
              if (range.e.r > bestCount) { bestCount = range.e.r; bestSheet = sname; }
            }
          });
          const ws = wb.Sheets[bestSheet];
          // First try normal header parsing
          // Smart header row finder — scans up to row 50 for the real table header
          // Strategy: the real header row must contain the word "date" as a standalone cell value
          const findHeaderRow = (ws) => {
            const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
            for (let r = 0; r <= Math.min(range.e.r, 50); r++) {
              const rowVals = [];
              let filledCells = 0;
              for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                const val = cell ? String(cell.v).trim() : "";
                rowVals.push(val);
                if (val !== "") filledCells++;
              }
              // Must have "date" as an exact or near-exact cell value (not inside "Opening Balance:")
              const hasDateCell = rowVals.some(v => /^date$/i.test(v.trim()) || /^(value\s*date|trans.*date|tran.*date|posting.*date)$/i.test(v.trim()));
              // AND must have at least 3 filled cells (rules out sparse info rows)
              if (hasDateCell && filledCells >= 3) return r;
            }
            // Fallback: look for any row with 5+ filled cells (likely the data table)
            for (let r = 0; r <= Math.min(range.e.r, 50); r++) {
              let filledCells = 0;
              const rowVals = [];
              for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                const val = cell ? String(cell.v).trim() : "";
                rowVals.push(val);
                if (val !== "") filledCells++;
              }
              const rowStr = rowVals.join(" ").toLowerCase();
              const hasFinancialKeywords = ["amount","credit","debit","narration","transaction"].filter(k => rowStr.includes(k)).length >= 2;
              if (hasFinancialKeywords && filledCells >= 4) return r;
            }
            return 0;
          };

          const headerRow = findHeaderRow(ws);
          let json = XLSX.utils.sheet_to_json(ws, { defval: "", raw: false, dateNF: "yyyy-mm-dd", range: headerRow });

          if (!json.length) { setError("Sheet appears empty or headers not found."); return; }
          const fields = Object.keys(json[0]);
          processRows(json, fields);
        } catch (err) {
          setError("Failed to parse Excel file. Ensure it is a valid .xlsx or .xls file.");
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      setError("Unsupported file type. Please upload a .csv, .xlsx, or .xls file.");
    }
  }, []);

  const handleDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  };

  const handleRun = () => {
    const result = analyzeTransactions(csvData.data, cols);
    if (!result) { setError("No valid transactions found after parsing."); return; }
    setAnalysis(result);
    setStage("analysis");
  };

  const handleGenerateBrief = async () => {
    setBriefLoading(true);
    setBrief("");
    setStage("brief");
    try {
      await generateBrief(analysis, (text) => setBrief(text));
    } catch (e) {
      setBrief("Error generating brief. Please try again.");
    }
    setBriefLoading(false);
  };

  const handlePrint = () => window.print();

  const risk = analysis ? riskLevel(analysis.riskScore) : null;

  return (
    <div className="min-h-screen bg-[#0a0c10] text-gray-100" style={{ fontFamily: "'Courier New', monospace" }}>
      {/* Print styles */}
      <style>{`
        @media print {
          .no-print { display: none !important; }
          body { background: white !important; color: black !important; }
          .print-area { background: white !important; color: black !important; }
          * { font-family: 'Courier New', monospace !important; }
        }
        .scan-line {
          position: fixed; top: 0; left: 0; right: 0;
          height: 2px; background: linear-gradient(90deg, transparent, #00ff9d, transparent);
          animation: scan 4s linear infinite; pointer-events: none; z-index: 50;
        }
        @keyframes scan { 0% { top: 0; opacity: 1; } 100% { top: 100vh; opacity: 0.3; } }
        .blink { animation: blink 1.2s step-end infinite; }
        @keyframes blink { 50% { opacity: 0; } }
        .brief-text { white-space: pre-wrap; line-height: 1.8; }
      `}</style>

      <div className="scan-line no-print" />

      {/* Header */}
      <header className="no-print border-b border-[#00ff9d]/20 px-6 py-4 flex items-center justify-between">
        <div>
          <div className="text-[#00ff9d] text-xs tracking-[0.3em] mb-1">FINANCIAL INTELLIGENCE UNIT</div>
          <div className="text-white text-xl font-bold tracking-wider">TRANSACTION PATTERN ANALYZER</div>
        </div>
        <div className="text-right text-xs text-gray-500">
          <div>STATUS: <span className="text-[#00ff9d]">ACTIVE</span></div>
          <div>{new Date().toISOString().split("T")[0]}</div>
        </div>
      </header>

      {/* Developer credit bar */}
      <div className="no-print bg-[#0d1117] border-b border-[#00ff9d]/10 px-6 py-2 flex items-center justify-between">
        <div className="text-xs text-gray-500">
          Built by <span className="text-[#00ff9d] font-bold">Yakubu Tanimu</span>
          <span className="text-gray-600 mx-2">|</span>
          <span className="text-gray-600">Criminology & Security Studies, NAUB</span>
        </div>
        <div className="flex items-center gap-4 text-xs">
          <a href="https://www.linkedin.com/in/yakubu-tanimu-723a3221a?utm_source=share&utm_campaign=share_via&utm_content=profile&utm_medium=android_app" target="_blank" rel="noreferrer"
            className="text-gray-500 hover:text-[#0077b5] transition-colors flex items-center gap-1">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><path d="M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.062 2.062 0 01-2.063-2.065 2.064 2.064 0 112.063 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0h.003z"/></svg>
            LinkedIn
          </a>
          <a href="https://github.com/yakubuta" target="_blank" rel="noreferrer"
            className="text-gray-500 hover:text-white transition-colors flex items-center gap-1">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><path d="M12 0C5.374 0 0 5.373 0 12c0 5.302 3.438 9.8 8.207 11.387.599.111.793-.261.793-.577v-2.234c-3.338.726-4.033-1.416-4.033-1.416-.546-1.387-1.333-1.756-1.333-1.756-1.089-.745.083-.729.083-.729 1.205.084 1.839 1.237 1.839 1.237 1.07 1.834 2.807 1.304 3.492.997.107-.775.418-1.305.762-1.604-2.665-.305-5.467-1.334-5.467-5.931 0-1.311.469-2.381 1.236-3.221-.124-.303-.535-1.524.117-3.176 0 0 1.008-.322 3.301 1.23A11.509 11.509 0 0112 5.803c1.02.005 2.047.138 3.006.404 2.291-1.552 3.297-1.23 3.297-1.23.653 1.653.242 2.874.118 3.176.77.84 1.235 1.911 1.235 3.221 0 4.609-2.807 5.624-5.479 5.929.43.372.823 1.102.823 2.222v3.293c0 .319.192.694.801.576C20.566 21.797 24 17.3 24 12c0-6.627-5.373-12-12-12z"/></svg>
            GitHub
          </a>
          <a href="https://www.facebook.com/share/1HUk922EmY/" target="_blank" rel="noreferrer"
            className="text-gray-500 hover:text-[#1877f2] transition-colors flex items-center gap-1">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><path d="M24 12.073c0-6.627-5.373-12-12-12s-12 5.373-12 12c0 5.994 4.388 10.954 10.125 11.854v-8.385H7.078v-3.47h3.047V9.43c0-3.007 1.792-4.669 4.533-4.669 1.312 0 2.686.235 2.686.235v2.953H15.83c-1.491 0-1.956.925-1.956 1.874v2.25h3.328l-.532 3.47h-2.796v8.385C19.612 23.027 24 18.062 24 12.073z"/></svg>
            Facebook
          </a>
        </div>
      </div>

      {/* Nav */}
      <div className="no-print flex gap-0 border-b border-[#00ff9d]/10 px-6">
        {["upload","preview","analysis","brief"].map((s, i) => (
          <div key={s} className={`px-4 py-2 text-xs tracking-widest border-b-2 transition-all ${
            stage === s ? "border-[#00ff9d] text-[#00ff9d]" : "border-transparent text-gray-600"
          }`}>
            {String(i+1).padStart(2,"0")}. {s.toUpperCase()}
          </div>
        ))}
      </div>

      <main className="p-6 max-w-6xl mx-auto">

        {/* ── UPLOAD ───────────────────────────────────────────── */}
        {stage === "upload" && (
          <div className="mt-12">
            <div className="text-center mb-8">
              <div className="text-[#00ff9d] text-xs tracking-[0.4em] mb-2">STEP 01 / INGEST</div>
              <div className="text-3xl text-white font-bold mb-2">Upload Bank Statement</div>
              <div className="text-gray-500 text-sm">Supports CSV and Excel (.xlsx/.xls) — Moniepoint, GTBank, Access, Zenith, UBA, Opay, etc.</div>
            </div>

            <div
              onDrop={handleDrop}
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onClick={() => document.getElementById("fileInput").click()}
              className={`border-2 border-dashed rounded cursor-pointer transition-all p-16 text-center ${
                dragOver ? "border-[#00ff9d] bg-[#00ff9d]/5" : "border-gray-700 hover:border-[#00ff9d]/50"
              }`}
            >
              <div className="text-5xl mb-4">⬆</div>
              <div className="text-white text-lg mb-2">Drop CSV or Excel file here</div>
              <div className="text-gray-500 text-sm">or click to browse (.csv, .xlsx, .xls)</div>
              <input id="fileInput" type="file" accept=".csv,.xlsx,.xls" className="hidden"
                onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
            </div>

            {error && <div className="mt-4 p-3 bg-red-900/30 border border-red-500/50 text-red-400 text-sm rounded">{error}</div>}

            <div className="mt-8 grid grid-cols-3 gap-4 text-xs">
              {[
                ["FLEXIBLE FORMAT","Auto-detects columns from any bank statement layout"],
                ["7 PATTERN FLAGS","Velocity, round-trips, pass-through, outliers + more"],
                ["AI BRIEF","Claude generates an analyst-grade intelligence brief"],
              ].map(([t, d]) => (
                <div key={t} className="border border-gray-800 rounded p-4">
                  <div className="text-[#00ff9d] font-bold mb-1 tracking-wider">{t}</div>
                  <div className="text-gray-400">{d}</div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── PREVIEW ──────────────────────────────────────────── */}
        {stage === "preview" && csvData && cols && (
          <div>
            <div className="text-[#00ff9d] text-xs tracking-[0.4em] mb-4 mt-4">STEP 02 / VERIFY COLUMN MAPPING</div>
            <div className="grid grid-cols-3 gap-3 mb-6">
              {Object.entries(cols).filter(([,v]) => v).map(([k, v]) => (
                <div key={k} className="border border-gray-800 rounded p-3">
                  <div className="text-gray-500 text-xs mb-1">{k.toUpperCase()}</div>
                  <div className="text-[#00ff9d] text-sm">{v}</div>
                </div>
              ))}
            </div>
            <div className="text-gray-500 text-xs mb-2">{csvData.data.length} rows detected</div>
            <div className="overflow-x-auto border border-gray-800 rounded mb-6 max-h-56 overflow-y-auto">
              <table className="text-xs w-full">
                <thead className="bg-gray-900 sticky top-0">
                  <tr>{csvData.meta.fields.slice(0,6).map(f => (
                    <th key={f} className="px-3 py-2 text-left text-gray-400 border-b border-gray-800">{f}</th>
                  ))}</tr>
                </thead>
                <tbody>
                  {csvData.data.slice(0,10).map((row, i) => (
                    <tr key={i} className="border-b border-gray-900 hover:bg-gray-900/50">
                      {csvData.meta.fields.slice(0,6).map(f => (
                        <td key={f} className="px-3 py-1.5 text-gray-300 truncate max-w-[120px]">{row[f]}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {error && <div className="mb-4 p-3 bg-red-900/30 border border-red-500/50 text-red-400 text-sm rounded">{error}</div>}
            <div className="flex gap-3">
              <button onClick={() => setStage("upload")} className="px-6 py-2 border border-gray-700 text-gray-400 text-sm rounded hover:border-gray-500 transition-all">
                ← Back
              </button>
              <button onClick={handleRun} className="px-8 py-2 bg-[#00ff9d] text-black font-bold text-sm rounded hover:bg-[#00ff9d]/90 transition-all tracking-wider">
                RUN ANALYSIS →
              </button>
            </div>
          </div>
        )}

        {/* ── ANALYSIS ─────────────────────────────────────────── */}
        {(stage === "analysis" || stage === "brief") && analysis && (
          <div>
            <div className="no-print flex items-center justify-between mb-6 mt-4">
              <div className="text-[#00ff9d] text-xs tracking-[0.4em]">STEP 03 / ANALYSIS RESULTS</div>
              <div className="flex gap-3">
                <button onClick={() => setStage("upload")} className="px-4 py-1.5 border border-gray-700 text-gray-400 text-xs rounded hover:border-gray-500">
                  New File
                </button>
                <button onClick={handlePrint} className="px-4 py-1.5 border border-[#00ff9d]/50 text-[#00ff9d] text-xs rounded hover:bg-[#00ff9d]/10">
                  Export PDF
                </button>
                {stage === "analysis" && (
                  <button onClick={handleGenerateBrief} className="px-6 py-1.5 bg-[#00ff9d] text-black font-bold text-xs rounded hover:bg-[#00ff9d]/90 tracking-wider">
                    GENERATE INTEL BRIEF →
                  </button>
                )}
              </div>
            </div>

            {/* Risk score banner */}
            <div className="border rounded p-5 mb-6 flex items-center justify-between" style={{ borderColor: risk.color + "40", background: risk.color + "10" }}>
              <div>
                <div className="text-xs tracking-widest mb-1" style={{ color: risk.color }}>COMPOSITE RISK SCORE</div>
                <div className="text-5xl font-bold" style={{ color: risk.color }}>{analysis.riskScore}</div>
                <div className="text-sm text-gray-400 mt-1">{analysis.flaggedTxns.length} flagged transactions</div>
              </div>
              <div className="text-right">
                <div className="text-4xl font-bold tracking-widest" style={{ color: risk.color }}>{risk.label}</div>
                <div className="text-xs text-gray-500 mt-1">{analysis.period.start.toDateString()} – {analysis.period.end.toDateString()}</div>
                <div className="text-xs text-gray-500">{analysis.period.days} days | {analysis.txns.length} transactions</div>
              </div>
            </div>

            {/* Stats grid */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
              {[
                ["TOTAL CREDIT", fmt(analysis.totalCredit), "#00ff9d"],
                ["TOTAL DEBIT", fmt(analysis.totalDebit), "#ef4444"],
                ["NET FLOW", fmt(analysis.netFlow), analysis.netFlow >= 0 ? "#00ff9d" : "#ef4444"],
                ["TURNOVER RATIO", analysis.turnoverRatio.toFixed(1) + "x", analysis.passThroughFlag ? "#ef4444" : "#eab308"],
              ].map(([label, value, color]) => (
                <div key={label} className="border border-gray-800 rounded p-4">
                  <div className="text-gray-500 text-xs mb-1 tracking-wider">{label}</div>
                  <div className="text-lg font-bold" style={{ color }}>{value}</div>
                </div>
              ))}
            </div>

            {/* Charts */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
              <div className="md:col-span-2 border border-gray-800 rounded p-4">
                <div className="text-xs text-gray-500 tracking-wider mb-3">MONTHLY FLOW</div>
                <ResponsiveContainer width="100%" height={160}>
                  <BarChart data={analysis.monthlyData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" />
                    <XAxis dataKey="month" tick={{ fontSize: 9, fill: "#6b7280" }} />
                    <YAxis tickFormatter={fmtShort} tick={{ fontSize: 9, fill: "#6b7280" }} />
                    <Tooltip formatter={(v) => fmt(v)} contentStyle={{ background: "#111", border: "1px solid #374151", fontSize: 11 }} />
                    <Bar dataKey="credit" fill="#00ff9d" name="Credit" />
                    <Bar dataKey="debit" fill="#ef4444" name="Debit" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div className="border border-gray-800 rounded p-4">
                <div className="text-xs text-gray-500 tracking-wider mb-3">FLAG BREAKDOWN</div>
                {analysis.flagBreakdown.length > 0 ? (
                  <ResponsiveContainer width="100%" height={160}>
                    <PieChart>
                      <Pie data={analysis.flagBreakdown} cx="50%" cy="50%" innerRadius={40} outerRadius={65} dataKey="value" label={({ name, value }) => `${value}`} labelLine={false}>
                        {analysis.flagBreakdown.map((entry, i) => (
                          <Cell key={i} fill={entry.color} />
                        ))}
                      </Pie>
                      <Tooltip contentStyle={{ background: "#111", border: "1px solid #374151", fontSize: 11 }} />
                    </PieChart>
                  </ResponsiveContainer>
                ) : (
                  <div className="flex items-center justify-center h-40 text-gray-600 text-sm">No flags detected</div>
                )}
                <div className="grid grid-cols-1 gap-1 mt-2">
                  {analysis.flagBreakdown.map((f) => (
                    <div key={f.name} className="flex items-center gap-2 text-xs">
                      <div className="w-2 h-2 rounded-full flex-shrink-0" style={{ background: f.color }} />
                      <span className="text-gray-400">{f.name}</span>
                      <span className="ml-auto font-bold" style={{ color: f.color }}>{f.value}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* Pass-through alert */}
            {analysis.passThroughFlag && (
              <div className="border border-red-500/40 bg-red-900/20 rounded p-4 mb-4 flex gap-3">
                <div className="text-red-400 text-lg">⚠</div>
                <div>
                  <div className="text-red-400 text-xs font-bold tracking-wider mb-1">PASS-THROUGH ACCOUNT INDICATOR</div>
                  <div className="text-gray-400 text-xs">Turnover ratio of {analysis.turnoverRatio.toFixed(1)}x suggests funds are being cycled through this account without accumulation — consistent with money mule, POS aggregator, or layering behavior.</div>
                </div>
              </div>
            )}

            {/* Flagged transactions */}
            {analysis.flaggedTxns.length > 0 && (
              <div className="border border-gray-800 rounded mb-6">
                <div className="px-4 py-3 border-b border-gray-800 text-xs text-gray-500 tracking-wider flex justify-between">
                  <span>FLAGGED TRANSACTIONS</span>
                  <span>{analysis.flaggedTxns.length} entries</span>
                </div>
                <div className="max-h-72 overflow-y-auto">
                  {analysis.flaggedTxns.map((f, i) => {
                    const fc = FLAG_COLORS[f.flag] || FLAG_COLORS.ROUND_NUMBER;
                    return (
                      <div key={i} className={`border-b border-gray-900 px-4 py-3 ${fc.bg} border-l-2 ${fc.border}`}>
                        <div className="flex items-start justify-between gap-4">
                          <div className="flex-1 min-w-0">
                            <div className="flex items-center gap-2 mb-1">
                              <span className={`text-[10px] font-bold tracking-wider px-1.5 py-0.5 rounded ${fc.text} border ${fc.border}`}>{fc.label}</span>
                              <span className="text-gray-500 text-xs">{f.date?.toDateString()}</span>
                            </div>
                            <div className="text-gray-300 text-xs truncate">{f.description || "—"}</div>
                            <div className={`text-xs mt-1 ${fc.text}`}>{f.detail}</div>
                          </div>
                          <div className="text-right flex-shrink-0">
                            <div className={`font-bold text-sm ${f.isCredit ? "text-[#00ff9d]" : "text-red-400"}`}>
                              {f.isCredit ? "+" : "-"}{fmt(f.amount)}
                            </div>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* ── BRIEF ─────────────────────────────────────────── */}
            {stage === "brief" && (
              <div ref={briefRef} className="border border-[#00ff9d]/30 rounded p-6 print-area">
                <div className="text-[#00ff9d] text-xs tracking-[0.4em] mb-4 no-print">INTELLIGENCE BRIEF</div>
                {briefLoading && !brief && (
                  <div className="text-[#00ff9d] text-sm">
                    Generating brief<span className="blink">_</span>
                  </div>
                )}
                {brief && (
                  <div className="brief-text text-sm text-gray-200 leading-relaxed">
                    {brief}
                    {briefLoading && <span className="text-[#00ff9d] blink">█</span>}
                  </div>
                )}
              </div>
            )}
          </div>
        )}

      </main>
    </div>
  );
}