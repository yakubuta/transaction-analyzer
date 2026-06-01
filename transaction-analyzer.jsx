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

// ── API call ───────────────────────────────────────────────────────────────
async function generateBrief(analysis, onChunk) {
  const summary = {
    period: `${analysis.period.start.toDateString()} – ${analysis.period.end.toDateString()} (${analysis.period.days} days)`,
    totalTransactions: analysis.txns.length,
    totalCredit: fmt(analysis.totalCredit),
    totalDebit: fmt(analysis.totalDebit),
    netFlow: fmt(analysis.netFlow),
    turnoverRatio: analysis.turnoverRatio.toFixed(1) + "x",
    riskScore: analysis.riskScore,
    riskLevel: riskLevel(analysis.riskScore).label,
    flaggedCount: analysis.flaggedTxns.length,
    passThroughFlag: analysis.passThroughFlag,
    highVelocityCount: analysis.velocityFlags.length,
    roundTripCount: analysis.roundTrips.length,
    roundNumberCount: analysis.roundFlags.length,
    lateNightCount: analysis.lateNight.length,
    largeOutlierCount: analysis.largeFlags.length,
    flaggedSamples: analysis.flaggedTxns.slice(0,5).map(f => ({
      flag: f.flag, date: f.date?.toDateString(), amount: fmt(f.amount), description: f.description, detail: f.detail
    })),
  };

  const prompt = `You are a senior financial intelligence analyst. Analyze the following transaction data summary and produce a structured intelligence brief.

TRANSACTION DATA:
${JSON.stringify(summary, null, 2)}

Write the brief in this EXACT structure:

CLASSIFICATION: RESTRICTED — FINANCIAL INTELLIGENCE BRIEF
REFERENCE: FIB-${Date.now().toString().slice(-6)}
DATE: ${new Date().toDateString().toUpperCase()}

BOTTOM LINE UP FRONT (BLUF):
[2-3 sentence executive summary of what this account looks like and the overall risk assessment]

KEY JUDGMENTS:
1. [Primary behavioral assessment]
2. [Transaction pattern assessment]
3. [Risk indicator summary]

INDICATORS OF CONCERN:
[For each flag type present, one paragraph explaining the specific pattern and its significance in the context of financial crime / AML / terrorism financing]

ACCOUNT BEHAVIOR PROFILE:
[2 paragraphs describing what type of business/individual this account likely belongs to based on the patterns, turnover ratio, and flow characteristics]

RECOMMENDED ACTIONS:
1. [Action 1]
2. [Action 2]
3. [Action 3]

ANALYST NOTE:
[Brief caveat about limitations of open-source financial analysis and what additional information would strengthen the assessment]

Be precise, clinical, and professional. Use intelligence community writing style. Do not use markdown formatting symbols like ** or ##.`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      stream: true,
      messages: [{ role: "user", content: prompt }],
    }),
  });

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let fullText = "";

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    const chunk = decoder.decode(value);
    const lines = chunk.split("\n").filter((l) => l.startsWith("data: "));
    for (const line of lines) {
      try {
        const data = JSON.parse(line.slice(6));
        if (data.type === "content_block_delta" && data.delta?.text) {
          fullText += data.delta.text;
          onChunk(fullText);
        }
      } catch {}
    }
  }
  return fullText;
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