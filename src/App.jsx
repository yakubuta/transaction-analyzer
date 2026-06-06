import React from "react";
import { useState, useCallback, useRef, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, PieChart, Pie, Cell
} from "recharts";
import { SpeedInsights } from "@vercel/speed-insights/react";

// ── helpers ────────────────────────────────────────────────────────────────
const fmt = (n) =>
  new Intl.NumberFormat("en-NG", { style: "currency", currency: "NGN", maximumFractionDigits: 0 }).format(n);

const fmtShort = (n) => {
  if (Math.abs(n) >= 1000000) return `₦${(n / 1000000).toFixed(1)}M`;
  if (Math.abs(n) >= 1000) return `₦${(n / 1000).toFixed(0)}K`;
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
    beneficiary: find("beneficiary", "beneficiary name", "receiver", "recipient", "destination"),
    source:      find("source", "source name", "sender", "originator", "source institution"),
    beneficiaryBank: find("beneficiary institution", "beneficiary bank", "dest bank", "receiving bank"),
    sourceBank:  find("source institution", "source bank", "sending bank", "originating bank"),
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

function analyzeTransactions(rows, cols, accountMeta = {}) {
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
      const beneficiary = cols.beneficiary ? (r[cols.beneficiary] || "").trim() : "";
      const source = cols.source ? (r[cols.source] || "").trim() : "";
      const beneficiaryBank = cols.beneficiaryBank ? (r[cols.beneficiaryBank] || "").trim() : "";
      const sourceBank = cols.sourceBank ? (r[cols.sourceBank] || "").trim() : "";
      return {
        id: i,
        date,
        dateStr: r[cols.date] || "",
        credit,
        debit,
        amount: credit || debit,
        isCredit: credit > 0,
        description: r[cols.description] || "",
        beneficiary,
        source,
        beneficiaryBank,
        sourceBank,
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

  // ── Flag: high-velocity ──────────────────────────────────────────────
  // Only flag DISTINCT days where >= 10 txns occurred (not every transaction in those days)
  const velocityDays = {};
  txns.forEach((t) => {
    const day = t.date.toISOString().split("T")[0];
    if (!velocityDays[day]) velocityDays[day] = [];
    velocityDays[day].push(t);
  });
  const velocityFlags = [];
  Object.entries(velocityDays).forEach(([day, dayTxns]) => {
    if (dayTxns.length >= 10) {
      // Flag only the highest-value transaction of that day as representative
      const rep = dayTxns.reduce((a, b) => b.amount > a.amount ? b : a);
      velocityFlags.push({
        ...rep,
        flag: "HIGH_VELOCITY",
        detail: `${dayTxns.length} transactions on ${day} — unusually high daily volume`,
      });
    }
  });

  // ── Flag: round-number transfers (multiples of 100k+, not just any 50k) ─
  const roundFlags = txns.filter(
    (t) => t.amount >= 100000 && t.amount % 100000 === 0
  ).map((t) => ({ ...t, flag: "ROUND_NUMBER", detail: `${fmt(t.amount)} — exact round figure` }));

  // ── Flag: same-day round-trip (credit ≈ debit ±5%, min 50k) ──────────
  const roundTrips = [];
  const dayMap = {};
  txns.forEach((t) => {
    const key = t.date.toISOString().split("T")[0];
    if (!dayMap[key]) dayMap[key] = { credits: [], debits: [] };
    if (t.credit > 0) dayMap[key].credits.push(t);
    else dayMap[key].debits.push(t);
  });
  // Only flag once per day, not every matching pair
  Object.entries(dayMap).forEach(([day, { credits, debits }]) => {
    let flagged = false;
    credits.forEach((c) => {
      if (flagged) return;
      debits.forEach((d) => {
        if (flagged) return;
        const diff = Math.abs(c.credit - d.debit) / c.credit;
        if (diff < 0.05 && c.credit >= 50000) {
          roundTrips.push({
            ...c,
            flag: "ROUND_TRIP",
            detail: `In: ${fmt(c.credit)} | Out: ${fmt(d.debit)} same day (${day})`,
          });
          flagged = true;
        }
      });
    });
  });

  // ── Flag: pass-through (turnover ratio > 10x) ────────────────────────
  const passThroughFlag = turnoverRatio > 10;

  // ── Flag: late-night txns (23:00–04:00) ────────────────���─────────────
  const lateNight = txns.filter((t) => {
    const h = t.date.getHours();
    return h >= 23 || h <= 4;
  }).map((t) => ({ ...t, flag: "LATE_NIGHT", detail: `Transaction at ${t.date.toTimeString().slice(0,5)}` }));

  // ── Flag: single large transaction (>5x avg) ─────────────────────────
  const largeFlags = txns.filter(
    (t) => t.amount > 5 * (t.isCredit ? avgCredit : avgDebit) && t.amount > 100000
  ).map((t) => ({ ...t, flag: "LARGE_OUTLIER", detail: `${fmt(t.amount)} — ${(t.amount / (t.isCredit ? avgCredit : avgDebit)).toFixed(1)}x avg` }));

  // ── Counterparty network analysis ─────────────────────────────────────
  const counterpartyMap = {};
  txns.forEach((t) => {
    // Priority: use dedicated beneficiary/source columns if available
    // For credits: counterparty is the source (who sent money)
    // For debits: counterparty is the beneficiary (who received money)
    let name = "";
    let bank = "";
    if (t.isCredit) {
      name = t.source || t.beneficiary || "";
      bank = t.sourceBank || t.beneficiaryBank || "";
    } else {
      name = t.beneficiary || t.source || "";
      bank = t.beneficiaryBank || t.sourceBank || "";
    }
    // Fallback: parse from description if no dedicated column
    if (!name) {
      const raw = (t.description || "").trim();
      if (!raw) return;
      name = raw.split(/[\/|,\-]/)[0].trim().slice(0, 40);
    }
    name = name.trim().slice(0, 50);
    if (!name || name === "" || name === "-") return;

    const key = name.toUpperCase();
    if (!counterpartyMap[key]) {
      counterpartyMap[key] = {
        name,
        bank,
        totalSent: 0,
        totalReceived: 0,
        txnCount: 0,
        firstDate: t.date,
        lastDate: t.date,
        firstAmount: t.amount,
        transactions: [],
      };
    }
    const cp = counterpartyMap[key];
    // Update bank if we now have it
    if (!cp.bank && bank) cp.bank = bank;
    if (t.credit > 0) cp.totalReceived += t.credit;
    else cp.totalSent += t.debit;
    cp.txnCount += 1;
    if (t.date < cp.firstDate) { cp.firstDate = t.date; cp.firstAmount = t.amount; }
    if (t.date > cp.lastDate) cp.lastDate = t.date;
    cp.transactions.push(t);
  });

  const counterparties = Object.values(counterpartyMap)
    .map((cp) => {
      let cpRisk = 0;
      const total = cp.totalSent + cp.totalReceived;
      // Single transaction with large amount
      if (cp.txnCount === 1 && total >= 100000) cpRisk += 30;
      // High total volume
      if (total >= 1000000) cpRisk += 20;
      else if (total >= 500000) cpRisk += 10;
      // Round number totals
      if (total >= 100000 && total % 100000 === 0) cpRisk += 15;
      // Only sends or only receives (one-directional)
      if (cp.totalSent === 0 || cp.totalReceived === 0) cpRisk += 10;
      // High frequency
      if (cp.txnCount >= 10) cpRisk += 10;
      // Keyword in name
      const suspectNames = ["unknown","cash","agent","transfer","payment","forex","crypto","bitcoin","invest"];
      if (suspectNames.some((k) => cp.name.toLowerCase().includes(k))) cpRisk += 15;
      cp.riskScore = Math.min(cpRisk, 100);
      return cp;
    })
    .sort((a, b) => b.riskScore - a.riskScore || (b.totalSent + b.totalReceived) - (a.totalSent + a.totalReceived));

  // Flag counterparties with only 1 transaction (first-time, high value)
  const firstTimeLargeFlags = counterparties
    .filter((cp) => cp.txnCount === 1 && (cp.totalSent + cp.totalReceived) >= 100000)
    .map((cp) => ({
      ...cp.transactions[0],
      flag: "FIRST_TIME_LARGE",
      detail: `First-ever transaction with ${cp.name}: ${fmt(cp.totalSent + cp.totalReceived)}`,
    }));

  // ── Structuring Detector ─────────────────────────────────────────────
  // Detects transactions deliberately kept just below round thresholds
  const THRESHOLDS = [5000000, 2000000, 1000000, 500000, 200000, 100000];
  const structuringFlags = [];
  const structuringSeen = new Set();

  // Group by counterparty
  Object.values(counterpartyMap).forEach((cp) => {
    if (cp.transactions.length < 2) return;
    const cpTxns = cp.transactions.sort((a, b) => a.date - b.date);

    // Check rolling 7-day windows
    for (let i = 0; i < cpTxns.length; i++) {
      const windowTxns = cpTxns.filter(
        (t) => Math.abs(t.date - cpTxns[i].date) <= 7 * 86400000
      );
      if (windowTxns.length < 2) continue;
      const windowTotal = windowTxns.reduce((s, t) => s + t.amount, 0);

      for (const threshold of THRESHOLDS) {
        // Total is within 5% below a threshold and each txn is below threshold
        const allBelow = windowTxns.every((t) => t.amount < threshold);
        const nearThreshold = windowTotal >= threshold * 0.85 && windowTotal < threshold * 1.05;
        if (allBelow && nearThreshold && windowTxns.length >= 2) {
          const key = `${cp.name}-${threshold}-${cpTxns[i].id}`;
          if (!structuringSeen.has(key)) {
            structuringSeen.add(key);
            structuringFlags.push({
              ...cpTxns[i],
              flag: "STRUCTURING",
              detail: `${windowTxns.length} txns totalling ${fmt(windowTotal)} to/from ${cp.name} within 7 days — near ₦${(threshold/1000).toFixed(0)}k threshold`,
            });
          }
          break;
        }
      }
    }
  });

  // Also detect same-day splits (regardless of counterparty)
  Object.entries(velocityDays).forEach(([day, dayTxns]) => {
    if (dayTxns.length < 3) return;
    const dayTotal = dayTxns.reduce((s, t) => s + t.amount, 0);
    for (const threshold of THRESHOLDS) {
      const allBelow = dayTxns.every((t) => t.amount < threshold);
      const nearThreshold = dayTotal >= threshold * 0.85 && dayTotal < threshold * 1.05;
      if (allBelow && nearThreshold) {
        const key = `day-${day}-${threshold}`;
        if (!structuringSeen.has(key)) {
          structuringSeen.add(key);
          const rep = dayTxns.reduce((a, b) => b.amount > a.amount ? b : a);
          structuringFlags.push({
            ...rep,
            flag: "STRUCTURING",
            detail: `${dayTxns.length} txns on ${day} total ${fmt(dayTotal)} — near ₦${(threshold/1000).toFixed(0)}k threshold (possible structuring)`,
          });
        }
        break;
      }
    }
  });

  // ── Dormancy & Sudden Activation Detector ────────────────────────────
  const dormancyFlags = [];
  for (let i = 1; i < txns.length; i++) {
    const gap = (txns[i].date - txns[i-1].date) / 86400000;
    if (gap >= 30) {
      // Check if activity after gap is significantly higher than before
      const before = txns.slice(Math.max(0, i-10), i);
      const after = txns.slice(i, Math.min(txns.length, i+10));
      const avgBefore = before.reduce((s,t) => s + t.amount, 0) / (before.length || 1);
      const avgAfter = after.reduce((s,t) => s + t.amount, 0) / (after.length || 1);
      if (avgAfter > avgBefore * 2 || txns[i].amount > 100000) {
        dormancyFlags.push({
          ...txns[i],
          flag: "DORMANCY_ACTIVATION",
          detail: `${Math.round(gap)}-day gap before this transaction — sudden reactivation with ${fmt(txns[i].amount)}`,
        });
      }
    }
  }

  // ── Keyword Flagging ──────────────────────────────────────────────────
  const SUSPICIOUS_KEYWORDS = [
    "loan","gift","charity","investment","urgent","crypto","bitcoin","wallet",
    "btc","usdt","forex","ponzi","scheme","refund","transfer fee","commission",
    "offshore","anonymous","cash out","withdraw all","emergency","bribe",
    "settlement","compensation","winnings","prize","lottery","inheritance",
  ];
  const keywordFlags = txns.filter((t) => {
    const desc = (t.description || "").toLowerCase();
    return SUSPICIOUS_KEYWORDS.some((k) => desc.includes(k));
  }).map((t) => {
    const desc = (t.description || "").toLowerCase();
    const matched = SUSPICIOUS_KEYWORDS.filter((k) => desc.includes(k));
    return { ...t, flag: "KEYWORD_FLAG", detail: `Suspicious keyword(s): ${matched.join(", ")}` };
  });

  const allFlags = [
    ...velocityFlags,
    ...roundFlags,
    ...roundTrips,
    ...lateNight,
    ...largeFlags,
    ...firstTimeLargeFlags,
    ...structuringFlags,
    ...dormancyFlags,
    ...keywordFlags,
  ];

  // deduplicate by id+flag
  const seen = new Set();
  const flaggedTxns = allFlags.filter((f) => {
    if (seen.has(`${f.id}-${f.flag}`)) return false;
    seen.add(`${f.id}-${f.flag}`);
    return true;
  });

  // ── Risk score ────────────────────────────────────────────────────────
  let riskScore = 0;
  if (velocityFlags.length > 0) riskScore += 20;
  if (roundTrips.length > 0) riskScore += 25;
  if (passThroughFlag) riskScore += 25;
  if (roundFlags.length > 3) riskScore += 10;
  if (largeFlags.length > 0) riskScore += 10;
  if (firstTimeLargeFlags.length > 0) riskScore += 10;
  if (structuringFlags.length > 0) riskScore += 20;
  if (dormancyFlags.length > 0) riskScore += 15;
  if (keywordFlags.length > 0) riskScore += 15;
  riskScore = Math.min(riskScore, 100);

  // ── Chart data ────────────────────────────────────────────────────────
  const monthMap2 = {};
  txns.forEach((t) => {
    const key = `${t.date.getFullYear()}-${String(t.date.getMonth() + 1).padStart(2,"0")}`;
    if (!monthMap2[key]) monthMap2[key] = { month: key, credit: 0, debit: 0 };
    monthMap2[key].credit += t.credit;
    monthMap2[key].debit += t.debit;
  });
  const monthlyData = Object.values(monthMap2).sort((a, b) => a.month.localeCompare(b.month));

  // ── Behavioral Timeline: hour-of-day and day-of-week heatmap ──────────
  const DAYS = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"];
  const hourGrid = {}; // key: "day-hour" -> {count, amount}
  txns.forEach((t) => {
    const dow = t.date.getDay();
    const hour = t.date.getHours();
    const key = `${dow}-${hour}`;
    if (!hourGrid[key]) hourGrid[key] = { day: DAYS[dow], dow, hour, count: 0, amount: 0 };
    hourGrid[key].count += 1;
    hourGrid[key].amount += t.amount;
  });
  const behaviorGrid = Object.values(hourGrid);
  const maxGridCount = Math.max(...behaviorGrid.map((g) => g.count), 1);

  // Hour distribution for bar chart
  const hourDist = Array.from({ length: 24 }, (_, h) => ({
    hour: `${String(h).padStart(2,"0")}:00`,
    count: txns.filter((t) => t.date.getHours() === h).length,
    amount: txns.filter((t) => t.date.getHours() === h).reduce((s,t) => s + t.amount, 0),
  }));

  // Day of week distribution
  const dowDist = DAYS.map((d, i) => ({
    day: d,
    count: txns.filter((t) => t.date.getDay() === i).length,
    amount: txns.filter((t) => t.date.getDay() === i).reduce((s,t) => s + t.amount, 0),
  }));

  const flagBreakdown = [
    { name: "High Velocity", value: velocityFlags.length, color: "#ef4444" },
    { name: "Round Numbers", value: roundFlags.length, color: "#f97316" },
    { name: "Round Trips", value: roundTrips.length, color: "#a855f7" },
    { name: "Late Night", value: lateNight.length, color: "#3b82f6" },
    { name: "Large Outliers", value: largeFlags.length, color: "#eab308" },
    { name: "First-Time Large", value: firstTimeLargeFlags.length, color: "#ec4899" },
    { name: "Structuring", value: structuringFlags.length, color: "#fb7185" },
    { name: "Dormancy", value: dormancyFlags.length, color: "#22d3ee" },
    { name: "Keywords", value: keywordFlags.length, color: "#fbbf24" },
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
    counterparties,
    firstTimeLargeFlags,
    structuringFlags,
    accountMeta,
    behaviorGrid,
    maxGridCount,
    hourDist,
    dowDist,
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
  FIRST_TIME_LARGE: { bg: "bg-pink-900/40", border: "border-pink-500/50", text: "text-pink-400", label: "FIRST TIME LARGE" },
  STRUCTURING: { bg: "bg-rose-900/40", border: "border-rose-400/50", text: "text-rose-300", label: "STRUCTURING" },
  DORMANCY_ACTIVATION: { bg: "bg-cyan-900/40", border: "border-cyan-500/50", text: "text-cyan-400", label: "DORMANCY" },
  KEYWORD_FLAG: { bg: "bg-amber-900/40", border: "border-amber-500/50", text: "text-amber-400", label: "KEYWORD" },
};

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

━━━━━━━━━���━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

KEY JUDGMENTS:
1. ${j1}

2. ${j2}

3. ${j3}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

INDICATORS OF CONCERN:
${indicators}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━��━━━━━━━━━━━━━━━

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
END OF BRIEF — ${ref} | SYSTEM: TRANSACTION PATTERN ANALYZER v1.0 | DEVELOPED BY: YAKUBU TANIMU — NAUB CRIMINOLOGY & SECURITY STUDIES`;

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
  const [searchQuery, setSearchQuery] = useState("");
  const [filterFlag, setFilterFlag] = useState("ALL");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const briefRef = useRef(null);

  const handleFile = useCallback((file) => {
    setError("");
    const name = file.name.toLowerCase();

        const processRows = (rows, fields, accountMeta = {}) => {
      if (!rows.length) { setError("No data found in file."); return; }
      // Filter out __EMPTY keys SheetJS generates for blank Excel columns
      const cleanFields = fields.filter(f => f && !String(f).startsWith("__EMPTY") && String(f).trim() !== "");
      const detected = detectColumns(cleanFields);
      if (!detected.date) {
        setError(`Could not detect a date column. Columns found: ${cleanFields.slice(0,8).join(" | ")}`);
        return;
      }
      setCsvData({ data: rows, meta: { fields: cleanFields, accountMeta } });
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
          // Extract account metadata from rows before the header
          const extractMetadata = (ws) => {
            const meta = {};
            const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
            const metaKeywords = {
              "account name": "accountName",
              "account no": "accountNumber",
              "account number": "accountNumber",
              "account type": "accountType",
              "currency": "currency",
              "branch": "branch",
              "opening balance": "openingBalance",
              "closing balance": "closingBalance",
              "available balance": "availableBalance",
              "period": "period",
              "from": "periodFrom",
              "to": "periodTo",
              "statement date": "statementDate",
              "customer name": "accountName",
              "customer id": "customerId",
              "bvn": "bvn",
              "email": "email",
              "phone": "phone",
              "address": "address",
              "sort code": "sortCode",
              "bank": "bank",
              "branch code": "branchCode",
            };
            for (let r = 0; r <= Math.min(range.e.r, 30); r++) {
              const rowVals = [];
              for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = ws[XLSX.utils.encode_cell({ r, c })];
                rowVals.push(cell ? String(cell.v).trim() : "");
              }
              // Check pairs: label | value
              for (let c = 0; c < rowVals.length - 1; c++) {
                const label = rowVals[c].toLowerCase().replace(/[:\s]+$/, "").trim();
                const value = rowVals[c + 1];
                if (label && value && metaKeywords[label]) {
                  meta[metaKeywords[label]] = value;
                }
              }
              // Also check if whole row is a title (single non-empty cell)
              const filled = rowVals.filter((v) => v !== "");
              if (filled.length === 1 && !meta.bankTitle) {
                meta.bankTitle = filled[0];
              }
            }
            return meta;
          };

          const accountMeta = extractMetadata(ws);

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
          processRows(json, fields, accountMeta);
        } catch {
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
    setIsLoading(true);
    setError("");
    setTimeout(() => {
      try {
        const result = analyzeTransactions(csvData.data, cols, csvData.meta.accountMeta || {});
        if (!result) { setError("No valid transactions found after parsing."); setIsLoading(false); return; }
        setAnalysis(result);
        setStage("analysis");
      } catch (err) {
        setError("Analysis failed: " + err.message);
      }
      setIsLoading(false);
    }, 50);
  };

  const handleGenerateBrief = async () => {
    setBriefLoading(true);
    setBrief("");
    setStage("brief");
    try {
      await generateBrief(analysis, (text) => setBrief(text));
    } catch {
      setBrief("Error generating brief. Please try again.");
    }
    setBriefLoading(false);
  };

  const handlePrint = () => window.print();

  const risk = analysis ? riskLevel(analysis.riskScore) : null;

  return (
    <div className="min-h-screen bg-[#0a0c10] text-gray-100" style={{ fontFamily: "'Courier New', monospace", position: "relative" }}>
      {/* Background watermark */}
      <div style={{
        position: "fixed", inset: 0, zIndex: 0, pointerEvents: "none",
        display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: "60px",
        overflow: "hidden",
      }}>
        {/* Shield logo */}
        <svg width="320" height="320" viewBox="0 0 320 320" style={{ opacity: 0.03, position: "absolute", top: "50%", left: "50%", transform: "translate(-50%, -60%)" }}>
          <path d="M160 20 L280 70 L280 160 C280 230 220 285 160 305 C100 285 40 230 40 160 L40 70 Z" fill="#00ff9d" stroke="#00ff9d" strokeWidth="2"/>
          <path d="M160 50 L255 90 L255 160 C255 215 212 262 160 278 C108 262 65 215 65 160 L65 90 Z" fill="none" stroke="#00ff9d" strokeWidth="1.5"/>
          <text x="160" y="155" textAnchor="middle" fontSize="28" fontWeight="bold" fill="#00ff9d" fontFamily="Courier New">FIU</text>
          <text x="160" y="185" textAnchor="middle" fontSize="11" fill="#00ff9d" fontFamily="Courier New" letterSpacing="3">NIGERIA</text>
          <line x1="100" y1="200" x2="220" y2="200" stroke="#00ff9d" strokeWidth="1"/>
          <text x="160" y="218" textAnchor="middle" fontSize="8" fill="#00ff9d" fontFamily="Courier New" letterSpacing="2">FINANCIAL INTELLIGENCE</text>
        </svg>
        {/* Repeated CLASSIFIED text rows */}
        {Array.from({ length: 12 }, (_, row) => (
          <div key={row} style={{
            position: "absolute",
            top: `${row * 120 - 100}px`,
            left: "-100px", right: "-100px",
            display: "flex", gap: "60px",
            transform: `rotate(-25deg)`,
            opacity: 0.025,
            whiteSpace: "nowrap",
          }}>
            {Array.from({ length: 8 }, (_, col) => (
              <span key={col} style={{
                color: "#00ff9d",
                fontSize: "22px",
                fontWeight: "bold",
                fontFamily: "Courier New",
                letterSpacing: "8px",
                flexShrink: 0,
              }}>CLASSIFIED</span>
            ))}
          </div>
        ))}
      </div>
      {/* All content above watermark */}
      <div style={{ position: "relative", zIndex: 1 }}>
      {/* Print styles */}
      <style>{`
        /* Force dark background everywhere including artifact viewer */
        html, body, #root {
          background: #0a0c10 !important;
          color: #f3f4f6 !important;
          min-height: 100vh !important;
        }
        @media print {
          html, body, #root { background: #0a0c10 !important; color: #f3f4f6 !important; }
          .no-print { display: none !important; }
          * { font-family: 'Courier New', monospace !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }
          .border { border-color: #374151 !important; }
        }
        .blink { animation: blink 1.2s step-end infinite; }
        @keyframes blink { 50% { opacity: 0; } }
        .brief-text { white-space: pre-wrap; line-height: 1.8; }
      `}</style>


      {/* Header */}
      <header className="no-print border-b border-[#00ff9d]/20 px-6 py-4 flex items-center justify-between">
        <div>
          <div className="text-[#00ff9d] text-xs tracking-[0.3em] mb-1">FINANCIAL INTELLIGENCE UNIT</div>
          <div className="text-white text-xl font-bold tracking-wider">TRANSACTION PATTERN ANALYZER</div>
          {analysis?.accountMeta?.bank || analysis?.accountMeta?.bankTitle ? (
            <div className="text-[#00ff9d]/60 text-xs mt-1 tracking-widest">
              {(analysis.accountMeta.bank || analysis.accountMeta.bankTitle).toUpperCase()}
            </div>
          ) : null}
        </div>
        <div className="text-right text-xs text-gray-500">
          <div>STATUS: <span className="text-[#00ff9d]">ACTIVE</span></div>
          <div>{new Date().toISOString().split("T")[0]}</div>
          {analysis?.accountMeta?.accountName && (
            <div className="text-white text-xs mt-1 font-bold">{analysis.accountMeta.accountName}</div>
          )}
          {analysis?.accountMeta?.accountNumber && (
            <div className="text-gray-500 text-[10px]">{analysis.accountMeta.accountNumber}</div>
          )}
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
              <div className="text-gray-500 text-sm">Supports Excel (.xlsx/.xls), CSV, and PDF — Moniepoint, GTBank, Access, Zenith, UBA, Opay, etc.</div>
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
              <div className="text-white text-lg mb-2">Drop your bank statement here</div>
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
              <button onClick={handleRun} disabled={isLoading}
                className="px-8 py-2 bg-[#00ff9d] text-black font-bold text-sm rounded hover:bg-[#00ff9d]/90 transition-all tracking-wider disabled:opacity-60 flex items-center gap-2">
                {isLoading ? (
                  <>
                    <svg className="animate-spin h-4 w-4" viewBox="0 0 24 24" fill="none">
                      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"/>
                      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z"/>
                    </svg>
                    ANALYZING...
                  </>
                ) : "RUN ANALYSIS →"}
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

            {/* Account Info Card */}
            {analysis.accountMeta && Object.keys(analysis.accountMeta).length > 0 && (
              <div className="border border-[#00ff9d]/20 rounded p-4 mb-6 bg-[#0d1117]">
                <div className="text-xs text-[#00ff9d] tracking-wider mb-3">SUBJECT ACCOUNT PROFILE</div>
                <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
                  {[
                    ["Account Name", analysis.accountMeta.accountName],
                    ["Account Number", analysis.accountMeta.accountNumber],
                    ["Account Type", analysis.accountMeta.accountType],
                    ["Currency", analysis.accountMeta.currency],
                    ["Bank / Institution", analysis.accountMeta.bank || analysis.accountMeta.bankTitle],
                    ["Branch", analysis.accountMeta.branch],
                    ["Sort Code", analysis.accountMeta.sortCode],
                    ["BVN", analysis.accountMeta.bvn],
                    ["Customer ID", analysis.accountMeta.customerId],
                    ["Email", analysis.accountMeta.email],
                    ["Phone", analysis.accountMeta.phone],
                    ["Statement Period", analysis.accountMeta.period],
                    ["Period From", analysis.accountMeta.periodFrom],
                    ["Period To", analysis.accountMeta.periodTo],
                    ["Opening Balance", analysis.accountMeta.openingBalance],
                    ["Closing Balance", analysis.accountMeta.closingBalance],
                    ["Available Balance", analysis.accountMeta.availableBalance],
                  ].filter(([, v]) => v && String(v).trim() !== "").map(([label, value]) => (
                    <div key={label} className="bg-gray-900/50 rounded p-2">
                      <div className="text-gray-500 text-[10px] tracking-wider mb-0.5">{label.toUpperCase()}</div>
                      <div className="text-white text-xs font-mono">{String(value)}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}

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
            <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mb-6">
              {[
                ["TOTAL TRANSACTIONS", analysis.txns.length.toLocaleString(), "#a78bfa"],
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

            {/* ── GLOBAL SEARCH ────────────────────────────────── */}
            <div className="border border-[#00ff9d]/20 rounded mb-6 bg-[#0d1117]">
              <div className="px-4 py-3 border-b border-gray-800 text-xs text-[#00ff9d] tracking-wider">GLOBAL SEARCH</div>
              <div className="p-4">
                <div className="flex items-center gap-2 mb-3">
                  <input
                    type="text"
                    placeholder="Search all transactions by name, amount, date, narration..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="flex-1 bg-gray-900 border border-gray-700 rounded px-3 py-2 text-xs text-gray-300 placeholder-gray-600 focus:outline-none focus:border-[#00ff9d]/50"
                  />
                  {searchQuery && (
                    <button onClick={() => setSearchQuery("")} className="text-gray-500 hover:text-white text-xs px-3 py-2 border border-gray-700 rounded">✕ Clear</button>
                  )}
                </div>
                {/* Date range filter */}
                <div className="flex items-center gap-2 mb-3">
                  <span className="text-[10px] text-gray-500 tracking-wider flex-shrink-0">DATE RANGE</span>
                  <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)}
                    className="bg-gray-900 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300 focus:outline-none focus:border-[#00ff9d]/50"/>
                  <span className="text-gray-600 text-xs">to</span>
                  <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)}
                    className="bg-gray-900 border border-gray-700 rounded px-2 py-1 text-xs text-gray-300 focus:outline-none focus:border-[#00ff9d]/50"/>
                  {(dateFrom || dateTo) && (
                    <button onClick={() => { setDateFrom(""); setDateTo(""); }} className="text-gray-500 hover:text-white text-[10px] px-2 py-1 border border-gray-700 rounded">Clear</button>
                  )}
                </div>
                {/* Filter by flag type */}
                <div className="flex flex-wrap gap-2 mb-3">
                  {["ALL", "FLAGGED", ...Object.keys(FLAG_COLORS)].map((f) => (
                    <button key={f} onClick={() => setFilterFlag(f)}
                      className={`text-[10px] px-2 py-1 rounded border tracking-wider transition-all ${filterFlag === f ? "border-[#00ff9d] text-[#00ff9d] bg-[#00ff9d]/10" : "border-gray-700 text-gray-500 hover:border-gray-500"}`}>
                      {f.replace(/_/g, " ")}
                    </button>
                  ))}
                </div>
                {/* Results */}
                {searchQuery || filterFlag !== "ALL" ? (() => {
                  const q = searchQuery.toLowerCase();
                  const flaggedIds = new Set(analysis.flaggedTxns.map((f) => f.id));
                  const fromDate = dateFrom ? new Date(dateFrom) : null;
                  const toDate = dateTo ? new Date(dateTo + "T23:59:59") : null;
                  const results = analysis.txns.filter((t) => {
                    const matchesFlag = filterFlag === "ALL" ? true
                      : filterFlag === "FLAGGED" ? flaggedIds.has(t.id)
                      : analysis.flaggedTxns.some((f) => f.id === t.id && f.flag === filterFlag);
                    const matchesQuery = !q ||
                      (t.description || "").toLowerCase().includes(q) ||
                      fmt(t.amount).includes(q) ||
                      t.date?.toDateString().toLowerCase().includes(q) ||
                      String(t.amount).includes(q);
                    const matchesDate = (!fromDate || t.date >= fromDate) && (!toDate || t.date <= toDate);
                    return matchesFlag && matchesQuery && matchesDate;
                  });
                  // Also search counterparties
                  const cpResults = q ? (analysis.scoredCounterparties || []).filter((cp) =>
                    cp.name.toLowerCase().includes(q)
                  ) : [];
                  return (
                    <div>
                      {cpResults.length > 0 && (
                        <div className="mb-3">
                          <div className="text-[10px] text-gray-500 tracking-wider mb-2">COUNTERPARTY MATCHES ({cpResults.length})</div>
                          {cpResults.map((cp, i) => {
                            const cpRiskColor = cp.riskScore >= 60 ? "#ef4444" : cp.riskScore >= 30 ? "#f97316" : "#22c55e";
                            return (
                              <div key={i} className="flex items-center justify-between border border-gray-800 rounded px-3 py-2 mb-1 bg-gray-900/40">
                                <div>
                                  <span className="text-white text-xs font-bold">{cp.name}</span>
                                  <span className="text-gray-500 text-[10px] ml-2">{cp.txnCount} txns</span>
                                </div>
                                <div className="text-right text-xs">
                                  <span className="font-bold mr-2" style={{ color: cpRiskColor }}>RISK {cp.riskScore}</span>
                                  <span className="text-[#00ff9d]">+{fmt(cp.totalReceived)}</span>
                                  <span className="text-red-400 ml-2">-{fmt(cp.totalSent)}</span>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      )}
                      <div className="text-[10px] text-gray-500 tracking-wider mb-2">
                        TRANSACTIONS ({results.length} of {analysis.txns.length})
                      </div>
                      {results.length === 0 && (
                        <div className="py-6 text-center text-gray-600 text-sm">No results found</div>
                      )}
                      <div className="max-h-72 overflow-y-auto">
                        {results.slice(0, 100).map((t, i) => {
                          const flags = analysis.flaggedTxns.filter((f) => f.id === t.id);
                          return (
                            <div key={i} className={`border-b border-gray-900 px-3 py-2.5 hover:bg-gray-900/40 ${flags.length ? "border-l-2 border-l-yellow-500/50" : ""}`}>
                              <div className="flex items-start justify-between gap-4">
                                <div className="flex-1 min-w-0">
                                  <div className="flex items-center gap-2 mb-0.5 flex-wrap">
                                    <span className="text-gray-500 text-[10px]">{t.date?.toDateString()}</span>
                                    {flags.map((f, fi) => {
                                      const fc = FLAG_COLORS[f.flag] || FLAG_COLORS.ROUND_NUMBER;
                                      return <span key={fi} className={`text-[9px] px-1 py-0.5 rounded border ${fc.border} ${fc.text}`}>{fc.label}</span>;
                                    })}
                                  </div>
                                  <div className="text-gray-300 text-xs truncate">{t.description || "—"}</div>
                                </div>
                                <div className={`font-bold text-sm flex-shrink-0 ${t.isCredit ? "text-[#00ff9d]" : "text-red-400"}`}>
                                  {t.isCredit ? "+" : "-"}{fmt(t.amount)}
                                </div>
                              </div>
                            </div>
                          );
                        })}
                        {results.length > 100 && (
                          <div className="px-4 py-2 text-center text-gray-600 text-xs">Showing first 100 of {results.length} results — refine your search</div>
                        )}
                      </div>
                    </div>
                  );
                })() : (
                  <div className="text-center text-gray-600 text-xs py-4">Type to search all {analysis.txns.length} transactions or filter by flag type above</div>
                )}
              </div>
            </div>

            {/* ── EMPTY STATE ──────────────────────────────────── */}
            {analysis.flaggedTxns.length === 0 && (
              <div className="border border-[#00ff9d]/20 rounded p-8 mb-6 text-center">
                <div className="text-4xl mb-3">✓</div>
                <div className="text-[#00ff9d] font-bold text-sm tracking-wider mb-1">NO SUSPICIOUS ACTIVITY DETECTED</div>
                <div className="text-gray-500 text-xs">No transactions triggered AML flags during this assessment period. This does not rule out all risk — manual review is still recommended for high-value transactions.</div>
              </div>
            )}

            {/* ── FLAGGED TRANSACTIONS ──────────────────────────── */}

            {/* ── BEHAVIORAL TIMELINE — RADIAL CLOCK ───────────── */}
            {analysis.hourDist && (
              <div className="border border-gray-800 rounded mb-6">
                <div className="px-4 py-3 border-b border-gray-800 text-xs text-gray-500 tracking-wider">
                  BEHAVIORAL TIMELINE — 24-HOUR ACTIVITY CLOCK
                </div>
                <div className="p-4 flex flex-col md:flex-row gap-6 items-center">
                  <div className="flex-shrink-0 flex flex-col items-center">
                    <svg width="220" height="220" viewBox="0 0 220 220" role="img" aria-label="24-hour radial clock showing transaction activity">
                      <circle cx="110" cy="110" r="100" fill="none" stroke="#1f2937" strokeWidth="1"/>
                      <circle cx="110" cy="110" r="65" fill="none" stroke="#1f2937" strokeWidth="0.5" strokeDasharray="4,4"/>
                      <circle cx="110" cy="110" r="30" fill="#0a0c10" stroke="#1f2937" strokeWidth="1"/>
                      {[0,3,6,9,12,15,18,21].map((h) => {
                        const angle = (h / 24) * 2 * Math.PI - Math.PI / 2;
                        const x = 110 + 112 * Math.cos(angle);
                        const y = 110 + 112 * Math.sin(angle);
                        return <text key={h} x={x.toFixed(1)} y={y.toFixed(1)} textAnchor="middle" dominantBaseline="middle" fontSize="9" fill="#4b5563">{String(h).padStart(2,"0")}h</text>;
                      })}
                      {analysis.hourDist.map((h, idx) => {
                        if (h.count === 0) return null;
                        const angle = (idx / 24) * 2 * Math.PI - Math.PI / 2;
                        const maxCount = Math.max(...analysis.hourDist.map((x) => x.count), 1);
                        const intensity = h.count / maxCount;
                        const r = 32 + intensity * 60;
                        const x = 110 + r * Math.cos(angle);
                        const y = 110 + r * Math.sin(angle);
                        const isNight = idx >= 23 || idx <= 4;
                        const dotR = Math.max(3, intensity * 10);
                        return (
                          <g key={idx}>
                            <line x1="110" y1="110" x2={x.toFixed(1)} y2={y.toFixed(1)} stroke={isNight ? "rgba(239,68,68,0.3)" : "rgba(0,255,157,0.2)"} strokeWidth="1"/>
                            <circle cx={x.toFixed(1)} cy={y.toFixed(1)} r={dotR.toFixed(1)} fill={isNight ? "#ef4444" : "#00ff9d"} opacity={0.3 + intensity * 0.7}/>
                          </g>
                        );
                      })}
                      <text x="110" y="107" textAnchor="middle" fontSize="10" fill="#6b7280">24h</text>
                      <text x="110" y="118" textAnchor="middle" fontSize="10" fill="#6b7280">clock</text>
                    </svg>
                    <div className="flex gap-4 mt-2 text-[10px]">
                      <span className="flex items-center gap-1"><span className="inline-block w-2 h-2 rounded-full bg-[#00ff9d]"/><span className="text-gray-500">Business hours</span></span>
                      <span className="flex items-center gap-1"><span className="inline-block w-2 h-2 rounded-full bg-red-500"/><span className="text-gray-500">Late night</span></span>
                    </div>
                  </div>
                  <div className="flex-1 w-full">
                    <div className="text-xs text-gray-500 mb-3 tracking-wider">ACTIVITY BY DAY OF WEEK</div>
                    {analysis.dowDist.map((d) => {
                      const maxCount = Math.max(...analysis.dowDist.map((x) => x.count), 1);
                      const pct = Math.max((d.count / maxCount) * 100, 2);
                      const isWeekend = d.day === "Sat" || d.day === "Sun";
                      return (
                        <div key={d.day} className="flex items-center gap-2 mb-2">
                          <div className="text-xs text-gray-500 w-8">{d.day}</div>
                          <div className="flex-1 h-5 bg-gray-900 rounded-sm overflow-hidden">
                            <div className="h-full rounded-sm" style={{ width: `${pct}%`, background: isWeekend ? "#f97316" : "#00ff9d", opacity: 0.8 }}/>
                          </div>
                          <div className="text-xs text-gray-500 w-8 text-right">{d.count}</div>
                        </div>
                      );
                    })}
                    {(() => {
                      const peakIdx = analysis.hourDist.reduce((maxI, h, i, arr) => h.count > arr[maxI].count ? i : maxI, 0);
                      const peak = analysis.hourDist[peakIdx];
                      const isNight = peakIdx >= 23 || peakIdx <= 4;
                      return (
                        <div className={`mt-4 p-3 rounded text-xs border ${isNight ? "border-red-500/30 bg-red-900/10 text-red-400" : "border-[#00ff9d]/20 bg-[#00ff9d]/5 text-[#00ff9d]"}`}>
                          Peak activity: <strong>{String(peakIdx).padStart(2,"0")}:00</strong> — {peak.count} transactions {isNight ? "⚠ Suspicious hour" : "✓ Normal business hour"}
                        </div>
                      );
                    })()}
                  </div>
                </div>
              </div>
            )}


            {/* ── STRUCTURING ALERTS ───────────────────────────── */}
            {analysis.structuringFlags && analysis.structuringFlags.length > 0 && (
              <div className="border border-rose-500/30 bg-rose-900/10 rounded mb-6">
                <div className="px-4 py-3 border-b border-rose-500/20 text-xs text-rose-400 tracking-wider flex justify-between">
                  <span>⚠ STRUCTURING ALERTS — POSSIBLE THRESHOLD EVASION</span>
                  <span>{analysis.structuringFlags.length} pattern{analysis.structuringFlags.length > 1 ? "s" : ""} detected</span>
                </div>
                <div className="max-h-56 overflow-y-auto">
                  {analysis.structuringFlags.map((f, i) => (
                    <div key={i} className="border-b border-rose-900/30 px-4 py-3">
                      <div className="text-rose-300 text-xs font-bold mb-1">STRUCTURING PATTERN #{i + 1}</div>
                      <div className="text-gray-300 text-xs">{f.detail}</div>
                      <div className="text-gray-500 text-[10px] mt-1">{f.date?.toDateString()}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* ── DORMANCY ALERTS ──────────────────────────────── */}
            {analysis.dormancyFlags && analysis.dormancyFlags.length > 0 && (
              <div className="border border-cyan-500/30 bg-cyan-900/10 rounded mb-6">
                <div className="px-4 py-3 border-b border-cyan-500/20 text-xs text-cyan-400 tracking-wider flex justify-between">
                  <span>⚠ DORMANCY & SUDDEN ACTIVATION</span>
                  <span>{analysis.dormancyFlags.length} instance{analysis.dormancyFlags.length > 1 ? "s" : ""}</span>
                </div>
                <div className="max-h-48 overflow-y-auto">
                  {analysis.dormancyFlags.map((f, i) => (
                    <div key={i} className="border-b border-cyan-900/30 px-4 py-3">
                      <div className="text-cyan-300 text-xs font-bold mb-1">REACTIVATION EVENT #{i + 1}</div>
                      <div className="text-gray-300 text-xs">{f.detail}</div>
                      <div className="text-gray-500 text-[10px] mt-1">{f.description}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* ── KEYWORD FLAGS ─────────────────────────────────── */}
            {analysis.keywordFlags && analysis.keywordFlags.length > 0 && (
              <div className="border border-amber-500/30 bg-amber-900/10 rounded mb-6">
                <div className="px-4 py-3 border-b border-amber-500/20 text-xs text-amber-400 tracking-wider flex justify-between">
                  <span>⚠ SUSPICIOUS KEYWORD FLAGS</span>
                  <span>{analysis.keywordFlags.length} transaction{analysis.keywordFlags.length > 1 ? "s" : ""}</span>
                </div>
                <div className="max-h-48 overflow-y-auto">
                  {analysis.keywordFlags.map((f, i) => (
                    <div key={i} className="border-b border-amber-900/30 px-4 py-3">
                      <div className="flex items-start justify-between gap-4">
                        <div className="flex-1 min-w-0">
                          <div className="text-amber-400 text-xs font-bold mb-1">{f.detail}</div>
                          <div className="text-gray-300 text-xs truncate">{f.description}</div>
                          <div className="text-gray-500 text-[10px] mt-1">{f.date?.toDateString()}</div>
                        </div>
                        <div className={`font-bold text-sm flex-shrink-0 ${f.isCredit ? "text-[#00ff9d]" : "text-red-400"}`}>
                          {f.isCredit ? "+" : "-"}{fmt(f.amount)}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* ── COUNTERPARTY NETWORK ─────────────────────────── */}
            {analysis.counterparties && analysis.counterparties.length > 0 && (
              <div className="border border-gray-800 rounded mb-6">
                <div className="px-4 py-3 border-b border-gray-800 text-xs text-gray-500 tracking-wider flex justify-between">
                  <span>COUNTERPARTY NETWORK ANALYSIS</span>
                  <span>{analysis.counterparties.length} unique counterparties</span>
                </div>
                <div className="max-h-80 overflow-y-auto">
                  {analysis.counterparties.slice(0, 50).map((cp, i) => {
                    const cpRiskColor = cp.riskScore >= 60 ? "#ef4444" : cp.riskScore >= 30 ? "#f97316" : "#22c55e";
                    const cpRiskLabel = cp.riskScore >= 60 ? "HIGH" : cp.riskScore >= 30 ? "MED" : "LOW";
                    return (
                      <div key={i} className="border-b border-gray-900 px-4 py-3 hover:bg-gray-900/40">
                        <div className="flex items-start justify-between gap-4">
                          <div className="flex-1 min-w-0">
                            <div className="flex items-center gap-2 mb-1 flex-wrap">
                              <span className="text-white text-xs font-bold truncate">{cp.name}</span>
                              <span className="text-[10px] px-1.5 py-0.5 rounded font-bold" style={{ color: cpRiskColor, border: `1px solid ${cpRiskColor}40` }}>{cpRiskLabel} {cp.riskScore}</span>
                              {cp.txnCount === 1 && (cp.totalSent + cp.totalReceived) >= 100000 && (
                                <span className="text-[10px] px-1.5 py-0.5 rounded border border-pink-500/50 text-pink-400">FIRST TIME</span>
                              )}
                            </div>
                            <div className="flex gap-4 text-xs text-gray-500 flex-wrap">
                              <span>{cp.txnCount} txn{cp.txnCount > 1 ? "s" : ""}</span>
                              <span>First: {cp.firstDate?.toLocaleDateString()}</span>
                              <span>Last: {cp.lastDate?.toLocaleDateString()}</span>
                            </div>
                            {cp.bank && <div className="text-[10px] text-gray-600 mt-0.5">via {cp.bank}</div>}
                            {/* Mini risk bar */}
                            <div className="mt-1.5 h-1 bg-gray-800 rounded-full overflow-hidden w-32">
                              <div className="h-full rounded-full" style={{ width: `${cp.riskScore}%`, background: cpRiskColor }}/>
                            </div>
                          </div>
                          <div className="text-right flex-shrink-0 text-xs">
                            {cp.totalReceived > 0 && <div className="text-[#00ff9d]">+{fmt(cp.totalReceived)}</div>}
                            {cp.totalSent > 0 && <div className="text-red-400">-{fmt(cp.totalSent)}</div>}
                            <div className="text-gray-500 mt-0.5">Net: {fmt(cp.totalReceived - cp.totalSent)}</div>
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
                <div className="border border-yellow-500/30 bg-yellow-900/10 rounded p-3 mb-4 text-xs text-yellow-400">
                  ⚠ DISCLAIMER: This brief is generated by an automated analytical system for investigative assistance only. All findings are probabilistic and must be verified by a qualified financial intelligence officer before any enforcement, legal, or administrative action is taken. Unauthorised dissemination of this document is prohibited.
                </div>
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
      <SpeedInsights />
    </div>
  );
}
