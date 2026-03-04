import { useState, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Legend,
  Cell,
} from "recharts";

// ─── CONFIG ───────────────────────────────────────────────────────────────────
const YEAR = 2026;
const MONTHS = [
  "",
  "Ene",
  "Feb",
  "Mar",
  "Abr",
  "May",
  "Jun",
  "Jul",
  "Ago",
  "Sep",
  "Oct",
  "Nov",
  "Dic",
];

const BCE_RATES_2026 = {
  1: 1.1738,
  2: 1.181,
};

const calcAvg = (rates) => {
  const vals = Object.values(rates);
  return vals.length > 0 ? vals.reduce((a, b) => a + b, 0) / vals.length : 1;
};

const BCE_AVG_2026 = calcAvg(BCE_RATES_2026);

const EMPRESAS_CONFIG = [
  {
    name: "Rentaltoken SL",
    color: "#00d4aa",
    parser: "holded",
    currency: "EUR",
  },
  { name: "Admin RNT SL", color: "#818cf8", parser: "holded", currency: "EUR" },
  {
    name: "Reental Tourist Homes",
    color: "#fb923c",
    parser: "holded",
    currency: "EUR",
  },
  {
    name: "Reental Rock Capital",
    color: "#f472b6",
    parser: "holded",
    currency: "EUR",
  },
  {
    name: "Reental America LLC",
    color: "#facc15",
    parser: "quickbooks",
    currency: "USD",
  },
];

// ─── REFRESCO BCE VÍA ANTHROPIC API ───────────────────────────────────────────
async function refreshRatesViaAI(year) {
  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      tools: [{ type: "web_search_20250305", name: "web_search" }],
      messages: [
        {
          role: "user",
          content: `Fetch the official monthly average EUR/USD exchange rates for ${year} from the ECB (European Central Bank) or Banco de España. Return ONLY a valid JSON object with the month numbers that have been officially published (do NOT include months without official data). Keys are month numbers as strings (1-12) and values are rates. Example: {"1":1.0385,"2":1.0446}. Use only confirmed BCE/BdE reference rates, no estimates or projections.`,
        },
      ],
    }),
  });
  if (!resp.ok) throw new Error(`API ${resp.status}`);
  const data = await resp.json();
  const text = data.content
    .filter((b) => b.type === "text")
    .map((b) => b.text)
    .join("");
  const match = text.match(/\{[\s\S]*\}/);
  if (!match) throw new Error("No JSON in response");
  const raw = JSON.parse(match[0]);
  const rates = {};
  for (const [k, v] of Object.entries(raw)) {
    const m = parseInt(k);
    const r = parseFloat(v);
    if (m >= 1 && m <= 12 && r > 0.5 && r < 2) rates[m] = r;
  }
  if (Object.keys(rates).length < 1) throw new Error("Insufficient data");
  return rates;
}

// ─── CATEGORÍAS ───────────────────────────────────────────────────────────────
const CATEGORY_RULES = [
  {
    category: "setup_fees",
    type: "income",
    prefixes: ["7050", "7051", "7052", "7053"],
    keywords: ["setup", "onboarding", "alta"],
  },
  {
    category: "rental_contracts",
    type: "income",
    prefixes: ["7520"],
    keywords: ["alquiler", "recurrente alquiler", "renta", "rinc"],
  },
  {
    category: "rnt_sales",
    type: "income",
    prefixes: ["7000"],
    keywords: [
      "venta mercan",
      "nft",
      "token",
      "rnt",
      "reental pro",
      "superreental",
      "upgrade",
    ],
  },
  {
    category: "properties_sale",
    type: "income",
    prefixes: ["7710", "7711"],
    keywords: ["venta inmueble", "plusvalia", "enajenacion"],
  },
  {
    category: "other_revenues",
    type: "income",
    prefixes: ["7400", "7590", "7750", "7760", "7800"],
    keywords: ["otros ingresos", "subvencion", "ingreso excepcional"],
  },
  {
    category: "rental_costs",
    type: "expense",
    prefixes: ["6280", "6281", "6282", "6000"],
    keywords: [
      "suministro",
      "luz",
      "agua",
      "gas",
      "electricidad",
      "iberdrola",
      "holaluz",
      "emasesa",
      "facsa",
      "fomento agricola",
    ],
  },
  {
    category: "rnt_expenses",
    type: "expense",
    prefixes: ["6290010", "6290020"],
    keywords: ["comunidad", "limpieza", "mantenimiento", "rd"],
  },
  {
    category: "properties_sale_cost",
    type: "expense",
    prefixes: ["6710", "6711"],
    keywords: ["coste venta inmueble"],
  },
  {
    category: "tech_costs",
    type: "expense",
    prefixes: ["6290040"],
    keywords: [
      "plataformas operativas",
      "google cloud",
      "webflow",
      "saas",
      "tech",
      "operations",
    ],
  },
  {
    category: "ga_personnel",
    type: "expense",
    prefixes: ["6220", "6230"],
    keywords: [
      "nomina",
      "personal",
      "salario",
      "autonomo",
      "proveedores finance",
      "proveedores rnt",
      "notario",
      "gestor",
      "registro",
      "finance",
      "ga",
    ],
  },
  {
    category: "ga_suppliers",
    type: "expense",
    prefixes: ["6250", "6260", "6280000", "6000"],
    keywords: ["seguro", "bancario", "comision banco", "ga"],
  },
  {
    category: "sales_marketing",
    type: "expense",
    prefixes: ["6270", "6284"],
    keywords: ["publicidad", "marketing", "digital publicidad", "sm"],
  },
  {
    category: "interests_token_holders",
    type: "expense",
    prefixes: ["6690"],
    keywords: [
      "token holder",
      "distribucion",
      "intereses token",
      "otros gastos financieros",
    ],
  },
  {
    category: "fx_gain_loss",
    type: "mixed",
    prefixes: ["7680", "6680"],
    keywords: ["diferencia cambio", "fx", "divisa"],
  },
  {
    category: "bank_interests",
    type: "mixed",
    prefixes: ["7620", "6623"],
    keywords: ["interes banco", "rendimiento"],
  },
  {
    category: "other_expenses",
    type: "expense",
    prefixes: ["6620", "6780", "6290000"],
    keywords: [
      "ibi",
      "impuesto",
      "no deducible",
      "herramientas digital",
      "otros gastos",
    ],
  },
];

const QB_PL_MAP = {
  "services income": { category: "setup_fees", type: "income" },
  "service income": { category: "setup_fees", type: "income" },
  "tokenization fees": { category: "setup_fees", type: "income" },
  "setup fees": { category: "setup_fees", type: "income" },
  sales: { category: "rnt_sales", type: "income" },
  "sales income": { category: "rnt_sales", type: "income" },
  "rental income": { category: "rental_contracts", type: "income" },
  "rent income": { category: "rental_contracts", type: "income" },
  "other income": { category: "other_revenues", type: "income" },
  "interest income": { category: "bank_interests", type: "income" },
  "accounting and administrative": {
    category: "ga_suppliers",
    type: "expense",
  },
  "bank service charges": { category: "ga_suppliers", type: "expense" },
  "bank charges": { category: "ga_suppliers", type: "expense" },
  "consulting fees": { category: "ga_personnel", type: "expense" },
  "legal fees": { category: "ga_suppliers", type: "expense" },
  marketing: { category: "sales_marketing", type: "expense" },
  advertising: { category: "sales_marketing", type: "expense" },
  "meals and entertainment": { category: "sales_marketing", type: "expense" },
  "other administrative expenses": {
    category: "ga_suppliers",
    type: "expense",
  },
  "professional fees": { category: "ga_personnel", type: "expense" },
  payroll: { category: "ga_personnel", type: "expense" },
  "salaries and wages": { category: "ga_personnel", type: "expense" },
  "travel and accomodations": { category: "sales_marketing", type: "expense" },
  "travel and accommodations": { category: "sales_marketing", type: "expense" },
  travel: { category: "sales_marketing", type: "expense" },
  software: { category: "tech_costs", type: "expense" },
  "software subscriptions": { category: "tech_costs", type: "expense" },
  technology: { category: "tech_costs", type: "expense" },
  insurance: { category: "ga_suppliers", type: "expense" },
  "office supplies": { category: "ga_suppliers", type: "expense" },
  utilities: { category: "rental_costs", type: "expense" },
  "interest expense": { category: "bank_interests", type: "expense" },
  taxes: { category: "other_expenses", type: "expense" },
  "income tax": { category: "other_expenses", type: "expense" },
  "other expenses": { category: "other_expenses", type: "expense" },
  "unrealized gain or loss": { category: "fx_gain_loss", type: "mixed" },
  "exchange gain or loss": { category: "fx_gain_loss", type: "mixed" },
  "foreign exchange": { category: "fx_gain_loss", type: "mixed" },
};
const QB_PL_SKIP = new Set([
  "income",
  "cost of sales",
  "gross profit",
  "expenses",
  "net earnings",
  "profit and loss",
  "distribution account",
]);
const QB_MONTH_NAMES = {
  january: 1,
  february: 2,
  march: 3,
  april: 4,
  may: 5,
  june: 6,
  july: 7,
  august: 8,
  september: 9,
  october: 10,
  november: 11,
  december: 12,
};
const ALL_CATS = [
  "setup_fees",
  "rental_contracts",
  "rnt_sales",
  "properties_sale",
  "other_revenues",
  "rental_costs",
  "rnt_expenses",
  "properties_sale_cost",
  "tech_costs",
  "ga_personnel",
  "ga_suppliers",
  "sales_marketing",
  "interests_token_holders",
  "fx_gain_loss",
  "bank_interests",
  "other_expenses",
];

function emptyMonths() {
  const r = {};
  for (let m = 1; m <= 12; m++) {
    r[m] = {};
    ALL_CATS.forEach((c) => (r[m][c] = 0));
  }
  return r;
}

// ─── PARSERS ──────────────────────────────────────────────────────────────────
function categorizeAccountHolded(accountNum, accountName) {
  const name = (accountName || "").toLowerCase(),
    acc = String(accountNum || "");
  for (const rule of CATEGORY_RULES) {
    if (rule.prefixes.some((p) => acc.startsWith(p))) return rule;
    if (rule.keywords.some((k) => name.includes(k))) return rule;
  }
  return null;
}

function parseExcelHolded(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), {
          type: "array",
          cellDates: true,
        });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
          header: 1,
          defval: null,
        });
        let hRow = -1;
        for (let i = 0; i < Math.min(10, rows.length); i++)
          if (
            rows[i]?.some((c) =>
              String(c || "")
                .toLowerCase()
                .includes("asiento")
            )
          ) {
            hRow = i;
            break;
          }
        if (hRow === -1)
          throw new Error("No se encontró la cabecera del libro diario");
        const H = rows[hRow].map((h) => String(h || "").toLowerCase());
        const iD = H.findIndex((h) => h.includes("fecha")),
          iA = H.findIndex(
            (h) => h.includes("cuenta") && !h.includes("nombre")
          ),
          iN = H.findIndex((h) => h.includes("nombre")),
          iDb = H.findIndex((h) => h.includes("debe")),
          iHb = H.findIndex((h) => h.includes("haber"));
        const result = emptyMonths(),
          unmatched = {};
        for (let i = hRow + 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || !row[iA]) continue;
          const debe = parseFloat(row[iDb] || 0),
            haber = parseFloat(row[iHb] || 0);
          let month = null;
          const dv = row[iD];
          if (dv instanceof Date) month = dv.getMonth() + 1;
          else if (typeof dv === "string") {
            const p = dv.split(/[\/\-]/);
            if (p.length >= 2) month = parseInt(p[1]);
          } else if (typeof dv === "number")
            month = new Date((dv - 25569) * 86400 * 1000).getMonth() + 1;
          if (!month || month < 1 || month > 12) continue;
          const f = String(row[iA] || "")[0];
          if (f !== "6" && f !== "7") continue;
          const rule = categorizeAccountHolded(row[iA], row[iN]);
          if (!rule) {
            const k = `${row[iA]} — ${row[iN]}`;
            unmatched[k] = (unmatched[k] || 0) + Math.abs(debe - haber);
            continue;
          }
          if (rule.type === "income")
            result[month][rule.category] += haber - debe;
          else if (rule.type === "expense")
            result[month][rule.category] += debe - haber;
          else result[month][rule.category] += haber - debe;
        }
        resolve({ data: result, unmatched });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Error leyendo el archivo"));
    reader.readAsArrayBuffer(file);
  });
}

function parseExcelQuickBooksPL(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), {
          type: "array",
          cellDates: false,
          raw: true,
        });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
          header: 1,
          defval: null,
        });
        const result = emptyMonths(),
          unmatched = {};
        const monthRe =
          /^(january|february|march|april|may|june|july|august|september|october|november|december)/i;
        let headerIdx = -1,
          monthCols = [];
        for (let i = 0; i < Math.min(20, rows.length); i++) {
          const row = rows[i];
          if (!row || !Array.isArray(row)) continue;
          const monthCount = row.filter(
            (c, ci) => ci > 0 && c && monthRe.test(String(c).trim())
          ).length;
          const isDist = /distribution.account/i.test(String(row[0] || ""));
          if (monthCount >= 1 && isDist) {
            headerIdx = i;
            break;
          }
          if (monthCount >= 2) {
            headerIdx = i;
            break;
          }
          if (monthCount === 1 && isDist) {
            headerIdx = i;
            break;
          }
        }
        if (headerIdx === -1) {
          for (let i = 0; i < Math.min(20, rows.length); i++) {
            const row = rows[i];
            if (!row || !Array.isArray(row)) continue;
            const isDist = /distribution.account/i.test(String(row[0] || ""));
            const hasMonthInNonFirstCol = row
              .slice(1)
              .some((c) => c && monthRe.test(String(c).trim()));
            if (isDist || hasMonthInNonFirstCol) {
              headerIdx = i;
              break;
            }
          }
        }
        if (headerIdx === -1)
          throw new Error(
            "No se encontró cabecera de meses en el P&L de QuickBooks"
          );
        const HR = rows[headerIdx];
        for (let c = 1; c < HR.length; c++) {
          const cs = String(HR[c] || "").trim();
          if (!cs || /^total/i.test(cs)) continue;
          const match = cs.match(monthRe);
          if (match) {
            const month = QB_MONTH_NAMES[match[1].toLowerCase()];
            if (month) monthCols.push({ col: c, month });
          }
        }
        if (!monthCols.length)
          throw new Error("No se detectaron columnas de mes en el P&L");
        for (let i = headerIdx + 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row) continue;
          const rawName = String(row[0] || "").trim();
          if (!rawName) continue;
          const key = rawName.toLowerCase().trim();
          if (QB_PL_SKIP.has(key)) continue;
          if (
            /^(total for|total |accrual|cash basis|net earnings|profit and loss)/i.test(
              rawName
            )
          )
            continue;
          const rule = QB_PL_MAP[key];
          for (const { col, month } of monthCols) {
            const raw = row[col];
            if (raw === null || raw === undefined || raw === "") continue;
            let val =
              typeof raw === "number"
                ? raw
                : parseFloat(String(raw).replace(/[$,\s=]/g, ""));
            if (isNaN(val) || val === 0) continue;
            if (!rule) {
              unmatched[rawName] = (unmatched[rawName] || 0) + Math.abs(val);
              continue;
            }
            if (rule.type === "income")
              result[month][rule.category] += Math.abs(val);
            else if (rule.type === "expense")
              result[month][rule.category] += Math.abs(val);
            else result[month][rule.category] += val;
          }
        }
        resolve({ data: result, unmatched });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Error leyendo el archivo"));
    reader.readAsArrayBuffer(file);
  });
}

// ─── CALCULATORS ─────────────────────────────────────────────────────────────
function calcIS(d) {
  const t = (k) => Object.values(d).reduce((s, m) => s + (m[k] || 0), 0);
  const gross =
    t("setup_fees") +
    t("rental_contracts") +
    t("rnt_sales") +
    t("properties_sale") +
    t("other_revenues");
  const deductions =
    t("rental_costs") + t("rnt_expenses") + t("properties_sale_cost");
  const net_rev = gross - deductions,
    opex =
      t("tech_costs") +
      t("ga_personnel") +
      t("ga_suppliers") +
      t("sales_marketing");
  const ebitda = net_rev - opex,
    fin =
      t("interests_token_holders") +
      t("fx_gain_loss") +
      t("bank_interests") +
      t("other_expenses");
  const net_profit = ebitda - fin;
  return {
    gross,
    deductions,
    net_rev,
    opex,
    ebitda,
    fin,
    net_profit,
    ...ALL_CATS.reduce((o, k) => {
      o[k] = t(k);
      return o;
    }, {}),
    gross_margin: gross > 0 ? (net_rev / gross) * 100 : 0,
    ebitda_margin: gross > 0 ? (ebitda / gross) * 100 : 0,
    net_margin: gross > 0 ? (net_profit / gross) * 100 : 0,
  };
}

const CONS_EXCLUSIONS = {
  "Rentaltoken SL": ["setup_fees"],
  "Admin RNT SL": ["setup_fees"],
  "Reental Tourist Homes": ["setup_fees"],
  "Reental Rock Capital": ["setup_fees"],
};

function buildConsolidated(empresas, fxRates, fxAvg) {
  const cons = emptyMonths();
  for (let m = 1; m <= 12; m++) {
    const rate = fxRates[m] || fxAvg;
    ALL_CATS.forEach((cat) => {
      cons[m][cat] = EMPRESAS_CONFIG.reduce((s, emp) => {
        if (CONS_EXCLUSIONS[emp.name]?.includes(cat)) return s;
        const val =
          emp.currency === "USD"
            ? empresas[emp.name]?.[m]?.[cat] || 0
            : (empresas[emp.name]?.[m]?.[cat] || 0) * rate;
        return s + val;
      }, 0);
    });
  }
  return cons;
}

function calcMonthly(d) {
  return Object.entries(d).map(([m, v]) => {
    const gross =
      (v.setup_fees || 0) +
      (v.rental_contracts || 0) +
      (v.rnt_sales || 0) +
      (v.properties_sale || 0) +
      (v.other_revenues || 0);
    const deductions =
      (v.rental_costs || 0) +
      (v.rnt_expenses || 0) +
      (v.properties_sale_cost || 0);
    const net_rev = gross - deductions,
      opex =
        (v.tech_costs || 0) +
        (v.ga_personnel || 0) +
        (v.ga_suppliers || 0) +
        (v.sales_marketing || 0);
    return {
      mes: MONTHS[parseInt(m)],
      gross,
      net_rev,
      ebitda: net_rev - opex,
      opex,
      gastos: deductions + opex,
    };
  });
}

// ─── EXPORT ───────────────────────────────────────────────────────────────────
function exportToExcel(consData, consIS, empresasData, year, fxRates, fxAvg) {
  const wb = XLSX.utils.book_new(),
    meses = MONTHS.slice(1);
  const knownMonths = Object.keys(fxRates).length;
  const hdr = {
    font: { bold: true, color: { rgb: "FFFFFF" }, name: "Arial", sz: 10 },
    fill: { fgColor: { rgb: "0D1F18" } },
    alignment: { horizontal: "center" },
  };
  const lS = {
    font: { name: "Arial", sz: 9, color: { rgb: "94A3B8" } },
    fill: { fgColor: { rgb: "0A0F1E" } },
    alignment: { indent: 1 },
  };
  const lB = {
    font: { bold: true, name: "Arial", sz: 10, color: { rgb: "F1F5F9" } },
    fill: { fgColor: { rgb: "111827" } },
  };
  const nc = (v, color, bg = "#0D1526", bold = false) => ({
    v,
    t: "n",
    z: '#,##0.00;(#,##0.00);"-"',
    s: {
      font: { bold, name: "Arial", sz: bold ? 10 : 9, color: { rgb: color } },
      fill: { fgColor: { rgb: bg } },
      alignment: { horizontal: "right" },
    },
  });
  const nG = (v) => nc(v, "00D4AA", "0D1526", true),
    nP = (v) => nc(v, "818CF8", "0D1526", true),
    nPk = (v) => nc(v, "F472B6", "0A0F1E");
  const nS = (v) => nc(v, "94A3B8", "0A0F1E"),
    nE = (v) => nc(v, "475569", "0A0F1E"),
    nEb = (v) => (v >= 0 ? nG(v) : nc(v, "EF4444", "2D0A0A", true));
  // Orange for total opex
  const nO = (v) => nc(v, "FB923C", "0D1526", true);
  const pct = (v) => ({
    v: v / 100,
    t: "n",
    z: "0.0%;(0.0%);-",
    s: {
      font: { name: "Arial", sz: 9, color: { rgb: "475569" } },
      fill: { fgColor: { rgb: "0D1526" } },
      alignment: { horizontal: "right" },
    },
  });
  const cell = (v, s) => ({ v, t: "s", s });
  const mv2 = (fn) => meses.map((_, i) => fn(consData[i + 1] || {}));
  const gv = (cat) => consIS[cat] || 0;
  const aoa = [];
  aoa.push([
    cell(`P&G Consolidado USD — Grupo Reental ${year}`, {
      font: { bold: true, name: "Arial", sz: 14, color: { rgb: "00D4AA" } },
    }),
  ]);
  aoa.push([
    cell(
      `Tipos de cambio: media mensual BCE/Banco de España ${year} · ${knownMonths} meses reales · Media parcial: ${fxAvg.toFixed(
        4
      )}`,
      { font: { name: "Arial", sz: 9, color: { rgb: "4A5568" } } }
    ),
  ]);
  aoa.push([]);
  aoa.push([
    cell("Tipo EUR/USD por mes:", {
      font: { bold: true, name: "Arial", sz: 9, color: { rgb: "374151" } },
    }),
    cell(`Media (${knownMonths}m)`, {
      font: { name: "Arial", sz: 8, color: { rgb: "374151" } },
      alignment: { horizontal: "right" },
    }),
    cell(""),
    cell(""),
    ...meses.map((_, i) =>
      cell(MONTHS[i + 1], {
        font: { name: "Arial", sz: 8, color: { rgb: "374151" } },
        alignment: { horizontal: "right" },
      })
    ),
  ]);
  aoa.push([
    cell(""),
    {
      v: fxAvg,
      t: "n",
      z: "0.0000",
      s: {
        font: { bold: true, name: "Arial", sz: 9, color: { rgb: "00D4AA" } },
        alignment: { horizontal: "right" },
      },
    },
    cell(""),
    cell(""),
    ...meses.map((_, i) => ({
      v: fxRates[i + 1] || fxAvg,
      t: "n",
      z: "0.0000",
      s: {
        font: {
          name: "Arial",
          sz: 8,
          color: { rgb: fxRates[i + 1] ? "374151" : "1e293b" },
        },
        alignment: { horizontal: "right" },
      },
    })),
  ]);
  aoa.push([]);
  aoa.push([
    cell("Línea P&G", hdr),
    cell("TOTAL USD", { ...hdr, alignment: { horizontal: "right" } }),
    cell("EQUIV EUR", { ...hdr, alignment: { horizontal: "right" } }),
    cell("% Rev", { ...hdr, alignment: { horizontal: "right" } }),
    ...meses.map((mo) =>
      cell(mo, { ...hdr, alignment: { horizontal: "right" } })
    ),
  ]);
  aoa.push([
    cell("GROSS REVENUES", {
      font: { bold: true, name: "Arial", sz: 10, color: { rgb: "00D4AA" } },
      fill: { fgColor: { rgb: "0D1F18" } },
    }),
    nG(consIS.gross),
    nE(consIS.gross / fxAvg),
    cell(""),
    ...mv2(
      (d) =>
        (d.setup_fees || 0) +
        (d.rental_contracts || 0) +
        (d.rnt_sales || 0) +
        (d.properties_sale || 0) +
        (d.other_revenues || 0)
    ).map(nG),
  ]);
  [
    ["  Set up fees / Tokenization fees", "setup_fees"],
    ["  Rental contracts", "rental_contracts"],
    ["  RNT Sales", "rnt_sales"],
    ["  Properties Sale Gain", "properties_sale"],
    ["  Other revenues", "other_revenues"],
  ].forEach(([lbl, cat]) =>
    aoa.push([
      cell(lbl, lS),
      nS(gv(cat)),
      nE(gv(cat) / fxAvg),
      pct(consIS.gross > 0 ? (gv(cat) / consIS.gross) * 100 : 0),
      ...meses.map((_, i) => nS(consData[i + 1]?.[cat] || 0)),
    ])
  );
  aoa.push([]);
  aoa.push([
    cell("DEDUCTIONS", {
      font: { bold: true, name: "Arial", sz: 10, color: { rgb: "F472B6" } },
      fill: { fgColor: { rgb: "1A0D14" } },
    }),
    nPk(-consIS.deductions),
    nE(-consIS.deductions / fxAvg),
    cell(""),
    ...mv2(
      (d) =>
        -(
          (d.rental_costs || 0) +
          (d.rnt_expenses || 0) +
          (d.properties_sale_cost || 0)
        )
    ).map(nPk),
  ]);
  [
    ["  Rental costs", "rental_costs"],
    ["  RNT expenses", "rnt_expenses"],
    ["  Properties Sale costs", "properties_sale_cost"],
  ].forEach(([lbl, cat]) =>
    aoa.push([
      cell(lbl, lS),
      nS(-gv(cat)),
      nE(-gv(cat) / fxAvg),
      cell(""),
      ...meses.map((_, i) => nS(-(consData[i + 1]?.[cat] || 0))),
    ])
  );
  aoa.push([]);
  aoa.push([
    cell("NET REVENUES", {
      font: { bold: true, name: "Arial", sz: 10, color: { rgb: "818CF8" } },
      fill: { fgColor: { rgb: "0D0E20" } },
    }),
    nP(consIS.net_rev),
    nE(consIS.net_rev / fxAvg),
    pct(consIS.gross_margin),
    ...mv2((d) => {
      const g =
        (d.setup_fees || 0) +
        (d.rental_contracts || 0) +
        (d.rnt_sales || 0) +
        (d.properties_sale || 0) +
        (d.other_revenues || 0);
      return (
        g -
        (d.rental_costs || 0) -
        (d.rnt_expenses || 0) -
        (d.properties_sale_cost || 0)
      );
    }).map(nP),
  ]);
  aoa.push([]);
  [
    ["  Tech costs", "tech_costs"],
    ["  G&A Personnel", "ga_personnel"],
    ["  G&A Suppliers", "ga_suppliers"],
    ["  Sales & Marketing", "sales_marketing"],
  ].forEach(([lbl, cat]) =>
    aoa.push([
      cell(lbl, lS),
      nS(-gv(cat)),
      nE(-gv(cat) / fxAvg),
      cell(""),
      ...meses.map((_, i) => nS(-(consData[i + 1]?.[cat] || 0))),
    ])
  );
  aoa.push([]);
  // TOTAL OPEX row
  aoa.push([
    cell("TOTAL OPEX", {
      font: { bold: true, name: "Arial", sz: 10, color: { rgb: "FB923C" } },
      fill: { fgColor: { rgb: "1A1000" } },
    }),
    nO(-consIS.opex),
    nE(-consIS.opex / fxAvg),
    pct(consIS.gross > 0 ? (consIS.opex / consIS.gross) * 100 : 0),
    ...mv2(
      (d) =>
        -(
          (d.tech_costs || 0) +
          (d.ga_personnel || 0) +
          (d.ga_suppliers || 0) +
          (d.sales_marketing || 0)
        )
    ).map(nO),
  ]);
  aoa.push([]);
  aoa.push([
    cell("EBITDA", {
      font: {
        bold: true,
        name: "Arial",
        sz: 11,
        color: { rgb: consIS.ebitda >= 0 ? "00D4AA" : "EF4444" },
      },
      fill: { fgColor: { rgb: consIS.ebitda >= 0 ? "0D2018" : "2D0A0A" } },
    }),
    nEb(consIS.ebitda),
    nE(consIS.ebitda / fxAvg),
    pct(consIS.ebitda_margin),
    ...mv2((d) => {
      const g =
          (d.setup_fees || 0) +
          (d.rental_contracts || 0) +
          (d.rnt_sales || 0) +
          (d.properties_sale || 0) +
          (d.other_revenues || 0),
        ded =
          (d.rental_costs || 0) +
          (d.rnt_expenses || 0) +
          (d.properties_sale_cost || 0),
        opex =
          (d.tech_costs || 0) +
          (d.ga_personnel || 0) +
          (d.ga_suppliers || 0) +
          (d.sales_marketing || 0);
      return g - ded - opex;
    }).map(nEb),
  ]);
  aoa.push([]);
  [
    ["  Interests TH", "interests_token_holders"],
    ["  FX Gain/Loss", "fx_gain_loss"],
    ["  Bank interests", "bank_interests"],
    ["  Other expenses", "other_expenses"],
  ].forEach(([lbl, cat]) =>
    aoa.push([
      cell(lbl, lS),
      nS(-gv(cat)),
      nE(-gv(cat) / fxAvg),
      cell(""),
      ...meses.map((_, i) => nS(-(consData[i + 1]?.[cat] || 0))),
    ])
  );
  aoa.push([]);
  aoa.push([
    cell("NET PROFIT / LOSS", {
      font: {
        bold: true,
        name: "Arial",
        sz: 11,
        color: { rgb: consIS.net_profit >= 0 ? "00D4AA" : "EF4444" },
      },
      fill: { fgColor: { rgb: consIS.net_profit >= 0 ? "0D2018" : "2D0A0A" } },
    }),
    nEb(consIS.net_profit),
    nE(consIS.net_profit / fxAvg),
    pct(consIS.net_margin),
    ...mv2((d) => {
      const g =
          (d.setup_fees || 0) +
          (d.rental_contracts || 0) +
          (d.rnt_sales || 0) +
          (d.properties_sale || 0) +
          (d.other_revenues || 0),
        ded =
          (d.rental_costs || 0) +
          (d.rnt_expenses || 0) +
          (d.properties_sale_cost || 0),
        opex =
          (d.tech_costs || 0) +
          (d.ga_personnel || 0) +
          (d.ga_suppliers || 0) +
          (d.sales_marketing || 0),
        fin =
          (d.interests_token_holders || 0) +
          (d.fx_gain_loss || 0) +
          (d.bank_interests || 0) +
          (d.other_expenses || 0);
      return g - ded - opex - fin;
    }).map(nEb),
  ]);
  const ws1 = XLSX.utils.aoa_to_sheet(aoa);
  ws1["!cols"] = [
    { wch: 32 },
    { wch: 16 },
    { wch: 14 },
    { wch: 8 },
    ...meses.map(() => ({ wch: 11 })),
  ];
  XLSX.utils.book_append_sheet(wb, ws1, "Consolidado USD");

  const aoa2 = [];
  aoa2.push([
    cell(`Tipos EUR/USD aplicados — ${year} — BCE/Banco de España`, {
      font: { bold: true, name: "Arial", sz: 13, color: { rgb: "00D4AA" } },
    }),
  ]);
  aoa2.push([
    cell(
      `Serie: EXR.M.USD.EUR.SP00.A · ${knownMonths} meses con dato real publicado`,
      { font: { name: "Arial", sz: 9, color: { rgb: "4A5568" } } }
    ),
  ]);
  aoa2.push([]);
  aoa2.push([
    cell("Mes", hdr),
    cell("EUR/USD (media mensual)", {
      ...hdr,
      alignment: { horizontal: "right" },
    }),
    cell("Estado", { ...hdr, alignment: { horizontal: "center" } }),
  ]);
  for (let m = 1; m <= 12; m++)
    aoa2.push([
      cell(MONTHS[m], lS),
      fxRates[m]
        ? {
            v: fxRates[m],
            t: "n",
            z: "0.0000",
            s: {
              font: {
                bold: true,
                name: "Arial",
                sz: 10,
                color: { rgb: "00D4AA" },
              },
              fill: { fgColor: { rgb: "0D1526" } },
              alignment: { horizontal: "right" },
            },
          }
        : {
            v: fxAvg,
            t: "n",
            z: "0.0000",
            s: {
              font: { name: "Arial", sz: 10, color: { rgb: "374151" } },
              fill: { fgColor: { rgb: "0D1526" } },
              alignment: { horizontal: "right" },
            },
          },
      cell(fxRates[m] ? "✓ Real BCE" : "~ Avg aplicado", {
        font: {
          name: "Arial",
          sz: 8,
          color: { rgb: fxRates[m] ? "00D4AA" : "374151" },
        },
        fill: { fgColor: { rgb: "0D1526" } },
        alignment: { horizontal: "center" },
      }),
    ]);
  aoa2.push([]);
  aoa2.push([
    cell(`Media parcial (${knownMonths}m reales)`, lB),
    {
      v: fxAvg,
      t: "n",
      z: "0.0000",
      s: {
        font: { bold: true, name: "Arial", sz: 11, color: { rgb: "00D4AA" } },
        fill: { fgColor: { rgb: "0D1526" } },
        alignment: { horizontal: "right" },
      },
    },
    cell("BCE oficial", {
      font: { name: "Arial", sz: 8, color: { rgb: "374151" } },
      fill: { fgColor: { rgb: "0D1526" } },
      alignment: { horizontal: "center" },
    }),
  ]);
  const ws2 = XLSX.utils.aoa_to_sheet(aoa2);
  ws2["!cols"] = [{ wch: 12 }, { wch: 22 }, { wch: 16 }];
  XLSX.utils.book_append_sheet(wb, ws2, "Tipos BCE 2026");
  XLSX.writeFile(
    wb,
    `PYG_Reental_${year}_USD_BCE_${new Date().toISOString().slice(0, 10)}.xlsx`
  );
}

// ─── UTILS ────────────────────────────────────────────────────────────────────
const fp = (n) => n.toFixed(1) + "%";
const fmtU = (v, compact = false) => {
  if (compact) {
    const a = Math.abs(v);
    const s =
      a >= 1e6
        ? (a / 1e6).toFixed(2) + "M"
        : a >= 1000
        ? (a / 1e3).toFixed(1) + "K"
        : a.toFixed(0);
    return (v < 0 ? "-$" : "$") + s;
  }
  return (
    (v < 0 ? "-$" : "$") +
    Math.abs(v).toLocaleString("es-ES", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })
  );
};
const fmtE = (v, rate) => {
  const e = v / rate;
  const a = Math.abs(e);
  const s =
    a >= 1e6
      ? (a / 1e6).toFixed(1) + "M"
      : a >= 1000
      ? (a / 1e3).toFixed(1) + "K"
      : a.toFixed(0);
  return (e < 0 ? "-€" : "€") + s;
};
const fmtN = (v, sym) =>
  (v < 0 ? `-${sym}` : `${sym}`) +
  Math.abs(v).toLocaleString("es-ES", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
const fC = (v) => {
  if (v === 0) return "—";
  const a = Math.abs(v);
  const s =
    a >= 1e6
      ? (a / 1e6).toFixed(1) + "M"
      : a >= 1000
      ? (a / 1e3).toFixed(1) + "K"
      : a.toFixed(0);
  return (v < 0 ? "-$" : "$") + s;
};
const tt = {
  background: "#0d1526",
  border: "1px solid #1e293b",
  borderRadius: 8,
  fontSize: 13,
  color: "#e2e8f0",
};

// ─── FX MODAL ────────────────────────────────────────────────────────────────
const FxModal = ({ fxRates, fxAvg, refreshing, onRefresh, onClose }) => {
  const knownMonths = Object.keys(fxRates).length;
  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        background: "#0a0f1ecc",
        zIndex: 100,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      }}
      onClick={onClose}
    >
      <div
        style={{
          background: "#111827",
          border: "1px solid #1e293b",
          borderRadius: 14,
          padding: 28,
          minWidth: 400,
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "flex-start",
            marginBottom: 18,
          }}
        >
          <div>
            <div
              style={{ fontWeight: 700, color: "#00d4aa", fontSize: "1.05em" }}
            >
              Tipos EUR/USD — BCE {YEAR}
            </div>
            <div style={{ fontSize: "0.72em", color: "#4a5568", marginTop: 2 }}>
              Media mensual · {knownMonths} meses reales publicados · Resto usa
              avg parcial
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <button
              onClick={onRefresh}
              disabled={refreshing}
              style={{
                background: refreshing ? "#111827" : "#00d4aa22",
                border: "1px solid #00d4aa44",
                color: refreshing ? "#4a5568" : "#00d4aa",
                borderRadius: 6,
                padding: "5px 10px",
                fontSize: "0.75em",
                cursor: refreshing ? "default" : "pointer",
                fontFamily: "monospace",
              }}
            >
              {refreshing ? "⏳ Actualizando…" : "↻ Actualizar vía AI"}
            </button>
            <button
              onClick={onClose}
              style={{
                background: "transparent",
                border: "none",
                color: "#4a5568",
                fontSize: 20,
                cursor: "pointer",
              }}
            >
              ✕
            </button>
          </div>
        </div>
        <table
          style={{
            width: "100%",
            borderCollapse: "collapse",
            fontSize: "0.9em",
          }}
        >
          <thead>
            <tr
              style={{
                borderBottom: "1px solid #1e293b",
                color: "#4a5568",
                fontSize: "0.78em",
              }}
            >
              <th
                style={{
                  padding: "6px 10px",
                  textAlign: "left",
                  fontWeight: 500,
                }}
              >
                Mes
              </th>
              <th
                style={{
                  padding: "6px 10px",
                  textAlign: "right",
                  fontWeight: 500,
                }}
              >
                EUR/USD
              </th>
              <th
                style={{
                  padding: "6px 10px",
                  textAlign: "center",
                  fontWeight: 500,
                }}
              >
                Fuente
              </th>
              <th
                style={{
                  padding: "6px 10px",
                  textAlign: "center",
                  fontWeight: 500,
                }}
              >
                vs avg
              </th>
            </tr>
          </thead>
          <tbody>
            {MONTHS.slice(1).map((mes, i) => {
              const m = i + 1,
                isReal = !!fxRates[m],
                rate = fxRates[m] || fxAvg;
              const diff = ((rate - fxAvg) / fxAvg) * 100;
              return (
                <tr
                  key={m}
                  style={{
                    borderBottom: "1px solid #0d1526",
                    opacity: isReal ? 1 : 0.45,
                  }}
                >
                  <td
                    style={{
                      padding: "7px 10px",
                      color: isReal ? "#f1f5f9" : "#4a5568",
                      fontWeight: 600,
                    }}
                  >
                    {mes}
                  </td>
                  <td
                    style={{
                      padding: "7px 10px",
                      textAlign: "right",
                      color: isReal ? "#00d4aa" : "#374151",
                      fontWeight: 700,
                      fontFamily: "monospace",
                    }}
                  >
                    {rate.toFixed(4)}
                  </td>
                  <td
                    style={{
                      padding: "7px 10px",
                      textAlign: "center",
                      fontSize: "0.72em",
                      color: isReal ? "#00d4aa" : "#374151",
                    }}
                  >
                    {isReal ? "✓ BCE real" : "~ avg parcial"}
                  </td>
                  <td
                    style={{
                      padding: "7px 10px",
                      textAlign: "center",
                      fontSize: "0.78em",
                      color: diff >= 0 ? "#00d4aa" : "#f472b6",
                    }}
                  >
                    {isReal
                      ? (diff >= 0 ? "+" : "") + diff.toFixed(1) + "%"
                      : "—"}
                  </td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr style={{ borderTop: "2px solid #1e293b" }}>
              <td
                style={{
                  padding: "8px 10px",
                  color: "#f1f5f9",
                  fontWeight: 700,
                }}
              >
                Media parcial
              </td>
              <td
                style={{
                  padding: "8px 10px",
                  textAlign: "right",
                  color: "#00d4aa",
                  fontWeight: 800,
                  fontFamily: "monospace",
                }}
              >
                {fxAvg.toFixed(4)}
              </td>
              <td
                style={{
                  padding: "8px 10px",
                  textAlign: "center",
                  fontSize: "0.72em",
                  color: "#374151",
                }}
              >
                {knownMonths}m reales
              </td>
              <td />
            </tr>
          </tfoot>
        </table>
        <div
          style={{
            marginTop: 12,
            fontSize: "0.7em",
            color: "#374151",
            textAlign: "center",
          }}
        >
          Fuente: BCE EXR.M.USD.EUR.SP00.A · Actualizable vía AI cuando BdE
          publique nuevos meses
        </div>
      </div>
    </div>
  );
};

// ─── IS TABLE ─────────────────────────────────────────────────────────────────
const ISTable = ({
  is,
  monthlyData,
  showNative,
  fxRates,
  fxAvg,
  entityCurrency,
}) => {
  const mv = (field, sign = 1) =>
    Array.from(
      { length: 12 },
      (_, i) => (monthlyData?.[i + 1]?.[field] || 0) * sign
    );
  const mc = (fn) =>
    Array.from({ length: 12 }, (_, i) => fn(monthlyData?.[i + 1] || {}));
  const mh = MONTHS.slice(1);
  const isEUR = entityCurrency === "EUR";
  const sym = isEUR ? "€" : "$";

  const ThRow = ({ label, value, color, vals, showPct, pct }) => (
    <tr style={{ borderBottom: "1px solid #0d1526", background: color + "11" }}>
      <td style={{ padding: "9px 14px", color, fontWeight: 700 }}>{label}</td>
      <td
        style={{
          padding: "9px 14px",
          textAlign: "right",
          color,
          fontWeight: 700,
        }}
      >
        {fmtU(isEUR ? value * fxAvg : value)}
      </td>
      {showNative && (
        <td
          style={{
            padding: "9px 14px",
            textAlign: "right",
            color: "#374151",
            fontSize: "0.82em",
          }}
        >
          {fmtN(value, sym)}
        </td>
      )}
      <td
        style={{
          padding: "9px 14px",
          textAlign: "right",
          color: "#475569",
          fontSize: "0.82em",
        }}
      >
        {showPct ? pct.toFixed(1) + "%" : ""}
      </td>
      {vals.map((v, i) => (
        <td
          key={i}
          style={{
            padding: "6px 10px",
            textAlign: "right",
            fontSize: "0.82em",
            color: v === 0 ? "#1e293b" : v > 0 ? color : "#ef4444",
            fontWeight: 700,
            whiteSpace: "nowrap",
          }}
        >
          {fC(isEUR ? v * (fxRates[i + 1] || fxAvg) : v)}
        </td>
      ))}
    </tr>
  );

  // ── NEW: Total OPEX subtotal row (orange, shown before EBITDA)
  const TotalOpexRow = () => {
    const totalOpex = is.opex || 0;
    const monthlyVals = mc(
      (d) =>
        -(
          (d.tech_costs || 0) +
          (d.ga_personnel || 0) +
          (d.ga_suppliers || 0) +
          (d.sales_marketing || 0)
        )
    );
    const color = "#fb923c";
    const pctVal =
      is.gross > 0 ? ((totalOpex / is.gross) * 100).toFixed(1) + "%" : "";
    return (
      <tr
        style={{ borderBottom: "2px solid #fb923c44", background: "#fb923c0d" }}
      >
        <td
          style={{
            padding: "10px 14px",
            color,
            fontWeight: 800,
            fontSize: "0.95em",
            letterSpacing: "0.04em",
          }}
        >
          <span style={{ borderLeft: `3px solid ${color}`, paddingLeft: 8 }}>
            TOTAL OPEX
          </span>
        </td>
        <td
          style={{
            padding: "10px 14px",
            textAlign: "right",
            color,
            fontWeight: 800,
          }}
        >
          {fmtU(isEUR ? -totalOpex * fxAvg : -totalOpex)}
        </td>
        {showNative && (
          <td
            style={{
              padding: "10px 14px",
              textAlign: "right",
              color: "#374151",
              fontSize: "0.82em",
            }}
          >
            {fmtN(-totalOpex, sym)}
          </td>
        )}
        <td
          style={{
            padding: "10px 14px",
            textAlign: "right",
            color: "#fb923c99",
            fontSize: "0.82em",
            fontWeight: 600,
          }}
        >
          {pctVal}
        </td>
        {monthlyVals.map((v, i) => (
          <td
            key={i}
            style={{
              padding: "6px 10px",
              textAlign: "right",
              fontSize: "0.82em",
              color: v === 0 ? "#1e293b" : "#fb923c",
              fontWeight: 700,
              whiteSpace: "nowrap",
            }}
          >
            {fC(isEUR ? v * (fxRates[i + 1] || fxAvg) : v)}
          </td>
        ))}
      </tr>
    );
  };

  const SubRow = ({ label, field, sign = 1 }) => (
    <tr style={{ borderBottom: "1px solid #0d1526" }}>
      <td
        style={{
          padding: "7px 14px 7px 32px",
          color: "#94a3b8",
          fontSize: "0.9em",
        }}
      >
        {label}
      </td>
      <td
        style={{
          padding: "7px 14px",
          textAlign: "right",
          color: "#94a3b8",
          fontSize: "0.9em",
        }}
      >
        {fmtU(isEUR ? is[field] * sign * fxAvg : is[field] * sign)}
      </td>
      {showNative && (
        <td
          style={{
            padding: "7px 14px",
            textAlign: "right",
            color: "#374151",
            fontSize: "0.78em",
          }}
        >
          {fmtN(is[field] * sign, sym)}
        </td>
      )}
      <td
        style={{
          padding: "7px 14px",
          textAlign: "right",
          color: "#475569",
          fontSize: "0.78em",
        }}
      >
        {is.gross > 0
          ? ((is[field] / is.gross) * 100 * sign).toFixed(1) + "%"
          : ""}
      </td>
      {mv(field, sign).map((v, i) => (
        <td
          key={i}
          style={{
            padding: "6px 10px",
            textAlign: "right",
            fontSize: "0.82em",
            color: v === 0 ? "#1e293b" : v > 0 ? "#94a3b8" : "#ef4444",
            whiteSpace: "nowrap",
          }}
        >
          {fC(isEUR ? v * (fxRates[i + 1] || fxAvg) : v)}
        </td>
      ))}
    </tr>
  );
  const Sep = () => (
    <tr>
      <td
        colSpan={99}
        style={{ borderBottom: "1px solid #1e293b", padding: "3px 0" }}
      />
    </tr>
  );

  return (
    <div style={{ overflowX: "auto" }}>
      <table
        style={{ width: "100%", borderCollapse: "collapse", minWidth: 1100 }}
      >
        <thead>
          <tr
            style={{
              color: "#4a5568",
              fontSize: "0.78em",
              textTransform: "uppercase",
              borderBottom: "2px solid #1e293b",
            }}
          >
            <th
              style={{
                padding: "10px 14px",
                textAlign: "left",
                fontWeight: 500,
                minWidth: 220,
              }}
            >
              Línea P&G
            </th>
            <th
              style={{
                padding: "10px 14px",
                textAlign: "right",
                fontWeight: 500,
                minWidth: 120,
              }}
            >
              USD
            </th>
            {showNative && (
              <th
                style={{
                  padding: "10px 14px",
                  textAlign: "right",
                  fontWeight: 500,
                  minWidth: 100,
                }}
              >
                {entityCurrency || "Orig."}
              </th>
            )}
            <th
              style={{
                padding: "10px 14px",
                textAlign: "right",
                fontWeight: 500,
                minWidth: 60,
              }}
            >
              % Rev
            </th>
            {mh.map((m, i) => (
              <th
                key={m}
                style={{
                  padding: "10px 10px",
                  textAlign: "right",
                  fontWeight: 500,
                  minWidth: 80,
                }}
              >
                <div style={{ color: fxRates[i + 1] ? "#374151" : "#1e293b" }}>
                  {m}
                </div>
                {isEUR && (
                  <div
                    style={{
                      color: fxRates[i + 1] ? "#1e293b" : "#111827",
                      fontSize: "0.72em",
                      fontWeight: 400,
                      fontFamily: "monospace",
                    }}
                  >
                    {(fxRates[i + 1] || fxAvg).toFixed(3)}
                  </div>
                )}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          <ThRow
            label="GROSS REVENUES"
            value={is.gross}
            color="#00d4aa"
            vals={mc(
              (d) =>
                (d.setup_fees || 0) +
                (d.rental_contracts || 0) +
                (d.rnt_sales || 0) +
                (d.properties_sale || 0) +
                (d.other_revenues || 0)
            )}
          />
          <SubRow label="Set up fees / Tokenization fees" field="setup_fees" />
          <SubRow label="Rental contracts" field="rental_contracts" />
          <SubRow label="RNT Sales" field="rnt_sales" />
          <SubRow label="Properties Sale Gain" field="properties_sale" />
          <SubRow label="Other revenues" field="other_revenues" />
          <Sep />
          <ThRow
            label="DEDUCTIONS & DIRECT COSTS"
            value={-is.deductions}
            color="#f472b6"
            vals={mc(
              (d) =>
                -(
                  (d.rental_costs || 0) +
                  (d.rnt_expenses || 0) +
                  (d.properties_sale_cost || 0)
                )
            )}
          />
          <SubRow label="Rental costs" field="rental_costs" sign={-1} />
          <SubRow label="RNT expenses" field="rnt_expenses" sign={-1} />
          <SubRow
            label="Properties Sale costs"
            field="properties_sale_cost"
            sign={-1}
          />
          <Sep />
          <ThRow
            label="NET REVENUES"
            value={is.net_rev}
            color="#818cf8"
            showPct
            pct={is.gross_margin}
            vals={mc((d) => {
              const g =
                (d.setup_fees || 0) +
                (d.rental_contracts || 0) +
                (d.rnt_sales || 0) +
                (d.properties_sale || 0) +
                (d.other_revenues || 0);
              return (
                g -
                (d.rental_costs || 0) -
                (d.rnt_expenses || 0) -
                (d.properties_sale_cost || 0)
              );
            })}
          />
          <Sep />
          <SubRow label="Tech costs" field="tech_costs" sign={-1} />
          <SubRow label="G&A Personnel" field="ga_personnel" sign={-1} />
          <SubRow label="G&A Suppliers" field="ga_suppliers" sign={-1} />
          <SubRow label="Sales & Marketing" field="sales_marketing" sign={-1} />
          <Sep />
          {/* ── TOTAL OPEX ROW — NEW ── */}
          <TotalOpexRow />
          <Sep />
          <ThRow
            label="EBITDA"
            value={is.ebitda}
            color={is.ebitda >= 0 ? "#00d4aa" : "#ef4444"}
            showPct
            pct={is.ebitda_margin}
            vals={mc((d) => {
              const g =
                  (d.setup_fees || 0) +
                  (d.rental_contracts || 0) +
                  (d.rnt_sales || 0) +
                  (d.properties_sale || 0) +
                  (d.other_revenues || 0),
                ded =
                  (d.rental_costs || 0) +
                  (d.rnt_expenses || 0) +
                  (d.properties_sale_cost || 0),
                opex =
                  (d.tech_costs || 0) +
                  (d.ga_personnel || 0) +
                  (d.ga_suppliers || 0) +
                  (d.sales_marketing || 0);
              return g - ded - opex;
            })}
          />
          <Sep />
          <SubRow
            label="Interests token holders"
            field="interests_token_holders"
            sign={-1}
          />
          <SubRow label="FX Gain / Loss" field="fx_gain_loss" sign={-1} />
          <SubRow label="Bank interests" field="bank_interests" sign={-1} />
          <SubRow label="Other expenses" field="other_expenses" sign={-1} />
          <Sep />
          <ThRow
            label="NET PROFIT / LOSS"
            value={is.net_profit}
            color={is.net_profit >= 0 ? "#00d4aa" : "#ef4444"}
            showPct
            pct={is.net_margin}
            vals={mc((d) => {
              const g =
                  (d.setup_fees || 0) +
                  (d.rental_contracts || 0) +
                  (d.rnt_sales || 0) +
                  (d.properties_sale || 0) +
                  (d.other_revenues || 0),
                ded =
                  (d.rental_costs || 0) +
                  (d.rnt_expenses || 0) +
                  (d.properties_sale_cost || 0),
                opex =
                  (d.tech_costs || 0) +
                  (d.ga_personnel || 0) +
                  (d.ga_suppliers || 0) +
                  (d.sales_marketing || 0),
                fin =
                  (d.interests_token_holders || 0) +
                  (d.fx_gain_loss || 0) +
                  (d.bank_interests || 0) +
                  (d.other_expenses || 0);
              return g - ded - opex - fin;
            })}
          />
        </tbody>
      </table>
    </div>
  );
};

// ─── FILE UPLOADER ────────────────────────────────────────────────────────────
const FileUploader = ({ onUpload, onClear, loadedFiles }) => {
  const [processing, setProcessing] = useState(null);
  const handleFile = useCallback(
    async (file, emp) => {
      if (!file) return;
      setProcessing(emp.name);
      try {
        const result =
          emp.parser === "quickbooks"
            ? await parseExcelQuickBooksPL(file)
            : await parseExcelHolded(file);
        onUpload(emp.name, result.data, result.unmatched);
      } catch (err) {
        alert(`Error procesando ${file.name}: ${err.message}`);
      } finally {
        setProcessing(null);
      }
    },
    [onUpload]
  );
  return (
    <div
      style={{
        background: "#111827",
        borderRadius: 14,
        border: "1px solid #1e293b",
        padding: 24,
        marginBottom: 20,
      }}
    >
      <div
        style={{
          fontSize: "0.75em",
          color: "#4a5568",
          textTransform: "uppercase",
          letterSpacing: "0.12em",
          marginBottom: 16,
        }}
      >
        📂 Cargar libros contables
      </div>
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit,minmax(250px,1fr))",
          gap: 12,
        }}
      >
        {EMPRESAS_CONFIG.map((emp) => {
          const loaded = !!loadedFiles[emp.name],
            inputId = `fi-${emp.name.replace(/\s+/g, "-")}`;
          return (
            <div key={emp.name}>
              <div
                onClick={() => document.getElementById(inputId).click()}
                style={{
                  background: loaded ? "#0d2018" : "#0a0f1e",
                  border: `2px dashed ${loaded ? emp.color : "#1e293b"}`,
                  borderRadius: 10,
                  padding: "14px 18px",
                  opacity: processing === emp.name ? 0.6 : 1,
                  cursor: "pointer",
                  userSelect: "none",
                }}
              >
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    marginBottom: 6,
                  }}
                >
                  <span
                    style={{
                      fontWeight: 700,
                      color: loaded ? emp.color : "#4a5568",
                      fontSize: "0.88em",
                    }}
                  >
                    {emp.name}
                  </span>
                  <div style={{ display: "flex", gap: 4 }}>
                    <span
                      style={{
                        fontSize: "0.65em",
                        color: "#374151",
                        background: "#1e293b",
                        padding: "1px 5px",
                        borderRadius: 3,
                      }}
                    >
                      {emp.currency}
                    </span>
                    <span
                      style={{
                        fontSize: "0.65em",
                        color: "#374151",
                        background: "#1e293b",
                        padding: "1px 5px",
                        borderRadius: 3,
                      }}
                    >
                      {emp.parser === "quickbooks" ? "QB" : "Holded"}
                    </span>
                    {loaded && <span style={{ color: emp.color }}>✓</span>}
                  </div>
                </div>
                <div style={{ fontSize: "0.72em", color: "#374151" }}>
                  {processing === emp.name
                    ? "⏳ Procesando..."
                    : emp.parser === "quickbooks"
                    ? "Reports → P&L → Export Excel"
                    : "Contabilidad → Libro diario → Excel"}
                </div>
              </div>
              <input
                id={inputId}
                type="file"
                accept=".xlsx,.xls"
                style={{ display: "none" }}
                onClick={(e) => {
                  e.target.value = null;
                }}
                onChange={(e) => handleFile(e.target.files[0], emp)}
              />
            </div>
          );
        })}
      </div>
      <div
        style={{ marginTop: 10, display: "flex", justifyContent: "flex-end" }}
      >
        <button
          onClick={onClear}
          style={{
            background: "transparent",
            border: "1px solid #1e293b",
            color: "#374151",
            borderRadius: 6,
            padding: "5px 12px",
            fontSize: "0.75em",
            cursor: "pointer",
            fontFamily: "monospace",
          }}
        >
          Limpiar todo
        </button>
      </div>
    </div>
  );
};

const KPI = ({ label, value, secondary, sub, color }) => (
  <div
    style={{
      background: "#111827",
      borderRadius: 12,
      padding: "20px 22px",
      borderLeft: `3px solid ${color}`,
      flex: 1,
      minWidth: 160,
    }}
  >
    <div
      style={{
        fontSize: "0.72em",
        color: "#4a5568",
        textTransform: "uppercase",
        letterSpacing: "0.1em",
        marginBottom: 8,
      }}
    >
      {label}
    </div>
    <div
      style={{
        fontSize: "1.9em",
        fontWeight: 800,
        color,
        letterSpacing: "-0.02em",
      }}
    >
      {value}
    </div>
    {secondary && (
      <div style={{ fontSize: "0.82em", color: "#374151", marginTop: 4 }}>
        {secondary}
      </div>
    )}
    {sub && (
      <div style={{ fontSize: "0.78em", color: "#374151", marginTop: 4 }}>
        {sub}
      </div>
    )}
  </div>
);

// ─── APP ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState("consolidado");
  const [empTab, setEmpTab] = useState(EMPRESAS_CONFIG[0].name);
  const [showNative, setShowNative] = useState(true);
  const [showUnmatched, setShowUnmatched] = useState(false);
  const [showFxModal, setShowFxModal] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const [fxRates, setFxRates] = useState({ ...BCE_RATES_2026 });
  const [fxSource, setFxSource] = useState(
    `BCE ${YEAR} · ${
      Object.keys(BCE_RATES_2026).length
    } meses reales · avg parcial`
  );
  const [empresasData, setEmpresasData] = useState(
    Object.fromEntries(EMPRESAS_CONFIG.map((e) => [e.name, emptyMonths()]))
  );
  const [loadedFiles, setLoadedFiles] = useState({});
  const [unmatchedAccounts, setUnmatchedAccounts] = useState({});

  const fxAvg = calcAvg(fxRates);
  const knownMonths = Object.keys(fxRates).length;

  const handleRefreshRates = async () => {
    setRefreshing(true);
    try {
      const rates = await refreshRatesViaAI(YEAR);
      setFxRates((prev) => ({ ...prev, ...rates }));
      const newCount = Object.keys({ ...fxRates, ...rates }).length;
      setFxSource(
        `BCE ${YEAR} · ${newCount} meses reales · actualizado ${new Date().toLocaleDateString(
          "es-ES"
        )}`
      );
    } catch (e) {
      alert(`No se pudo actualizar: ${e.message}`);
    } finally {
      setRefreshing(false);
    }
  };

  const handleUpload = useCallback((name, data, unmatched) => {
    setEmpresasData((p) => ({ ...p, [name]: data }));
    setLoadedFiles((p) => ({ ...p, [name]: true }));
    if (Object.keys(unmatched).length > 0)
      setUnmatchedAccounts((p) => ({ ...p, [name]: unmatched }));
  }, []);

  const handleClear = () => {
    setEmpresasData(
      Object.fromEntries(EMPRESAS_CONFIG.map((e) => [e.name, emptyMonths()]))
    );
    setLoadedFiles({});
    setUnmatchedAccounts({});
  };

  const consData = buildConsolidated(empresasData, fxRates, fxAvg);
  const consIS = calcIS(consData);
  const consMonthly = calcMonthly(consData);
  const totalLoaded = Object.keys(loadedFiles).length;
  const empConfig =
    EMPRESAS_CONFIG.find((e) => e.name === empTab) || EMPRESAS_CONFIG[0];
  const empNativeData = empresasData[empTab] || emptyMonths();
  const empIS = calcIS(empNativeData);
  const empMonthly = calcMonthly(empNativeData);
  const empSym = empConfig.currency === "EUR" ? "€" : "$";
  const allUnmatched = Object.entries(unmatchedAccounts)
    .flatMap(([emp, acc]) =>
      Object.entries(acc).map(([k, v]) => ({ emp, acc: k, total: v }))
    )
    .sort((a, b) => b.total - a.total);

  const Tab = ({ id, label }) => (
    <button
      onClick={() => setTab(id)}
      style={{
        padding: "9px 22px",
        borderRadius: 8,
        border: "none",
        cursor: "pointer",
        fontFamily: "monospace",
        fontSize: "0.9em",
        fontWeight: 700,
        background: tab === id ? "#00d4aa" : "#1a2234",
        color: tab === id ? "#0a0f1e" : "#4a5568",
      }}
    >
      {label}
    </button>
  );

  return (
    <div
      style={{
        minHeight: "100vh",
        background: "#0a0f1e",
        color: "#e2e8f0",
        fontFamily: "monospace",
      }}
    >
      <style>{`*{box-sizing:border-box;margin:0;padding:0}::-webkit-scrollbar{width:4px}::-webkit-scrollbar-thumb{background:#1e293b;border-radius:4px}.hr:hover td{background:#ffffff05}`}</style>
      {showFxModal && (
        <FxModal
          fxRates={fxRates}
          fxAvg={fxAvg}
          refreshing={refreshing}
          onRefresh={handleRefreshRates}
          onClose={() => setShowFxModal(false)}
        />
      )}

      {/* HEADER */}
      <div
        style={{
          borderBottom: "1px solid #111827",
          padding: "14px 28px",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          position: "sticky",
          top: 0,
          background: "#0a0f1e",
          zIndex: 10,
          flexWrap: "wrap",
          gap: 10,
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div
            style={{
              width: 36,
              height: 36,
              borderRadius: 9,
              background: "linear-gradient(135deg,#00d4aa,#818cf8)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: 18,
              fontWeight: 800,
              color: "#0a0f1e",
            }}
          >
            R
          </div>
          <div>
            <div
              style={{ fontWeight: 800, color: "#f1f5f9", fontSize: "1.05em" }}
            >
              P&G Consolidado · Grupo Reental {YEAR}
            </div>
            <div style={{ fontSize: "0.72em", color: "#4a5568" }}>
              5 entidades · USD ·{" "}
              <span style={{ color: "#00d4aa" }}>{fxSource}</span>
            </div>
          </div>
        </div>
        <div
          style={{
            display: "flex",
            gap: 8,
            alignItems: "center",
            flexWrap: "wrap",
          }}
        >
          <button
            onClick={() => setShowFxModal(true)}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              background: "#111827",
              border: "1px solid #00d4aa33",
              borderRadius: 8,
              padding: "6px 14px",
              cursor: "pointer",
              fontFamily: "monospace",
            }}
          >
            <span
              style={{
                width: 7,
                height: 7,
                borderRadius: "50%",
                background: "#00d4aa",
                boxShadow: "0 0 6px #00d4aa",
              }}
            />
            <span
              style={{
                fontSize: "0.72em",
                color: "#4a5568",
                textTransform: "uppercase",
                letterSpacing: "0.08em",
              }}
            >
              BCE
            </span>
            <span
              style={{ fontWeight: 700, color: "#00d4aa", fontSize: "0.9em" }}
            >
              avg ({knownMonths}m) {fxAvg.toFixed(4)}
            </span>
            <span style={{ fontSize: "0.7em", color: "#374151" }}>▾</span>
          </button>
          {totalLoaded > 0 && (
            <div
              style={{
                background: "#00d4aa22",
                color: "#00d4aa",
                borderRadius: 20,
                padding: "4px 12px",
                fontSize: "0.75em",
                fontWeight: 700,
                display: "flex",
                alignItems: "center",
                gap: 6,
              }}
            >
              <span
                style={{
                  width: 7,
                  height: 7,
                  borderRadius: "50%",
                  background: "#00d4aa",
                  boxShadow: "0 0 8px #00d4aa",
                }}
              />
              {totalLoaded}/{EMPRESAS_CONFIG.length}
            </div>
          )}
          {allUnmatched.length > 0 && (
            <button
              onClick={() => setShowUnmatched(!showUnmatched)}
              style={{
                background: "#f59e0b22",
                color: "#f59e0b",
                border: "1px solid #f59e0b44",
                borderRadius: 7,
                padding: "5px 12px",
                fontSize: "0.78em",
                fontWeight: 700,
                cursor: "pointer",
                fontFamily: "monospace",
              }}
            >
              ⚠ {allUnmatched.length} sin mapear
            </button>
          )}
          <button
            onClick={() => setShowNative(!showNative)}
            style={{
              background: showNative ? "#818cf822" : "#111827",
              color: showNative ? "#818cf8" : "#4a5568",
              border: `1px solid ${showNative ? "#818cf844" : "#1e293b"}`,
              borderRadius: 7,
              padding: "7px 14px",
              fontSize: "0.82em",
              fontWeight: 700,
              cursor: "pointer",
              fontFamily: "monospace",
            }}
          >
            {showNative ? "USD+EUR" : "USD"}
          </button>
          <button
            onClick={() =>
              exportToExcel(
                consData,
                consIS,
                empresasData,
                YEAR,
                fxRates,
                fxAvg
              )
            }
            style={{
              background: totalLoaded > 0 ? "#00d4aa22" : "#111827",
              color: totalLoaded > 0 ? "#00d4aa" : "#1e293b",
              border: `1px solid ${totalLoaded > 0 ? "#00d4aa44" : "#1e293b"}`,
              borderRadius: 7,
              padding: "7px 14px",
              fontSize: "0.82em",
              fontWeight: 700,
              cursor: totalLoaded > 0 ? "pointer" : "default",
              fontFamily: "monospace",
            }}
          >
            ↓ Excel USD
          </button>
        </div>
      </div>

      <div style={{ padding: "24px", maxWidth: 1300, margin: "0 auto" }}>
        <FileUploader
          onUpload={handleUpload}
          onClear={handleClear}
          loadedFiles={loadedFiles}
        />

        {showUnmatched && allUnmatched.length > 0 && (
          <div
            style={{
              background: "#1a1200",
              border: "1px solid #3a2800",
              borderRadius: 10,
              padding: 20,
              marginBottom: 20,
            }}
          >
            <div
              style={{ color: "#f59e0b", fontWeight: 700, marginBottom: 12 }}
            >
              ⚠ Cuentas sin categorizar
            </div>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fill,minmax(400px,1fr))",
                gap: 6,
              }}
            >
              {allUnmatched.slice(0, 30).map((item, i) => (
                <div
                  key={i}
                  style={{
                    fontSize: "0.78em",
                    display: "flex",
                    justifyContent: "space-between",
                  }}
                >
                  <span style={{ color: "#6b5000" }}>
                    [{item.emp.split(" ")[0]}] {item.acc}
                  </span>
                  <span style={{ color: "#f59e0b", fontWeight: 700 }}>
                    {item.total.toFixed(0)}
                  </span>
                </div>
              ))}
            </div>
          </div>
        )}

        <div
          style={{
            display: "flex",
            gap: 8,
            marginBottom: 24,
            flexWrap: "wrap",
          }}
        >
          <Tab id="consolidado" label="📊 Consolidado USD" />
          <Tab id="empresa" label="🏢 Por entidad" />
          <Tab id="mensual" label="📅 Mes a mes" />
          <Tab id="matriz" label="🔀 Entidad × Mes" />
        </div>

        {/* ── CONSOLIDADO ── */}
        {tab === "consolidado" && (
          <div>
            <div
              style={{
                background: "#0d1526",
                border: "1px solid #1e293b",
                borderRadius: 10,
                padding: "10px 18px",
                marginBottom: 16,
                display: "flex",
                alignItems: "center",
                gap: 10,
                fontSize: "0.8em",
                color: "#475569",
                flexWrap: "wrap",
              }}
            >
              <span style={{ color: "#00d4aa", fontSize: "1.2em" }}>$</span>
              Consolidado en <strong style={{ color: "#f1f5f9" }}>USD</strong> ·
              Tipo mensual BCE por cada mes ·
              <span style={{ color: "#818cf8" }}>
                avg parcial {knownMonths}m = {fxAvg.toFixed(4)}
              </span>{" "}
              · para meses sin dato BCE aún ·
              <button
                onClick={() => setShowFxModal(true)}
                style={{
                  background: "none",
                  border: "none",
                  color: "#818cf8",
                  cursor: "pointer",
                  fontFamily: "monospace",
                  fontSize: "1em",
                  padding: 0,
                  textDecoration: "underline",
                }}
              >
                ver detalle →
              </button>
            </div>
            <div
              style={{
                display: "flex",
                gap: 12,
                flexWrap: "wrap",
                marginBottom: 20,
              }}
            >
              <KPI
                label="Gross Revenues"
                value={fmtU(consIS.gross)}
                secondary={
                  showNative ? `≈ ${fmtE(consIS.gross, fxAvg)}` : undefined
                }
                color="#00d4aa"
              />
              <KPI
                label="Net Revenues"
                value={fmtU(consIS.net_rev)}
                secondary={
                  showNative ? `≈ ${fmtE(consIS.net_rev, fxAvg)}` : undefined
                }
                sub={"Margen " + fp(consIS.gross_margin)}
                color="#818cf8"
              />
              {/* ── NEW KPI: Total OPEX ── */}
              <KPI
                label="Total OPEX"
                value={fmtU(-consIS.opex)}
                secondary={
                  showNative ? `≈ ${fmtE(-consIS.opex, fxAvg)}` : undefined
                }
                sub={
                  consIS.gross > 0
                    ? fp((consIS.opex / consIS.gross) * 100) + " s/ingresos"
                    : undefined
                }
                color="#fb923c"
              />
              <KPI
                label="EBITDA"
                value={fmtU(consIS.ebitda)}
                secondary={
                  showNative ? `≈ ${fmtE(consIS.ebitda, fxAvg)}` : undefined
                }
                sub={"Margen " + fp(consIS.ebitda_margin)}
                color={consIS.ebitda >= 0 ? "#00d4aa" : "#ef4444"}
              />
              <KPI
                label="Net Profit"
                value={fmtU(consIS.net_profit)}
                secondary={
                  showNative ? `≈ ${fmtE(consIS.net_profit, fxAvg)}` : undefined
                }
                sub={"Margen " + fp(consIS.net_margin)}
                color={consIS.net_profit >= 0 ? "#fb923c" : "#ef4444"}
              />
            </div>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))",
                gap: 12,
                marginBottom: 20,
              }}
            >
              {EMPRESAS_CONFIG.map(({ name, color, currency }) => {
                const eis = calcIS(empresasData[name] || emptyMonths());
                let grossUSD = 0;
                for (let m = 1; m <= 12; m++) {
                  const md = empresasData[name]?.[m] || {};
                  const r = fxRates[m] || fxAvg;
                  const fx2 = currency === "USD" ? 1 : r;
                  grossUSD +=
                    ((md.setup_fees || 0) +
                      (md.rental_contracts || 0) +
                      (md.rnt_sales || 0) +
                      (md.properties_sale || 0) +
                      (md.other_revenues || 0)) *
                    fx2;
                }
                return (
                  <div
                    key={name}
                    style={{
                      background: "#111827",
                      borderRadius: 12,
                      padding: "16px 18px",
                      borderTop: `3px solid ${color}`,
                    }}
                  >
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        marginBottom: 4,
                      }}
                    >
                      <span
                        style={{
                          fontWeight: 700,
                          color: "#f1f5f9",
                          fontSize: "0.85em",
                        }}
                      >
                        {name}
                      </span>
                      <div
                        style={{
                          display: "flex",
                          gap: 4,
                          alignItems: "center",
                        }}
                      >
                        <span
                          style={{
                            fontSize: "0.65em",
                            color: "#374151",
                            background: "#1e293b",
                            padding: "1px 5px",
                            borderRadius: 3,
                          }}
                        >
                          {currency}
                        </span>
                        {loadedFiles[name] && (
                          <span
                            style={{
                              width: 5,
                              height: 5,
                              borderRadius: "50%",
                              background: color,
                              boxShadow: `0 0 5px ${color}`,
                            }}
                          />
                        )}
                      </div>
                    </div>
                    <div
                      style={{
                        display: "grid",
                        gridTemplateColumns: "1fr 1fr",
                        gap: 6,
                        marginTop: 10,
                      }}
                    >
                      {[
                        ["Gross USD", fmtU(grossUSD, true), "#00d4aa"],
                        [
                          "EBITDA%",
                          fp(eis.ebitda_margin),
                          eis.ebitda_margin >= 0 ? "#818cf8" : "#ef4444",
                        ],
                        [
                          "Net P/L",
                          fmtU(
                            eis.net_profit * (currency === "USD" ? 1 : fxAvg),
                            true
                          ),
                          eis.net_profit >= 0 ? "#fb923c" : "#ef4444",
                        ],
                        [
                          "Margen",
                          fp(eis.net_margin),
                          eis.net_margin >= 0 ? "#94a3b8" : "#ef4444",
                        ],
                      ].map(([l, v, c], i) => (
                        <div
                          key={i}
                          style={{
                            background: "#0a0f1e",
                            borderRadius: 8,
                            padding: "8px 10px",
                          }}
                        >
                          <div
                            style={{
                              fontSize: "0.67em",
                              color: "#4a5568",
                              textTransform: "uppercase",
                              marginBottom: 2,
                            }}
                          >
                            {l}
                          </div>
                          <div
                            style={{
                              fontSize: "0.88em",
                              fontWeight: 700,
                              color: c,
                            }}
                          >
                            {v}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
            <div
              style={{
                background: "#111827",
                borderRadius: 14,
                border: "1px solid #1e293b",
                padding: "20px 24px",
              }}
            >
              <div
                style={{
                  fontSize: "0.75em",
                  color: "#4a5568",
                  textTransform: "uppercase",
                  letterSpacing: "0.12em",
                  marginBottom: 20,
                }}
              >
                Income Statement — Consolidado USD · {YEAR}
              </div>
              <ISTable
                is={consIS}
                monthlyData={consData}
                showNative={showNative}
                fxRates={fxRates}
                fxAvg={fxAvg}
                entityCurrency={null}
              />
            </div>
          </div>
        )}

        {/* ── POR ENTIDAD ── */}
        {tab === "empresa" && (
          <div>
            <div
              style={{
                background: "#0d1526",
                border: "1px solid #1e293b",
                borderRadius: 10,
                padding: "10px 18px",
                marginBottom: 16,
                fontSize: "0.8em",
                color: "#475569",
              }}
            >
              Moneda <strong style={{ color: "#f1f5f9" }}>nativa</strong> ·
              Columna USD usa avg parcial BCE{" "}
              <strong style={{ color: "#00d4aa" }}>{fxAvg.toFixed(4)}</strong> (
              {knownMonths}m reales)
            </div>
            <div
              style={{
                display: "flex",
                gap: 8,
                marginBottom: 20,
                flexWrap: "wrap",
              }}
            >
              {EMPRESAS_CONFIG.map(({ name, color, currency }) => (
                <button
                  key={name}
                  onClick={() => setEmpTab(name)}
                  style={{
                    padding: "8px 18px",
                    borderRadius: 8,
                    cursor: "pointer",
                    fontFamily: "monospace",
                    fontSize: "0.85em",
                    fontWeight: 700,
                    border: `1px solid ${empTab === name ? color : "#1e293b"}`,
                    background: empTab === name ? color + "22" : "#111827",
                    color: empTab === name ? color : "#4a5568",
                  }}
                >
                  {name.split(" ")[0]}{" "}
                  <span style={{ fontSize: "0.7em", color: "#374151" }}>
                    ({currency})
                  </span>{" "}
                  {loadedFiles[name] ? "●" : ""}
                </button>
              ))}
            </div>
            <div
              style={{
                display: "flex",
                gap: 12,
                flexWrap: "wrap",
                marginBottom: 20,
              }}
            >
              <KPI
                label={`Gross (${empConfig.currency})`}
                value={
                  empSym +
                  Math.abs(empIS.gross).toLocaleString("es-ES", {
                    maximumFractionDigits: 0,
                  })
                }
                secondary={`≈ ${fmtU(
                  empIS.gross * (empConfig.currency === "EUR" ? fxAvg : 1),
                  true
                )}`}
                color={empConfig.color}
              />
              <KPI
                label="Net Revenues"
                value={
                  empSym +
                  Math.abs(empIS.net_rev).toLocaleString("es-ES", {
                    maximumFractionDigits: 0,
                  })
                }
                sub={"Margen " + fp(empIS.gross_margin)}
                color="#818cf8"
              />
              {/* ── NEW KPI: Total OPEX per entity ── */}
              <KPI
                label="Total OPEX"
                value={
                  empSym +
                  Math.abs(empIS.opex).toLocaleString("es-ES", {
                    maximumFractionDigits: 0,
                  })
                }
                sub={
                  empIS.gross > 0
                    ? fp((empIS.opex / empIS.gross) * 100) + " s/ingresos"
                    : undefined
                }
                color="#fb923c"
              />
              <KPI
                label="EBITDA"
                value={
                  empSym +
                  Math.abs(empIS.ebitda).toLocaleString("es-ES", {
                    maximumFractionDigits: 0,
                  })
                }
                sub={fp(empIS.ebitda_margin)}
                color={empIS.ebitda >= 0 ? "#00d4aa" : "#ef4444"}
              />
              <KPI
                label="Net Profit"
                value={
                  empSym +
                  Math.abs(empIS.net_profit).toLocaleString("es-ES", {
                    maximumFractionDigits: 0,
                  })
                }
                sub={fp(empIS.net_margin)}
                color={empIS.net_profit >= 0 ? "#fb923c" : "#ef4444"}
              />
            </div>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr",
                gap: 16,
                marginBottom: 20,
              }}
            >
              {["Ingresos vs Gastos", "EBITDA mensual"].map((title, ti) => (
                <div
                  key={ti}
                  style={{
                    background: "#111827",
                    borderRadius: 14,
                    border: "1px solid #1e293b",
                    padding: 20,
                  }}
                >
                  <div
                    style={{
                      fontSize: "0.75em",
                      color: "#4a5568",
                      textTransform: "uppercase",
                      letterSpacing: "0.12em",
                      marginBottom: 16,
                    }}
                  >
                    {title} · {empConfig.currency}
                  </div>
                  <ResponsiveContainer width="100%" height={220}>
                    {ti === 0 ? (
                      <BarChart
                        data={empMonthly}
                        barGap={4}
                        margin={{ top: 5, right: 5, left: 5, bottom: 5 }}
                      >
                        <CartesianGrid
                          strokeDasharray="3 3"
                          stroke="#111827"
                          vertical={false}
                        />
                        <XAxis
                          dataKey="mes"
                          tick={{ fill: "#4a5568", fontSize: 11 }}
                          axisLine={false}
                          tickLine={false}
                        />
                        <YAxis
                          tickFormatter={(v) =>
                            empSym + Math.abs(v / 1000).toFixed(0) + "K"
                          }
                          tick={{ fill: "#4a5568", fontSize: 10 }}
                          axisLine={false}
                          tickLine={false}
                          width={65}
                        />
                        <Tooltip
                          contentStyle={tt}
                          formatter={(v, n) => [
                            empSym +
                              Math.abs(v).toLocaleString("es-ES", {
                                maximumFractionDigits: 0,
                              }),
                            n === "gross"
                              ? "Ingresos"
                              : n === "opex"
                              ? "OPEX"
                              : "Gastos",
                          ]}
                        />
                        <Bar
                          dataKey="gross"
                          fill="#00d4aa"
                          radius={[4, 4, 0, 0]}
                          name="Ingresos"
                        />
                        <Bar
                          dataKey="opex"
                          fill="#fb923c"
                          radius={[4, 4, 0, 0]}
                          name="OPEX"
                        />
                      </BarChart>
                    ) : (
                      <BarChart
                        data={empMonthly}
                        margin={{ top: 5, right: 5, left: 5, bottom: 5 }}
                      >
                        <CartesianGrid
                          strokeDasharray="3 3"
                          stroke="#111827"
                          vertical={false}
                        />
                        <XAxis
                          dataKey="mes"
                          tick={{ fill: "#4a5568", fontSize: 11 }}
                          axisLine={false}
                          tickLine={false}
                        />
                        <YAxis
                          tickFormatter={(v) =>
                            empSym + Math.abs(v / 1000).toFixed(0) + "K"
                          }
                          tick={{ fill: "#4a5568", fontSize: 10 }}
                          axisLine={false}
                          tickLine={false}
                          width={65}
                        />
                        <Tooltip
                          contentStyle={tt}
                          formatter={(v) => [
                            empSym +
                              Math.abs(v).toLocaleString("es-ES", {
                                maximumFractionDigits: 0,
                              }),
                            "EBITDA",
                          ]}
                        />
                        <Bar dataKey="ebitda" radius={[4, 4, 0, 0]}>
                          {empMonthly.map((m, i) => (
                            <Cell
                              key={i}
                              fill={m.ebitda >= 0 ? empConfig.color : "#ef4444"}
                            />
                          ))}
                        </Bar>
                      </BarChart>
                    )}
                  </ResponsiveContainer>
                </div>
              ))}
            </div>
            <div
              style={{
                background: "#111827",
                borderRadius: 14,
                border: "1px solid #1e293b",
                padding: "20px 24px",
              }}
            >
              <ISTable
                is={empIS}
                monthlyData={empNativeData}
                showNative={showNative}
                fxRates={fxRates}
                fxAvg={fxAvg}
                entityCurrency={empConfig.currency}
              />
            </div>
          </div>
        )}

        {/* ── MES A MES ── */}
        {tab === "mensual" && (
          <div>
            <div
              style={{
                background: "#111827",
                borderRadius: 14,
                border: "1px solid #1e293b",
                padding: 20,
                marginBottom: 16,
              }}
            >
              <div
                style={{
                  fontSize: "0.75em",
                  color: "#4a5568",
                  textTransform: "uppercase",
                  marginBottom: 16,
                }}
              >
                Gross Revenues USD por entidad
              </div>
              <ResponsiveContainer width="100%" height={280}>
                <LineChart margin={{ top: 5, right: 20, left: 10, bottom: 5 }}>
                  <CartesianGrid
                    strokeDasharray="3 3"
                    stroke="#111827"
                    vertical={false}
                  />
                  <XAxis
                    dataKey="mes"
                    type="category"
                    allowDuplicatedCategory={false}
                    tick={{ fill: "#4a5568", fontSize: 12 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <YAxis
                    tickFormatter={(v) =>
                      "$" + Math.abs(v / 1000).toFixed(0) + "K"
                    }
                    tick={{ fill: "#4a5568", fontSize: 11 }}
                    axisLine={false}
                    tickLine={false}
                    width={70}
                  />
                  <Tooltip
                    contentStyle={tt}
                    formatter={(v, n) => [
                      "$" +
                        Math.abs(v).toLocaleString("es-ES", {
                          maximumFractionDigits: 0,
                        }),
                      n,
                    ]}
                  />
                  <Legend wrapperStyle={{ fontSize: 12, color: "#4a5568" }} />
                  {EMPRESAS_CONFIG.map(({ name, color, currency }) => {
                    const d = calcMonthly(
                      empresasData[name] || emptyMonths()
                    ).map((m, i) => ({
                      ...m,
                      gross:
                        m.gross *
                        (currency === "USD" ? 1 : fxRates[i + 1] || fxAvg),
                    }));
                    return (
                      <Line
                        key={name}
                        data={d}
                        type="monotone"
                        dataKey="gross"
                        name={name}
                        stroke={color}
                        strokeWidth={2}
                        dot={{ fill: color, r: 3 }}
                      />
                    );
                  })}
                </LineChart>
              </ResponsiveContainer>
            </div>
            <div
              style={{
                background: "#111827",
                borderRadius: 14,
                border: "1px solid #1e293b",
                padding: 20,
                marginBottom: 16,
              }}
            >
              <div
                style={{
                  fontSize: "0.75em",
                  color: "#4a5568",
                  textTransform: "uppercase",
                  marginBottom: 16,
                }}
              >
                EBITDA USD consolidado mensual
              </div>
              <ResponsiveContainer width="100%" height={200}>
                <BarChart
                  data={consMonthly}
                  margin={{ top: 5, right: 10, left: 10, bottom: 5 }}
                >
                  <CartesianGrid
                    strokeDasharray="3 3"
                    stroke="#111827"
                    vertical={false}
                  />
                  <XAxis
                    dataKey="mes"
                    tick={{ fill: "#4a5568", fontSize: 12 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <YAxis
                    tickFormatter={(v) =>
                      "$" + Math.abs(v / 1000).toFixed(0) + "K"
                    }
                    tick={{ fill: "#4a5568", fontSize: 11 }}
                    axisLine={false}
                    tickLine={false}
                    width={70}
                  />
                  <Tooltip
                    contentStyle={tt}
                    formatter={(v) => [
                      "$" +
                        Math.abs(v).toLocaleString("es-ES", {
                          maximumFractionDigits: 0,
                        }),
                      "EBITDA",
                    ]}
                  />
                  <Bar dataKey="ebitda" radius={[4, 4, 0, 0]}>
                    {consMonthly.map((m, i) => (
                      <Cell
                        key={i}
                        fill={m.ebitda >= 0 ? "#818cf8" : "#ef4444"}
                      />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div
              style={{
                background: "#111827",
                borderRadius: 14,
                border: "1px solid #1e293b",
                padding: "20px 24px",
              }}
            >
              <table
                style={{
                  width: "100%",
                  borderCollapse: "collapse",
                  fontSize: "0.9em",
                }}
              >
                <thead>
                  <tr
                    style={{
                      color: "#4a5568",
                      fontSize: "0.8em",
                      textTransform: "uppercase",
                      borderBottom: "2px solid #1e293b",
                    }}
                  >
                    {[
                      "Mes",
                      "EUR/USD BCE",
                      "Gross USD",
                      "Net Rev",
                      "Total OPEX",
                      "EBITDA",
                      "Margen",
                    ].map((h, i) => (
                      <th
                        key={i}
                        style={{
                          padding: "10px 14px",
                          textAlign: i === 0 ? "left" : "right",
                          fontWeight: 500,
                          color: h === "Total OPEX" ? "#fb923c" : undefined,
                        }}
                      >
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {consMonthly.map((m, i) => {
                    const rate = fxRates[i + 1] || fxAvg;
                    const isReal = !!fxRates[i + 1];
                    return (
                      <tr
                        key={i}
                        className="hr"
                        style={{ borderBottom: "1px solid #0a0f1e" }}
                      >
                        <td
                          style={{
                            padding: "9px 14px",
                            color: "#f1f5f9",
                            fontWeight: 600,
                          }}
                        >
                          {m.mes}
                        </td>
                        <td
                          style={{
                            padding: "9px 14px",
                            textAlign: "right",
                            color: isReal ? "#00d4aa" : "#374151",
                            fontFamily: "monospace",
                            fontSize: "0.88em",
                          }}
                        >
                          {rate.toFixed(4)}
                          {!isReal && (
                            <span
                              style={{
                                fontSize: "0.75em",
                                color: "#1e293b",
                                marginLeft: 4,
                              }}
                            >
                              ~avg
                            </span>
                          )}
                        </td>
                        <td
                          style={{
                            padding: "9px 14px",
                            textAlign: "right",
                            color: "#00d4aa",
                            fontWeight: 700,
                          }}
                        >
                          {"$" +
                            m.gross.toLocaleString("es-ES", {
                              maximumFractionDigits: 0,
                            })}
                        </td>
                        <td
                          style={{
                            padding: "9px 14px",
                            textAlign: "right",
                            color: "#818cf8",
                          }}
                        >
                          {"$" +
                            m.net_rev.toLocaleString("es-ES", {
                              maximumFractionDigits: 0,
                            })}
                        </td>
                        {/* ── NEW: Total OPEX column in monthly table ── */}
                        <td
                          style={{
                            padding: "9px 14px",
                            textAlign: "right",
                            color: "#fb923c",
                            fontWeight: 700,
                          }}
                        >
                          {"-$" +
                            m.opex.toLocaleString("es-ES", {
                              maximumFractionDigits: 0,
                            })}
                        </td>
                        <td
                          style={{
                            padding: "9px 14px",
                            textAlign: "right",
                            color: m.ebitda >= 0 ? "#818cf8" : "#ef4444",
                            fontWeight: 700,
                          }}
                        >
                          {"$" +
                            m.ebitda.toLocaleString("es-ES", {
                              maximumFractionDigits: 0,
                            })}
                        </td>
                        <td
                          style={{
                            padding: "9px 14px",
                            textAlign: "right",
                            color: "#4a5568",
                          }}
                        >
                          {m.gross > 0 ? fp((m.ebitda / m.gross) * 100) : "—"}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                <tfoot>
                  <tr style={{ borderTop: "2px solid #1e293b" }}>
                    <td
                      style={{
                        padding: "10px 14px",
                        color: "#f1f5f9",
                        fontWeight: 800,
                      }}
                    >
                      TOTAL
                    </td>
                    <td
                      style={{
                        padding: "10px 14px",
                        textAlign: "right",
                        color: "#374151",
                        fontSize: "0.8em",
                        fontFamily: "monospace",
                      }}
                    >
                      avg ({knownMonths}m) {fxAvg.toFixed(4)}
                    </td>
                    <td
                      style={{
                        padding: "10px 14px",
                        textAlign: "right",
                        color: "#00d4aa",
                        fontWeight: 800,
                      }}
                    >
                      {fmtU(consIS.gross)}
                    </td>
                    <td
                      style={{
                        padding: "10px 14px",
                        textAlign: "right",
                        color: "#818cf8",
                        fontWeight: 800,
                      }}
                    >
                      {fmtU(consIS.net_rev)}
                    </td>
                    {/* ── Total OPEX footer ── */}
                    <td
                      style={{
                        padding: "10px 14px",
                        textAlign: "right",
                        color: "#fb923c",
                        fontWeight: 800,
                      }}
                    >
                      {"-$" +
                        consIS.opex.toLocaleString("es-ES", {
                          maximumFractionDigits: 0,
                        })}
                    </td>
                    <td
                      style={{
                        padding: "10px 14px",
                        textAlign: "right",
                        color: consIS.ebitda >= 0 ? "#818cf8" : "#ef4444",
                        fontWeight: 800,
                      }}
                    >
                      {fmtU(consIS.ebitda)}
                    </td>
                    <td
                      style={{
                        padding: "10px 14px",
                        textAlign: "right",
                        color: "#4a5568",
                        fontWeight: 800,
                      }}
                    >
                      {fp(consIS.ebitda_margin)}
                    </td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>
        )}

        {/* ── ENTIDAD × MES ── */}
        {tab === "matriz" && (
          <div>
            {["gross", "ebitda", "net_profit"].map((metric) => {
              const labels = {
                gross: "Gross Revenues USD",
                ebitda: "EBITDA USD",
                net_profit: "Net Profit USD",
              };
              const mc = {
                gross: "#00d4aa",
                ebitda: "#818cf8",
                net_profit: "#fb923c",
              };
              return (
                <div
                  key={metric}
                  style={{
                    background: "#111827",
                    borderRadius: 14,
                    border: "1px solid #1e293b",
                    padding: "20px 24px",
                    marginBottom: 16,
                    overflowX: "auto",
                  }}
                >
                  <div
                    style={{
                      fontSize: "0.75em",
                      color: mc[metric],
                      textTransform: "uppercase",
                      letterSpacing: "0.12em",
                      marginBottom: 16,
                      fontWeight: 700,
                    }}
                  >
                    {labels[metric]} · por entidad y mes
                  </div>
                  <table
                    style={{
                      width: "100%",
                      borderCollapse: "collapse",
                      fontSize: "0.85em",
                      minWidth: 700,
                    }}
                  >
                    <thead>
                      <tr style={{ borderBottom: "2px solid #1e293b" }}>
                        <th
                          style={{
                            padding: "8px 14px",
                            textAlign: "left",
                            color: "#4a5568",
                            fontWeight: 500,
                            fontSize: "0.78em",
                            textTransform: "uppercase",
                          }}
                        >
                          Mes
                        </th>
                        <th
                          style={{
                            padding: "8px 14px",
                            textAlign: "right",
                            color: "#374151",
                            fontWeight: 500,
                            fontSize: "0.75em",
                          }}
                        >
                          EUR/USD
                        </th>
                        {EMPRESAS_CONFIG.map((e) => (
                          <th
                            key={e.name}
                            style={{
                              padding: "8px 14px",
                              textAlign: "right",
                              color: e.color,
                              fontWeight: 600,
                              fontSize: "0.78em",
                            }}
                          >
                            {e.name.split(" ")[0]}
                            <span
                              style={{ color: "#374151", fontSize: "0.8em" }}
                            >
                              {" "}
                              ({e.currency})
                            </span>
                          </th>
                        ))}
                        <th
                          style={{
                            padding: "8px 14px",
                            textAlign: "right",
                            color: "#f1f5f9",
                            fontWeight: 700,
                            fontSize: "0.78em",
                            textTransform: "uppercase",
                          }}
                        >
                          Total USD
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {Array.from({ length: 12 }, (_, i) => i + 1).map((m) => {
                        const rate = fxRates[m] || fxAvg;
                        const isReal = !!fxRates[m];
                        const rowVals = EMPRESAS_CONFIG.map((e) => {
                          const is = calcIS({
                            [m]: empresasData[e.name]?.[m] || {},
                          });
                          return (
                            (is[metric] || 0) *
                            (e.currency === "USD" ? 1 : rate)
                          );
                        });
                        const total = rowVals.reduce((a, b) => a + b, 0),
                          hasData = rowVals.some((v) => v !== 0);
                        return (
                          <tr
                            key={m}
                            className="hr"
                            style={{
                              borderBottom: "1px solid #0a0f1e",
                              opacity: hasData ? 1 : 0.3,
                            }}
                          >
                            <td
                              style={{
                                padding: "8px 14px",
                                color: "#f1f5f9",
                                fontWeight: 600,
                              }}
                            >
                              {MONTHS[m]}
                            </td>
                            <td
                              style={{
                                padding: "8px 14px",
                                textAlign: "right",
                                color: isReal ? "#374151" : "#1e293b",
                                fontFamily: "monospace",
                                fontSize: "0.8em",
                              }}
                            >
                              {rate.toFixed(4)}
                              {!isReal && (
                                <span style={{ color: "#111827" }}>~</span>
                              )}
                            </td>
                            {rowVals.map((v, i) => (
                              <td
                                key={i}
                                style={{
                                  padding: "8px 14px",
                                  textAlign: "right",
                                  color:
                                    v === 0
                                      ? "#1e293b"
                                      : v > 0
                                      ? EMPRESAS_CONFIG[i].color
                                      : "#ef4444",
                                  fontWeight: v !== 0 ? 600 : 400,
                                }}
                              >
                                {v === 0
                                  ? "—"
                                  : "$" + Math.abs(v / 1000).toFixed(1) + "K"}
                              </td>
                            ))}
                            <td
                              style={{
                                padding: "8px 14px",
                                textAlign: "right",
                                color:
                                  total > 0
                                    ? "#f1f5f9"
                                    : total < 0
                                    ? "#ef4444"
                                    : "#1e293b",
                                fontWeight: 700,
                              }}
                            >
                              {total === 0
                                ? "—"
                                : "$" + Math.abs(total / 1000).toFixed(1) + "K"}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                    <tfoot>
                      <tr style={{ borderTop: "2px solid #1e293b" }}>
                        <td
                          style={{
                            padding: "10px 14px",
                            color: "#f1f5f9",
                            fontWeight: 800,
                          }}
                        >
                          TOTAL
                        </td>
                        <td
                          style={{
                            padding: "10px 14px",
                            textAlign: "right",
                            color: "#374151",
                            fontSize: "0.8em",
                            fontFamily: "monospace",
                          }}
                        >
                          avg ({knownMonths}m) {fxAvg.toFixed(4)}
                        </td>
                        {EMPRESAS_CONFIG.map((e) => {
                          const is = calcIS(
                            empresasData[e.name] || emptyMonths()
                          );
                          const v =
                            (is[metric] || 0) *
                            (e.currency === "USD" ? 1 : fxAvg);
                          return (
                            <td
                              key={e.name}
                              style={{
                                padding: "10px 14px",
                                textAlign: "right",
                                color: v >= 0 ? e.color : "#ef4444",
                                fontWeight: 800,
                              }}
                            >
                              {"$" + (v / 1000).toFixed(1) + "K"}
                            </td>
                          );
                        })}
                        <td
                          style={{
                            padding: "10px 14px",
                            textAlign: "right",
                            color: consIS[metric] >= 0 ? "#f1f5f9" : "#ef4444",
                            fontWeight: 800,
                          }}
                        >
                          {"$" + (consIS[metric] / 1000).toFixed(1) + "K"}
                        </td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              );
            })}
          </div>
        )}

        <div
          style={{
            marginTop: 32,
            textAlign: "center",
            color: "#1e293b",
            fontSize: "0.7em",
          }}
        >
          P&G Grupo Reental · BCE EXR.M.USD.EUR.SP00.A · {YEAR} · avg (
          {knownMonths}m) {fxAvg.toFixed(4)}
        </div>
      </div>
    </div>
  );
}
