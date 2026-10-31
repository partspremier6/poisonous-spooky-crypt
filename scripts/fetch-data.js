// ============================================================================
// scripts/fetch-data.js — v12.1
//
// v12.1 fix (user-reported):
//   Dashboard showed "2,017 public chargers" while L2 (30,741) + DC fast
//   (8,479) summed to 39,220. The prose regex for "Total chargers" had
//   grabbed a stray year (2017). The province-sum fallback never fired
//   because the province table yielded <5 validated rows this run.
//
//   FIX: post-parse RECONCILIATION with arithmetic invariants, so no single
//   bad grab survives:
//     • If level2 & level3 both exist → totalChargers := level2 + level3
//       whenever they disagree by >15% (components win).
//     • If bevs/phevs exist → totalEvs := bevs + phevs + fcev whenever
//       totalEvs is missing or smaller than the component sum.
//     • Ratio recomputed last, from reconciled totals.
//   Plus: the fully parsed KPI set is printed to the run log for debugging.
//
// v12 recap: permutation-validated TC province rows (ratio×chargers≈EVs),
// AFDC ev-charging-units rows grouped by station (docs: port-count fields
// repeat the station total across unit rows), KPI sanity bounds.
// ============================================================================

"use strict";

const fs = require("fs");
const path = require("path");
const os = require("os");
const { execFileSync } = require("child_process");

const ROOT = process.cwd();
const DATA_DIR = path.join(ROOT, "data");
const HISTORY_DIR = path.join(DATA_DIR, "history");

const TC_URL =
  "https://tc.canada.ca/en/road-transportation/innovative-technologies/" +
  "electric-vehicles/canada-electric-vehicle-dashboard";
const AFDC_HOST = "https://developer.nlr.gov";
const AFDC_API_KEY = process.env.AFDC_API_KEY || "DEMO_KEY";

const UA = "CanadaEVDashboard/1.0 (GitHub Actions; personal use)";
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ---------------------------------------------------------------- utilities

async function httpGet(url, { asBuffer = false, timeoutMs = 120000, retries = 2 } = {}) {
  let lastErr;
  for (let attempt = 0; attempt <= retries; attempt++) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), timeoutMs);
    try {
      const res = await fetch(url, { headers: { "User-Agent": UA }, signal: ctrl.signal });
      clearTimeout(timer);
      if (!res.ok) {
        const body = (await res.text()).substring(0, 200);
        throw new Error("HTTP " + res.status + (body ? " — " + body.replace(/\s+/g, " ") : ""));
      }
      return asBuffer ? Buffer.from(await res.arrayBuffer()) : await res.text();
    } catch (e) {
      clearTimeout(timer);
      lastErr = e;
      if (attempt < retries) { console.log("  retry: " + e.message); await sleep(5000 * (attempt + 1)); }
    }
  }
  throw lastErr;
}

async function downloadToFile(url, tmpFile, { timeoutMs = 180000 } = {}) {
  try {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), timeoutMs);
    const res = await fetch(url, { headers: { "User-Agent": UA }, signal: ctrl.signal });
    clearTimeout(timer);
    if (res.ok) {
      fs.writeFileSync(tmpFile, Buffer.from(await res.arrayBuffer()));
      return { method: "fetch" };
    }
    const body = (await res.text()).substring(0, 300);
    throw new Error("HTTP " + res.status + (body ? " — " + body.replace(/\s+/g, " ") : ""));
  } catch (fetchErr) {
    console.log("  fetch failed (" + fetchErr.message + ") — trying curl");
  }
  try {
    const codeOut = execFileSync(
      "curl",
      ["-sS", "--retry", "2", "--retry-delay", "5", "--connect-timeout", "30",
       "--max-time", "180", "--http1.1", "--retry-all-errors", "-4",
       "-A", UA, "-o", tmpFile, "-w", "%{http_code}", url],
      { encoding: "utf8", timeout: 600000, stdio: ["ignore", "pipe", "pipe"] }
    ).trim();
    const code = parseInt(codeOut, 10);
    let bodyExcerpt = "";
    try { bodyExcerpt = fs.readFileSync(tmpFile, "utf8").substring(0, 300).replace(/\s+/g, " "); } catch (_) {}
    if (code === 429) {
      throw Object.assign(new Error("HTTP 429 (rate limited) — " + bodyExcerpt), { rateLimited: true });
    }
    if (code !== 200) throw new Error("HTTP " + code + (bodyExcerpt ? " — " + bodyExcerpt : ""));
    if (!fs.existsSync(tmpFile) || fs.statSync(tmpFile).size === 0) {
      throw new Error("curl reported 200 but the file is empty");
    }
    return { method: "curl" };
  } catch (curlErr) {
    if (curlErr.rateLimited) throw curlErr;
    const status = curlErr.status || curlErr.code || "unknown";
    const stderr = String(curlErr.stderr || curlErr.message || "").trim()
      .split("\n").slice(-3).join(" | ").substring(0, 300);
    throw new Error(`curl failed (exit ${status}): ${stderr}`);
  }
}

function num(v) {
  if (v == null) return null;
  const m = String(v).replace(/,/g, "").match(/-?\d+(\.\d+)?/);
  return m ? Math.round(parseFloat(m[0])) : null;
}
function pct(v) {
  if (v == null) return null;
  const m = String(v).match(/[\d.]+/);
  return m ? parseFloat(m[0]) : null;
}
function cellText(html) {
  return String(html || "").replace(/&nbsp;/gi, " ").replace(/&#160;/g, " ")
    .replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
}
function parseTables(html) {
  const tables = [];
  const tRe = /<table[^>]*>([\s\S]*?)<\/table>/gi;
  let tm;
  while ((tm = tRe.exec(html))) {
    const rows = [], rRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    let rm;
    while ((rm = rRe.exec(tm[1]))) {
      const cells = [], cRe = /<t[hd][^>]*>([\s\S]*?)<\/t[hd]>/gi;
      let cm;
      while ((cm = cRe.exec(rm[1]))) cells.push(cellText(cm[1]));
      if (cells.length) rows.push(cells);
    }
    if (rows.length) tables.push(rows);
  }
  return tables;
}
function parseCsvLine(line) {
  const cells = [], cur = [];
  let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQ) {
      if (ch === '"') { if (line[i + 1] === '"') { cur.push('"'); i++; } else inQ = false; }
      else cur.push(ch);
    } else if (ch === '"') inQ = true;
    else if (ch === ",") { cells.push(cur.join("")); cur.length = 0; }
    else cur.push(ch);
  }
  cells.push(cur.join(""));
  return cells;
}

// Assign 3 unknown numbers to (ratio, chargers, evs) by internal consistency:
// the right assignment satisfies ratio * chargers ~= evs (within 15%) and
// evs > chargers. Returns null if no assignment fits.
function assignProvinceNumbers(nums) {
  if (!nums || nums.length < 2) return null;
  if (nums.length === 2) {
    const [a, b] = nums;
    if (a > b && b > 0) return { evPerCharger: null, totalChargers: b, totalEvs: a };
    if (b > a && a > 0) return { evPerCharger: null, totalChargers: a, totalEvs: b };
    return null;
  }
  let best = null, bestScore = Infinity;
  for (let i = 0; i < 3; i++) {
    const rest = [0, 1, 2].filter((j) => j !== i);
    const r = nums[i];
    for (const [ci, ei] of [rest, [rest[1], rest[0]]]) {
      const chargers = nums[ci], evs = nums[ei];
      if (!(evs > chargers) || chargers <= 0 || r <= 0) continue;
      const expected = r * chargers;
      const dev = Math.abs(evs - expected) / expected;
      if (dev <= 0.15 && dev < bestScore) {
        bestScore = dev;
        best = { evPerCharger: r, totalChargers: chargers, totalEvs: evs };
      }
    }
  }
  return best;
}

// Enforce arithmetic invariants across the KPI set so no single bad grab
// survives. Mutates and returns the kpis object.
function reconcileKpis(k) {
  // Charger family: components beat the total when they disagree badly.
  if (k.level2 != null && k.level3 != null) {
    const sum = k.level2 + k.level3;
    if (k.totalChargers == null || Math.abs(k.totalChargers - sum) / sum > 0.15) {
      if (k.totalChargers != null) {
        console.log("  [reconcile] totalChargers " + k.totalChargers +
          " disagreed with L2+L3 sum " + sum + " — using component sum");
      }
      k.totalChargers = sum;
    }
  }
  // EV family: total can never be less than its components.
  if ((k.bevs != null || k.phevs != null) && k.totalEvs != null) {
    const comp = (k.bevs || 0) + (k.phevs || 0) + (k.fcevs || 0);
    if (comp > 0 && k.totalEvs < comp) {
      console.log("  [reconcile] totalEvs " + k.totalEvs + " < component sum " + comp + " — using component sum");
      k.totalEvs = comp;
    }
  }
  // Sanity bounds on the ratio; recompute from reconciled totals.
  const plausibleRatio = (r) => r != null && r >= 5 && r <= 100;
  if (!plausibleRatio(k.evPerCharger)) k.evPerCharger = null;
  if (k.evPerCharger == null && k.totalEvs && k.totalChargers) {
    k.evPerCharger = Math.round(k.totalEvs / k.totalChargers);
  }
  return k;
}

// ------------------------------------------- source 1: Transport Canada

async function fetchTransportCanada() {
  const html = await httpGet(TC_URL);
  const updatedMatch = html.match(/[Dd]ata was updated in ([A-Z][a-z]+ \d{4})/);
  const updated = updatedMatch ? updatedMatch[1] : null;

  const text = html.replace(/<[^>]*>/g, " ");
  const grab = (re) => { const m = text.match(re); return m ? num(m[1]) : null; };
  const ratioProse =
    text.match(/(\d+)\s*EVs?\s*[:,]?\s*(?:to|per)?\s*(?:every\s*)?1\s*charger/i) ||
    text.match(/(?:ratio of|ratio:?)\s*(\d+)\s*(?:to\s*1|:\s*1|EVs)/i);

  const kpis = {
    totalEvs: grab(/Total EVs?:?\s*([\d,]+)/),
    bevs: grab(/BEVs?:?\s*([\d,]+)/),
    phevs: grab(/PHEVs?:?\s*([\d,]+)/),
    totalChargers: grab(/Total chargers:?\s*([\d,]+)/),
    level2: grab(/Level 2:?\s*([\d,]+)/),
    level3: grab(/Level 3:?\s*([\d,]+)/),
    evPerCharger: ratioProse ? parseInt(ratioProse[1], 10) : null,
  };

  const tables = parseTables(html);
  const header = (t) => t[0].map((c) => String(c || "").toLowerCase());

  // ---- fleet table (categories) ----
  let fleet = null;
  for (const t of tables) {
    const h = header(t);
    if (h.some((c) => /categor/.test(c))) {
      fleet = t.slice(1).map((r) => ({ category: r[0], count: num(r[1]), sharePct: pct(r[2]) }))
        .filter((r) => r.category && r.count != null);
      if (fleet.length) break;
    }
  }

  // ---- market share (annual) ----
  let marketShare = null;
  for (const t of tables) {
    const h = header(t);
    if (h.includes("year") && (h.includes("bev") || h.includes("zev"))) {
      marketShare = t.slice(1).map((r) => ({
        period: r[0], bev: pct(r[1]), phev: pct(r[2]), fcev: pct(r[3]), total: pct(r[4]),
      })).filter((x) => x.total != null);
      break;
    }
  }

  // ---- charger growth series (chart only — not used for KPIs) ----
  let chargerGrowth = null;
  for (const t of tables) {
    const h = header(t);
    const iL2 = h.findIndex((c) => /level\s*2/.test(c));
    const iL3 = h.findIndex((c) => /level\s*3|dc\s*fast/.test(c));
    const iYr = h.findIndex((c) => /year|period|date/.test(c));
    if (iL2 >= 0 && iL3 >= 0 && iYr >= 0) {
      chargerGrowth = t.slice(1).map((r) => ({
        period: String(r[iYr] || ""), level2: num(r[iL2]), level3: num(r[iL3]),
      })).filter((x) => x.period && (x.level2 != null || x.level3 != null));
      if (chargerGrowth.length > 1) break;
    }
  }

  // ---- province table: role discovery by permutation validation ----
  let provinces = null;
  for (const t of tables) {
    const h = header(t);
    if (!(h.some((c) => /province|jurisdiction|territor/.test(c)))) continue;
    const iProv = h.findIndex((c) => /province|jurisdiction|territor/.test(c));
    const parsed = [];
    for (const r of t.slice(1)) {
      const name = cellText(r[iProv] || "");
      if (!name || /canada/i.test(name) || /total/i.test(name)) continue;
      const nums = r.slice(iProv + 1).map(num).filter((v) => v != null && v > 0);
      const a = assignProvinceNumbers(nums);
      if (a) parsed.push({ province: name, ...a });
    }
    if (parsed.length >= 5) { provinces = parsed; break; }
  }

  // ---- fill gaps from tables before reconciling ----
  if (fleet) {
    const sum = (re) => fleet.filter((r) => re.test(r.category)).reduce((a, r) => a + (r.count || 0), 0);
    const fBev = sum(/bev/i), fPhev = sum(/phev/i), fFcev = sum(/fcev/i);
    if (kpis.bevs == null && fBev) kpis.bevs = fBev;
    if (kpis.phevs == null && fPhev) kpis.phevs = fPhev;
    if (kpis.fcevs == null) kpis.fcevs = fFcev || null;
    if (kpis.totalEvs == null && (fBev || fPhev)) kpis.totalEvs = fBev + fPhev + fFcev;
  }
  if (provinces) {
    const pCh = provinces.reduce((a, p) => a + (p.totalChargers || 0), 0);
    const pEv = provinces.reduce((a, p) => a + (p.totalEvs || 0), 0);
    if (kpis.totalChargers == null && pCh > 10000) kpis.totalChargers = pCh;
    if (kpis.totalEvs == null && pEv > 500000) kpis.totalEvs = pEv;
  }
  if ((kpis.level2 == null || kpis.level3 == null) && chargerGrowth) {
    const last = chargerGrowth[chargerGrowth.length - 1];
    const tot = (last.level2 || 0) + (last.level3 || 0);
    if (tot > 0 && kpis.totalChargers && Math.abs(tot - kpis.totalChargers) / kpis.totalChargers < 0.1) {
      if (kpis.level2 == null) kpis.level2 = last.level2;
      if (kpis.level3 == null) kpis.level3 = last.level3;
    }
  }

  // ---- reconciliation: arithmetic invariants have the final word ----
  reconcileKpis(kpis);

  console.log("  [kpis] " + JSON.stringify(kpis));
  if (provinces) console.log("  [provinces] " + provinces.length + " rows parsed");
  else console.log("  [provinces] none passed validation");

  return { updated, kpis, marketShare, fleet, provinces, chargerGrowth, url: TC_URL };
}

// ------------------------------------------- source 2: AFDC

const PROV_NAMES = {
  AB: "Alberta", BC: "British Columbia", MB: "Manitoba", NB: "New Brunswick",
  NL: "Newfoundland and Labrador", NS: "Nova Scotia", NT: "Northwest Territories",
  NU: "Nunavut", ON: "Ontario", PE: "Prince Edward Island", QC: "Quebec",
  SK: "Saskatchewan", YT: "Yukon",
};
const provName = (s) => PROV_NAMES[String(s || "").toUpperCase().trim()] || String(s || "").trim();

// ev-charging-units flavor: group rows by station (port-count fields repeat
// the station-level total across unit rows), take max per station, sum.
function aggregateUnitsCsv(raw) {
  const lines = raw.split("\n");
  if (lines.length < 3) throw new Error("AFDC CSV too short to be data (" + lines.length + " lines)");
  const firstLine = lines[0].replace(/^\uFEFF/, "");
  const head = firstLine.split(",").map((h) => h.trim());
  const lower = head.map((h) => h.toLowerCase());
  const col = (re) => lower.findIndex((h) => re.test(h));
  const iProv = col(/^state$|^province$/);
  const iL2 = col(/level\s*2.*(evse|num|count)|evse.*level\s*2/);
  const iDC = col(/dc\s*fast.*(count|num)|ev_dc_fast_num/);
  const iL1 = col(/level\s*1.*(evse|num|count)/);
  const iNet = col(/^ev[_ ]network$/);
  const iKw = col(/power output/i);
  const iId = col(/^id$|^station.*id$|^fuel.*id$/);
  if (iProv < 0) throw new Error("No province/state column. Headers: " + head.slice(0, 14).join(" | "));
  if (iL2 < 0 && iDC < 0) throw new Error("No charger-level columns. Headers: " + head.slice(0, 14).join(" | "));

  const stations = new Map();
  let rawRows = 0;
  for (let r = 1; r < lines.length; r++) {
    const line = lines[r];
    if (!line.trim()) continue;
    const f = parseCsvLine(line);
    if (f.length < 3) continue;
    const prov = provName(f[iProv]);
    if (!prov) continue;
    rawRows++;
    const id = iId >= 0 && f[iId] ? f[iId] : null;
    const key = id != null ? "id:" + id : "row:" + r;
    let st = stations.get(key);
    if (!st) {
      st = { prov, network: iNet >= 0 ? ((f[iNet] || "Non-networked").trim() || "Non-networked") : null,
             maxL2: 0, maxDC: 0, maxL1: 0, maxKw: 0 };
      stations.set(key, st);
    }
    if (iL2 >= 0) st.maxL2 = Math.max(st.maxL2, parseFloat(f[iL2]) || 0);
    if (iDC >= 0) st.maxDC = Math.max(st.maxDC, parseFloat(f[iDC]) || 0);
    if (iL1 >= 0) st.maxL1 = Math.max(st.maxL1, parseFloat(f[iL1]) || 0);
    if (iKw >= 0) st.maxKw = Math.max(st.maxKw, parseFloat(f[iKw]) || 0);
  }

  const byProv = {}, byNetwork = {};
  let l2 = 0, dc = 0, l1 = 0, totalKw = 0;
  for (const st of stations.values()) {
    l2 += st.maxL2; dc += st.maxDC; l1 += st.maxL1; totalKw += st.maxKw;
    const p = (byProv[st.prov] = byProv[st.prov] || { level2: 0, dcFast: 0, kw: 0 });
    p.level2 += st.maxL2; p.dcFast += st.maxDC; p.kw += st.maxKw;
    if (st.network) byNetwork[st.network] = (byNetwork[st.network] || 0) + 1;
  }
  if (!Object.keys(byProv).length) throw new Error("AFDC CSV parsed but no Canadian province rows found");

  return {
    source: "US DOE Alternative Fuels Data Center (covers Canada)",
    flavor: "units",
    stationCount: stations.size,
    rawRowCount: rawRows,
    headerLine: firstLine.substring(0, 400),
    totalLocations: stations.size,
    level1Units: Math.round(l1),
    level2Units: Math.round(l2),
    dcFastUnits: Math.round(dc),
    totalKw: Math.round(totalKw),
    byProvince: Object.entries(byProv).map(([province, v]) => ({
      province, level2: Math.round(v.level2), dcFast: Math.round(v.dcFast), kw: Math.round(v.kw),
    })).sort((a, b) => (b.level2 + b.dcFast) - (a.level2 + a.dcFast)),
    topNetworks: Object.entries(byNetwork).sort((a, b) => b[1] - a[1]).slice(0, 10)
      .map(([network, count]) => ({ network, count })),
    url: "https://afdc.energy.gov/stations/states",
  };
}

// All Stations flavor: one row per station — direct sums.
function aggregateStationsCsv(raw) {
  const lines = raw.split("\n");
  if (lines.length < 3) throw new Error("AFDC CSV too short to be data (" + lines.length + " lines)");
  const firstLine = lines[0].replace(/^\uFEFF/, "");
  const head = firstLine.split(",").map((h) => h.trim());
  const lower = head.map((h) => h.toLowerCase());
  const col = (re) => lower.findIndex((h) => re.test(h));
  const iProv = col(/^state$|^province$/);
  const iL2 = col(/level\s*2.*(evse|num|count)|evse.*level\s*2/);
  const iDC = col(/dc\s*fast.*(count|num)/);
  const iNet = col(/^ev[_ ]network$/);
  const iKw = col(/power output/i);
  if (iProv < 0) throw new Error("No province/state column. Headers: " + head.slice(0, 14).join(" | "));

  const byProv = {}, byNetwork = {};
  let l2 = 0, dc = 0, totalKw = 0, locations = 0;
  for (let r = 1; r < lines.length; r++) {
    const line = lines[r];
    if (!line.trim()) continue;
    const f = parseCsvLine(line);
    if (f.length < 3) continue;
    const prov = provName(f[iProv]);
    if (!prov) continue;
    locations++;
    const n2 = iL2 >= 0 ? (parseFloat(f[iL2]) || 0) : 0;
    const nd = iDC >= 0 ? (parseFloat(f[iDC]) || 0) : 0;
    const kw = iKw >= 0 ? (parseFloat(f[iKw]) || 0) : 0;
    l2 += n2; dc += nd; totalKw += kw;
    const p = (byProv[prov] = byProv[prov] || { level2: 0, dcFast: 0, kw: 0 });
    p.level2 += n2; p.dcFast += nd; p.kw += kw;
    if (iNet >= 0) {
      const net = (f[iNet] || "Non-networked").trim() || "Non-networked";
      byNetwork[net] = (byNetwork[net] || 0) + 1;
    }
  }
  if (!Object.keys(byProv).length) throw new Error("AFDC CSV parsed but no Canadian province rows found");

  return {
    source: "US DOE Alternative Fuels Data Center (covers Canada)",
    flavor: "stations",
    stationCount: locations,
    rawRowCount: locations,
    headerLine: firstLine.substring(0, 400),
    totalLocations: locations,
    level2Units: Math.round(l2),
    dcFastUnits: Math.round(dc),
    totalKw: Math.round(totalKw),
    byProvince: Object.entries(byProv).map(([province, v]) => ({
      province, level2: Math.round(v.level2), dcFast: Math.round(v.dcFast), kw: Math.round(v.kw),
    })).sort((a, b) => (b.level2 + b.dcFast) - (a.level2 + a.dcFast)),
    topNetworks: Object.entries(byNetwork).sort((a, b) => b[1] - a[1]).slice(0, 10)
      .map(([network, count]) => ({ network, count })),
    url: "https://afdc.energy.gov/stations/states",
  };
}

async function fetchAfdcWithRetries(buildUrl, tmpFile, aggregator) {
  const dl = await downloadToFile(buildUrl(), tmpFile);
  let raw = fs.readFileSync(tmpFile, "utf8").replace(/^\uFEFF/, "");
  if (/429|rate limit|OVER_RATE_LIMIT/i.test(raw.substring(0, 300))) {
    console.log("  rate limited — backing off 60s");
    await sleep(60000);
    await downloadToFile(buildUrl(), tmpFile);
    raw = fs.readFileSync(tmpFile, "utf8").replace(/^\uFEFF/, "");
  }
  const parsed = aggregator(raw);
  parsed.downloadedVia = dl.method;
  return parsed;
}

async function fetchAfdc() {
  if (!AFDC_API_KEY) throw new Error("AFDC_API_KEY not set");
  const enc = encodeURIComponent;
  const tmpFile = path.join(os.tmpdir(), "afdc-units.csv");
  const errs = [];

  try {
    const unitsUrl = () =>
      AFDC_HOST + "/api/alt-fuel-stations/v1/ev-charging-units.csv?access=public&api_key=" + enc(AFDC_API_KEY) +
      "&cards_accepted=all&country=CA&download=true&ev_charging_level=2%2Cdc_fast" +
      "&ev_connector_type=all&ev_network=all&facility_type=all&fuel_type=ELEC" +
      "&funding_sources=all&limit=all&owner_type=all&state=all&status=E&utf8_bom=true";
    const r = await fetchAfdcWithRetries(unitsUrl, tmpFile, aggregateUnitsCsv);
    console.log("  (ev-charging-units via " + r.downloadedVia + ": " + r.rawRowCount + " rows → " + r.stationCount + " stations)");
    r.lastUpdated = await fetchAfdcLastUpdated();
    return r;
  } catch (e) {
    errs.push("ev-charging-units: " + e.message);
    console.log("  units endpoint failed — " + e.message);
  }

  try {
    const allUrl = () =>
      AFDC_HOST + "/api/alt-fuel-stations/v1.csv?access=public&api_key=" + enc(AFDC_API_KEY) +
      "&country=CA&download=true&ev_charging_level=2%2Cdc_fast&fuel_type=ELEC" +
      "&limit=all&status=E&utf8_bom=true";
    const r = await fetchAfdcWithRetries(allUrl, tmpFile, aggregateStationsCsv);
    console.log("  (all-stations via " + r.downloadedVia + ": " + r.stationCount + " stations)");
    r.lastUpdated = await fetchAfdcLastUpdated();
    return r;
  } catch (e) {
    errs.push("all-stations: " + e.message);
  }

  throw new Error("AFDC failed both endpoints — " + errs.join(" || "));
}

async function fetchAfdcLastUpdated() {
  try {
    const txt = await httpGet(
      AFDC_HOST + "/api/alt-fuel-stations/v1/last-updated?api_key=" + encodeURIComponent(AFDC_API_KEY)
    );
    const j = JSON.parse(txt);
    return j.last_updated || null;
  } catch (_) { return null; }
}

// --------------------------------------------------------------- pipeline

function rebuildDataJs() {
  const files = fs.readdirSync(HISTORY_DIR).filter((f) => f.endsWith(".json")).sort();
  if (!files.length) return null;
  const history = files.map((f) => {
    const j = JSON.parse(fs.readFileSync(path.join(HISTORY_DIR, f), "utf-8"));
    return { file: f, fetchedAt: j.fetchedAt, kpis: j.kpis, sources: j.sources };
  });
  const latest = JSON.parse(fs.readFileSync(path.join(HISTORY_DIR, files[files.length - 1]), "utf-8"));
  const js = "// AUTO-GENERATED by fetch-data.js\n" +
    "window.EV_HISTORY = " + JSON.stringify(history) + ";\n" +
    "window.EV_LATEST = " + JSON.stringify(latest) + ";\n";
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(path.join(DATA_DIR, "data.js"), js);
  return { snapshots: files.length, latestAt: latest.fetchedAt };
}

async function main() {
  fs.mkdirSync(HISTORY_DIR, { recursive: true });
  const ts = new Date();
  const snapshot = { fetchedAt: ts.toISOString(), sources: {}, kpis: {}, data: {} };

  const jobs = [
    ["transportCanada", "Transport Canada - EV Dashboard", fetchTransportCanada],
    ["afdc", "US DOE AFDC - Public EV charging (Canada)", fetchAfdc],
  ];

  for (const [key, name, fn] of jobs) {
    process.stdout.write("Fetching " + name + " ... ");
    const rec = { name, lastRefreshed: snapshot.fetchedAt, ok: false, error: null };
    try {
      snapshot.data[key] = await fn();
      rec.ok = true;
      console.log("ok");
    } catch (e) {
      rec.error = String(e.message || e);
      console.log("FAILED - " + rec.error);
    }
    snapshot.sources[key] = rec;
  }

  const tc = snapshot.data.transportCanada || {};
  const afdc = snapshot.data.afdc || {};
  const ms = (tc.marketShare || []).slice(-1)[0] || {};

  snapshot.kpis = {
    totalEvs: tc.kpis?.totalEvs, bevs: tc.kpis?.bevs, phevs: tc.kpis?.phevs,
    totalChargers: tc.kpis?.totalChargers, evPerCharger: tc.kpis?.evPerCharger,
    marketSharePct: ms.total, marketShareQuarter: ms.period,
    afdcLocations: afdc.totalLocations, afdcLevel2: afdc.level2Units,
    afdcDcFast: afdc.dcFastUnits, afdcKw: afdc.totalKw,
  };

  const fname = "snapshot-" + ts.toISOString().replace(/[:.]/g, "-") + ".json";
  fs.writeFileSync(path.join(HISTORY_DIR, fname), JSON.stringify(snapshot, null, 2));
  console.log("Saved data/history/" + fname);
  const stats = rebuildDataJs();
  if (stats) console.log("data.js rebuilt - " + stats.snapshots + " snapshot(s)");

  const k = snapshot.kpis;
  if (k.totalEvs && k.totalChargers) {
    console.log("Sanity: " + k.totalEvs + " EVs / " + k.totalChargers + " chargers = " +
      (k.totalEvs / k.totalChargers).toFixed(1) + ":1");
  }
  if (afdc.totalLocations) {
    console.log("Sanity: AFDC " + ((afdc.level2Units || 0) + (afdc.dcFastUnits || 0)) +
      " units across " + afdc.totalLocations + " stations");
  }

  const failures = Object.values(snapshot.sources).filter((s) => !s.ok);
  if (failures.length) console.warn("NOTE: " + failures.length + " source(s) failed; others committed.");
  console.log("Done.");
}

main().catch((e) => { console.error("Fatal error:", e); process.exit(1); });
