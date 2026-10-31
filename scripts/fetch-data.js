// ============================================================================
// scripts/fetch-data.js — v10 (streamlined)
//
// What works (final stable sources):
//   • Transport Canada (HTML tables) — primary source for ALL EV + charger KPIs
//   • AFDC /v1.csv (Canada only) — province charging units + networks
//   • Open Canada XLSX — funded projects (best-effort, graceful degradation)
//
// What we DROPPED:
//   • StatCan WDS/CSV — kept causing failures and timeouts
//     → ZEV market share now derived from TC's annual ZEV percentages
//
// Runtime: ~2–3 minutes total
// ============================================================================

"use strict";

const fs = require("fs");
const path = require("path");
const os = require("os");
const { execFileSync } = require("child_process");
const AdmZip = require("adm-zip");
const XLSX = require("xlsx");

const ROOT = process.cwd();
const DATA_DIR = path.join(ROOT, "data");
const HISTORY_DIR = path.join(DATA_DIR, "history");

const TC_URL =
  "https://tc.canada.ca/en/road-transportation/innovative-technologies/" +
  "electric-vehicles/canada-electric-vehicle-dashboard";
const CKAN_API_URL =
  "https://open.canada.ca/data/en/api/3/action/package_show" +
  "?id=706a6d1a-df8e-40e3-8e0c-94211025c528";
const AFDC_HOST = "https://developer.nrel.gov";
const AFDC_API_KEY =
  process.env.AFDC_API_KEY || "ZCWrO9fKZWtFzxsMrqfGlCcBnCx1ibQSqvJfhA1y";

const UA = "CanadaEVDashboard/1.0 (GitHub Actions; personal use)";

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
      if (attempt < retries) {
        console.log("  retry " + (attempt + 1) + " after: " + e.message);
        await new Promise((r) => setTimeout(r, 5000 * (attempt + 1)));
      }
    }
  }
  throw lastErr;
}

async function downloadToFile(url, tmpFile, { timeoutMs = 300000 } = {}) {
  try {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), timeoutMs);
    const res = await fetch(url, { headers: { "User-Agent": UA }, signal: ctrl.signal });
    clearTimeout(timer);
    if (res.ok) {
      fs.writeFileSync(tmpFile, Buffer.from(await res.arrayBuffer()));
      return { method: "fetch", httpCode: res.status };
    }
    const body = (await res.text()).substring(0, 200);
    throw new Error("HTTP " + res.status + (body ? " — " + body.replace(/\s+/g, " ") : ""));
  } catch (fetchErr) {
    console.log("  fetch failed (" + fetchErr.message + "), falling back to curl");
  }
  try {
    const httpCode = execFileSync(
      "curl",
      ["-sS", "--retry", "3", "--max-time", String(Math.round(timeoutMs / 1000)), "-A", UA, "-o", tmpFile, "-w", "%{http_code}", url],
      { encoding: "utf8", timeout: timeoutMs + 60000, stdio: ["ignore", "pipe", "pipe"] }
    ).trim();
    const code = parseInt(httpCode, 10);
    if (code !== 200) {
      let bodyExcerpt = "";
      try { bodyExcerpt = fs.readFileSync(tmpFile, "utf8").substring(0, 200); } catch (_) {}
      throw new Error("HTTP " + code + (bodyExcerpt ? " — " + bodyExcerpt.replace(/\s+/g, " ") : ""));
    }
    if (!fs.existsSync(tmpFile) || fs.statSync(tmpFile).size === 0) {
      throw new Error("curl succeeded but produced an empty file");
    }
    return { method: "curl", httpCode: code };
  } catch (curlErr) {
    const stderr = (curlErr.stderr || "").toString().substring(0, 200);
    throw new Error("curl failed: " + String(curlErr.message || "").substring(0, 150) + (stderr ? " | " + stderr : ""));
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
  return String(html || "").replace(/&nbsp;/gi, " ").replace(/&#160;/g, " ").replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
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

// ------------------------------------------- source 1: Transport Canada (PRIMARY)

async function fetchTransportCanada() {
  const html = await httpGet(TC_URL);
  const updatedMatch = html.match(/[Dd]ata was updated in ([A-Z][a-z]+ \d{4})/);
  const updated = updatedMatch ? updatedMatch[1] : null;

  // 1. Extract headline KPIs from prose
  const text = html.replace(/<[^>]*>/g, " ");
  const grab = (re) => { const m = text.match(re); return m ? num(m[1]) : null; };
  const ratio = text.match(/(\d+)\s*EVs?\s*(?:to|per)\s*(?:every\s*)?1\s*charger/i);

  const kpis = {
    totalEvs: grab(/Total EVs?:?\s*([\d,]+)/),
    bevs: grab(/BEVs?:?\s*([\d,]+)/),
    phevs: grab(/PHEVs?:?\s*([\d,]+)/),
    totalChargers: grab(/Total chargers:?\s*([\d,]+)/),
    level2: grab(/Level 2:?\s*([\d,]+)/),
    level3: grab(/Level 3:?\s*([\d,]+)/),
    evPerCharger: ratio ? parseInt(ratio[1], 10) : null,
  };

  const tables = parseTables(html);
  const header = (t) => t[0].map((c) => String(c || "").toLowerCase());
  const hasHeader = (t, ...words) => words.every((w) => header(t).some((c) => c.includes(w)));

  // 2. Market share (annual ZEV trend from TC's annual table)
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

  // 3. Fleet composition
  let fleet = null;
  for (const t of tables) {
    if (hasHeader(t, "vehicle category") || hasHeader(t, "category")) {
      const rows = t.slice(1).map((r) => ({ category: r[0], count: num(r[1]), sharePct: pct(r[2]) }));
      fleet = rows.filter((r) => r.category && r.count != null);
      if (fleet.length) break;
    }
  }
  if (fleet) {
    const sum = (re) => fleet.filter((r) => re.test(r.category)).reduce((a, r) => a + (r.count || 0), 0);
    const fBev = sum(/bev/i), fPhev = sum(/phev/i), fFcev = sum(/fcev/i);
    if (kpis.bevs == null && fBev) kpis.bevs = fBev;
    if (kpis.phevs == null && fPhev) kpis.phevs = fPhev;
    if (kpis.totalEvs == null && (fBev || fPhev)) kpis.totalEvs = fBev + fPhev + fFcev;
  }

  // 4. Charger growth over time
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
  if (chargerGrowth) {
    const last = chargerGrowth[chargerGrowth.length - 1];
    if (kpis.level2 == null && last.level2 != null) kpis.level2 = last.level2;
    if (kpis.level3 == null && last.level3 != null) kpis.level3 = last.level3;
  }
  if (kpis.totalChargers == null && kpis.level2 != null && kpis.level3 != null) {
    kpis.totalChargers = kpis.level2 + kpis.level3;
  }
  if (kpis.evPerCharger == null && kpis.totalEvs && kpis.totalChargers) {
    kpis.evPerCharger = Math.round(kpis.totalEvs / kpis.totalChargers);
  }

  // 5. Province-level data
  let provinces = null;
  for (const t of tables) {
    const h = header(t);
    if (h.some((c) => /province|jurisdiction/.test(c)) && h.some((c) => /charger/.test(c))) {
      const iProv = h.findIndex((c) => /province|jurisdiction/.test(c));
      provinces = t.slice(1).map((r) => {
        const rest = r.slice(iProv + 1);
        const nums = rest.map(num).filter((v) => v != null);
        const ratioM = (rest.find((c) => /per/.test(c)) || "").match(/(\d+)\s*EVs?\s*per/i);
        return {
          province: cellText(r[iProv]),
          evPerCharger: ratioM ? parseInt(ratioM[1], 10) : null,
          totalChargers: nums.length ? nums[nums.length - 2] : null,
          totalEvs: nums.length ? nums[nums.length - 1] : null,
        };
      }).filter((p) => p.province && !/canada/i.test(p.province));
      if (provinces.length >= 5) break;
    }
  }

  return { updated, kpis, marketShare, fleet, provinces, chargerGrowth, url: TC_URL };
}

// ------------------------------------------- source 2: AFDC charging stations

const PROV_NAMES = {
  AB: "Alberta", BC: "British Columbia", MB: "Manitoba", NB: "New Brunswick",
  NL: "Newfoundland and Labrador", NS: "Nova Scotia", NT: "Northwest Territories",
  NU: "Nunavut", ON: "Ontario", PE: "Prince Edward Island", QC: "Quebec",
  SK: "Saskatchewan", YT: "Yukon",
};
const provName = (s) => PROV_NAMES[String(s || "").toUpperCase().trim()] || String(s || "").trim();

async function fetchAfdc() {
  if (!AFDC_API_KEY) throw new Error("AFDC_API_KEY not set");
  const enc = encodeURIComponent;
  const url = AFDC_HOST + "/api/alt-fuel-stations/v1.csv?access=public&api_key=" + enc(AFDC_API_KEY) +
    "&country=CA&fuel_type=ELEC&status=E&limit=all&utf8_bom=true";
  const tmpCsv = path.join(os.tmpdir(), "afdc-stations.csv");
  const dl = await downloadToFile(url, tmpCsv, { timeoutMs: 300000 });
  const raw = fs.readFileSync(tmpCsv, "utf8").replace(/^\uFEFF/, "");
  const rows = raw.split("\n").map((line) => {
    const cells = [], cur = [], inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQ) {
        if (ch === '"') { if (line[i + 1] === '"') { cur.push('"'); i++; } else inQ = false; } else cur.push(ch);
      } else if (ch === '"') inQ = true;
      else if (ch === ",") { cells.push(cur.join("")); cur = []; }
      else cur.push(ch);
    }
    cells.push(cur.join(""));
    return cells;
  });
  if (rows.length < 2) throw new Error("AFDC returned no data rows");
  const head = rows[0].map((h) => String(h).trim().toLowerCase());
  const col = (re) => head.findIndex((h) => re.test(h));
  const iProv = col(/^state$|^province$/);
  const iL2 = col(/level\s*2.*evse|evse.*level\s*2/);
  const iDC = col(/dc\s*fast.*count/i);
  const iNet = col(/^ev network$/);
  if (iProv < 0 || (iL2 < 0 && iDC < 0)) throw new Error("Unexpected AFDC columns: " + head.slice(0, 15).join(" | "));
  const byProv = {}, byNetwork = {};
  let locations = 0, l2Units = 0, dcUnits = 0;
  for (let r = 1; r < rows.length; r++) {
    const f = rows[r], prov = provName(f[iProv]);
    if (!prov) continue;
    locations++;
    const n2 = iL2 >= 0 ? parseInt(f[iL2] || "0", 10) || 0 : 0;
    const nd = iDC >= 0 ? parseInt(f[iDC] || "0", 10) || 0 : 0;
    l2Units += n2; dcUnits += nd;
    const p = (byProv[prov] = byProv[prov] || { locations: 0, level2: 0, dcFast: 0 });
    p.locations++; p.level2 += n2; p.dcFast += nd;
    if (iNet >= 0) {
      const net = (f[iNet] || "Non-networked").trim() || "Non-networked";
      byNetwork[net] = (byNetwork[net] || 0) + 1;
    }
  }
  if (!locations) throw new Error("AFDC CSV parsed but no province data found");
  return {
    source: "US DOE Alternative Fuels Data Center (covers Canada)",
    downloadedVia: dl.method,
    totalLocations: locations,
    level2Units: l2Units, dcFastUnits: dcUnits,
    byProvince: Object.entries(byProv).map(([province, v]) => ({ province, ...v })).sort((a, b) => b.locations - a.locations),
    topNetworks: Object.entries(byNetwork).sort((a, b) => b[1] - a[1]).slice(0, 10).map(([network, count]) => ({ network, count })),
    url: "https://afdc.energy.gov/stations/states",
  };
}

// ------------------------------------------- source 3: Open Canada (best-effort)

async function fetchChargingProjects() {
  const meta = JSON.parse(await httpGet(CKAN_API_URL));
  const resources = (meta.result && meta.result.resources) || [];
  const xlsxRes = resources.filter((r) => /^xls/i.test(String(r.format || "")));
  if (!xlsxRes.length) throw new Error("No XLSX resource in the Open Canada package");
  const res = xlsxRes.find((r) => /eng/i.test(String(r.language || ""))) || xlsxRes[xlsxRes.length - 1];
  const buf = await httpGet(res.url, { asBuffer: true, timeoutMs: 180000 });
  const wb = XLSX.read(buf, { type: "buffer" });
  const findCol = (kws, head) => head.findIndex((c) => kws.some((k) => new RegExp("\\b" + k, "i").test(String(c))));
  let lastErr = "no worksheets with recognizable columns";
  for (const sheetName of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true });
    for (let hr = 0; hr < Math.min(10, rows.length); hr++) {
      const head = (rows[hr] || []).map((c) => String(c));
      if (!head.some((c) => c && c !== "undefined")) continue;
      const iProv = findCol(["province", "territor", "jurisdiction"], head);
      const iCnt = findCol(["charger", "borne", "port", "connector", "unit"], head);
      if (iProv < 0 || iCnt < 0) {
        lastErr = "sheet '" + sheetName + "' row " + (hr + 1) + ": " + head.slice(0, 10).join(" | ");
        continue;
      }
      const byProv = {}, byStatus = {};
      let total = 0, projects = 0;
      for (let r = hr + 1; r < rows.length; r++) {
        const provRaw = rows[r][iProv];
        if (provRaw == null || String(provRaw).trim() === "") continue;
        const m = String(rows[r][iCnt] ?? "").match(/\d+/);
        const cnt = m ? parseFloat(m[0]) : 0;
        total += cnt; projects++;
        byProv[String(provRaw).trim()] = (byProv[String(provRaw).trim()] || 0) + cnt;
      }
      if (projects) {
        return {
          resourceName: "NRCan funded projects (ZEVIP/EVAFIDI/CHRI)",
          totalFundedChargers: Math.round(total),
          projectCount: projects,
          chargersByProvince: Object.entries(byProv).map(([p, v]) => ({ province: p, chargers: Math.round(v) })).sort((a, b) => b.chargers - a.chargers),
          url: "https://open.canada.ca/data/en/dataset/706a6d1a-df8e-40e3-8e0c-94211025c528",
        };
      }
    }
  }
  throw new Error("Downloaded XLSX but no usable data. " + lastErr);
}

// --------------------------------------------------------------- pipeline

function rebuildDataJs() {
  const files = fs.readdirSync(HISTORY_DIR).filter((f) => f.endsWith(".json")).sort();
  if (!files.length) return null;
  const history = files.map((f) => JSON.parse(fs.readFileSync(path.join(HISTORY_DIR, f), "utf-8")));
  const latest = history[history.length - 1];
  const js = "window.EV_HISTORY = " + JSON.stringify(history) + ";\nwindow.EV_LATEST = " + JSON.stringify(latest) + ";\n";
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
    ["afdc", "US DOE AFDC - Public EV charging stations (Canada)", fetchAfdc],
    ["chargingProjects", "Open Gov - NRCan funded projects", fetchChargingProjects],
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
  const cp = snapshot.data.chargingProjects || {};
  const afdc = snapshot.data.afdc || {};

  // Derive ZEV market share from TC's annual data (dropped StatCan)
  const ms = (tc.marketShare || []).slice(-1)[0] || {};
  snapshot.kpis = {
    totalEvs: tc.kpis?.totalEvs, bevs: tc.kpis?.bevs, phevs: tc.kpis?.phevs,
    totalChargers: tc.kpis?.totalChargers, evPerCharger: tc.kpis?.evPerCharger,
    marketSharePct: ms.total, marketShareQuarter: ms.period, // derived from TC, not StatCan
    fundedChargers: cp.totalFundedChargers, fundedProjects: cp.projectCount,
    afdcLocations: afdc.totalLocations, afdcLevel2: afdc.level2Units, afdcDcFast: afdc.dcFastUnits,
  };

  const fname = "snapshot-" + ts.toISOString().replace(/[:.]/g, "-") + ".json";
  fs.writeFileSync(path.join(HISTORY_DIR, fname), JSON.stringify(snapshot, null, 2));
  console.log("Saved data/history/" + fname);
  const stats = rebuildDataJs();
  if (stats) console.log("data.js rebuilt - " + stats.snapshots + " snapshot(s)");

  const failures = Object.values(snapshot.sources).filter((s) => !s.ok);
  if (failures.length) console.warn("NOTE: " + failures.length + " source(s) failed; others committed.");
  console.log("Done.");
}

main().catch((e) => { console.error("Fatal error:", e); process.exit(1); });
