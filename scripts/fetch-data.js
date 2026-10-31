// ============================================================================
// FILE: scripts/fetch-data.js   (FIXED VERSION - Nov 2026)
// ----------------------------------------------------------------------------
// Fixes in this version:
//   1. StatCan: the bulk CSV is ~hundreds of MB (covers every census
//      subdivision). Loading it as one string exceeded Node's ~512MB string
//      limit. Now parsed as a byte-stream, line by line, keeping only
//      Canada-level rows -> tiny memory footprint.
//   2. Open Canada: the CKAN API base URL was wrong (open.canada.ca/api/...
//      returns 404). Correct base is open.canada.ca/data/en/api/3/action/...
//      The script now tries several known bases, and if none respond it
//      falls back to scraping the dataset's own page for the direct XLSX
//      download link.
//
// Replace the entire content of scripts/fetch-data.js in the GitHub repo
// with this file, commit, then run the workflow from the Actions tab.
// ============================================================================

"use strict";

const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");
const XLSX = require("xlsx");

const ROOT = process.cwd();
const HISTORY_DIR = path.join(ROOT, "data", "history");

const TC_URL =
  "https://tc.canada.ca/en/road-transportation/innovative-technologies/" +
  "electric-vehicles/canada-electric-vehicle-dashboard";
const STATCAN_CSV_URL =
  "https://www150.statcan.gc.ca/n1/tbl/csv/20100025-eng.zip";
const OPENCANADA_PACKAGE_ID = "706a6d1a-df8e-40e3-8e0c-94211025c528";
const OPENCANADA_DATASET_PAGE =
  "https://open.canada.ca/data/en/dataset/" + OPENCANADA_PACKAGE_ID;
// Known CKAN API base URLs for open.canada.ca, tried in order
const CKAN_BASES = [
  "https://open.canada.ca/data/en/api/3/action/package_show",
  "https://open.canada.ca/data/fr/api/3/action/package_show",
  "https://open.canada.ca/data/api/3/action/package_show",
];

const UA = "LocalEVDashboard/1.0 (auto-fetched via GitHub Actions)";

// ---------------------------------------------------------------- utilities

async function getText(url) {
  const res = await fetch(url, { headers: { "User-Agent": UA } });
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  return res.text();
}

async function getBuffer(url) {
  const res = await fetch(url, { headers: { "User-Agent": UA } });
  if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
  return Buffer.from(await res.arrayBuffer());
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

function decodeEntities(s) {
  return s.replace(/&nbsp;/gi, " ").replace(/&#160;/g, " ").trim();
}

function cellText(html) {
  return decodeEntities(html.replace(/<[^>]*>/g, " ")).replace(/\s+/g, " ");
}

function parseTables(html) {
  const tables = [];
  const tRe = /<table[^>]*>([\s\S]*?)<\/table>/gi;
  let tm;
  while ((tm = tRe.exec(html))) {
    const rows = [];
    const rRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    let rm;
    while ((rm = rRe.exec(tm[1]))) {
      const cells = [];
      const cRe = /<t[hd][^>]*>([\s\S]*?)<\/t[hd]>/gi;
      let cm;
      while ((cm = cRe.exec(rm[1]))) cells.push(cellText(cm[1]));
      if (cells.length) rows.push(cells);
    }
    if (rows.length) tables.push(rows);
  }
  return tables;
}

// Parse ONE line of CSV (handles quoted fields containing commas).
// Used instead of parsing the whole file as a single string.
function parseCsvLine(line) {
  const out = [];
  let cur = "", inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQ) {
      if (ch === '"') {
        if (line[i + 1] === '"') { cur += '"'; i++; }
        else inQ = false;
      } else cur += ch;
    } else if (ch === '"') inQ = true;
    else if (ch === ",") { out.push(cur); cur = ""; }
    else cur += ch;
  }
  out.push(cur);
  return out;
}

// Iterate lines of a (possibly huge) Buffer WITHOUT building one giant
// string: decode byte-aligned chunks (cut only at newline bytes, which can
// never split a multi-byte UTF-8 character mid-sequence).
function* iterateLines(buffer) {
  const CHUNK = 16 * 1024 * 1024; // 16 MB at a time
  let start = 0;
  let carry = "";
  while (start < buffer.length) {
    let end = Math.min(start + CHUNK, buffer.length);
    if (end < buffer.length) {
      // extend/shrink to a real newline byte so we decode whole lines
      const nlByte = buffer.indexOf(10, end);
      end = nlByte === -1 ? buffer.length : nlByte;
    }
    const text = carry + buffer.toString("utf8", start, end);
    start = end + 1;
    let from = 0, nl;
    while ((nl = text.indexOf("\n", from)) !== -1) {
      yield text.slice(from, nl);
      from = nl + 1;
    }
    carry = text.slice(from);
  }
  if (carry) yield carry;
}

// ------------------------------------------------------ source 1: Transport Canada

async function fetchTransportCanada() {
  const html = await getText(TC_URL);
  const updatedMatch = html.match(/[Dd]ata was updated in ([A-Z][a-z]+ \d{4})/);
  const updated = updatedMatch ? updatedMatch[1] : null;
  const text = cellText(html);
  const grab = (re) => { const m = text.match(re); return m ? num(m[1]) : null; };
  const ratio = text.match(/ratio[^\d]{0,40}(\d+)\D{1,15}(\d+)/);

  const kpis = {
    totalEvs: grab(/Total EVs:?\s*([\d,]+)/),
    bevs: grab(/BEVs:?\s*([\d,]+)/),
    phevs: grab(/PHEVs:?\s*([\d,]+)/),
    totalChargers: grab(/Total chargers:?\s*([\d,]+)/),
    level2: grab(/Level 2:?\s*([\d,]+)/),
    level3: grab(/Level 3:?\s*([\d,]+)/),
    evPerCharger: ratio ? Math.round(parseInt(ratio[1], 10) / parseInt(ratio[2], 10)) : null,
  };

  const tables = parseTables(html);
  const header = (t) => t[0].map((c) => c.toLowerCase());

  let marketShare = null;
  for (const t of tables) {
    const h = header(t);
    if (h.includes("year") && h.includes("bev")) {
      marketShare = t.slice(1).map((r) => ({
        period: r[0], bev: pct(r[1]), phev: pct(r[2]), fcev: pct(r[3]), total: pct(r[4]),
      })).filter((x) => x.total != null);
      break;
    }
  }

  let fleet = null;
  for (const t of tables) {
    const h = header(t);
    if (h.some((c) => c.includes("vehicle category")) && h.some((c) => c.includes("market share"))) {
      fleet = t.slice(1).map((r) => ({ category: r[0], count: num(r[1]), sharePct: pct(r[2]) }));
      break;
    }
  }

  let imhzev = null;
  for (const t of tables) {
    const h = header(t);
    if (h.includes("year") && h.includes("qc") && h.includes("total")) {
      imhzev = t.slice(1).map((r) => ({
        year: r[0], qc: num(r[1]), bc: num(r[2]), on: num(r[3]), rest: num(r[4]), total: num(r[5]),
      }));
      break;
    }
  }

  return { updated, kpis, marketShare, fleet, imhzev, url: TC_URL };
}

// ------------------------------------------------------------- source 2: StatCan

async function fetchStatcan() {
  const buf = await getBuffer(STATCAN_CSV_URL);
  const zip = new AdmZip(buf);
  const entry = zip
    .getEntries()
    .find((e) => /\.csv$/i.test(e.entryName) && !/meta/i.test(e.entryName));
  if (!entry) throw new Error("CSV not found inside StatCan zip archive");

  // The unzipped CSV can exceed 500 MB (all census subdivisions), which is
  // larger than Node's maximum string length - so we never materialize it
  // as a single string. Extract as a raw Buffer and walk it line by line.
  const data = entry.getData();

  const series = {};
  let headerParsed = false;
  let iRef, iGeo, iVeh, iFuel, iUom, iVal;
  let kept = 0;

  for (const rawLine of iterateLines(data)) {
    const line = rawLine.charCodeAt(rawLine.length - 1) === 13
      ? rawLine.slice(0, -1) : rawLine;
    if (!line) continue;

    if (!headerParsed) {
      const head = parseCsvLine(line);
      const idx = (name) => head.findIndex((h) => h.trim() === name);
      iRef = idx("REF_DATE"); iGeo = idx("GEO"); iVeh = idx("Vehicle type");
      iFuel = idx("Fuel type"); iUom = idx("UOM"); iVal = idx("VALUE");
      if (iGeo < 0 || iVal < 0 || iFuel < 0) {
        throw new Error("Unexpected StatCan CSV header: " + line.substring(0, 200));
      }
      headerParsed = true;
      continue;
    }

    // cheap pre-filter: only lines whose GEO is Canada (field is quoted in
    // StatCan bulk CSVs). Skips millions of sub-national rows instantly.
    if (!line.includes('"Canada"')) continue;

    const row = parseCsvLine(line);
    const g = (i) => (i >= 0 ? row[i] : "");
    if (g(iGeo) !== "Canada") continue;
    if (iUom >= 0 && !/number/i.test(g(iUom))) continue;
    if (!/^total/i.test(g(iVeh))) continue; // "Total, new motor vehicles"

    const val = parseFloat(g(iVal));
    if (!isFinite(val)) continue;

    const quarter = g(iRef).trim();
    (series[quarter] = series[quarter] || {})[g(iFuel).trim()] = val;
    kept++;
  }

  if (!kept) throw new Error("Parsed StatCan CSV but found 0 Canada-level rows");

  const registrations = Object.keys(series).sort().map((q) => {
    const d = series[q];
    const bev = d["Battery electric"] || 0;
    const phev = d["Plug-in hybrid electric"] || 0;
    const totalEntry = Object.keys(d).find((k) => /^total/i.test(k));
    const total = totalEntry ? d[totalEntry] : null;
    const zev = bev + phev;
    return {
      quarter: q, bev, phev, total,
      zevSharePct: total ? +(100 * zev / total).toFixed(1) : null,
    };
  });

  return {
    tableId: "20-10-0025-01",
    rowsKept: kept,
    lastQuarter: registrations.length
      ? registrations[registrations.length - 1].quarter
      : null,
    registrations,
    url: "https://www150.statcan.gc.ca/t1/tbl1/en/tv.action?pid=2010002501",
  };
}

// -------------------------------------------- source 3: Open Canada charging projects

async function findXlsxResource() {
  // 1. Try the known CKAN API bases (correct one is /data/en/api/3/action/)
  for (const base of CKAN_BASES) {
    try {
      const txt = await getText(base + "?id=" + OPENCANADA_PACKAGE_ID);
      let meta;
      try { meta = JSON.parse(txt); } catch (e) { continue; }
      const resources = (meta.result && meta.result.resources) || [];
      const xlsxRes = resources.filter((r) => /^xls/i.test(String(r.format || "")));
      if (xlsxRes.length) {
        const res = xlsxRes.find((r) => /eng/i.test(String(r.language || ""))) ||
                    xlsxRes[xlsxRes.length - 1];
        return { url: res.url, via: base };
      }
    } catch (e) {
      // wrong base or temporary failure - try the next one
    }
  }

  // 2. Fallback: scrape the dataset's public page for a direct XLSX link
  const html = await getText(OPENCANADA_DATASET_PAGE);
  const matches = [...html.matchAll(/href="([^"]+\.xlsx[^"]*)"/gi)];
  if (!matches.length) {
    throw new Error("Could not locate the XLSX resource via the CKAN API or the dataset page");
  }
  const href = matches[0][1].replace(/&amp;/g, "&");
  return { url: new URL(href, OPENCANADA_DATASET_PAGE).href, via: "dataset page" };
}

async function fetchChargingProjects() {
  const { url: xlsxUrl, via } = await findXlsxResource();

  const buf = await getBuffer(xlsxUrl);
  const wb = XLSX.read(buf, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  const head = rows[0].map((c) => String(c));

  const findCol = (keywords) =>
    head.findIndex((c) => keywords.some((k) => c.toLowerCase().includes(k)));

  const iProv = findCol(["province", "prov", "state"]);
  const iCnt = findCol(["charger", "borne", "number of"]);
  const iStatus = findCol(["status"]);

  if (iProv < 0 || iCnt < 0)
    throw new Error(
      "Downloaded the XLSX but could not find province/charger columns. Header: " +
        head.slice(0, 15).join(" | ")
    );

  const byProv = {}, byStatus = {};
  let total = 0, projects = 0;
  for (let r = 1; r < rows.length; r++) {
    const provRaw = rows[r][iProv];
    if (provRaw == null || String(provRaw).trim() === "") continue;
    const cntMatch = String(rows[r][iCnt]).match(/\d+(\.\d+)?/);
    const cnt = cntMatch ? parseFloat(cntMatch[0]) : 0;
    total += cnt;
    projects++;
    const p = String(provRaw).trim();
    byProv[p] = (byProv[p] || 0) + cnt;
    if (iStatus >= 0) {
      const s = String(rows[r][iStatus] || "").trim();
      if (s) byStatus[s] = (byStatus[s] || 0) + 1;
    }
  }

  const chargersByProvince = Object.entries(byProv)
    .map(([province, chargers]) => ({ province, chargers: Math.round(chargers) }))
    .sort((a, b) => b.chargers - a.chargers);

  return {
    resourceName: "Projects funded by the EVAFIDI, ZEVIP and CHRI programs (NRCan)",
    resourceUrl: xlsxUrl,
    resourceVia: via,
    totalFundedChargers: Math.round(total),
    projectCount: projects,
    chargersByProvince,
    projectsByStatus: byStatus,
    url: OPENCANADA_DATASET_PAGE,
  };
}

// ------------------------------------------------------------------ pipeline

function rebuildDataJs() {
  const files = fs.readdirSync(HISTORY_DIR).filter((f) => f.endsWith(".json")).sort();
  if (!files.length) return null;

  const history = files.map((f) => {
    const j = JSON.parse(fs.readFileSync(path.join(HISTORY_DIR, f), "utf-8"));
    return { file: f, fetchedAt: j.fetchedAt, kpis: j.kpis, sources: j.sources };
  });

  const latest = JSON.parse(
    fs.readFileSync(path.join(HISTORY_DIR, files[files.length - 1]), "utf-8")
  );
  const js =
    "// AUTO-GENERATED by fetch-data.js\n" +
    "window.EV_HISTORY = " + JSON.stringify(history) + ";\n" +
    "window.EV_LATEST = " + JSON.stringify(latest) + ";\n";

  const dataDir = path.join(ROOT, "data");
  if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });
  fs.writeFileSync(path.join(dataDir, "data.js"), js);
  return { snapshots: files.length, latestAt: latest.fetchedAt };
}

async function main() {
  fs.mkdirSync(HISTORY_DIR, { recursive: true });
  const ts = new Date();
  const snapshot = { fetchedAt: ts.toISOString(), sources: {}, kpis: {}, data: {} };

  const jobs = [
    ["transportCanada", "Transport Canada - Canada's Electric Vehicle Dashboard", fetchTransportCanada],
    ["statcan", "Statistics Canada - New motor vehicle registrations (20-10-0025-01)", fetchStatcan],
    ["chargingProjects", "Open Government - NRCan ZEVIP/EVAFIDI/CHRI funded projects", fetchChargingProjects],
  ];

  for (const [key, name, fn] of jobs) {
    process.stdout.write("Fetching " + name + " ... ");
    const rec = { name, lastRefreshed: snapshot.fetchedAt, ok: false, error: null };
    try {
      snapshot.data[key] = await fn();
      rec.ok = true;
      console.log("ok");
    } catch (e) {
      rec.error = String((e && e.message) || e);
      console.log("FAILED - " + rec.error);
    }
    snapshot.sources[key] = rec;
  }

  const tc = snapshot.data.transportCanada || {};
  const cp = snapshot.data.chargingProjects || {};
  const regs = (snapshot.data.statcan || {}).registrations || [];
  const lastReg = regs.length ? regs[regs.length - 1] : {};

  snapshot.kpis = {
    totalEvs: tc.kpis ? tc.kpis.totalEvs : null,
    bevs: tc.kpis ? tc.kpis.bevs : null,
    phevs: tc.kpis ? tc.kpis.phevs : null,
    totalChargers: tc.kpis ? tc.kpis.totalChargers : null,
    evPerCharger: tc.kpis ? tc.kpis.evPerCharger : null,
    marketSharePct: lastReg.zevSharePct != null ? lastReg.zevSharePct : null,
    marketShareQuarter: lastReg.quarter || null,
    fundedChargers: cp.totalFundedChargers != null ? cp.totalFundedChargers : null,
    fundedProjects: cp.projectCount != null ? cp.projectCount : null,
  };

  const fname = "snapshot-" + ts.toISOString().replace(/[:.]/g, "-") + ".json";
  fs.writeFileSync(path.join(HISTORY_DIR, fname), JSON.stringify(snapshot, null, 2));
  console.log("Saved " + path.join("data", "history", fname));

  rebuildDataJs();
  console.log("Done.");

  const failures = Object.values(snapshot.sources).filter((s) => !s.ok);
  if (failures.length) console.warn("NOTE: " + failures.length + " source(s) failed this run.");
}

main().catch((e) => { console.error("Fatal error:", e); process.exit(1); });
