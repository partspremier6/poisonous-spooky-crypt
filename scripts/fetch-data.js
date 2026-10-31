// scripts/fetch-data.js
// Fetches Canadian EV data from three government sources
// Runs on GitHub Actions, commits results back to the repo

"use strict";

const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");
const XLSX = require("xlsx");
const { execSync } = require("child_process");

const ROOT = process.cwd();
const HISTORY_DIR = path.join(ROOT, "data", "history");

const TC_URL =
  "https://tc.canada.ca/en/road-transportation/innovative-technologies/" +
  "electric-vehicles/canada-electric-vehicle-dashboard";
const STATCAN_CSV_URL =
  "https://www150.statcan.gc.ca/n1/tbl/csv/20100025-eng.zip";
const CKAN_API_URL =
  "https://open.canada.ca/api/3/action/package_show" +
  "?id=706a6d1a-df8e-40e3-8e0c-94211025c528";

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

function parseCsv(text) {
  const rows = [];
  let row = [], cur = "", inQ = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"') {
        if (text[i + 1] === '"') { cur += '"'; i++; }
        else inQ = false;
      } else cur += ch;
    } else if (ch === '"') inQ = true;
    else if (ch === ",") { row.push(cur); cur = ""; }
    else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(cur); cur = "";
      if (row.some((c) => c !== "")) rows.push(row);
      row = [];
    } else cur += ch;
  }
  row.push(cur);
  if (row.some((c) => c !== "")) rows.push(row);
  return rows;
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
  const entry = zip.getEntries().find((e) => /\.csv$/i.test(e.entryName) && !/meta/i.test(e.entryName));
  if (!entry) throw new Error("CSV not found inside StatCan zip archive");

  const text = zip.readAsText(entry, "utf-8").replace(/^\uFEFF/, "");
  const rows = parseCsv(text);
  const head = rows[0];
  const idx = (name) => head.findIndex((h) => h.trim() === name);
  const iRef = idx("REF_DATE"), iGeo = idx("GEO"), iVeh = idx("Vehicle type"),
        iFuel = idx("Fuel type"), iUom = idx("UOM"), iVal = idx("VALUE");

  const series = {};
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const g = (i) => (i >= 0 ? row[i] : "");
    if (g(iGeo) !== "Canada") continue;
    if (iUom >= 0 && !/number/i.test(g(iUom))) continue;
    if (!/^total/i.test(g(iVeh))) continue;
    const val = parseFloat(g(iVal));
    if (!isFinite(val)) continue;
    const quarter = g(iRef).trim();
    (series[quarter] = series[quarter] || {})[g(iFuel).trim()] = val;
  }

  const registrations = Object.keys(series).sort().map((q) => {
    const d = series[q];
    const bev = d["Battery electric"] || 0;
    const phev = d["Plug-in hybrid electric"] || 0;
    const totalEntry = Object.keys(d).find((k) => /^total/i.test(k));
    const total = totalEntry ? d[totalEntry] : null;
    const zev = bev + phev;
    return { quarter: q, bev, phev, total, zevSharePct: total ? +(100 * zev / total).toFixed(1) : null };
  });

  return { tableId: "20-10-0025-01", lastQuarter: registrations.length ? registrations[registrations.length - 1].quarter : null, registrations, url: "https://www150.statcan.gc.ca/t1/tbl1/en/tv.action?pid=2010002501" };
}

// -------------------------------------------- source 3: Open Canada charging projects

async function fetchChargingProjects() {
  const meta = JSON.parse(await getText(CKAN_API_URL));
  const resources = (meta.result && meta.result.resources) || [];
  const xlsxRes = resources.filter((r) => /^xls/i.test(String(r.format || "")));
  if (!xlsxRes.length) throw new Error("No XLSX resource found in the Open Canada package");

  const res = xlsxRes.find((r) => /eng/i.test(String(r.language || ""))) || xlsxRes[xlsxRes.length - 1];
  const buf = await getBuffer(res.url);
  const wb = XLSX.read(buf, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  const head = rows[0].map((c) => String(c));

  const findCol = (keywords) => head.findIndex((c) => keywords.some((k) => c.toLowerCase().includes(k)));
  const iProv = findCol(["province", "prov", "state"]);
  const iCnt = findCol(["charger", "borne", "number of"]);
  const iStatus = findCol(["status"]);

  if (iProv < 0 || iCnt < 0) throw new Error("Could not find province/charger columns.");

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
    if (iStatus >= 0) { const s = String(rows[r][iStatus] || "").trim(); if (s) byStatus[s] = (byStatus[s] || 0) + 1; }
  }

  const chargersByProvince = Object.entries(byProv).map(([province, chargers]) => ({ province, chargers: Math.round(chargers) })).sort((a, b) => b.chargers - a.chargers);

  return { resourceName: "Projects funded by the EVAFIDI, ZEVIP and CHRI programs (NRCan)", resourceUrl: res.url, totalFundedChargers: Math.round(total), projectCount: projects, chargersByProvince, projectsByStatus: byStatus, url: "https://open.canada.ca/data/en/dataset/706a6d1a-df8e-40e3-8e0c-94211025c528" };
}

// ------------------------------------------------------------------ pipeline

function rebuildDataJs() {
  const files = fs.readdirSync(HISTORY_DIR).filter((f) => f.endsWith(".json")).sort();
  if (!files.length) return null;

  const history = files.map((f) => {
    const j = JSON.parse(fs.readFileSync(path.join(HISTORY_DIR, f), "utf-8"));
    return { file: f, fetchedAt: j.fetchedAt, kpis: j.kpis, sources: j.sources };
  });

  const latest = JSON.parse(fs.readFileSync(path.join(HISTORY_DIR, files[files.length - 1]), "utf-8"));
  const js = "// AUTO-GENERATED by fetch-data.js\nwindow.EV_HISTORY = " + JSON.stringify(history) + ";\nwindow.EV_LATEST = " + JSON.stringify(latest) + ";\n";

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
    totalEvs: tc.kpis ? tc.kpis.totalEvs : null, bevs: tc.kpis ? tc.kpis.bevs : null, phevs: tc.kpis ? tc.kpis.phevs : null,
    totalChargers: tc.kpis ? tc.kpis.totalChargers : null, evPerCharger: tc.kpis ? tc.kpis.evPerCharger : null,
    marketSharePct: lastReg.zevSharePct != null ? lastReg.zevSharePct : null, marketShareQuarter: lastReg.quarter || null,
    fundedChargers: cp.totalFundedChargers != null ? cp.totalFundedChargers : null, fundedProjects: cp.projectCount != null ? cp.projectCount : null,
  };

  const fname = "snapshot-" + ts.toISOString().replace(/[:.]/g, "-") + ".json";
  fs.writeFileSync(path.join(HISTORY_DIR, fname), JSON.stringify(snapshot, null, 2));
  console.log("Saved " + path.join("data", "history", fname));

  rebuildDataJs();
  console.log("Done.");

  const failures = Object.values(snapshot.sources).filter((s) => !s.ok);
  if (failures.length) console.warn("NOTE: " + failures.length + " source(s) failed this run.");
  process.exit(failures.length > 0 ? 1 : 0);
}

main().catch((e) => { console.error("Fatal error:", e); process.exit(1); });
