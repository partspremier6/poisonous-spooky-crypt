// ============================================================================
// scripts/fetch-data.js — v8
// Canada EV Dashboard data fetcher (runs on GitHub Actions)
//
// v8 fixes:
//   - StatCan: strip the UTF-8 BOM from the header line (this was the
//     "Unexpected StatCan header: ﻿REF_DATE" failure)
//   - AFDC: use the canonical station-level endpoint (/v1.csv) with a known
//     column schema, plus a curl fallback when Node's fetch fails to connect
//   - Open Gov XLSX: scan the first 10 rows for the real header row
//     (presentation sheets bury headers under titles), broader column matching
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
const STATCAN_CSV_URL =
  "https://www150.statcan.gc.ca/n1/tbl/csv/20100025-eng.zip";
const CKAN_API_URL =
  "https://open.canada.ca/data/en/api/3/action/package_show" +
  "?id=706a6d1a-df8e-40e3-8e0c-94211025c528";
const AFDC_HOST = "https://developer.nrel.gov";
   const AFDC_API_KEY = "ZCWrO9fKZWtFzxsMrqfGlCcBnCx1ibQSqvJfhA1y";
const UA = "CanadaEVDashboard/1.0 (GitHub Actions; personal use)";

// ---------------------------------------------------------------- utilities

async function httpGet(url, { asBuffer = false, timeoutMs = 120000, retries = 2 } = {}) {
  let lastErr;
  for (let attempt = 0; attempt <= retries; attempt++) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), timeoutMs);
    try {
      const res = await fetch(url, {
        headers: { "User-Agent": UA },
        signal: ctrl.signal,
      });
      clearTimeout(timer);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return asBuffer ? Buffer.from(await res.arrayBuffer()) : await res.text();
    } catch (e) {
      clearTimeout(timer);
      lastErr = e;
      if (attempt < retries) {
        console.log(`  retry ${attempt + 1} after: ${e.message}`);
        await new Promise((r) => setTimeout(r, 5000 * (attempt + 1)));
      }
    }
  }
  throw lastErr;
}

// Robust download: fetch first, then curl as fallback (fetch sometimes fails
// to CONNECT against certain hosts on GitHub runners — "fetch failed").
function robustDownload(url, tmpFile, timeoutMs = 300000) {
  try {
    const buf = await0 => null; // placeholder, see below
  } catch (e) {}
}

async function downloadToFile(url, tmpFile, { timeoutMs = 300000 } = {}) {
  // 1) Try Node fetch
  try {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), timeoutMs);
    const res = await fetch(url, {
      headers: { "User-Agent": UA },
      signal: ctrl.signal,
    });
    clearTimeout(timer);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buf = Buffer.from(await res.arrayBuffer());
    fs.writeFileSync(tmpFile, buf);
    return { method: "fetch" };
  } catch (fetchErr) {
    console.log(`  fetch failed (${fetchErr.message}), falling back to curl`);
  }
  // 2) curl fallback (preinstalled on ubuntu runners)
  try {
    execFileSync(
      "curl",
      [
        "-sSL", "--fail",
        "--retry", "3",
        "--max-time", String(Math.round(timeoutMs / 1000)),
        "-A", UA,
        "-o", tmpFile,
        url,
      ],
      { stdio: ["ignore", "inherit", "inherit"], timeout: timeoutMs + 60000 }
    );
    if (!fs.existsSync(tmpFile) || fs.statSync(tmpFile).size === 0) {
      throw new Error("curl produced an empty file");
    }
    return { method: "curl" };
  } catch (curlErr) {
    throw new Error(
      `download failed via both fetch and curl: ${curlErr.message}`
    );
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
function decodeEntities(s) {
  return s
    .replace(/&nbsp;/gi, " ").replace(/&#160;/g, " ")
    .replace(/&amp;/gi, "&").replace(/&lt;/gi, "<").replace(/&gt;/gi, ">")
    .replace(/&apos;/gi, "'").replace(/&quot;/gi, '"')
    .trim();
}
function cellText(html) {
  return decodeEntities(String(html).replace(/<[^>]*>/g, " ")).replace(/\s+/g, " ");
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
        if (text[i + 1] === '"') { cur += '"'; i++; } else inQ = false;
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

// Stream CSV lines from a Buffer without building one giant string.
function streamCsvLines(buf, lineCb) {
  const CHUNK = 1 << 20;
  let pos = 0;
  let leftover = null;
  let isFirstLine = true;
  while (pos < buf.length) {
    let nl = -1;
    const end = Math.min(pos + CHUNK, buf.length);
    for (let i = end - 1; i >= pos; i--) { if (buf[i] === 10) { nl = i; break; } }
    if (nl === -1) { pos = end; continue; }
    const slice = buf.slice(pos, nl + 1);
    const piece = leftover ? Buffer.concat([leftover, slice]) : slice;
    const s = piece.toString("utf-8");
    const lines = s.split("\n");
    const carry = lines.pop();
    for (const ln of lines) {
      let t = ln.replace(/\r$/, "");
      // FIX (v8): strip UTF-8 BOM from the very first line if present
      if (isFirstLine) { t = t.replace(/^\uFEFF/, ""); isFirstLine = false; }
      if (t.trim() !== "") lineCb(t);
    }
    leftover = carry && carry.length ? Buffer.from(carry, "utf-8") : null;
    pos = nl + 1;
  }
  if (leftover && leftover.length) {
    let t = leftover.toString("utf-8").replace(/\r$/, "");
    if (isFirstLine) t = t.replace(/^\uFEFF/, "");
    if (t.trim() !== "") lineCb(t);
  }
}

// ------------------------------------------- source 1: Transport Canada

async function fetchTransportCanada() {
  const html = await httpGet(TC_URL);
  const updatedMatch = html.match(/[Dd]ata was updated in ([A-Z][a-z]+ \d{4})/);
  const updated = updatedMatch ? updatedMatch[1] : null;

  const text = cellText(html);
  const grab = (re) => { const m = text.match(re); return m ? num(m[1]) : null; };
  let kpis = {
    totalEvs: grab(/Total EVs?:?\s*([\d,]+)/),
    bevs: grab(/BEVs?:?\s*([\d,]+)/),
    phevs: grab(/PHEVs?:?\s*([\d,]+)/),
    totalChargers: grab(/Total chargers:?\s*([\d,]+)/),
    level2: grab(/Level 2:?\s*([\d,]+)/),
    level3: grab(/Level 3:?\s*([\d,]+)/),
    evPerCharger: null,
  };
  const ratio = text.match(/(\d+)\s*EVs?\s*(?:to|per)\s*(?:every\s*)?1\s*charger/i);
  if (ratio) kpis.evPerCharger = parseInt(ratio[1], 10);

  const tables = parseTables(html);
  const header = (t) => t[0].map((c) => c.toLowerCase());
  const hasHeader = (t, ...words) => words.every((w) => header(t).some((c) => c.includes(w)));

  let marketShare = null;
  for (const t of tables) {
    if (hasHeader(t, "year") && hasHeader(t, "bev")) {
      marketShare = t.slice(1).map((r) => ({
        period: r[0], bev: pct(r[1]), phev: pct(r[2]), fcev: pct(r[3]), total: pct(r[4]),
      })).filter((x) => x.total != null);
      break;
    }
  }

  let fleet = null;
  for (const t of tables) {
    if (hasHeader(t, "vehicle category") || hasHeader(t, "category")) {
      const rows = t.slice(1).map((r) => ({ category: r[0], count: num(r[1]), sharePct: pct(r[2]) }));
      const named = rows.filter((r) => r.category && r.count != null);
      if (named.length) { fleet = named; break; }
    }
  }
  if (fleet) {
    const sum = (re) => fleet.filter((r) => re.test(r.category)).reduce((a, r) => a + (r.count || 0), 0);
    const fBev = sum(/bev/i), fPhev = sum(/phev/i), fFcev = sum(/fcev/i);
    if (kpis.bevs == null && fBev) kpis.bevs = fBev;
    if (kpis.phevs == null && fPhev) kpis.phevs = fPhev;
    if (kpis.totalEvs == null && (fBev || fPhev)) kpis.totalEvs = fBev + fPhev + fFcev;
  }

  let chargerGrowth = null;
  for (const t of tables) {
    const h = header(t);
    const iL2 = h.findIndex((c) => /level\s*2/.test(c));
    const iL3 = h.findIndex((c) => /level\s*3|dc\s*fast/.test(c));
    const iYr = h.findIndex((c) => /year|period|date/.test(c));
    if (iL2 >= 0 && iL3 >= 0 && iYr >= 0) {
      const g = t.slice(1).map((r) => ({
        period: String(r[iYr] || ""), level2: num(r[iL2]), level3: num(r[iL3]),
      })).filter((x) => x.period && (x.level2 != null || x.level3 != null));
      if (g.length > 1) { chargerGrowth = g; break; }
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

  let provinces = null;
  for (const t of tables) {
    const h = header(t);
    if (h.some((c) => /province|jurisdiction/.test(c)) && h.some((c) => /charger/.test(c))) {
      const iProv = h.findIndex((c) => /province|jurisdiction/.test(c));
      const rows = t.slice(1).map((r) => {
        const rest = r.slice(iProv + 1);
        const nums = rest.map(num).filter((v) => v != null);
        const ratioM = (rest.find((c) => /per/.test(c)) || "").match(/(\d+)\s*EVs?\s*per/i);
        return {
          province: cellText(r[iProv] || ""),
          evPerCharger: ratioM ? parseInt(ratioM[1], 10) : null,
          totalChargers: nums.length ? nums[nums.length - 2] : null,
          totalEvs: nums.length ? nums[nums.length - 1] : null,
        };
      }).filter((p) => p.province && !/canada/i.test(p.province));
      if (rows.length >= 5) { provinces = rows; break; }
    }
  }

  return { updated, kpis, marketShare, fleet, provinces, chargerGrowth, url: TC_URL };
}

// ------------------------------------------- source 2: Statistics Canada

async function fetchStatcan() {
  const tmpZip = path.join(os.tmpdir(), "statcan-20100025.zip");
  const dl = await downloadToFile(STATCAN_CSV_URL, tmpZip, { timeoutMs: 300000 });
  console.log(`  (downloaded via ${dl.method})`);

  const zip = new AdmZip(tmpZip);
  const entry = zip.getEntries().find((e) => /\.csv$/i.test(e.entryName) && !/meta/i.test(e.entryName));
  if (!entry) throw new Error("CSV entry not found inside StatCan zip");
  const csvBuf = entry.getData();

  let head = null;
  const idxOf = (name) => head.findIndex((h) => h === name);
  let iRef = -1, iGeo = -1, iVeh = -1, iFuel = -1, iUom = -1, iVal = -1;

  const canadaRows = [];
  const seenUom = {}, seenVeh = {};
  let headerProblem = null;

  streamCsvLines(csvBuf, (line) => {
    if (!head) {
      head = parseCsv(line)[0];
      iRef = idxOf("REF_DATE"); iGeo = idxOf("GEO"); iVeh = idxOf("Vehicle type");
      iFuel = idxOf("Fuel type"); iUom = idxOf("UOM"); iVal = idxOf("VALUE");
      if (iRef < 0 || iGeo < 0 || iFuel < 0 || iVal < 0) {
        // Diagnostic: show what we DID find before failing
        headerProblem =
          "Unexpected StatCan header. Looking for REF_DATE/GEO/'Fuel type'/VALUE. " +
          "Columns found: " + head.slice(0, 15).join(" | ") +
          " — REF_DATE:" + iRef + " GEO:" + iGeo + " Fuel:" + iFuel + " VALUE:" + iVal;
        return;
      }
      return;
    }
    if (headerProblem) return;
    if (!/Canada/.test(line)) return;
    const f = parseCsv(line);
    if (f.length <= Math.max(iRef, iGeo, iVal)) return;
    const uom = iUom >= 0 ? f[iUom] : "";
    const veh = iVeh >= 0 ? f[iVeh] : "";
    if (uom) seenUom[uom] = (seenUom[uom] || 0) + 1;
    if (veh) seenVeh[veh] = (seenVeh[veh] || 0) + 1;
    canadaRows.push({ ref: f[iRef], fuel: f[iFuel], uom, veh, val: parseFloat(f[iVal]) });
  });

  if (headerProblem) throw new Error(headerProblem);

  if (!canadaRows.length) {
    throw new Error(
      "StatCan CSV parsed but no Canada rows found. UOMs seen: " +
        Object.keys(seenUom).slice(0, 5).join("; ") +
        " | Vehicle types seen: " + Object.keys(seenVeh).slice(0, 5).join("; ")
    );
  }

  const numberUom = Object.keys(seenUom).find((u) => /number/i.test(u));
  const totalVeh = Object.keys(seenVeh).find((v) => /^total/i.test(v)) ||
                   Object.keys(seenVeh).sort((a, b) => a.length - b.length)[0];

  const usable = canadaRows.filter((r) =>
    (!numberUom || r.uom === numberUom) && (!totalVeh || r.veh === totalVeh));

  const series = {};
  for (const r of usable) {
    if (!isFinite(r.val)) continue;
    (series[r.ref] = series[r.ref] || {})[r.fuel] = r.val;
  }

  const findKey = (obj, re) => Object.keys(obj).find((k) => re.test(k));
  const registrations = Object.keys(series).sort().map((q) => {
    const d = series[q];
    const bevKey = findKey(d, /battery/i);
    const phevKey = findKey(d, /plug-?in/i);
    const totKey = findKey(d, /^total/i);
    const bev = bevKey ? d[bevKey] : 0;
    const phev = phevKey ? d[phevKey] : 0;
    const total = totKey ? d[totKey] : null;
    return { quarter: q, bev, phev, total,
      zevSharePct: total ? +(100 * (bev + phev) / total).toFixed(1) : null };
  }).filter((r) => r.bev || r.phev);

  if (!registrations.length) {
    throw new Error(
      "Canada rows found but no BEV/PHEV fuel types. Fuels seen: " +
        Object.keys(series[Object.keys(series)[0]] || {}).join("; ")
    );
  }

  return {
    tableId: "20-10-0025-01",
    lastQuarter: registrations[registrations.length - 1].quarter,
    registrations,
    url: "https://www150.statcan.gc.ca/t1/tbl1/en/tv.action?pid=2010002501",
  };
}

// ------------------------------------------- source 3: AFDC charging stations

// Map US/Canadian "State" codes & names to readable province names
const PROV_NAMES = {
  AB: "Alberta", BC: "British Columbia", MB: "Manitoba", NB: "New Brunswick",
  NL: "Newfoundland and Labrador", NS: "Nova Scotia", NT: "Northwest Territories",
  NU: "Nunavut", ON: "Ontario", PE: "Prince Edward Island", QC: "Quebec",
  SK: "Saskatchewan", YT: "Yukon",
};
const provName = (s) => {
  const t = String(s || "").trim();
  return PROV_NAMES[t.toUpperCase()] || t;
};

async function fetchAfdc() {
  if (!AFDC_API_KEY) {
    throw new Error("AFDC_API_KEY not set (env var or hardcoded constant)");
  }
  const params =
    "?access=public&api_key=" + encodeURIComponent(AFDC_API_KEY) +
    "&country=CA&ev_charging_level=" + encodeURIComponent("2,dc_fast") +
    "&ev_connector_type=all&ev_network=all&facility_type=all" +
    "&fuel_type=ELEC&funding_sources=all&limit=all&state=all&status=E" +
    "&utf8_bom=true";
  const url = AFDC_HOST + "/api/alt-fuel-stations/v1.csv" + params;

  const tmpCsv = path.join(os.tmpdir(), "afdc-stations.csv");
  const dl = await downloadToFile(url, tmpCsv, { timeoutMs: 300000 });
  console.log(`  (downloaded via ${dl.method})`);

  const text = fs.readFileSync(tmpCsv, "utf-8").replace(/^\uFEFF/, "");
  const rows = parseCsv(text);
  if (rows.length < 2) throw new Error("AFDC returned no data rows");

  const head = rows[0].map((h) => String(h).trim());
  const col = (re) => head.findIndex((h) => re.test(h));

  // Canonical station-level schema: "State" holds the province,
  // "EV Level2 EVSE Num" and "EV DC Fast Count" hold unit counts,
  // "EV Network" the operator network.
  const iProv = col(/^state$|^province$/i);
  const iL2 = col(/level\s*2.*evse|evse.*level\s*2/i);
  const iDC = col(/dc\s*fast.*count/i);
  const iNet = col(/^ev network$/i);
  if (iProv < 0 || (iL2 < 0 && iDC < 0)) {
    throw new Error(
      "Unexpected AFDC columns: " + head.slice(0, 20).join(" | ")
    );
  }

  const byProv = {}, byNetwork = {};
  let locations = 0, l2Units = 0, dcUnits = 0;
  for (let r = 1; r < rows.length; r++) {
    const f = rows[r];
    const prov = provName(f[iProv]);
    if (!prov) continue;
    locations++;
    const n2 = iL2 >= 0 ? parseInt(f[iL2], 10) || 0 : 0;
    const nd = iDC >= 0 ? parseInt(f[iDC], 10) || 0 : 0;
    l2Units += n2; dcUnits += nd;
    const p = (byProv[prov] = byProv[prov] || { locations: 0, level2: 0, dcFast: 0 });
    p.locations++; p.level2 += n2; p.dcFast += nd;
    if (iNet >= 0) {
      const net = (f[iNet] || "Non-networked").trim() || "Non-networked";
      byNetwork[net] = (byNetwork[net] || 0) + 1;
    }
  }

  if (!locations) throw new Error("AFDC CSV parsed but all rows lacked a province field");

  return {
    source: "US DOE Alternative Fuels Data Center (covers Canada)",
    downloadedVia: dl.method,
    totalLocations: locations,
    level2Units: l2Units,
    dcFastUnits: dcUnits,
    byProvince: Object.entries(byProv)
      .map(([province, v]) => ({ province, ...v }))
      .sort((a, b) => b.locations - a.locations),
    topNetworks: Object.entries(byNetwork)
      .sort((a, b) => b[1] - a[1]).slice(0, 10)
      .map(([network, count]) => ({ network, count })),
    url: "https://afdc.energy.gov/stations/states",
  };
}

// -------------------------- source 4: Open Canada funded projects (best-effort)

async function fetchChargingProjects() {
  const meta = JSON.parse(await httpGet(CKAN_API_URL));
  const resources = (meta.result && meta.result.resources) || [];
  const xlsxRes = resources.filter((r) => /^xls/i.test(String(r.format || "")));
  if (!xlsxRes.length) throw new Error("No XLSX resource in the Open Canada package");

  const res = xlsxRes.find((r) => /eng/i.test(String(r.language || ""))) || xlsxRes[xlsxRes.length - 1];
  const buf = await httpGet(res.url, { asBuffer: true, timeoutMs: 180000 });
  const wb = XLSX.read(buf, { type: "buffer" });

  const findCol = (kws, head) => head.findIndex((c) =>
    kws.some((k) => new RegExp("\\b" + k, "i").test(String(c))));

  let lastErr = "no worksheets with recognizable columns";
  for (const sheetName of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true });
    if (!rows.length) continue;

    // v8: the real header row may not be row 1 — presentation sheets often
    // have a title above it. Scan the first 10 rows for one that matches.
    for (let hr = 0; hr < Math.min(10, rows.length); hr++) {
      const head = (rows[hr] || []).map((c) => String(c));
      if (!head.some((c) => c && c !== "undefined")) continue;

      const iProv = findCol(["province", "territor", "jurisdiction", "region", "location", "city"], head);
      const iCnt = findCol(["charger", "borne", "port", "connector", "unit", "station"], head);
      if (iProv < 0 || iCnt < 0) {
        lastErr = `sheet '${sheetName}' row ${hr + 1} headers: ` + head.slice(0, 12).join(" | ");
        continue; // try next candidate header row in this sheet
      }

      const iStatus = findCol(["status"], head);
      const byProv = {}, byStatus = {};
      let total = 0, projects = 0;
      for (let r = hr + 1; r < rows.length; r++) {
        const provRaw = rows[r][iProv];
        if (provRaw == null || String(provRaw).trim() === "") continue;
        const m = String(rows[r][iCnt] ?? "").match(/\d+(\.\d+)?/);
        const cnt = m ? parseFloat(m[0]) : 0;
        total += cnt; projects++;
        const p = String(provRaw).trim();
        byProv[p] = (byProv[p] || 0) + cnt;
        if (iStatus >= 0) {
          const s = String(rows[r][iStatus] || "").trim();
          if (s) byStatus[s] = (byStatus[s] || 0) + 1;
        }
      }
      if (projects) {
        return {
          resourceName: "Projects funded by the EVAFIDI, ZEVIP and CHRI programs (NRCan)",
          headerRowFoundAt: hr + 1,
          totalFundedChargers: Math.round(total),
          projectCount: projects,
          chargersByProvince: Object.entries(byProv)
            .map(([province, chargers]) => ({ province, chargers: Math.round(chargers) }))
            .sort((a, b) => b.chargers - a.chargers),
          projectsByStatus: byStatus,
          url: "https://open.canada.ca/data/en/dataset/706a6d1a-df8e-40e3-8e0c-94211025c528",
        };
      }
      lastErr = `sheet '${sheetName}': found header-like row ${hr + 1} but no data rows under it`;
    }
  }
  throw new Error("Downloaded XLSX but no usable columns. " + lastErr);
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

  const js =
    "// AUTO-GENERATED by fetch-data.js\n" +
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
    ["statcan", "StatCan - New motor vehicle registrations (20-10-0025-01)", fetchStatcan],
    ["afdc", "US DOE AFDC - Public EV charging stations (Canada)", fetchAfdc],
    ["chargingProjects", "Open Gov - ZEVIP/EVAFIDI/CHRI funded projects", fetchChargingProjects],
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
  const afdc = snapshot.data.afdc || {};
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
    afdcLocations: afdc.totalLocations != null ? afdc.totalLocations : null,
    afdcLevel2: afdc.level2Units != null ? afdc.level2Units : null,
    afdcDcFast: afdc.dcFastUnits != null ? afdc.dcFastUnits : null,
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
