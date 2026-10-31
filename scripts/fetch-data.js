// ============================================================================
// FILE: scripts/fetch-data.js   (v5 - Nov 2026)
// ----------------------------------------------------------------------------
// Changes in this version:
//   1. Open Canada XLSX: headers have changed to "Organization | Initiative | ..."
//      - Try all worksheets, not just the first one
//      - Broaden column search: look for any geography/province/territory/jurisdiction
//        and any count/port/charger/connector column
//      - Print full header on failure so we can diagnose
//   2. StatCan WDS API with CSV fallback (as in v4)
//   3. Transport Canada unchanged
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
const STATCAN_WDS_BASE = "https://www150.statcan.gc.ca/t1/wds/rest";
const STATCAN_PRODUCT_ID = 2010002501;
const STATCAN_CSV_URL = "https://www150.statcan.gc.ca/n1/tbl/csv/20100025-eng.zip";
const OPENCANADA_PACKAGE_ID = "706a6d1a-df8e-40e3-8e0c-94211025c528";
const OPENCANADA_DATASET_PAGE =
  "https://open.canada.ca/data/en/dataset/" + OPENCANADA_PACKAGE_ID;
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

async function wdsPost(endpoint, body) {
  const res = await fetch(STATCAN_WDS_BASE + endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json", "User-Agent": UA },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`WDS ${endpoint} HTTP ${res.status}`);
  return res.json();
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

function* iterateLines(buffer) {
  const CHUNK = 16 * 1024 * 1024;
  let start = 0;
  let carry = "";
  while (start < buffer.length) {
    let end = Math.min(start + CHUNK, buffer.length);
    if (end < buffer.length) {
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

// ------------------------------------------------- source 2: StatCan (WDS API)

function flattenMembers(dim) {
  const out = [];
  const walk = (members) => {
    for (const m of members || []) {
      out.push(m);
      if (Array.isArray(m.member) && m.member.length) walk(m.member);
    }
  };
  walk(dim.member);
  return out;
}

function memberCoordId(m, fallbackIndex) {
  return m.memberId != null ? m.memberId
       : m.memberPosition != null ? m.memberPosition
       : fallbackIndex + 1;
}

async function fetchStatcanWds() {
  const metaResult = await wdsPost("/getCubeMetadata", [
    { productId: STATCAN_PRODUCT_ID },
  ]);
  if (!Array.isArray(metaResult) || metaResult[0].status !== "SUCCESS") {
    throw new Error("WDS metadata request failed: " + JSON.stringify(metaResult).substring(0, 200));
  }
  const meta = metaResult[0].object;
  console.log("  WDS table: " + (meta.cubeTitleEn || "?") +
              " (" + (meta.nbSeriesCube || "?") + " series)");

  const dims = (meta.dimension || []).filter(
    (d) => !/ref.?date|\bperiod\b|\btime\b/i.test(d.dimensionNameEn || "")
  );
  if (dims.length < 2) {
    throw new Error("WDS metadata: unexpected dimensions " +
      (meta.dimension || []).map((d) => d.dimensionNameEn).join(", "));
  }

  const withMembers = dims.map((d) => ({ dim: d, members: flattenMembers(d) }));

  const geoHolder = withMembers.find((h) => /geograph/i.test(h.dim.dimensionNameEn));
  const vehHolder = withMembers.find((h) => /vehicle\s*type/i.test(h.dim.dimensionNameEn));
  const fuelHolder = withMembers.find((h) => /fuel\s*type/i.test(h.dim.dimensionNameEn));

  if (!geoHolder || !fuelHolder) {
    throw new Error("WDS metadata: could not find geography/fuel-type dimensions. " +
      "Dimensions are: " + withMembers.map((h) => h.dim.dimensionNameEn).join(", "));
  }

  const nameOf = (m) => m.memberNameEn || m.memberNameFr || "";

  const canadaMember = geoHolder.members.find(
    (m) => /^(total,?\s*)?canada$/i.test(nameOf(m).trim())
  ) || geoHolder.members.find((m) => /\bcanada\b/i.test(nameOf(m)));
  if (!canadaMember) throw new Error("WDS: no 'Canada' member in geography dimension");

  let vehMember = null;
  if (vehHolder) {
    vehMember = vehHolder.members.find((m) => /^total/i.test(nameOf(m))) ||
                vehHolder.members[0];
  }

  const chosen = withMembers.map((h) => {
    if (h === geoHolder) return { holder: h, member: canadaMember };
    if (vehHolder && h === vehHolder) return { holder: h, member: vehMember };
    const pick = h.members.find((m) => /^total/i.test(nameOf(m))) || h.members[0];
    return { holder: h, member: pick };
  });

  const fuelMembers = fuelHolder.members;
  const N_PERIODS = 48;

  const requests = fuelMembers.map((fm) => ({
    productId: STATCAN_PRODUCT_ID,
    coordinateId: chosen.map((c) =>
      c.holder === fuelHolder ? memberCoordId(fm, fuelMembers.indexOf(fm))
                              : memberCoordId(c.member, 0)
    ).join("."),
    recentNPeriods: N_PERIODS,
  }));

  let dataResults;
  try {
    dataResults = await wdsPost("/getDataFromCubePidCoordAndLatestNPeriods", requests);
  } catch (e) {
    requests.forEach((r) => (r.recentNPeriods = 20));
    dataResults = await wdsPost("/getDataFromCubePidCoordAndLatestNPeriods", requests);
  }

  const series = {};
  let okCount = 0;
  dataResults.forEach((resObj, i) => {
    const fuelName = nameOf(fuelMembers[i]).trim();
    if (resObj.status !== "SUCCESS" || !resObj.object) return;
    okCount++;
    for (const point of resObj.object.vectorDataPoint || []) {
      const quarter = String(point.refPer || "").trim();
      const value = parseFloat(point.value);
      if (!quarter || !isFinite(value)) continue;
      (series[quarter] = series[quarter] || {})[fuelName] = value;
    }
  });

  if (!okCount) {
    throw new Error("WDS data request returned no successful series. Response: " +
      JSON.stringify(dataResults).substring(0, 300));
  }

  const fuels = Object.keys(series).flatMap((q) => Object.keys(series[q]));
  const fuelNames = [...new Set(fuels)];
  const bevLabel = fuelNames.find((f) => /battery\s*electric/i.test(f));
  const phevLabel = fuelNames.find((f) => /plug[- ]?in\s*hybrid/i.test(f));
  const totalFuel = fuelNames.find((f) => /^total/i.test(f));

  if (!bevLabel) {
    throw new Error("WDS: no battery-electric fuel type among: " +
      fuelNames.slice(0, 12).join(" ; "));
  }

  const registrations = Object.keys(series).sort().map((q) => {
    const d = series[q];
    const bev = d[bevLabel] || 0;
    const phev = (phevLabel && d[phevLabel]) || 0;
    let total = totalFuel ? d[totalFuel] : null;
    if (total == null) {
      total = Object.entries(d)
        .filter(([f]) => f !== totalFuel)
        .reduce((s, [, v]) => s + v, 0);
    }
    const zev = bev + phev;
    return {
      quarter: q, bev, phev, total,
      zevSharePct: total ? +(100 * zev / total).toFixed(1) : null,
    };
  });

  return {
    tableId: "20-10-0025-01",
    via: "WDS API",
    fuelLabels: { bev: bevLabel, phev: phevLabel, total: totalFuel || "(sum of fuels)" },
    rowsKept: registrations.length,
    lastQuarter: registrations.length
      ? registrations[registrations.length - 1].quarter
      : null,
    registrations,
    url: "https://www150.statcan.gc.ca/t1/tbl1/en/tv.action?pid=2010002501",
  };
}

async function fetchStatcanCsv() {
  const buf = await getBuffer(STATCAN_CSV_URL);
  const zip = new AdmZip(buf);
  const entry = zip
    .getEntries()
    .find((e) => /\.csv$/i.test(e.entryName) && !/meta/i.test(e.entryName));
  if (!entry) throw new Error("CSV not found inside StatCan zip archive");

  const data = entry.getData();
  const canadaRows = [];
  let iRef, iGeo, iVeh, iFuel, iUom, iVal;
  let headerParsed = false;

  for (const rawLine of iterateLines(data)) {
    const line = rawLine.charCodeAt(rawLine.length - 1) === 13
      ? rawLine.slice(0, -1) : rawLine;
    if (!line) continue;

    if (!headerParsed) {
      const head = parseCsvLine(line);
      const idx = (name) => head.findIndex((h) => h.trim() === name);
      iRef = idx("REF_DATE"); iGeo = idx("GEO"); iVeh = idx("Vehicle type");
      iFuel = idx("Fuel type"); iUom = idx("UOM"); iVal = idx("VALUE");
      if (iGeo < 0 || iVal < 0 || iFuel < 0 || iRef < 0) {
        throw new Error("Unexpected StatCan CSV header. Columns: " + head.slice(0, 20).join(" | "));
      }
      headerParsed = true;
      continue;
    }

    if (!line.includes("Canada")) continue;
    const row = parseCsvLine(line);
    if ((row[iGeo] || "").trim() !== "Canada") continue;
    canadaRows.push(row);
  }

  if (!canadaRows.length) throw new Error("CSV fallback: no rows with GEO='Canada'");

  const vehTypes = [...new Set(canadaRows.map((r) => (r[iVeh] || "").trim()))];
  const uoms = iUom >= 0 ? [...new Set(canadaRows.map((r) => (r[iUom] || "").trim()))] : [];
  const vehTotal = vehTypes.find((v) => /^total/i.test(v)) || null;
  let uomPick = null;
  if (uoms.length > 1) uomPick = uoms.find((u) => /number|unit|count|vehicle/i.test(u)) || uoms[0];
  else if (uoms.length === 1) uomPick = uoms[0];

  const series = {};
  const fuelSeen = new Set();
  for (const row of canadaRows) {
    if (vehTotal && (row[iVeh] || "").trim() !== vehTotal) continue;
    if (iUom >= 0 && uomPick && (row[iUom] || "").trim() !== uomPick) continue;
    const fuel = (row[iFuel] || "").trim();
    const val = parseFloat(row[iVal]);
    if (!isFinite(val)) continue;
    fuelSeen.add(fuel);
    (series[(row[iRef] || "").trim()] = series[(row[iRef] || "").trim()] || {})[fuel] = val;
  }

  const fuels = [...fuelSeen];
  const bevLabel = fuels.find((f) => /battery\s*electric/i.test(f));
  const phevLabel = fuels.find((f) => /plug[- ]?in\s*hybrid/i.test(f));
  const totalFuelLabel = fuels.find((f) => /^total/i.test(f));

  if (!bevLabel) throw new Error("CSV fallback: no battery-electric fuel type. Seen: " + fuels.join(" ; "));

  const registrations = Object.keys(series).sort().map((q) => {
    const d = series[q];
    const bev = d[bevLabel] || 0;
    const phev = (phevLabel && d[phevLabel]) || 0;
    let total = totalFuelLabel && d[totalFuelLabel] != null ? d[totalFuelLabel] : null;
    if (total == null) total = Object.entries(d).reduce((s, [, v]) => s + v, 0);
    const zev = bev + phev;
    return { quarter: q, bev, phev, total, zevSharePct: total ? +(100 * zev / total).toFixed(1) : null };
  });

  return {
    tableId: "20-10-0025-01",
    via: "bulk CSV fallback",
    fuelLabels: { bev: bevLabel, phev: phevLabel, total: totalFuelLabel || "(sum of fuels)" },
    rowsKept: registrations.length,
    lastQuarter: registrations.length ? registrations[registrations.length - 1].quarter : null,
    registrations,
    url: "https://www150.statcan.gc.ca/t1/tbl1/en/tv.action?pid=2010002501",
  };
}

async function fetchStatcan() {
  try {
    console.log("  trying WDS API ...");
    return await fetchStatcanWds();
  } catch (e) {
    console.warn("  WDS failed (" + e.message + ") - falling back to bulk CSV");
    return await fetchStatcanCsv();
  }
}

// -------------------------------------------- source 3: Open Canada charging projects (v5 - flexible)

async function findXlsxResource() {
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
    } catch (e) { /* try next base */ }
  }

  const html = await getText(OPENCANADA_DATASET_PAGE);
  const matches = [...html.matchAll(/href="([^"]+\.xlsx[^"]*)"/gi)];
  if (!matches.length) {
    throw new Error("Could not locate the XLSX resource via the CKAN API or the dataset page");
  }
  const href = matches[0][1].replace(/&amp;/g, "&");
  return { url: new URL(href, OPENCANADA_DATASET_PAGE).href, via: "dataset page" };
}

function findColumn(head, keywords) {
  // First try exact substring match
  const idx = head.findIndex((c) => keywords.some((k) => c.toLowerCase().includes(k)));
  if (idx >= 0) return idx;
  // Fall back to word-boundary match
  for (let i = 0; i < head.length; i++) {
    const col = head[i].toLowerCase();
    if (keywords.some((k) => {
      const re = new RegExp("\\b" + k.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b", "i");
      return re.test(col);
    })) return i;
  }
  return -1;
}

async function fetchChargingProjectsSheet(ws, xlsxUrl) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  if (!rows.length) return null;

  const head = rows[0].map((c) => String(c).trim());

  // Broader keyword sets for province and charger count
  const provKeywords = ["province", "prov", "territory", "jurisdiction", "region", "location", "area"];
  const cntKeywords = ["charger", "port", "connector", "borne", "number of", "count", "units"];

  const iProv = findColumn(head, provKeywords);
  const iCnt = findColumn(head, cntKeywords);

  const iStatus = findColumn(head, ["status", "operational", "active"]);

  if (iProv < 0 || iCnt < 0) {
    // Return null so the caller tries another sheet
    return null;
  }

  const byProv = {}, byStatus = {};
  let total = 0, projects = 0;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length <= Math.max(iProv, iCnt)) continue;
    const provRaw = row[iProv];
    if (provRaw == null || String(provRaw).trim() === "" || String(provRaw) === "Canada") continue;
    const cntMatch = String(row[iCnt] || "").match(/\d+(\.\d+)?/);
    const cnt = cntMatch ? parseFloat(cntMatch[0]) : 0;
    total += cnt;
    projects++;
    const p = String(provRaw).trim();
    byProv[p] = (byProv[p] || 0) + cnt;
    if (iStatus >= 0 && row[iStatus]) {
      const s = String(row[iStatus]).trim();
      if (s) byStatus[s] = (byStatus[s] || 0) + 1;
    }
  }

  if (!total) return null;

  const chargersByProvince = Object.entries(byProv)
    .map(([province, chargers]) => ({ province, chargers: Math.round(chargers) }))
    .sort((a, b) => b.chargers - a.chargers);

  return {
    totalFundedChargers: Math.round(total),
    projectCount: projects,
    chargersByProvince,
    projectsByStatus: byStatus,
  };
}

async function fetchChargingProjects() {
  const { url: xlsxUrl, via } = await findXlsxResource();

  const buf = await getBuffer(xlsxUrl);
  const wb = XLSX.read(buf, { type: "buffer" });

  // Try each worksheet until we find one with usable data
  let result = null;
  let lastError = null;
  let lastHeader = null;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    result = await fetchChargingProjectsSheet(ws, xlsxUrl);
    if (result && result.totalFundedChargers) {
      console.log("  Found data in sheet: " + sheetName);
      break;
    }
    if (ws) {
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
      if (rows.length && rows[0]) {
        lastHeader = rows[0].map((c) => String(c)).slice(0, 15).join(" | ");
      }
    }
    lastError = "No usable province/charger columns found in sheet " + sheetName;
  }

  if (!result) {
    throw new Error(
      "Downloaded the XLSX but could not find province/charger columns in any sheet. " +
        "Sheet names: [" + wb.SheetNames.join(", ") + "]" +
        (lastHeader ? ". Last sheet header: " + lastHeader : "")
    );
  }

  return {
    resourceName: "Projects funded by the EVAFIDI, ZEVIP and CHRI programs (NRCan)",
    resourceUrl: xlsxUrl,
    resourceVia: via,
    ...result,
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
    fundedProjects: cp.projectCount != null ? cp.totalFundedProjects : null,
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
