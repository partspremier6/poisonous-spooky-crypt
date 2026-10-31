// ============================================================================
// FILE: scripts/fetch-data.js   (v6 - Nov 2026)
// ----------------------------------------------------------------------------
// Changes in this version:
//   1. Transport Canada KPIs now have a TABLE-BASED FALLBACK:
//      - primary: parse the headline prose (robust multi-pattern regexes)
//      - fallback: compute Total EVs from the fleet table (sum of BEV/PHEV/
//        FCEV rows) and Level 2/Level 3/total chargers from the cumulative
//        charger tables (last row = most recent, currently March 2026)
//      - ratio computed as totalEvs / totalChargers if the prose doesn't give it
//      Reason: the KPI strip on the page may be assembled by JavaScript, so a
//      plain HTTP fetch can miss it - but the tables are plain HTML.
//   2. NEW DATA: per-province table from Transport Canada (EVs per public
//      charger, total public chargers, total EV count per province/territory)
//      -> snapshot.data.transportCanada.provinces
//   3. NEW DATA: cumulative charger growth by year
//      -> snapshot.data.transportCanada.chargerTrend
//   4. Bug fix: fundedProjects KPI referenced a non-existent field.
//   5. StatCan WDS API + CSV fallback and the flexible Open Canada XLSX
//      parser from v5 are unchanged.
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
  return s
    .replace(/&nbsp;/gi, " ")
    .replace(/&#160;/g, " ")
    .replace(/&amp;/gi, "&")
    .replace(/\u00a0/g, " ")
    .trim();
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

function grabNum(text, patterns) {
  for (const re of patterns) {
    const m = text.match(re);
    if (m) {
      const v = num(m[1]);
      if (v != null) return v;
    }
  }
  return null;
}

async function fetchTransportCanada() {
  const html = await getText(TC_URL);
  const updatedMatch = html.match(/[Dd]ata was updated in ([A-Z][a-z]+ \d{4})/);
  const updated = updatedMatch ? updatedMatch[1] : null;
  const text = cellText(html);
  const tables = parseTables(html);
  const header = (t) => t[0].map((c) => c.toLowerCase());

  // ---- Tables ---------------------------------------------------------

  // Fleet: "Vehicle category and type | Number of vehicles | Market share"
  let fleet = null;
  for (const t of tables) {
    const h = header(t);
    if (h.some((c) => c.includes("vehicle category")) &&
        h.some((c) => c.includes("number of vehicles"))) {
      fleet = t.slice(1).map((r) => ({
        category: r[0], count: num(r[1]), sharePct: pct(r[2]),
      }));
      break;
    }
  }
  // older header variant
  if (!fleet) {
    for (const t of tables) {
      const h = header(t);
      if (h.some((c) => c.includes("vehicle category")) &&
          h.some((c) => c.includes("market share"))) {
        fleet = t.slice(1).map((r) => ({
          category: r[0], count: num(r[1]), sharePct: pct(r[2]),
        }));
        break;
      }
    }
  }

  // Cumulative chargers by year: "Year | Level 2 | Level 3"
  let chargerTrend = null;
  for (const t of tables) {
    const h = header(t);
    if (h[0] === "year" && h.includes("level 2") && h.includes("level 3")) {
      chargerTrend = t.slice(1).map((r) => ({
        period: r[0], level2: num(r[1]), level3: num(r[2]),
        total: num(r[1]) != null && num(r[2]) != null ? num(r[1]) + num(r[2]) : null,
      }));
      break;
    }
  }

  // Total chargers by year: "Year | Total chargers"
  let totalChargerTrend = null;
  for (const t of tables) {
    const h = header(t);
    if (h[0] === "year" && h.includes("total chargers") && !h.includes("level 2")) {
      totalChargerTrend = t.slice(1).map((r) => ({
        period: r[0], totalChargers: num(r[1]),
      }));
      break;
    }
  }

  // Annual market share: "Year | BEV% | PHEV% | FCEV% | Total%" style
  let marketShare = null;
  for (const t of tables) {
    const h = header(t);
    if (h.includes("year") && h.some((c) => c.includes("bev")) &&
        !h.includes("level 2")) {
      marketShare = t.slice(1).map((r) => ({
        period: r[0], bev: pct(r[1]), phev: pct(r[2]), fcev: pct(r[3]), total: pct(r[4]),
      })).filter((x) => x.total != null);
      if (marketShare.length) break;
      marketShare = null;
    }
  }

  // Province table: "Province/territory | [stats cell]"
  // stats cell reads like: "EVs per public charger: 36  Total public
  // chargers: 13,081  Total EV count: 473,385"
  let provinces = null;
  for (const t of tables) {
    const h = header(t);
    if (h.some((c) => c.includes("province")) &&
        t.slice(1).some((r) => /EVs per public charger/i.test(r[1] || ""))) {
      provinces = t.slice(1).map((r) => {
        const stats = String(r[1] || "");
        const g = (re) => {
          const m = stats.match(re);
          return m ? num(m[1]) : null;
        };
        return {
          province: String(r[0] || "").trim(),
          evsPerCharger: g(/EVs per public charger:?\s*([\d,.]+)/i),
          publicChargers: g(/Total public chargers?:?\s*([\d,.]+)/i),
          evCount: g(/Total EV count:?\s*([\d,.]+)/i),
        };
      }).filter((p) => p.province);
      break;
    }
  }

  // ---- KPIs: prose first, table fallback second -----------------------

  let kpis = {
    totalEvs: grabNum(text, [/Total EVs?:?\s*([\d,]+)/i, /Total (?:electric vehicles|EVs)[^\d]{0,40}([\d,]{6,})/i]),
    bevs: grabNum(text, [/BEVs?:?\s*([\d,]+)/i]),
    phevs: grabNum(text, [/PHEVs?:?\s*([\d,]+)/i]),
    totalChargers: grabNum(text, [/Total chargers?:?\s*([\d,]+)/i]),
    level2: grabNum(text, [/Level 2:?\s*([\d,]+)/i]),
    level3: grabNum(text, [/Level 3:?\s*([\d,]+)/i]),
    evPerCharger: null,
  };

  // ratio from prose: "ratio: 27 EVs: 1 charger"
  const ratioM = text.match(/ratio[^0-9]{0,40}?(\d+)\s*EVs?\s*[:,]?\s*(\d+)\s*charger/i);
  if (ratioM) {
    kpis.evPerCharger = num(ratioM[1]) != null ? num(ratioM[1]) :
      (num(ratioM[2]) ? Math.round(num(ratioM[1]) / num(ratioM[2])) : null);
  }

  // Fallbacks from tables (KPI strip may be JS-rendered; tables are not)
  if (kpis.level2 == null && chargerTrend && chargerTrend.length) {
    const last = chargerTrend[chargerTrend.length - 1];
    kpis.level2 = last.level2;
    kpis.level3 = last.level3;
  }
  if (kpis.totalChargers == null) {
    if (totalChargerTrend && totalChargerTrend.length) {
      kpis.totalChargers = totalChargerTrend[totalChargerTrend.length - 1].totalChargers;
    } else if (kpis.level2 != null && kpis.level3 != null) {
      kpis.totalChargers = kpis.level2 + kpis.level3;
    }
  }
  if (kpis.bevs == null && fleet) {
    kpis.bevs = fleet.filter((r) => /^light-duty BEVs?$/i.test(r.category))
      .reduce((s, r) => s + (r.count || 0), 0) || null;
  }
  if (kpis.phevs == null && fleet) {
    kpis.phevs = fleet.filter((r) => /^light-duty PHEVs?$/i.test(r.category))
      .reduce((s, r) => s + (r.count || 0), 0) || null;
  }
  if (kpis.totalEvs == null && fleet) {
    // sum every EV row (exclude the "All vehicles" denominator rows)
    const evRows = fleet.filter((r) =>
      /(^light-duty (BEV|PHEV|FCEV))|(^medium and heavy-duty (BEV|FCEV))/i.test(r.category)
    );
    if (evRows.length) {
      kpis.totalEvs = evRows.reduce((s, r) => s + (r.count || 0), 0);
    }
  }
  if (kpis.evPerCharger == null && kpis.totalEvs && kpis.totalChargers) {
    kpis.evPerCharger = Math.round(kpis.totalEvs / kpis.totalChargers);
  }
  // If prose gave everything except one field, patch from per-province sums
  if (provinces && provinces.length) {
    if (kpis.totalChargers == null) {
      kpis.totalChargers = provinces.reduce((s, p) => s + (p.publicChargers || 0), 0);
    }
    if (kpis.totalEvs == null) {
      kpis.totalEvs = provinces.reduce((s, p) => s + (p.evCount || 0), 0);
    }
  }

  kpis.source = kpis.totalEvs != null && kpis.totalChargers != null
    ? "parsed" : "partial";

  let imhzev = null;
  for (const t of tables) {
    const h = header(t);
    if (h.includes("year") && h.includes("qc") && h.includes("total") &&
        !h.includes("level 2")) {
      imhzev = t.slice(1).map((r) => ({
        year: r[0], qc: num(r[1]), bc: num(r[2]), on: num(r[3]), rest: num(r[4]), total: num(r[5]),
      }));
      break;
    }
  }

  return { updated, kpis, marketShare, fleet, chargerTrend, provinces, imhzev, url: TC_URL };
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

  const fuelNames = [...new Set(Object.keys(series).flatMap((q) => Object.keys(series[q])))];
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

// -------------------------------------------- source 3: Open Canada charging projects

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
  const idx = head.findIndex((c) => keywords.some((k) => c.toLowerCase().includes(k)));
  if (idx >= 0) return idx;
  for (let i = 0; i < head.length; i++) {
    const col = head[i].toLowerCase();
    if (keywords.some((k) => {
      const re = new RegExp("\\b" + k.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "\\b", "i");
      return re.test(col);
    })) return i;
  }
  return -1;
}

async function fetchChargingProjectsSheet(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  if (!rows.length) return null;

  const head = rows[0].map((c) => String(c).trim());

  const provKeywords = ["province", "prov", "territory", "jurisdiction", "region", "location", "area"];
  const cntKeywords = ["charger", "port", "connector", "borne", "number of", "count", "units"];

  const iProv = findColumn(head, provKeywords);
  const iCnt = findColumn(head, cntKeywords);
  const iStatus = findColumn(head, ["status", "operational", "active"]);

  if (iProv < 0 || iCnt < 0) return null;

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

  let result = null;
  let lastHeader = null;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    result = await fetchChargingProjectsSheet(ws);
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
    result = null;
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
