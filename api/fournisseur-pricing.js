import xlsx from "xlsx";

// ============ SHAREPOINT CONFIG ============
const TENANT_ID     = process.env.MICROSOFT_TENANT_ID;
const CLIENT_ID     = process.env.MICROSOFT_CLIENT_ID;
const CLIENT_SECRET = process.env.MICROSOFT_CLIENT_SECRET;
const DRIVE_ID      = process.env.SP_DRIVE_ID;
const FOLDER_PATH   = process.env.SP_FOLDER_PATH || "PARTAGE/Team/3. Fournisseurs/Data Fournisseurs";
const FILE_NAME     = "Chiffres_fournisseurs.xlsx";
const SHEET_NAME    = "Selection fournisseur";

// ============ CACHE MÉMOIRE (5 minutes) ============
let _cache = null;
let _cacheAt = 0;
const CACHE_TTL_MS = 5 * 60 * 1000;

// ============ TOKEN MICROSOFT ============
async function getMsToken() {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type:    "client_credentials",
        client_id:     CLIENT_ID,
        client_secret: CLIENT_SECRET,
        scope:         "https://graph.microsoft.com/.default"
      })
    }
  );
  const data = await res.json();
  if (!data.access_token) throw new Error("Impossible d'obtenir le token Microsoft Graph");
  return data.access_token;
}

// ============ TÉLÉCHARGEMENT EXCEL DEPUIS SHAREPOINT ============
async function downloadExcelFromSharePoint() {
  const token = await getMsToken();

  const fileRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/root:/${FOLDER_PATH}/${FILE_NAME}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const fileMeta = await fileRes.json();

  if (!fileMeta["@microsoft.graph.downloadUrl"]) {
    throw new Error(`Fichier SharePoint introuvable : ${FILE_NAME}`);
  }

  const dlRes = await fetch(fileMeta["@microsoft.graph.downloadUrl"]);
  const arrayBuffer = await dlRes.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

// ============ UTILITAIRES (inchangés) ============
function safeNumber(value) {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;

  let s = String(value)
    .replace(/\u00A0/g, " ")
    .replace(/€/g, "")
    .replace(/MWh/gi, "")
    .trim();

  if (!s) return null;
  s = s.replace(/\s+/g, "");

  const hasComma = s.includes(",");
  const hasDot   = s.includes(".");

  if (hasComma && hasDot) {
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) {
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      s = s.replace(/,/g, "");
    }
  } else if (hasComma) {
    s = s.replace(",", ".");
  }

  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function round(value, digits = 2) {
  if (value === null || value === undefined || Number.isNaN(value)) return null;
  return Number(Number(value).toFixed(digits));
}

function slugify(str = "") {
  return String(str)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/&/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function normalizeSupplierName(raw = "") {
  const cleaned = String(raw).trim();
  const key = slugify(cleaned);

  const aliases = new Map([
    ["endesa energia succursal france", "ENDESA"],
    ["sefe energy sas", "SEFE"],
    ["totalenergies electricite et gaz france", "TOTALENERGIES"],
    ["alterna energie", "ALTERNA"],
    ["engie", "ENGIE"],
    ["energem", "ENERGEM"],
    ["mint", "MINT"],
    ["ilek", "ILEK"],
    ["gedia energies services", "GEDIA"],
    ["enovos luxembourg s a", "ENOVOS"],
    ["primeo energie france", "PRIMEO"],
    ["ohm energie", "OHM"],
    ["dyneff s a s", "DYNEFF"],
    ["bcm energy", "BCM ENERGY"],
    ["hellio solutions", "HELLIO"],
    ["la bellenergie", "LA BELLENERGIE"],
    ["societe d approvisionnement et de vente d energies", "SAVE"],
    ["geg source d energies", "GEG"],
    ["synelva", "SYNELVA"]
  ]);

  return aliases.get(key) || cleaned.toUpperCase();
}

function extractDateRange(text = "") {
  const match = String(text).match(/\((\d{2}\/\d{2}\/\d{4})\s+to\s+(\d{2}\/\d{2}\/\d{4})\)/i);
  if (!match) return { startDate: null, endDate: null };
  return { startDate: match[1], endDate: match[2] };
}

function extractGeneratedAt(text = "") {
  const match = String(text).match(/As of (\d{4}-\d{2}-\d{2}) (\d{2}:\d{2}:\d{2})/i);
  if (!match) return { generatedAt: null, updatedAtLabel: null };
  const [, isoDate, time] = match;
  const [year, month, day] = isoDate.split("-");
  return {
    generatedAt: `${isoDate}T${time}`,
    updatedAtLabel: `${day}/${month}/${year}`
  };
}

function buildMarketAverage(rows) {
  const valid = rows.filter(
    r => Number.isFinite(r.price30j) && Number.isFinite(r.avgVolume) && Number.isFinite(r.offerCount)
  );
  const denominator = valid.reduce((sum, r) => sum + r.avgVolume * r.offerCount, 0);
  if (!denominator) return null;
  return valid.reduce((sum, r) => sum + r.price30j * r.avgVolume * r.offerCount, 0) / denominator;
}

// ============ PARSING DU WORKBOOK ============
function parseWorkbook(buffer) {
  const workbook = xlsx.read(buffer, { type: "buffer", cellDates: false });
  const sheet = workbook.Sheets[SHEET_NAME];
  if (!sheet) throw new Error(`Onglet introuvable: ${SHEET_NAME}`);

  const matrix = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: false });

  const generatedInfo = extractGeneratedAt(matrix?.[2]?.[1] || "");
  const periodInfo    = extractDateRange(matrix?.[7]?.[1] || "");

  const rows = xlsx.utils.sheet_to_json(sheet, { range: 14, defval: null, raw: true });

  function normalizeKey(str = "") {
    return String(str).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  }

  const keys        = Object.keys(rows[0] || {});
  const segmentKey  = keys.find(k => normalizeKey(k).includes("compteur segment"));
  const supplierKey = keys.find(k => normalizeKey(k).includes("nom fournisseur"));
  const priceKey    = keys.find(k => normalizeKey(k).includes("prix moyen") && normalizeKey(k).includes("average"));
  const volumeKey   = keys.find(k => normalizeKey(k).includes("volume du compteur"));
  const countKey    = keys.find(k => normalizeKey(k).includes("record count"));

  if (!segmentKey || !supplierKey || !priceKey || !volumeKey || !countKey) {
    throw new Error("Colonnes Excel introuvables ou format inattendu.");
  }

  let currentSegment = null;
  const normalizedRows = [];

  for (const row of rows) {
    const rawSegment  = row[segmentKey];
    const rawSupplier = row[supplierKey];

    if (rawSegment && String(rawSegment).trim()) {
      currentSegment = String(rawSegment).trim();
    }
    if (!rawSupplier || !String(rawSupplier).trim()) continue;
    if (
      String(currentSegment).toLowerCase().includes("subtotal") ||
      String(rawSupplier).toLowerCase().includes("subtotal")
    ) continue;

    normalizedRows.push({
      segment:     currentSegment,
      supplierRaw: String(rawSupplier).trim(),
      supplier:    normalizeSupplierName(rawSupplier),
      price30j:    safeNumber(row[priceKey]),
      avgVolume:   safeNumber(row[volumeKey]),
      offerCount:  safeNumber(row[countKey])
    });
  }

  const bySegment = new Map();
  for (const row of normalizedRows) {
    if (!bySegment.has(row.segment)) bySegment.set(row.segment, []);
    bySegment.get(row.segment).push(row);
  }

  const segments = {};
  for (const [segment, segmentRows] of bySegment.entries()) {
    const marketAverage = buildMarketAverage(segmentRows);
    segments[segment] = segmentRows
      .map(row => {
        const deltaPct =
          Number.isFinite(marketAverage) && Number.isFinite(row.price30j) && marketAverage !== 0
            ? ((row.price30j - marketAverage) / marketAverage) * 100
            : null;
        return { ...row, marketAverage: round(marketAverage, 2), deltaPct: round(deltaPct, 1) };
      })
      .sort((a, b) => {
        if (a.price30j === null) return 1;
        if (b.price30j === null) return -1;
        return a.price30j - b.price30j;
      });
  }

  return {
    meta: {
      fileName: FILE_NAME,
      sheetName: SHEET_NAME,
      ...periodInfo,
      ...generatedInfo,
      segments: Object.keys(segments).sort()
    },
    segments
  };
}

// ============ CHARGEMENT AVEC CACHE ============
async function loadWorkbookData() {
  const now = Date.now();
  if (_cache && now - _cacheAt < CACHE_TTL_MS) return _cache;

  const buffer = await downloadExcelFromSharePoint();
  _cache = parseWorkbook(buffer);
  _cacheAt = now;
  return _cache;
}

// ============ HANDLER ============
export default async function handler(req, res) {
  if (req.method !== "GET") {
    return res.status(405).json({ message: "Méthode non autorisée" });
  }

  try {
    const data      = await loadWorkbookData();
    const segment   = String(req.query.segment    || "").trim().toUpperCase();
    const fournisseur = String(req.query.fournisseur || "").trim();

    if (!segment) {
      return res.status(200).json(data);
    }

    const segmentRows = data.segments[segment] || [];

    if (!fournisseur) {
      return res.status(200).json({ meta: data.meta, segment, rows: segmentRows });
    }

    const supplierNeedle = normalizeSupplierName(fournisseur);
    const row =
      segmentRows.find(r => r.supplier === supplierNeedle) ||
      segmentRows.find(r => slugify(r.supplierRaw) === slugify(fournisseur)) ||
      null;

    return res.status(200).json({ meta: data.meta, segment, fournisseur: supplierNeedle, row, rows: segmentRows });

  } catch (error) {
    console.error("fournisseur-pricing error:", error);
    return res.status(500).json({ message: error.message || "Erreur serveur" });
  }
}
