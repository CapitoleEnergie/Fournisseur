const path = require("path");
const xlsx = require("xlsx");

const EXCEL_FILE = path.join(process.cwd(), "data", "regles_panel_fournisseurs.xlsx");
const PANEL_SHEET = "Fournisseurs Panel";

const PANEL_PRIORITY = {
  "gold premium": 1,
  "gold": 2,
  "silver": 3,
  "bronze": 4,
  "": 99,
  "non classe": 99
};

function normalizeText(value = "") {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

module.exports = function handler(req, res) {
  try {
    const workbook = xlsx.readFile(EXCEL_FILE);
    const sheet = workbook.Sheets[PANEL_SHEET];

    if (!sheet) {
      return res.status(404).json({ message: `Onglet introuvable : ${PANEL_SHEET}` });
    }

    const data = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    if (data.length < 2) {
      return res.status(200).json({ suppliers: [] });
    }

    const suppliers = [];
    for (let i = 1; i < data.length; i++) {
      const rawSupplier = normalizeText(data[i][0]);
      const rawPanel = normalizeText(data[i][1]);
      if (!rawSupplier) continue;

      suppliers.push({
        supplier: rawSupplier,
        panel: rawPanel || "",
        panelPriority: PANEL_PRIORITY[rawPanel.toLowerCase()] ?? 99
      });
    }

    // Tri par priorité panel
    suppliers.sort((a, b) => a.panelPriority - b.panelPriority || a.supplier.localeCompare(b.supplier));

    return res.status(200).json({ suppliers });
  } catch (err) {
    console.error("fournisseur-panel error:", err);
    return res.status(500).json({ message: err.message || "Erreur serveur" });
  }
};
