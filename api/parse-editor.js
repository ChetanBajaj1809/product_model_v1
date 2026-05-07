import formidable from "formidable";
import * as XLSX from "xlsx";
import fs from "fs";

export const config = {
  api: { bodyParser: false },
};

function cleanHeader(h) {
  return String(h).split("\n")[0].trim();
}

function val(x) {
  if (x === null || x === undefined) return "";
  const s = String(x).trim();
  return ["nan", "nat", "inf", "none", "undefined", "null"].includes(s.toLowerCase()) ? "" : s;
}

function parseSheet(wb, sheetName, wantedCols) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];

  const raw = XLSX.utils.sheet_to_json(ws, { defval: "", raw: true });
  if (!raw.length) return [];

  // Map raw XLSX keys -> cleaned header names
  const rawKeys = Object.keys(raw[0]);
  const keyMap = {};
  const seen = {};
  for (const rk of rawKeys) {
    let cleaned = cleanHeader(rk);
    if (seen[cleaned]) { seen[cleaned]++; cleaned += `_${seen[cleaned]}`; }
    else seen[cleaned] = 1;
    keyMap[rk] = cleaned;
  }

  // Build index: cleaned name -> raw key (first occurrence wins for duplicates)
  const cleanedToRaw = {};
  for (const [rk, ck] of Object.entries(keyMap)) {
    if (!cleanedToRaw[ck]) cleanedToRaw[ck] = rk;
  }

  return raw.map(row => {
    const out = {};
    for (const col of wantedCols) {
      const rk = cleanedToRaw[col];
      out[col] = rk ? val(row[rk]) : "";
    }
    return out;
  });
}

function findSheet(names, ...keywords) {
  return names.find(s => keywords.every(k => s.toLowerCase().includes(k.toLowerCase()))) || null;
}

const CLAUSES_COLS = [
  "Clause code","Name","Clause Type","Category","Description",
  "Existence","Covered Party Type","Premium bearing? (as per PM)",
  "Schedule?","Offering","Business Rules (If Applicable)",
  "Form Number and Edition (Informational Only)",
  "State Availbility as per SBT for (MI, NM, NV, KS, AZ, OK)",
  "Conditional Existence  as per SBT","Effective Date","Expiration Date",
];

const TERMS_COLS = [
  "Related Clause Code","Related Clause Name","Term Code","Term Name",
  "Term Type","Value Type",
  "Type of Field (User Entered, Drop Down or Radio Button or Contact or Other - Define Other)",
  "Required","Schedule?","Default Value","Minimum Value","Maximum Value",
  "State Availability","Business Rule (If Applicable)",
  "Effective Date","Expiration Date",
];

const OPTIONS_COLS = [
  "Related Clause Term Code","Related Clause Term","Related Clause Name",
  "Term Code","Option Description (to be displayed)","Term Value",
  "State Availability","Business Rule","Effective Date","Expiration Date",
];

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const form = formidable({ maxFileSize: 20 * 1024 * 1024 });

  form.parse(req, (err, _fields, files) => {
    if (err) return res.status(400).json({ error: "Upload failed: " + err.message });

    const fileArr = files.file;
    const file = Array.isArray(fileArr) ? fileArr[0] : fileArr;
    if (!file) return res.status(400).json({ error: "No file received" });

    try {
      const buffer = fs.readFileSync(file.filepath);
      const wb = XLSX.read(buffer, { type: "buffer" });
      const names = wb.SheetNames;

      const clausesSheet = findSheet(names,"Clauses","Small Business") || findSheet(names,"Clause","Small") || findSheet(names,"Clause");
      const termsSheet   = findSheet(names,"Terms","Small Business")   || findSheet(names,"Term","Small")   || findSheet(names,"Term");
      const optionsSheet = findSheet(names,"Options","Small Business") || findSheet(names,"Option","Small") || findSheet(names,"Option");

      const missing = [];
      if (!clausesSheet) missing.push("Clauses");
      if (!termsSheet)   missing.push("Terms");
      if (!optionsSheet) missing.push("Options");
      if (missing.length) return res.status(422).json({ error: `Cannot find sheets: ${missing.join(", ")}` });

      res.status(200).json({
        clauses: parseSheet(wb, clausesSheet, CLAUSES_COLS),
        terms:   parseSheet(wb, termsSheet,   TERMS_COLS),
        options: parseSheet(wb, optionsSheet, OPTIONS_COLS),
      });
    } catch (e) {
      res.status(500).json({ error: "Parse error: " + e.message });
    }
  });
}
