import formidable from "formidable";
import * as XLSX from "xlsx";
import fs from "fs";

export const config = {
  api: { bodyParser: false },
};

function cleanHeader(h) {
  return String(h).split("\n")[0].trim();
}

function findSheet(sheetNames, ...keywords) {
  return sheetNames.find((s) =>
    keywords.every((k) => s.toLowerCase().includes(k.toLowerCase()))
  ) || null;
}

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const form = formidable({ maxFileSize: 20 * 1024 * 1024 });

  form.parse(req, (err, fields, files) => {
    if (err) return res.status(400).json({ error: "Upload failed: " + err.message });

    const fileArr = files.file;
    const file = Array.isArray(fileArr) ? fileArr[0] : fileArr;
    if (!file) return res.status(400).json({ error: "No file received" });

    let edits;
    try {
      const raw = Array.isArray(fields.edits) ? fields.edits[0] : fields.edits;
      edits = JSON.parse(raw);
    } catch (e) {
      return res.status(400).json({ error: "Invalid edits JSON: " + e.message });
    }

    try {
      const buffer = fs.readFileSync(file.filepath);
      const wb = XLSX.read(buffer, { type: "buffer" });
      const sheetNames = wb.SheetNames;

      const clausesSheet = findSheet(sheetNames, "Clauses", "Small Business") || findSheet(sheetNames, "Clause", "Small") || findSheet(sheetNames, "Clause");
      const termsSheet   = findSheet(sheetNames, "Terms",   "Small Business") || findSheet(sheetNames, "Term",   "Small") || findSheet(sheetNames, "Term");
      const optionsSheet = findSheet(sheetNames, "Options", "Small Business") || findSheet(sheetNames, "Option", "Small") || findSheet(sheetNames, "Option");
      const rulesSheet   = findSheet(sheetNames, "PM Business Rules") || findSheet(sheetNames, "Business Rule");

      // Helper: update a sheet with edited rows
      // edits[sheetKey] = array of row objects (full rows, in order)
      function applyEdits(sheetName, rows) {
        if (!sheetName || !rows) return;
        const ws = wb.Sheets[sheetName];
        if (!ws) return;

        // Read current headers from row 1 (preserve original multiline headers)
        const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
        const headerRow = [];
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
          headerRow.push(cell ? cell.v : "");
        }

        // Build clean header -> original header mapping
        const cleanToOrig = {};
        const cleanHeaders = [];
        headerRow.forEach((h, i) => {
          const clean = cleanHeader(String(h));
          if (!cleanToOrig[clean]) {
            cleanToOrig[clean] = h;
          }
          cleanHeaders.push(clean);
        });

        // Rebuild worksheet: keep row 1 as-is, replace data rows
        const newData = [headerRow];
        for (const row of rows) {
          const dataRow = cleanHeaders.map((ch) => {
            const v = row[ch];
            return v !== undefined && v !== null ? v : "";
          });
          newData.push(dataRow);
        }

        const newWs = XLSX.utils.aoa_to_sheet(newData);
        // Preserve column widths and other sheet properties
        if (ws["!cols"]) newWs["!cols"] = ws["!cols"];
        if (ws["!merges"]) newWs["!merges"] = ws["!merges"];
        wb.Sheets[sheetName] = newWs;
      }

      if (edits.clauses) applyEdits(clausesSheet, edits.clauses);
      if (edits.terms)   applyEdits(termsSheet,   edits.terms);
      if (edits.options) applyEdits(optionsSheet,  edits.options);
      if (edits.rules)   applyEdits(rulesSheet,    edits.rules);

      // Write workbook to buffer
      const outBuffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", "attachment; filename=\"Product_Model_Updated.xlsx\"");
      res.status(200).send(Buffer.from(outBuffer));

    } catch (e) {
      res.status(500).json({ error: "Update error: " + e.message });
    }
  });
}
