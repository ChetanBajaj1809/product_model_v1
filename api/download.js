import * as XLSX from "xlsx";

export default function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const { clauses, terms, options, clausesCols, termsCols, optionsCols } = req.body;

    const wb = XLSX.utils.book_new();

    function makeSheet(rows, cols) {
      const data = [cols, ...rows.map(row => cols.map(c => row[c] ?? ""))];
      return XLSX.utils.aoa_to_sheet(data);
    }

    XLSX.utils.book_append_sheet(wb, makeSheet(clauses, clausesCols),  "Clauses - Small Business");
    XLSX.utils.book_append_sheet(wb, makeSheet(terms,   termsCols),    "Terms - Small Business");
    XLSX.utils.book_append_sheet(wb, makeSheet(options, optionsCols),  "Options - Small Business");

    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=Product_Model_Updated.xlsx");
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: "Export failed: " + e.message });
  }
}
