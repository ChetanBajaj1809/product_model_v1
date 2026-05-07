import formidable from "formidable";
import * as XLSX from "xlsx";
import fs from "fs";

export const config = {
  api: { bodyParser: false },
};

function val(x) {
  if (x === null || x === undefined) return "";
  const s = String(x).trim();
  return ["nan", "nat", "inf", "none", "undefined", "null"].includes(s.toLowerCase()) ? "" : s;
}

// Clean a header: take the first line, trim whitespace
function cleanHeader(h) {
  return String(h).split("\n")[0].trim();
}

// Parse a sheet into row objects, with cleaned single-line header keys
function parseSheet(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return { rows: [], headers: [] };

  const raw = XLSX.utils.sheet_to_json(ws, { defval: "", raw: true });
  if (!raw.length) return { rows: [], headers: [] };

  // Build a mapping from raw XLSX key -> cleaned header name
  // Use a suffix (_2, _3 ...) if cleaned names collide, to preserve all columns
  const rawKeys = Object.keys(raw[0]);
  const keyMap = {};
  const seen = {};
  for (const rk of rawKeys) {
    let cleaned = cleanHeader(rk);
    if (seen[cleaned]) {
      seen[cleaned]++;
      cleaned = `${cleaned}_${seen[cleaned]}`;
    } else {
      seen[cleaned] = 1;
    }
    keyMap[rk] = cleaned;
  }

  const rows = raw.map((row) => {
    const out = {};
    for (const [rk, ck] of Object.entries(keyMap)) {
      out[ck] = val(row[rk]);
    }
    return out;
  });

  return { rows, headers: Object.values(keyMap) };
}

// Find the right sheet name (case-insensitive partial match on all keywords)
function findSheet(sheetNames, ...keywords) {
  return sheetNames.find((s) =>
    keywords.every((k) => s.toLowerCase().includes(k.toLowerCase()))
  ) || null;
}

// Find the actual header key in a row that best matches the target string.
// Tries exact match first, then case-insensitive exact, then case-insensitive contains.
function findKey(headers, target) {
  const t = target.toLowerCase();
  return (
    headers.find((h) => h === target) ||
    headers.find((h) => h.toLowerCase() === t) ||
    headers.find((h) => h.toLowerCase().includes(t)) ||
    null
  );
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const form = formidable({ maxFileSize: 20 * 1024 * 1024 });

  form.parse(req, (err, _fields, files) => {
    if (err) return res.status(400).json({ error: "Upload failed: " + err.message });

    const fileArr = files.file;
    const file = Array.isArray(fileArr) ? fileArr[0] : fileArr;
    if (!file) return res.status(400).json({ error: "No file received" });

    try {
      const buffer = fs.readFileSync(file.filepath);
      const wb = XLSX.read(buffer, { type: "buffer" });
      const sheetNames = wb.SheetNames;

      // Prefer exact "Small Business" tabs; fall back to any tab containing the keyword
      const clausesSheet = findSheet(sheetNames, "Clauses", "Small Business")
                        || findSheet(sheetNames, "Clause", "Small")
                        || findSheet(sheetNames, "Clause")
                        || findSheet(sheetNames, "Coverage");
      const termsSheet   = findSheet(sheetNames, "Terms", "Small Business")
                        || findSheet(sheetNames, "Term", "Small")
                        || findSheet(sheetNames, "Term");
      const optionsSheet = findSheet(sheetNames, "Options", "Small Business")
                        || findSheet(sheetNames, "Option", "Small")
                        || findSheet(sheetNames, "Option");

      const rulesSheet   = findSheet(sheetNames, "PM Business Rules")
                        || findSheet(sheetNames, "Business Rule");

      const missing = [];
      if (!clausesSheet) missing.push("Clauses");
      if (!termsSheet)   missing.push("Terms");
      if (!optionsSheet) missing.push("Options");
      if (!rulesSheet)   missing.push("PM Business Rules");
      if (missing.length) {
        return res.status(422).json({
          error: `Could not find sheets: ${missing.join(", ")}. Available: ${sheetNames.join(", ")}`,
        });
      }

      const { rows: clauses, headers: clauseHeaders }   = parseSheet(wb, clausesSheet);
      const { rows: terms,   headers: termHeaders }     = parseSheet(wb, termsSheet);
      const { rows: options, headers: optionHeaders }   = parseSheet(wb, optionsSheet);
      const { rows: rulesRaw, headers: ruleHeadersRaw } = rulesSheet ? parseSheet(wb, rulesSheet) : { rows: [], headers: [] };
      const { rows: rules,   headers: ruleHeaders }     = parseSheet(wb, rulesSheet);

      // ── Resolve exact column keys ──────────────────────────────────────────
      // Clauses
      const CL_CODE = findKey(clauseHeaders, "Clause code") || findKey(clauseHeaders, "Clause Code");
      const CL_NAME = findKey(clauseHeaders, "Name");
      const CL_TYPE = findKey(clauseHeaders, "Clause Type");
      const CL_CAT  = findKey(clauseHeaders, "Category");
      const CL_DESC = findKey(clauseHeaders, "Description");
      const CL_EXIST= findKey(clauseHeaders, "Existence");
      const CL_PARTY= findKey(clauseHeaders, "Covered Party");
      const CL_PREM = findKey(clauseHeaders, "Premium bearing");
      const CL_OFFER= findKey(clauseHeaders, "Offering");
      const CL_STATE= findKey(clauseHeaders, "State Availbility as per SBT") || findKey(clauseHeaders, "State Availability as per SBT");
      const CL_FORM = findKey(clauseHeaders, "Form Number");
      const CL_COND = findKey(clauseHeaders, "Conditional Existence");
      const CL_BR   = findKey(clauseHeaders, "Business Rules (If Applicable)") || findKey(clauseHeaders, "Business Rules");

      // PM Business Rules
      const R_NUM   = findKey(ruleHeaders, "Business Rule Number");
      const R_NAME  = findKey(ruleHeaders, "Business Rule Name");
      const R_DESC  = findKey(ruleHeaders, "Rule Description");
      const R_CAT   = findKey(ruleHeaders, "Business Rule Category");

      // Terms — JOIN KEY: Terms.Related Clause Code = Clauses.Clause code
      const T_REL_CLAUSE_CODE = findKey(termHeaders, "Related Clause Code");
      const T_NAME   = findKey(termHeaders, "Term Name");
      const T_CODE   = findKey(termHeaders, "Term Code");          // JOIN KEY to Options
      const T_SCHED  = findKey(termHeaders, "Schedule");
      const T_TYPE   = findKey(termHeaders, "Term Type");
      const T_VTYPE  = findKey(termHeaders, "Value Type");

      // Options — JOIN KEY: Options.Related Clause Term Code = Terms.Term Code
      const O_REL_TERM_CODE = findKey(optionHeaders, "Related Clause Term Code");
      const O_DESC = findKey(optionHeaders, "Option Description") || findKey(optionHeaders, "Term Description");

      if (!CL_CODE) return res.status(422).json({ error: `Cannot find 'Clause code' column in ${clausesSheet}. Found: ${clauseHeaders.join(", ")}` });
      if (!T_REL_CLAUSE_CODE) return res.status(422).json({ error: `Cannot find 'Related Clause Code' column in ${termsSheet}` });
      if (!T_CODE) return res.status(422).json({ error: `Cannot find 'Term Code' column in ${termsSheet}` });
      if (!O_REL_TERM_CODE) return res.status(422).json({ error: `Cannot find 'Related Clause Term Code' column in ${optionsSheet}` });
      if (!O_DESC) return res.status(422).json({ error: `Cannot find 'Option Description' column in ${optionsSheet}` });

      // ── Build lookups ──────────────────────────────────────────────────────

      // Terms.Related Clause Code -> [term rows]
      const termsByClauseCode = {};
      for (const trow of terms) {
        const ccode = trow[T_REL_CLAUSE_CODE];
        if (!ccode) continue;
        if (!termsByClauseCode[ccode]) termsByClauseCode[ccode] = [];
        termsByClauseCode[ccode].push(trow);
      }

      // Terms.Term Code -> [option descriptions]
      const optionsByTermCode = {};
      for (const orow of options) {
        const tcode = orow[O_REL_TERM_CODE];
        const desc  = orow[O_DESC];
        if (!tcode || !desc) continue;
        if (!optionsByTermCode[tcode]) optionsByTermCode[tcode] = [];
        optionsByTermCode[tcode].push(desc);
      }

      // Business Rule Number -> rule object
      // JOIN KEY: extract number after # and match regardless of prefix (CA_PM_BR#16 = Test_PM_BR#16)
      const rulesByNumber = {};
      for (const rrow of rules) {
        const rnum = R_NUM ? rrow[R_NUM] : "";
        if (!rnum) continue;
        const match = rnum.match(/#(\d+)/);
        if (match) rulesByNumber[match[1]] = rrow;
      }

      // ── Build coverages ────────────────────────────────────────────────────
      const coverages = [];

      for (const row of clauses) {
        const ccode = row[CL_CODE];
        const cname = CL_NAME ? row[CL_NAME] : "";
        if (!ccode && !cname) continue;

        const cov = {
          name:                  cname,
          code:                  ccode,
          type:                  CL_TYPE  ? row[CL_TYPE]  : "",
          category:              CL_CAT   ? row[CL_CAT]   : "",
          description:           CL_DESC  ? row[CL_DESC]  : "",
          existence:             CL_EXIST ? row[CL_EXIST] : "",
          covered_party:         CL_PARTY ? row[CL_PARTY] : "",
          premium_bearing:       CL_PREM  ? row[CL_PREM]  : "",
          offering:              CL_OFFER ? (row[CL_OFFER] || "").split("\n")[0] : "",
          state_availability:    CL_STATE ? row[CL_STATE] : "",
          form_number:           CL_FORM  ? row[CL_FORM]  : "",
          conditional_existence: CL_COND  ? row[CL_COND]  : "",
          terms: [],
          rules: [],
        };

        // JOIN: Clauses."Business Rules (If Applicable)" -> PM Business Rules."Business Rule Number"
        // A clause can reference multiple rules e.g. "CA_PM_BR#16\nCA_PM_BR#17"
        const brRef = CL_BR ? row[CL_BR] : "";
        if (brRef) {
          const nums = [...brRef.matchAll(/#(\d+)/g)].map(m => m[1]);
          for (const num of nums) {
            const rrow = rulesByNumber[num];
            if (rrow) {
              cov.rules.push({
                number:      R_NUM  ? rrow[R_NUM]  : "",
                name:        R_NAME ? rrow[R_NAME] : "",
                description: R_DESC ? rrow[R_DESC] : "",
                category:    R_CAT  ? rrow[R_CAT]  : "",
              });
            }
          }
        }

        // JOIN: Clauses.Clause code -> Terms.Related Clause Code
        const covTermRows = termsByClauseCode[ccode] || [];
        for (const trow of covTermRows) {
          const tcode = trow[T_CODE];
          cov.terms.push({
            name:       T_NAME  ? trow[T_NAME]  : "",
            code:       tcode,
            schedule:   T_SCHED ? trow[T_SCHED] : "",
            term_type:  T_TYPE  ? trow[T_TYPE]  : "",
            value_type: T_VTYPE ? trow[T_VTYPE] : "",
            // JOIN: Terms.Term Code -> Options.Related Clause Term Code
            options: tcode ? (optionsByTermCode[tcode] || []) : [],
          });
        }

        coverages.push(cov);
      }

      res.status(200).json({
        coverages,
        // Raw sheet data for the editor (all columns, all rows)
        rawSheets: {
          clauses:       clauses,
          clauseHeaders: clauseHeaders,
          terms:         terms,
          termHeaders:   termHeaders,
          options:       options,
          optionHeaders: optionHeaders,
          rules:         rulesRaw,
          ruleHeaders:   ruleHeadersRaw,
        },
        meta: {
          sheets: { clausesSheet, termsSheet, optionsSheet },
          columns: { CL_CODE, T_REL_CLAUSE_CODE, T_CODE, O_REL_TERM_CODE, O_DESC },
          counts: {
            coverages: coverages.length,
            terms:   coverages.reduce((a, c) => a + c.terms.length, 0),
            options: coverages.reduce((a, c) => c.terms.reduce((b, t) => b + t.options.length, a), 0),
            rules:   coverages.reduce((a, c) => a + c.rules.length, 0),
          },
        },
      });
    } catch (e) {
      res.status(500).json({ error: "Parse error: " + e.message });
    }
  });
}
