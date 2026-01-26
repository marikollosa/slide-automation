"use client";

import React, { useMemo, useState } from "react";
import JSZip from "jszip";
import * as XLSX from "xlsx";

type SlideTypeId = "org_change" | "new_tools" | "cw_risk_assessment";

type SlideType = {
  id: SlideTypeId;
  label: string;
  description?: string;
};

type CellSpec =
  | { type: "cell"; ref: string }
  | { type: "join"; refs: string[]; joinWith: string }
  | { type: "const"; value: string }
  | {
      type: "month_year";
      ref: string;
      format?: "Mon YYYY" | "MMMM YYYY" | "MM/YYYY";
    };

type SlideMapping = Record<number, Record<string, CellSpec>>;

/**
 * -----------------------------
 * MAPPINGS
 * -----------------------------
 */

const ORG_CHANGE_MAPPING: SlideMapping = {
  1: {
    "NAME OF PROJECT": { type: "cell", ref: "F2" },
    "TYPE OF PROJECT": { type: "cell", ref: "K2" },
  },
  3: {
    "[Description]": { type: "cell", ref: "M2" },
  },
  4: {
    "[L2/L3]": { type: "cell", ref: "I2" },
    "[Owner]": { type: "cell", ref: "G2" },
    "[Lead]": { type: "cell", ref: "H2" },
    "[Comms]": { type: "cell", ref: "J2" },
  },
  5: {
    "[Date]": { type: "cell", ref: "N2" },
    "[Phases]": { type: "cell", ref: "P2" },
  },
  6: {
    "[1]": { type: "cell", ref: "Q2" },
    "[2]": { type: "cell", ref: "R2" },
    "[3]": { type: "join", refs: ["S2", "T2"], joinWith: " " },
    "[4]": { type: "cell", ref: "V2" },
  },
  7: {
    "[1]": { type: "cell", ref: "W2" },
  },
  8: {
    "[1]": { type: "cell", ref: "L2" },
    "[2]": { type: "join", refs: ["AA2", "AB2"], joinWith: " " },
    "[3]": { type: "cell", ref: "Y2" },
  },
  9: {
    "[1]": { type: "cell", ref: "Z2" },
    "[2]": { type: "cell", ref: "AD2" },
  },
  10: {
    "[1]": { type: "cell", ref: "AF2" },
    "[2]": { type: "cell", ref: "AG2" },
  },
  11: {
    "[1]": { type: "cell", ref: "DG2" },
    "[2]": { type: "cell", ref: "DI2" },
  },
  12: {
    "[1]": { type: "cell", ref: "AH2" },
  },
};

const NEW_TOOLS_MAPPING: SlideMapping = {
  1: {
    "NAME OF PROJECT": { type: "cell", ref: "F2" },
    "TYPE OF PROJECT": { type: "cell", ref: "K2" },
  },
  3: {
    "[1]": { type: "cell", ref: "BZ2" },
  },
  4: {
    "[1]": { type: "cell", ref: "I2" },
    "[2]": { type: "const", value: "N/A" },
    "[3]": { type: "cell", ref: "G2" },
    "[4]": { type: "cell", ref: "H2" },
    "[5]": { type: "cell", ref: "J2" },
  },
  5: {
    "[1]": { type: "cell", ref: "CA2" },
    "[2]": { type: "const", value: "N/A" },
  },
  6: {
    "[1]": { type: "const", value: "N/A" },
    "[2]": { type: "const", value: "N/A" },
    "[3]": { type: "const", value: "N/A" },
    "[4]": { type: "const", value: "N/A" },
  },
  7: {
    "[1]": { type: "const", value: "N/A" },
  },
  8: {
    "[1]": { type: "cell", ref: "BW2" },
    "[2]": { type: "join", refs: ["CD2", "CE2"], joinWith: " " },
    "[3]": { type: "cell", ref: "CC2" },
  },
  9: {
    "[1]": { type: "cell", ref: "BX2" },
    "[2]": { type: "cell", ref: "CH2" },
    "[3]": { type: "cell", ref: "CI2" },
    "[4]": { type: "cell", ref: "BN2" },
  },
  10: {
    "[1]": { type: "cell", ref: "CJ2" },
    "[2]": { type: "cell", ref: "CK2" },
    "[3]": { type: "cell", ref: "CL2" },
    "[4]": { type: "cell", ref: "CM2" },
    "[5]": { type: "cell", ref: "CN2" },
    "[6]": { type: "cell", ref: "CO2" },
    "[7]": { type: "cell", ref: "CP2" },
  },
  11: {
    "[1]": { type: "cell", ref: "CQ2" },
    "[2]": { type: "cell", ref: "CR2" },
    "[3]": { type: "cell", ref: "CS2" },
    "[4]": { type: "cell", ref: "CT2" },
    "[5]": { type: "cell", ref: "CU2" },
    "[6]": { type: "cell", ref: "CV2" },
    "[7]": { type: "cell", ref: "CX2" },
  },
  12: {
    "[1]": { type: "cell", ref: "CY2" },
    "[2]": { type: "cell", ref: "CZ2" },
    "[3]": { type: "cell", ref: "DB2" },
    "[4]": { type: "cell", ref: "DC2" },
  },
  13: {
    "[1]": { type: "cell", ref: "DG2" },
    "[2]": { type: "cell", ref: "DI2" },
  },
};

const CW_RISK_ASSESSMENT_MAPPING: SlideMapping = {
  1: {
    "[1]": { type: "cell", ref: "G2" },
    // Month + year only from B2 (handles true Excel dates OR strings like M/D/YY HH:MM:SS)
    "[2]": { type: "month_year", ref: "B2", format: "Mon YYYY" },
  },

  3: {
    "[1]": { type: "cell", ref: "K2" },
    "[2]": { type: "cell", ref: "K3" },
    "[3]": { type: "cell", ref: "K4" },
    "[4]": { type: "cell", ref: "K5" },
  },

  4: {
    "[A1]": { type: "cell", ref: "H2" },
    "[B1]": { type: "cell", ref: "H3" },
    "[C1]": { type: "cell", ref: "H4" },
    "[D1]": { type: "cell", ref: "H5" },

    "[A2]": { type: "cell", ref: "I2" },
    "[B2]": { type: "cell", ref: "I3" },
    "[C2]": { type: "cell", ref: "I4" },
    "[D2]": { type: "cell", ref: "I5" },

    "[A3]": { type: "cell", ref: "J2" },
    "[B3]": { type: "cell", ref: "J3" },
    "[C3]": { type: "cell", ref: "J4" },
    "[D3]": { type: "cell", ref: "J5" },

    "[A4]": { type: "cell", ref: "K2" },
    "[B4]": { type: "cell", ref: "K3" },
    "[C4]": { type: "cell", ref: "K4" },
    "[D4]": { type: "cell", ref: "K5" },
  },

  5: {
    "[A1]": { type: "cell", ref: "L2" },
    "[B1]": { type: "cell", ref: "L3" },
    "[C1]": { type: "cell", ref: "L4" },
    "[D1]": { type: "cell", ref: "L5" },

    "[A2]": { type: "cell", ref: "M2" },
    "[B2]": { type: "cell", ref: "M3" },
    "[C2]": { type: "cell", ref: "M4" },
    "[D2]": { type: "cell", ref: "M5" },

    "[A3]": { type: "cell", ref: "N2" },
    "[B3]": { type: "cell", ref: "N3" },
    "[C3]": { type: "cell", ref: "N4" },
    "[D3]": { type: "cell", ref: "N5" },

    "[A4]": { type: "cell", ref: "O2" },
    "[B4]": { type: "cell", ref: "O3" },
    "[C4]": { type: "cell", ref: "O4" },
    "[D4]": { type: "cell", ref: "O5" },

    "[A5]": { type: "cell", ref: "P2" },
    "[B5]": { type: "cell", ref: "P3" },
    "[C5]": { type: "cell", ref: "P4" },
    "[D5]": { type: "cell", ref: "P5" },

    "[A6]": { type: "cell", ref: "Q2" },
    "[B6]": { type: "cell", ref: "Q3" },
    "[C6]": { type: "cell", ref: "Q4" },
    "[D6]": { type: "cell", ref: "Q5" },
  },
};

const SLIDE_DEFS: Array<SlideType & { mapping: SlideMapping }> = [
  {
    id: "org_change",
    label: "Organization Change",
    description:
      "Upload the org change PPTX template + Excel file to generate the filled deck.",
    mapping: ORG_CHANGE_MAPPING,
  },
  {
    id: "new_tools",
    label: "New Tools / Surveys / Trainings",
    description:
      "Upload the New Tools/Surveys/Trainings template + Excel file to generate the filled deck.",
    mapping: NEW_TOOLS_MAPPING,
  },
  {
    id: "cw_risk_assessment",
    label: "CW Risk Assessment",
    description:
      "Upload the CW Risk Assessment template + Excel file to generate the filled deck.",
    mapping: CW_RISK_ASSESSMENT_MAPPING,
  },
];

export default function Page() {
  const [slideType, setSlideType] = useState<SlideTypeId>(SLIDE_DEFS[0].id);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);

  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);

  const [resetKey, setResetKey] = useState(0);

  const selectedSlideDef = useMemo(
    () => SLIDE_DEFS.find((s) => s.id === slideType),
    [slideType]
  );

  const canSubmit = useMemo(() => {
    return Boolean(slideType && templateFile && excelFile && !isSubmitting);
  }, [slideType, templateFile, excelFile, isSubmitting]);

  function handleReset() {
    setError(null);
    setSuccessMsg(null);
    setSlideType(SLIDE_DEFS[0].id);
    setTemplateFile(null);
    setExcelFile(null);
    setIsSubmitting(false);
    setResetKey((k) => k + 1);
  }

  async function handleGenerate() {
    setError(null);
    setSuccessMsg(null);

    if (!templateFile || !excelFile) {
      setError("Upload both a PPTX template and an Excel file.");
      return;
    }

    if (!templateFile.name.toLowerCase().endsWith(".pptx")) {
      setError("Template must be a .pptx file.");
      return;
    }

    const excelLower = excelFile.name.toLowerCase();
    if (!(excelLower.endsWith(".xlsx") || excelLower.endsWith(".xls"))) {
      setError("Excel must be a .xlsx or .xls file.");
      return;
    }

    const mapping = selectedSlideDef?.mapping;
    if (!mapping) {
      setError("No mapping found for this slide type.");
      return;
    }

    if (Object.keys(mapping).length === 0) {
      setError(
        `The "${selectedSlideDef?.label ?? "selected"}" slide type mapping is empty.`
      );
      return;
    }

    setIsSubmitting(true);

    try {
      // ---- Read Excel in-browser ----
      const excelArrayBuf = await excelFile.arrayBuffer();
      const workbook = XLSX.read(excelArrayBuf, { type: "array", cellDates: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      const getCellObj = (ref: string) => {
        const cell = (sheet as any)?.[ref];
        return cell ?? null;
      };

      const getCell = (ref: string): string => {
        const cell = getCellObj(ref);
        if (!cell || cell.v == null) return "N/A";
        const v = String(cell.w ?? cell.v).trim();
        return v === "" ? "N/A" : v;
      };

      const formatMonthYear = (
        ref: string,
        format: "Mon YYYY" | "MMMM YYYY" | "MM/YYYY" = "Mon YYYY"
      ): string => {
        const cell = getCellObj(ref);
        if (!cell || cell.v == null) return "N/A";

        let month: number | null = null; // 1-12
        let year: number | null = null;

        // 1) Date instance (because cellDates:true)
        if (cell.v instanceof Date && !isNaN(cell.v.getTime())) {
          month = cell.v.getMonth() + 1;
          year = cell.v.getFullYear();
        }

        // 2) Excel serial number date
        if ((month == null || year == null) && typeof cell.v === "number") {
          const parsed = XLSX.SSF.parse_date_code(cell.v);
          if (parsed && parsed.m && parsed.y) {
            month = parsed.m;
            year = parsed.y;
          }
        }

        // 3) String like M/D/YY HH:MM:SS (or M/D/YYYY ...)
        if (month == null || year == null) {
          const s = String(cell.w ?? cell.v).trim();
          const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
          if (m) {
            month = Number(m[1]);
            const yRaw = m[3];
            year = yRaw.length === 2 ? 2000 + Number(yRaw) : Number(yRaw);
          }
        }

        if (!month || !year) return "N/A";

        const monShort = [
          "Jan",
          "Feb",
          "Mar",
          "Apr",
          "May",
          "Jun",
          "Jul",
          "Aug",
          "Sep",
          "Oct",
          "Nov",
          "Dec",
        ];
        const monLong = [
          "January",
          "February",
          "March",
          "April",
          "May",
          "June",
          "July",
          "August",
          "September",
          "October",
          "November",
          "December",
        ];

        if (format === "MM/YYYY") {
          const mm = String(month).padStart(2, "0");
          return `${mm}/${year}`;
        }
        if (format === "MMMM YYYY") {
          return `${monLong[month - 1]} ${year}`;
        }
        return `${monShort[month - 1]} ${year}`;
      };

      const resolveSpec = (spec: CellSpec): string => {
        if (spec.type === "cell") return getCell(spec.ref);
        if (spec.type === "const") return spec.value;

        if (spec.type === "month_year") {
          return formatMonthYear(spec.ref, spec.format ?? "Mon YYYY");
        }

        // join
        const parts = spec.refs
          .map(getCell)
          .map((v) => String(v ?? "").trim())
          .filter((v) => v.length > 0 && v !== "N/A");

        return parts.length ? parts.join(spec.joinWith) : "N/A";
      };

      // ---- Read PPTX (zip) in-browser ----
      const pptxArrayBuf = await templateFile.arrayBuffer();
      const zip = await JSZip.loadAsync(pptxArrayBuf);

      for (const [slideNumStr, placeholders] of Object.entries(mapping)) {
        const slideNum = Number(slideNumStr);
        const slidePath = `ppt/slides/slide${slideNum}.xml`;
        const file = zip.file(slidePath);
        if (!file) continue;

        let xml = await file.async("string");

        for (const [needle, spec] of Object.entries(placeholders)) {
          const value = escapeXml(resolveSpec(spec));
          xml = xml.split(needle).join(value);
        }

        zip.file(slidePath, xml);
      }

      const outArrayBuffer = await zip.generateAsync({ type: "arraybuffer" });
      const outBlob = new Blob([outArrayBuffer], {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      });

      downloadBlob(outBlob, `generated_${slideType}.pptx`);
      setSuccessMsg("Generated! Your download should start automatically.");
    } catch (e: any) {
      setError(e?.message || "Something went wrong generating the slides.");
    } finally {
      setIsSubmitting(false);
    }
  }

  return (
    <main style={styles.page}>
      <div style={styles.card}>
        <h1 style={styles.h1}>Slide Automation</h1>
        <p style={styles.sub}>
          Select a slide type, upload a PPTX template + Excel data, and generate a
          filled deck.
        </p>

        <div style={styles.section}>
          <label style={styles.label}>Slide type</label>
          <select
            value={slideType}
            onChange={(e) => setSlideType(e.target.value as SlideTypeId)}
            style={styles.select}
            disabled={isSubmitting}
          >
            {SLIDE_DEFS.map((t) => (
              <option key={t.id} value={t.id}>
                {t.label}
              </option>
            ))}
          </select>

          {selectedSlideDef?.description && (
            <div style={styles.helperText}>{selectedSlideDef.description}</div>
          )}
        </div>

        <div key={resetKey}>
          <div style={styles.section}>
            <label style={styles.label}>Template (.pptx)</label>
            <input
              ref={(el) => {
                if (el) (window as any).templateInput = el;
              }}
              type="file"
              accept=".pptx"
              onChange={(e) => setTemplateFile(e.target.files?.[0] ?? null)}
              disabled={isSubmitting}
              style={{ display: "none" }}
            />
            <button
              type="button"
              onClick={() => (window as any).templateInput?.click()}
              disabled={isSubmitting}
              style={{
                ...styles.fileButton,
                background: templateFile
                  ? "rgba(34,197,94,0.12)"
                  : "rgba(96,125,255,0.12)",
                border: templateFile
                  ? "2px solid rgba(34,197,94,0.5)"
                  : "2px dashed rgba(96,125,255,0.5)",
              }}
            >
              <span style={{ fontSize: 20, marginRight: 8 }}>
                {templateFile ? "✓" : "📎"}
              </span>
              {templateFile ? templateFile.name : "Choose Template (.pptx)"}
            </button>
          </div>

          <div style={styles.section}>
            <label style={styles.label}>Data (.xlsx)</label>
            <input
              ref={(el) => {
                if (el) (window as any).excelInput = el;
              }}
              type="file"
              accept=".xlsx,.xls"
              onChange={(e) => setExcelFile(e.target.files?.[0] ?? null)}
              disabled={isSubmitting}
              style={{ display: "none" }}
            />
            <button
              type="button"
              onClick={() => (window as any).excelInput?.click()}
              disabled={isSubmitting}
              style={{
                ...styles.fileButton,
                background: excelFile
                  ? "rgba(34,197,94,0.12)"
                  : "rgba(96,125,255,0.12)",
                border: excelFile
                  ? "2px solid rgba(34,197,94,0.5)"
                  : "2px dashed rgba(96,125,255,0.5)",
              }}
            >
              <span style={{ fontSize: 20, marginRight: 8 }}>
                {excelFile ? "✓" : "📊"}
              </span>
              {excelFile ? excelFile.name : "Choose Data (.xlsx)"}
            </button>
          </div>
        </div>

        {error && <div style={styles.error}>{error}</div>}
        {successMsg && <div style={styles.success}>{successMsg}</div>}

        <button
          onClick={handleGenerate}
          disabled={!canSubmit}
          style={{
            ...styles.button,
            opacity: canSubmit ? 1 : 0.5,
            cursor: canSubmit ? "pointer" : "not-allowed",
          }}
        >
          {isSubmitting ? "Generating…" : "Generate & Download"}
        </button>

        <button
          type="button"
          onClick={handleReset}
          disabled={isSubmitting}
          style={{
            ...styles.secondaryButton,
            opacity: isSubmitting ? 0.6 : 1,
            cursor: isSubmitting ? "not-allowed" : "pointer",
          }}
        >
          Start Over
        </button>

        <div style={styles.note}>
          This version runs fully in the browser (GitHub Pages-friendly).
        </div>
      </div>
    </main>
  );
}

function escapeXml(s: string) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function downloadBlob(blob: Blob, filename: string) {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  window.URL.revokeObjectURL(url);
}

const styles: Record<string, React.CSSProperties> = {
  page: {
    minHeight: "100vh",
    display: "grid",
    placeItems: "center",
    padding: 24,
    fontFamily:
      "ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial",
    background: "#0b0b10",
    color: "#f4f4f5",
  },
  card: {
    width: "min(720px, 100%)",
    background: "#12121a",
    border: "1px solid #232334",
    borderRadius: 16,
    padding: 24,
    boxShadow: "0 10px 30px rgba(0,0,0,0.35)",
  },
  h1: { margin: 0, fontSize: 28, letterSpacing: -0.3 },
  sub: { marginTop: 8, marginBottom: 20, color: "#b8b8c7", lineHeight: 1.4 },
  section: { marginBottom: 16 },
  label: { display: "block", marginBottom: 8, fontWeight: 600 },
  helperText: { marginTop: 8, color: "#b8b8c7", fontSize: 13, lineHeight: 1.3 },
  select: {
    width: "100%",
    padding: "10px 12px",
    borderRadius: 10,
    border: "1px solid #2b2b3f",
    background: "#0e0e16",
    color: "#f4f4f5",
  },
  fileButton: {
    width: "100%",
    padding: "16px 14px",
    borderRadius: 10,
    border: "2px dashed #2b2b3f",
    background: "#0e0e16",
    color: "#f4f4f5",
    fontWeight: 500,
    cursor: "pointer",
    transition: "all 0.2s ease",
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-start",
    fontSize: 14,
  },
  error: {
    padding: 12,
    borderRadius: 12,
    background: "rgba(239,68,68,0.12)",
    border: "1px solid rgba(239,68,68,0.3)",
    color: "#fecaca",
    marginBottom: 12,
  },
  success: {
    padding: 12,
    borderRadius: 12,
    background: "rgba(34,197,94,0.12)",
    border: "1px solid rgba(34,197,94,0.3)",
    color: "#bbf7d0",
    marginBottom: 12,
  },
  button: {
    width: "100%",
    padding: "12px 14px",
    borderRadius: 12,
    border: "1px solid #2b2b3f",
    background: "#1a1a27",
    color: "#f4f4f5",
    fontWeight: 700,
  },
  secondaryButton: {
    width: "100%",
    padding: "12px 14px",
    borderRadius: 12,
    border: "1px solid #2b2b3f",
    background: "transparent",
    color: "#f4f4f5",
    fontWeight: 700,
    marginTop: 10,
  },
  note: { marginTop: 16, fontSize: 14, color: "#b8b8c7", lineHeight: 1.4 },
};
