"use client";

import React, { useMemo, useState } from "react";

type SlideType = {
  id: string;
  label: string;
  description?: string;
};

const SLIDE_TYPES: SlideType[] = [
  {
    id: "org_change",
    label: "Organization Change",
    description:
      "Upload the org change PPTX template + Excel file to generate the filled deck.",
  },
  {
    id: "new_tools",
    label: "New Tools / Surveys / Trainings",
    description:
      "Upload the New Tools/Surveys/Trainings template + Excel file to generate the filled deck.",
  },
];

export default function Page() {
  const [slideType, setSlideType] = useState<string>(SLIDE_TYPES[0].id);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);

  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);

  // Used to force-remount file inputs so the browser clears selected files
  const [resetKey, setResetKey] = useState(0);

  const selectedSlideType = useMemo(
    () => SLIDE_TYPES.find((s) => s.id === slideType),
    [slideType]
  );

  const canSubmit = useMemo(() => {
    return Boolean(slideType && templateFile && excelFile && !isSubmitting);
  }, [slideType, templateFile, excelFile, isSubmitting]);

  function handleReset() {
    setError(null);
    setSuccessMsg(null);

    setSlideType(SLIDE_TYPES[0].id);
    setTemplateFile(null);
    setExcelFile(null);

    setIsSubmitting(false);

    // Force file inputs to re-mount (clears the native file picker state)
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

    const formData = new FormData();
    formData.append("slideType", slideType);
    formData.append("template", templateFile);
    formData.append("excel", excelFile);

    setIsSubmitting(true);

    try {
      const res = await fetch("/api/generate", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const msg = await safeReadText(res);
        throw new Error(msg || `Generate failed (HTTP ${res.status})`);
      }

      const blob = await res.blob();

      const contentDisposition = res.headers.get("content-disposition");
      const suggestedName =
        parseFilenameFromContentDisposition(contentDisposition);
      const filename = suggestedName || `generated_${slideType}.pptx`;

      downloadBlob(blob, filename);
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
          Select a slide type, upload a PPTX template + Excel data, and generate
          a filled deck.
        </p>

        <div style={styles.section}>
          <label style={styles.label}>Slide type</label>
          <select
            value={slideType}
            onChange={(e) => setSlideType(e.target.value)}
            style={styles.select}
            disabled={isSubmitting}
          >
            {SLIDE_TYPES.map((t) => (
              <option key={t.id} value={t.id}>
                {t.label}
              </option>
            ))}
          </select>

          {selectedSlideType?.description && (
            <div style={styles.helperText}>{selectedSlideType.description}</div>
          )}
        </div>

        {/* key={resetKey} forces inputs to remount and clears selected files */}
        <div key={resetKey}>
          <div style={styles.section}>
            <label style={styles.label}>Template (.pptx)</label>
            <input
              type="file"
              accept=".pptx"
              onChange={(e) => setTemplateFile(e.target.files?.[0] ?? null)}
              disabled={isSubmitting}
            />
            {templateFile && (
              <div style={styles.fileMeta}>
                Selected: <strong>{templateFile.name}</strong>
              </div>
            )}
          </div>

          <div style={styles.section}>
            <label style={styles.label}>Data (.xlsx)</label>
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={(e) => setExcelFile(e.target.files?.[0] ?? null)}
              disabled={isSubmitting}
            />
            {excelFile && (
              <div style={styles.fileMeta}>
                Selected: <strong>{excelFile.name}</strong>
              </div>
            )}
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
          Tip: Use <strong>Start Over</strong> to clear uploads and generate
          another deck.
        </div>
      </div>
    </main>
  );
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

async function safeReadText(res: Response) {
  try {
    return await res.text();
  } catch {
    return "";
  }
}

function parseFilenameFromContentDisposition(cd: string | null) {
  if (!cd) return null;
  const match = cd.match(/filename\*?=(?:UTF-8'')?"?([^"]+)"?/i);
  if (!match?.[1]) return null;
  return decodeURIComponent(match[1]);
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
  fileMeta: { marginTop: 8, color: "#b8b8c7", fontSize: 14 },
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