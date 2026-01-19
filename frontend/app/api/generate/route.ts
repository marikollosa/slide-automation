import JSZip from "jszip";
import * as XLSX from "xlsx";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type CellSpec =
  | { type: "cell"; ref: string }
  | { type: "join"; refs: string[]; joinWith: string }
  | { type: "const"; value: string };

type SlideMapping = Record<number, Record<string, CellSpec>>;

/** Organization Change */
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
};

/** New Tools / Surveys / Trainings */
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
    "[2]": { type: "join", refs: ["CD2", "CE2"], joinWith: " " }, // change to "\n" if you want a line break
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

export async function POST(req: Request) {
  try {
    const form = await req.formData();
    const slideType = String(form.get("slideType") ?? "org_change");
    const template = form.get("template");
    const excel = form.get("excel");

    if (!(template instanceof File)) {
      return new Response("Missing 'template' file upload.", { status: 400 });
    }
    if (!(excel instanceof File)) {
      return new Response("Missing 'excel' file upload.", { status: 400 });
    }

    const mapping: SlideMapping =
      slideType === "new_tools" ? NEW_TOOLS_MAPPING : ORG_CHANGE_MAPPING;

    // ---- Excel ----
    const excelBuf = Buffer.from(await excel.arrayBuffer());
    const workbook = XLSX.read(excelBuf, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const getCell = (ref: string): string => {
        const cell = sheet?.[ref] as any;

         // If cell is missing or truly empty
        if (!cell || cell.v == null || String(cell.v).trim() === "") {
            return "N/A";
    }

  const value = String(cell.w ?? cell.v).trim();
  return value === "" ? "N/A" : value;
};

    const resolveSpec = (spec: CellSpec): string => {
      if (spec.type === "cell") return getCell(spec.ref);
      if (spec.type === "const") return spec.value;

      const parts = spec.refs
        .map(getCell)
        .map((v) => String(v ?? "").trim())
        .filter((v) => v.length > 0);

      return parts.join(spec.joinWith);
    };

    // ---- PPTX (zip) ----
    const pptxBuf = Buffer.from(await template.arrayBuffer());
    const zip = await JSZip.loadAsync(pptxBuf);

    // ✅ Use the selected mapping here
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

    return new Response(outArrayBuffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Disposition": `attachment; filename="generated.pptx"`,
      },
    });
  } catch (err: any) {
    return new Response(`Generate failed: ${err?.message ?? String(err)}`, { status: 500 });
  }
}

function escapeXml(s: string) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}