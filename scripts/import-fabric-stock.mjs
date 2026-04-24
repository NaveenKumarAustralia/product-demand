import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { basename, extname, join, posix, resolve } from "node:path";
import { tmpdir } from "node:os";

const workbookPath = process.argv[2];
if (!workbookPath) {
  console.error("Usage: node scripts/import-fabric-stock.mjs /path/to/FABRIC.xlsx");
  process.exit(1);
}

const rootDir = resolve(new URL("..", import.meta.url).pathname);
const outputDataPath = join(rootDir, "app", "fabric-stock-data.ts");
const outputImageDir = join(rootDir, "public", "fabric-stock", "images");
const extractDir = join(tmpdir(), `fabric-stock-${Date.now()}`);

const sheetIdByName = new Map([
  ["on order", "759049382"],
  ["fabrics samples under considera", "390729206"],
  ["fabric samples under consideration", "390729206"],
  ["60x60 printed", "1829736341"],
  ["4x40 printed", "0"],
  ["40x40 printed", "0"],
  ["velvet", "128048837"],
  ["rayon", "1670965992"],
  ["thick cord", "1949735348"],
  ["new fabric on order", "1128806463"],
  ["gorge{double fabric}", "2123625779"],
  ["gorge (double fabric)", "2123625779"],
  ["plain 40x40", "1972131020"],
  ["plain 60x60", "362283931"],
  ["cotton drill", "1240512146"],
  ["thick self black", "2008557105"],
  ["seersucker", "1939488240"],
  ["random fabrics bits and bobs", "246387032"],
  ["voil", "700159838"],
  ["fabric on order", "2030572043"],
]);

const kindByName = new Map([
  ["on order", "order"],
  ["fabrics samples under considera", "order"],
  ["fabric samples under consideration", "order"],
  ["new fabric on order", "wide-order"],
  ["fabric on order", "wide-order"],
  ["cotton drill", "simple-stock"],
  ["thick self black", "simple-stock"],
  ["seersucker", "simple-stock"],
  ["random fabrics bits and bobs", "random"],
]);

const displayNameByName = new Map([
  ["fabrics samples under considera", "Fabric Samples under consideration"],
  ["4x40 printed", "40x40 Printed"],
  ["gorge{double fabric}", "Gorge (Double Fabric)"],
  ["plain 40x40", "Plain 40x40"],
  ["plain 60x60", "Plain 60x60"],
  ["seersucker", "Seersucker"],
  ["voil", "Voil"],
]);

function xmlDecode(value = "") {
  return value
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function attr(xml, name) {
  const match = xml.match(new RegExp(`\\b${name}="([^"]*)"`));
  return match ? xmlDecode(match[1]) : "";
}

function readXml(path) {
  return readFileSync(path, "utf8");
}

function parseRelationships(path) {
  if (!existsSync(path)) return new Map();
  const rels = new Map();
  const xml = readXml(path);
  for (const match of xml.matchAll(/<Relationship\b([^>]*)\/>/g)) {
    rels.set(attr(match[1], "Id"), attr(match[1], "Target"));
  }
  return rels;
}

function parseSharedStrings() {
  const path = join(extractDir, "xl", "sharedStrings.xml");
  if (!existsSync(path)) return [];
  const xml = readXml(path);
  return [...xml.matchAll(/<si\b[\s\S]*?<\/si>/g)].map(([si]) => {
    const parts = [...si.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((match) => xmlDecode(match[1]));
    return parts.join("");
  });
}

function columnIndex(ref) {
  const letters = (ref.match(/[A-Z]+/)?.[0] ?? "A").split("");
  return letters.reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0) - 1;
}

function excelSerialToDate(value) {
  const serial = Number(value);
  if (!Number.isFinite(serial)) return "";
  const utc = Math.round((serial - 25569) * 86400 * 1000);
  const date = new Date(utc);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toLocaleDateString("en-AU", { day: "2-digit", month: "short", year: "numeric", timeZone: "UTC" });
}

function cleanName(name) {
  return name.trim().replace(/\s+/g, " ");
}

function keyName(name) {
  return cleanName(name).toLowerCase();
}

function slug(name) {
  return keyName(name).replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "sheet";
}

function jsString(value) {
  return JSON.stringify(value);
}

function parseWorkbookSheets() {
  const workbook = readXml(join(extractDir, "xl", "workbook.xml"));
  const workbookRels = parseRelationships(join(extractDir, "xl", "_rels", "workbook.xml.rels"));
  return [...workbook.matchAll(/<sheet\b([^>]*)\/>/g)].map((match, index) => {
    const name = attr(match[1], "name");
    const relId = attr(match[1], "r:id");
    const target = workbookRels.get(relId);
    return {
      index: index + 1,
      name,
      state: attr(match[1], "state") || "visible",
      worksheetPath: target ? join(extractDir, "xl", target) : "",
    };
  }).filter((sheet) => sheet.worksheetPath);
}

function drawingPathForSheet(sheetPath) {
  const relPath = join(
    sheetPath.slice(0, sheetPath.lastIndexOf("/")),
    "_rels",
    `${basename(sheetPath)}.rels`,
  );
  const rels = parseRelationships(relPath);
  const drawingTarget = [...rels.values()].find((target) => target.includes("drawings/drawing"));
  if (!drawingTarget) return "";
  return posix.normalize(posix.join("xl/worksheets", drawingTarget)).replace(/^xl\//, "");
}

function parseImageAnchors(sheetPath) {
  const drawingRelPath = drawingPathForSheet(sheetPath);
  if (!drawingRelPath) return new Map();

  const drawingFullPath = join(extractDir, "xl", drawingRelPath);
  if (!existsSync(drawingFullPath)) return new Map();

  const drawingRels = parseRelationships(join(
    drawingFullPath.slice(0, drawingFullPath.lastIndexOf("/")),
    "_rels",
    `${basename(drawingFullPath)}.rels`,
  ));
  const drawingXml = readXml(drawingFullPath);
  const anchors = new Map();

  for (const anchor of drawingXml.matchAll(/<xdr:(?:oneCellAnchor|twoCellAnchor)\b[\s\S]*?<\/xdr:(?:oneCellAnchor|twoCellAnchor)>/g)) {
    const block = anchor[0];
    const col = Number(block.match(/<xdr:col>(\d+)<\/xdr:col>/)?.[1]);
    const row = Number(block.match(/<xdr:row>(\d+)<\/xdr:row>/)?.[1]);
    const relId = block.match(/r:embed="([^"]+)"/)?.[1] ?? "";
    const target = drawingRels.get(relId);
    if (!Number.isFinite(row) || !Number.isFinite(col) || !target) continue;
    const mediaPath = posix.normalize(posix.join("xl/drawings", target));
    anchors.set(`${row + 1}:${col}`, mediaPath);
  }

  return anchors;
}

function parseRows(sheetPath, sharedStrings, imageAnchors, sheetSlug, imageUrls) {
  const xml = readXml(sheetPath);
  const rows = new Map();

  for (const rowMatch of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const rowNumber = Number(attr(rowMatch[1], "r"));
    if (!Number.isFinite(rowNumber)) continue;
    const row = [];
    for (const cellMatch of rowMatch[2].matchAll(/<c\b([^>]*)\/>|<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const cellAttrs = cellMatch[1] ?? cellMatch[2] ?? "";
      const cellBody = cellMatch[3] ?? "";
      const ref = attr(cellAttrs, "r");
      const index = columnIndex(ref);
      const type = attr(cellAttrs, "t");
      const rawValue = cellBody.match(/<v>([\s\S]*?)<\/v>/)?.[1] ?? "";
      if (!rawValue) {
        row[index] = "";
      } else if (type === "s") {
        row[index] = sharedStrings[Number(rawValue)] ?? "";
      } else {
        row[index] = xmlDecode(rawValue).replace(/\.0$/, "");
      }
    }
    rows.set(rowNumber, row);
  }

  for (const [cellKey, mediaPath] of imageAnchors.entries()) {
    const [rowText, colText] = cellKey.split(":");
    const rowNumber = Number(rowText);
    const col = Number(colText);
    const sourceExt = extname(mediaPath).toLowerCase() || ".jpg";
    const sourceBase = basename(mediaPath, sourceExt);
    const outputName = `${sourceBase}.jpg`;
    imageUrls.set(mediaPath, `/fabric-stock/images/${outputName}`);
    const row = rows.get(rowNumber) ?? [];
    const imageCell = `/fabric-stock/images/${outputName}`;
    if (String(row[col] ?? "").trim()) {
      const nearbyEmptyCol = [col - 1, col + 1, col - 2, col + 2]
        .find((candidate) => candidate >= 0 && !String(row[candidate] ?? "").trim());
      row[nearbyEmptyCol ?? col] = imageCell;
    } else {
      row[col] = imageCell;
    }
    rows.set(rowNumber, row);
  }

  const sortedRows = [...rows.entries()].sort((a, b) => a[0] - b[0]);
  const headerEntry = sortedRows.find(([, row]) => row.some((value) => String(value ?? "").trim()));
  const headerRowNumber = headerEntry?.[0] ?? 1;
  const headers = (headerEntry?.[1] ?? []).map((value, index) => cleanName(String(value || columnLabel(index))));
  const dataRows = sortedRows
    .filter(([rowNumber]) => rowNumber > headerRowNumber)
    .map(([rowNumber, row]) => ({ rowNumber, row }));

  const lastIndex = Math.min(13, Math.max(
    headers.findLastIndex((header) => header && !/^[A-Z]+$/.test(header)),
    ...dataRows.map(({ row }) => row.findLastIndex((value) => String(value ?? "").trim())),
  ));
  const usableColumns = Math.max(1, lastIndex + 1);
  const normalizedHeaders = Array.from({ length: usableColumns }, (_, index) => headers[index] || columnLabel(index));

  const dateColumns = new Set(
    normalizedHeaders
      .map((header, index) => /date|eta|received/i.test(header) ? index : -1)
      .filter((index) => index >= 0),
  );

  const normalizedRows = dataRows
    .map(({ row }) => Array.from({ length: usableColumns }, (_, index) => {
      const value = String(row[index] ?? "").trim();
      return dateColumns.has(index) && /^\d{5}(?:\.\d+)?$/.test(value) ? excelSerialToDate(value) : value;
    }))
    .filter((row) => row.some((value) => value.trim()));

  return { headers: normalizedHeaders, rows: normalizedRows };
}

function headerPreset(kind, currentHeaders) {
  if (kind === "order") {
    return [
      "Supplier", "Fabric Type", "Picture", "Planned Release Date", "Name",
      "Quantity Ordered", "Quantity Received", "Order Date", "Status", "ETA",
    ].slice(0, currentHeaders.length);
  }
  if (kind === "stock") {
    return [
      "Supplier", "Fabric Type", "Fabric", "Name", "Cost per Meter", "Meters in Stock",
      "Cut Pieces", "Received / Date", "Products", "Notes", "K", "L",
    ].slice(0, currentHeaders.length);
  }
  if (kind === "simple-stock") {
    return [
      "Fabric", "Name", "Price", "Meters Available", "Additional Quantity",
      "Meters Received", "Products", "Notes", "Supplier",
    ].slice(0, currentHeaders.length);
  }
  return currentHeaders;
}

function columnLabel(index) {
  let n = index + 1;
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function totalQuantity(headers, rows) {
  const quantityIndex = headers.findIndex((header) => /meters in stock|meters available|quantity ordered/i.test(header));
  if (quantityIndex < 0) return null;
  const total = rows.reduce((sum, row) => sum + (Number(String(row[quantityIndex]).replace(/,/g, "")) || 0), 0);
  return total || null;
}

function compressImages(imageUrls) {
  mkdirSync(outputImageDir, { recursive: true });
  const usedNames = new Set();

  for (const [mediaPath] of imageUrls.entries()) {
    const sourcePath = join(extractDir, mediaPath);
    if (!existsSync(sourcePath)) continue;
    const outputName = `${basename(mediaPath, extname(mediaPath))}.jpg`;
    if (usedNames.has(outputName)) continue;
    usedNames.add(outputName);
    const outputPath = join(outputImageDir, outputName);
    execFileSync("sips", [
      "-s", "format", "jpeg",
      "-s", "formatOptions", "68",
      "-Z", "700",
      sourcePath,
      "--out", outputPath,
    ], { stdio: "ignore" });
  }
}

rmSync(extractDir, { recursive: true, force: true });
mkdirSync(extractDir, { recursive: true });
rmSync(outputImageDir, { recursive: true, force: true });
execFileSync("unzip", ["-q", "-o", resolve(workbookPath), "-d", extractDir]);

const sharedStrings = parseSharedStrings();
const imageUrls = new Map();
const sheets = parseWorkbookSheets().map((sheet) => {
  const normalizedName = keyName(sheet.name);
  const sheetSlug = slug(sheet.name);
  const parsed = parseRows(sheet.worksheetPath, sharedStrings, parseImageAnchors(sheet.worksheetPath), sheetSlug, imageUrls);
  const name = displayNameByName.get(normalizedName) ?? cleanName(sheet.name);
  const kind = kindByName.get(normalizedName) ?? "stock";
  const headers = headerPreset(kind, parsed.headers);
  return {
    gid: sheetIdByName.get(normalizedName) ?? sheetSlug,
    name,
    kind,
    headers,
    rows: parsed.rows,
    rowCount: parsed.rows.length,
    totalQuantity: totalQuantity(headers, parsed.rows),
  };
});

compressImages(imageUrls);

const output = `export type FabricStockSheet = {
  gid: string;
  name: string;
  kind: string;
  headers: string[];
  rows: string[][];
  rowCount: number;
  totalQuantity: number | null;
};

export const fabricStockSheets: FabricStockSheet[] = [
${sheets.map((sheet) => `  {
    gid: ${jsString(sheet.gid)},
    name: ${jsString(sheet.name)},
    kind: ${jsString(sheet.kind)},
    headers: ${jsString(sheet.headers)},
    rows: ${jsString(sheet.rows)},
    rowCount: ${sheet.rowCount},
    totalQuantity: ${sheet.totalQuantity == null ? "null" : String(sheet.totalQuantity)},
  }`).join(",\n")}
];
`;

writeFileSync(outputDataPath, output);
rmSync(extractDir, { recursive: true, force: true });

console.log(`Imported ${sheets.length} sheets.`);
console.log(`Compressed ${imageUrls.size} image references into ${outputImageDir}.`);
