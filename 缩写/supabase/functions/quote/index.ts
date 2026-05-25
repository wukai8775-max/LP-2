import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import * as XLSX from "xlsx";

type RequestedProduct = {
  input: string;
  quantity: number;
};

type ProductRecord = {
  abbreviation: string;
  productName: string;
  specification: string;
  price: string | number;
  sheetName: string;
  rowNumber: number;
};

type MatchedProduct = {
  request: RequestedProduct;
  product: ProductRecord;
  matchType: string;
  score: number;
};

type ColumnMap = {
  abbreviation: number;
  productName: number;
  specification: number;
  price: number;
};

const headerAliases: Record<keyof ColumnMap, string[]> = {
  abbreviation: ["abbreviation", "abbr", "abb eviatio", "产品缩写", "缩写", "简称", "型号"],
  productName: ["product name", "product", "产品全称", "产品名称", "品名", "名称", "description"],
  specification: ["specification", "specificatio", "spec", "产品规格", "规格"],
  price: ["price ($/box)", "price", "unit price", "单价", "价格", "销售价", "报价"]
};

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};

Deno.serve(async (request: Request) => {
  if (request.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  if (request.method !== "POST") {
    return json({ error: "Only POST is supported." }, 405);
  }

  try {
    const formData = await request.formData();
    const formulaFile = formData.get("formulaFile");
    const products = String(formData.get("products") || "");

    if (!products.trim()) {
      return json({ error: "请输入需要生成报价的产品。" }, 400);
    }

    const workbookBytes =
      formulaFile instanceof File && formulaFile.size > 0
        ? new Uint8Array(await formulaFile.arrayBuffer())
        : await readEmbeddedTemplate();

    const result = generateFromFormulaWorkbook(workbookBytes, products);
    const stamp = timestamp();

    return json({
      fileName: `报价单_${stamp}.xlsx`,
      fileBase64: bytesToBase64(result.buffer),
      matched: result.matched.map((item) => ({
        input: item.request.input,
        quantity: item.request.quantity,
        abbreviation: item.product.abbreviation,
        productName: item.product.productName,
        specification: item.product.specification,
        price: item.product.price,
        sheetName: item.product.sheetName,
        matchType: item.matchType
      })),
      unmatched: result.unmatched
    });
  } catch (error) {
    console.error(error);
    return json({ error: error instanceof Error ? error.message : "报价单生成失败。" }, 500);
  }
});

async function readEmbeddedTemplate() {
  const templateUrl = new URL("./sales-price-template.b64", import.meta.url);
  const base64 = (await Deno.readTextFile(templateUrl)).replace(/\s+/g, "");
  return base64ToBytes(base64);
}

function generateFromFormulaWorkbook(workbookBytes: Uint8Array, productInput: string) {
  const records = readPriceWorkbook(workbookBytes);
  if (records.length === 0) {
    throw new Error("Excel 中没有识别到产品数据，请检查表头是否包含产品名称、简称、规格等字段。");
  }

  const requested = parseProductInput(productInput);
  if (requested.length === 0) {
    throw new Error("请输入需要报价的产品。");
  }

  const { matched, unmatched } = matchProducts(requested, records);
  const workbook = XLSX.read(workbookBytes, {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
    bookVBA: true
  });

  const matchedKeys = new Set(matched.map((item) => recordKey(item.product)));
  hideUnselectedProductRows(workbook, records, matchedKeys);
  writeRequestedQuantities(workbook, matched);
  appendUnmatchedInfo(workbook, unmatched);

  const output = XLSX.write(workbook, {
    type: "array",
    bookType: "xlsx",
    cellStyles: true,
    compression: true
  });

  return {
    buffer: output instanceof Uint8Array ? output : new Uint8Array(output),
    matched,
    unmatched
  };
}

function readPriceWorkbook(workbookBytes: Uint8Array): ProductRecord[] {
  const workbook = XLSX.read(workbookBytes, {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellNF: true
  });

  const records: ProductRecord[] = [];

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet?.["!ref"]) continue;

    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: false, defval: "" });
    const headerInfo = findHeader(rows);
    if (!headerInfo) continue;

    for (let index = headerInfo.rowIndex + 1; index < rows.length; index += 1) {
      const row = rows[index] || [];
      const abbreviation = cell(row, headerInfo.columns.abbreviation);
      const productName = cell(row, headerInfo.columns.productName);
      const specification = cell(row, headerInfo.columns.specification);
      const price = cell(row, headerInfo.columns.price);

      if (!abbreviation && !productName && !specification) continue;

      records.push({
        abbreviation,
        productName,
        specification,
        price,
        sheetName,
        rowNumber: index + 1
      });
    }
  }

  return records;
}

function parseProductInput(input: string): RequestedProduct[] {
  return input
    .split(/[\n\r,，;；、]+/g)
    .map((line) => line.trim())
    .filter(Boolean)
    .map(parseLine);
}

function parseLine(line: string): RequestedProduct {
  const normalized = line.replace(/\s+/g, " ").trim();
  const quantityMatch = normalized.match(/^(.*?)(?:\s*[*xX×]\s*|\s+)(\d+(?:\.\d+)?)$/);

  if (!quantityMatch) {
    return { input: normalized, quantity: 1 };
  }

  const name = quantityMatch[1].trim();
  const quantity = Number(quantityMatch[2]);

  return {
    input: name || normalized,
    quantity: Number.isFinite(quantity) && quantity > 0 ? quantity : 1
  };
}

function matchProducts(requested: RequestedProduct[], records: ProductRecord[]) {
  const matched: MatchedProduct[] = [];
  const unmatched: RequestedProduct[] = [];
  const rowMap = buildSearchMap(records);

  for (const request of requested) {
    const query = compactText(request.input);
    const direct = rowMap.get(query);

    if (direct) {
      matched.push({ request, product: direct, matchType: "exact", score: 100 });
      continue;
    }

    const fuzzy = findBestFuzzyMatch(request.input, records);
    if (fuzzy) {
      matched.push({ request, ...fuzzy });
    } else {
      unmatched.push(request);
    }
  }

  return { matched, unmatched };
}

function buildSearchMap(records: ProductRecord[]) {
  const map = new Map<string, ProductRecord>();

  for (const record of records) {
    for (const value of [record.abbreviation, record.productName, record.specification]) {
      const key = compactText(value);
      if (key && !map.has(key)) map.set(key, record);
    }
  }

  return map;
}

function findBestFuzzyMatch(input: string, records: ProductRecord[]) {
  let best: { product: ProductRecord; matchType: string; score: number } | null = null;

  for (const record of records) {
    const candidates = [
      scoreCandidate(input, record.abbreviation, "abbreviation"),
      scoreCandidate(input, record.productName, "productName"),
      scoreCandidate(input, record.specification, "specification")
    ].sort((a, b) => b.score - a.score);
    const candidate = candidates[0];

    if (candidate.score >= 70 && (!best || candidate.score > best.score)) {
      best = { product: record, matchType: candidate.type, score: candidate.score };
    }
  }

  return best;
}

function scoreCandidate(input: string, candidate: unknown, type: string) {
  const queryCompact = compactText(input);
  const candidateCompact = compactText(candidate);

  if (!queryCompact || !candidateCompact) return { type, score: 0 };
  if (candidateCompact === queryCompact) return { type, score: 95 };
  if (candidateCompact.includes(queryCompact)) return { type, score: 82 };
  if (queryCompact.includes(candidateCompact) && candidateCompact.length >= 3) return { type, score: 76 };

  return { type, score: 0 };
}

function findHeader(rows: unknown[][]): { rowIndex: number; columns: ColumnMap } | null {
  const scanLimit = Math.min(rows.length, 80);

  for (let rowIndex = 0; rowIndex < scanLimit; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const columns = findColumns(row);
    const hasProductColumn = columns.abbreviation >= 0 || columns.productName >= 0 || columns.specification >= 0;

    if (hasProductColumn) {
      const fallback = firstTextColumn(row);
      return {
        rowIndex,
        columns: {
          abbreviation: columns.abbreviation,
          productName: columns.productName >= 0 ? columns.productName : fallback,
          specification: columns.specification >= 0 ? columns.specification : columns.productName >= 0 ? columns.productName : fallback,
          price: columns.price
        }
      };
    }
  }

  return null;
}

function findColumns(row: unknown[]): ColumnMap {
  return {
    abbreviation: findColumn(row, headerAliases.abbreviation),
    productName: findColumn(row, headerAliases.productName),
    specification: findColumn(row, headerAliases.specification),
    price: findColumn(row, headerAliases.price)
  };
}

function findColumn(row: unknown[], aliases: string[]) {
  return row.findIndex((value) => {
    const text = normalizeText(value);
    return aliases.some((alias) => text === normalizeText(alias) || text.includes(normalizeText(alias)));
  });
}

function firstTextColumn(row: unknown[]) {
  const index = row.findIndex((value) => String(value ?? "").trim());
  return index >= 0 ? index : 0;
}

function cell(row: unknown[], index: number) {
  if (index < 0) return "";
  return String(row[index] ?? "").trim();
}

function hideUnselectedProductRows(workbook: XLSX.WorkBook, records: ProductRecord[], matchedKeys: Set<string>) {
  for (const record of records) {
    const sheet = workbook.Sheets[record.sheetName];
    if (!sheet) continue;
    if (!sheet["!rows"]) sheet["!rows"] = [];

    const rowIndex = record.rowNumber - 1;
    if (!sheet["!rows"][rowIndex]) sheet["!rows"][rowIndex] = {};
    sheet["!rows"][rowIndex].hidden = !matchedKeys.has(recordKey(record));
  }
}

function writeRequestedQuantities(
  workbook: XLSX.WorkBook,
  matchedProducts: Array<{ product: ProductRecord; request: RequestedProduct }>
) {
  const quantityByRecord = new Map<string, number>();
  matchedProducts.forEach((item) => {
    quantityByRecord.set(recordKey(item.product), item.request.quantity || 1);
  });

  for (const { product: record } of matchedProducts) {
    const sheet = workbook.Sheets[record.sheetName];
    if (!sheet?.["!ref"]) continue;

    const qtyColumn = findNearbyHeaderColumn(sheet, record.rowNumber, ["qty", "quantity", "数量", "件数"]);
    if (qtyColumn < 0) continue;

    const address = XLSX.utils.encode_cell({ r: record.rowNumber - 1, c: qtyColumn });
    const existing = sheet[address] || { t: "n" };
    sheet[address] = {
      ...existing,
      t: "n",
      v: quantityByRecord.get(recordKey(record)) || 1
    };
  }
}

function findNearbyHeaderColumn(sheet: XLSX.WorkSheet, productRowNumber: number, aliases: string[]) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const firstRow = Math.max(range.s.r, productRowNumber - 5);
  const lastRow = Math.min(range.e.r, productRowNumber - 1);

  for (let row = lastRow; row >= firstRow; row -= 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const value = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
      const text = normalizeText(value?.w ?? value?.v ?? "");
      if (aliases.some((alias) => text === normalizeText(alias) || text.includes(normalizeText(alias)))) return col;
    }
  }

  return -1;
}

function appendUnmatchedInfo(workbook: XLSX.WorkBook, unmatched: RequestedProduct[]) {
  if (unmatched.length === 0) return;

  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  if (!sheet) return;

  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const col = range.s.c;
  const startRow = range.e.r + 2;

  sheet[XLSX.utils.encode_cell({ r: startRow, c: col })] = {
    t: "s",
    v: "未匹配到的产品如下，已跳过："
  };

  unmatched.forEach((item, index) => {
    sheet[XLSX.utils.encode_cell({ r: startRow + index + 1, c: col })] = {
      t: "s",
      v: item.input
    };
    sheet[XLSX.utils.encode_cell({ r: startRow + index + 1, c: col + 1 })] = {
      t: "n",
      v: item.quantity
    };
  });

  sheet["!ref"] = XLSX.utils.encode_range({
    s: range.s,
    e: { r: startRow + unmatched.length, c: range.e.c }
  });
}

function recordKey(record: ProductRecord) {
  return `${record.sheetName}::${record.rowNumber}`;
}

function normalizeText(value: unknown): string {
  return String(value ?? "")
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function compactText(value: unknown): string {
  return normalizeText(value).replace(/[\s._\-/\\()（）]+/g, "");
}

function json(payload: unknown, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      ...corsHeaders,
      "Content-Type": "application/json"
    }
  });
}

function timestamp() {
  const date = new Date();
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  return `${yyyy}${mm}${dd}_${hh}${mi}`;
}

function base64ToBytes(base64: string) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function bytesToBase64(bytes: Uint8Array) {
  const chunkSize = 0x8000;
  let binary = "";
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  return btoa(binary);
}
