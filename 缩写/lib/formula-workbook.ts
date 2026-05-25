import * as XLSX from "xlsx";
import { matchProducts } from "./match";
import { parseProductInput } from "./input";
import { readPriceWorkbook } from "./price-workbook";
import type { ProductRecord, RequestedProduct } from "./types";

export async function generateFromFormulaWorkbook({
  workbookBuffer,
  productInput
}: {
  workbookBuffer: Buffer;
  productInput: string;
}) {
  const records = readPriceWorkbook(workbookBuffer);
  if (records.length === 0) {
    throw new Error("Excel 中没有识别到产品数据，请检查表头是否包含产品名称、简称、规格等字段。");
  }

  const requested = parseProductInput(productInput);
  if (requested.length === 0) {
    throw new Error("请输入需要报价的产品。");
  }

  const { matched, unmatched } = matchProducts(requested, records);
  const workbook = XLSX.read(workbookBuffer, {
    type: "buffer",
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

  const buffer = XLSX.write(workbook, {
    type: "buffer",
    bookType: "xlsx",
    cellStyles: true,
    compression: true
  });

  return {
    buffer: Buffer.from(buffer),
    matched,
    unmatched
  };
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
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
      const text = String(cell?.w ?? cell?.v ?? "").replace(/\s+/g, " ").trim().toLowerCase();
      if (aliases.some((alias) => text === alias.toLowerCase() || text.includes(alias.toLowerCase()))) return col;
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
