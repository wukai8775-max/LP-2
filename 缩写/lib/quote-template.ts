import ExcelJS from "exceljs";
import type { MatchedProduct, QuoteGenerationResult, RequestedProduct } from "./types";

type TemplateMap = {
  worksheet: ExcelJS.Worksheet;
  headerRow: number;
  startRow: number;
  endRow: number;
  productCol: number;
  specificationCol?: number;
  quantityCol: number;
  priceCol: number;
  abbreviationCol?: number;
  remarkCol?: number;
};

const quoteHeaderAliases = {
  product: ["product name", "product", "产品全称", "产品名称", "品名", "名称", "description"],
  abbreviation: ["abbreviation", "abbr", "产品缩写", "缩写", "简称"],
  specification: ["specification", "specificatio", "spec", "产品规格", "规格"],
  quantity: ["qty", "quantity", "数量", "pcs"],
  price: ["price ($/box)", "price", "unit price", "单价", "价格"],
  remark: ["remark", "remarks", "备注"]
};

export async function fillQuoteTemplate({
  templateBuffer,
  matched,
  unmatched,
  customerName,
  remark
}: {
  templateBuffer: Buffer;
  matched: MatchedProduct[];
  unmatched: RequestedProduct[];
  customerName: string;
  remark: string;
}): Promise<QuoteGenerationResult> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);

  workbook.creator = workbook.creator || "Excel Quote Generator";
  workbook.lastModifiedBy = "Excel Quote Generator";
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;
  (workbook.calcProperties as { forceFullCalc?: boolean }).forceFullCalc = true;

  const templateMap = findTemplateMap(workbook);
  fillCustomerFields(templateMap.worksheet, customerName, remark);
  const capacityUnmatched = fillProducts(templateMap, matched);
  const finalUnmatched = [...unmatched, ...capacityUnmatched];
  appendUnmatchedNote(templateMap.worksheet, finalUnmatched);

  const output = await workbook.xlsx.writeBuffer();
  return {
    buffer: Buffer.from(output),
    matched: matched.slice(0, matched.length - capacityUnmatched.length),
    unmatched: finalUnmatched
  };
}

function findTemplateMap(workbook: ExcelJS.Workbook): TemplateMap {
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("报价模板中没有可用工作表。");
  }

  for (const sheet of workbook.worksheets) {
    const detected = detectQuoteTable(sheet);
    if (detected) return detected;
  }

  return {
    worksheet,
    headerRow: 11,
    startRow: 12,
    endRow: 35,
    productCol: 2,
    specificationCol: 3,
    quantityCol: 6,
    priceCol: 7,
    abbreviationCol: 1,
    remarkCol: 4
  };
}

function detectQuoteTable(worksheet: ExcelJS.Worksheet): TemplateMap | null {
  const maxRow = Math.min(worksheet.rowCount || 80, 80);
  const maxCol = Math.min(worksheet.columnCount || 30, 30);

  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const productCol = findColumn(row, maxCol, quoteHeaderAliases.product);
    const quantityCol = findColumn(row, maxCol, quoteHeaderAliases.quantity);
    const priceCol = findColumn(row, maxCol, quoteHeaderAliases.price);

    if (productCol > 0 && quantityCol > 0 && priceCol > 0) {
      return {
        worksheet,
        headerRow: rowNumber,
        startRow: rowNumber + 1,
        endRow: findTableEndRow(worksheet, rowNumber + 1),
        productCol,
        specificationCol: findColumn(row, maxCol, quoteHeaderAliases.specification) || undefined,
        quantityCol,
        priceCol,
        abbreviationCol: findColumn(row, maxCol, quoteHeaderAliases.abbreviation) || undefined,
        remarkCol: findColumn(row, maxCol, quoteHeaderAliases.remark) || undefined
      };
    }
  }

  return null;
}

function findColumn(row: ExcelJS.Row, maxCol: number, aliases: string[]) {
  for (let col = 1; col <= maxCol; col += 1) {
    const text = normalizeCellText(row.getCell(col).value);
    if (aliases.some((alias) => text === normalizeCellText(alias) || text.includes(normalizeCellText(alias)))) {
      return col;
    }
  }
  return 0;
}

function findTableEndRow(worksheet: ExcelJS.Worksheet, startRow: number) {
  const fallbackEnd = Math.min(startRow + 29, Math.max(worksheet.rowCount, startRow + 10));
  for (let rowNumber = startRow; rowNumber <= Math.min(worksheet.rowCount, startRow + 80); rowNumber += 1) {
    const rowText = rowToText(worksheet.getRow(rowNumber));
    if (rowText && /合计|subtotal|total|运费|手续费|bank|shipping/i.test(rowText)) {
      return Math.max(startRow, rowNumber - 1);
    }
  }
  return fallbackEnd;
}

function fillCustomerFields(worksheet: ExcelJS.Worksheet, customerName: string, remark: string) {
  if (customerName) {
    setValueNextToLabel(worksheet, ["客户", "customer", "client", "company"], customerName);
  }
  if (remark) {
    setValueNextToLabel(worksheet, ["备注", "remark", "remarks", "note"], remark);
  }
}

function setValueNextToLabel(worksheet: ExcelJS.Worksheet, aliases: string[], value: string) {
  for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, 40); rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    for (let col = 1; col <= Math.min(worksheet.columnCount || 20, 20); col += 1) {
      const text = normalizeCellText(row.getCell(col).value);
      if (aliases.some((alias) => text.includes(normalizeCellText(alias)))) {
        const target = row.getCell(col + 1);
        if (!isFormulaCell(target)) target.value = value;
        return;
      }
    }
  }
}

function fillProducts(templateMap: TemplateMap, matched: MatchedProduct[]): RequestedProduct[] {
  const capacity = Math.max(0, templateMap.endRow - templateMap.startRow + 1);
  const accepted = matched.slice(0, capacity);
  const overflow = matched.slice(capacity).map((item) => item.request);

  accepted.forEach((item, index) => {
    const row = templateMap.worksheet.getRow(templateMap.startRow + index);
    setCell(row, templateMap.abbreviationCol, item.product.abbreviation);
    setCell(row, templateMap.productCol, item.product.productName || item.product.abbreviation);
    setCell(row, templateMap.specificationCol, item.product.specification);
    setCell(row, templateMap.quantityCol, item.request.quantity);
    setCell(row, templateMap.priceCol, toNumberOrText(item.product.price));
    if (templateMap.remarkCol) {
      setCell(row, templateMap.remarkCol, item.product.sheetName);
    }
    row.commit();
  });

  return overflow;
}

function setCell(row: ExcelJS.Row, col: number | undefined, value: string | number) {
  if (!col) return;
  const cell = row.getCell(col);
  if (!isFormulaCell(cell)) cell.value = value;
}

function appendUnmatchedNote(worksheet: ExcelJS.Worksheet, unmatched: RequestedProduct[]) {
  if (unmatched.length === 0) return;

  const startRow = Math.max(worksheet.rowCount + 2, 1);
  const title = worksheet.getCell(startRow, 1);
  title.value = "未匹配产品：";
  title.font = { ...(title.font || {}), bold: true, color: { argb: "FF9A3412" } };

  unmatched.forEach((item, index) => {
    worksheet.getCell(startRow + index + 1, 1).value = `${index + 1}. ${item.input}`;
    worksheet.getCell(startRow + index + 1, 2).value = item.quantity;
  });
}

function isFormulaCell(cell: ExcelJS.Cell) {
  return Boolean(cell.value && typeof cell.value === "object" && "formula" in cell.value);
}

function normalizeCellText(value: unknown): string {
  if (value == null) return "";
  if (typeof value === "object") {
    const richText = (value as { richText?: Array<{ text: string }> }).richText;
    if (richText) return richText.map((item) => item.text).join("").trim().toLowerCase();
    const formulaResult = (value as { result?: unknown }).result;
    if (formulaResult != null) return String(formulaResult).trim().toLowerCase();
  }
  return String(value).replace(/\s+/g, " ").trim().toLowerCase();
}

function rowToText(row: ExcelJS.Row) {
  const values: string[] = [];
  row.eachCell({ includeEmpty: false }, (cell) => values.push(normalizeCellText(cell.value)));
  return values.join(" ");
}

function toNumberOrText(value: string | number) {
  if (typeof value === "number") return value;
  const cleaned = String(value).replace(/[$,￥¥]/g, "").trim();
  const number = Number(cleaned);
  return Number.isFinite(number) && cleaned !== "" ? number : value;
}
