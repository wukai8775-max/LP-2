import * as XLSX from "xlsx";
import type { ProductRecord } from "./types";
import { normalizeText } from "./match";

type ColumnMap = {
  abbreviation: number;
  productName: number;
  specification: number;
  price: number;
};

const headerAliases: Record<keyof ColumnMap, string[]> = {
  abbreviation: ["abbreviation", "abbr", "产品缩写", "缩写", "简称", "产品简称", "型号"],
  productName: ["product name", "product", "产品全称", "产品名称", "品名", "名称", "描述", "description"],
  specification: ["specification", "spec", "产品规格", "规格", "尺寸"],
  price: ["price ($/box)", "price", "unit price", "单价", "价格", "销售价", "报价"]
};

export function readPriceWorkbook(buffer: Buffer): ProductRecord[] {
  const workbook = XLSX.read(buffer, {
    type: "buffer",
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
