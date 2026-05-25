import { readFile } from "fs/promises";
import path from "path";
import { NextResponse } from "next/server";
import { generateFromFormulaWorkbook } from "../../../lib/formula-workbook";

export const runtime = "nodejs";

const defaultTemplatePath = path.join(process.cwd(), "assets", "sales-price-template.xls");

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const formulaFile = formData.get("formulaFile");
    const products = String(formData.get("products") || "");

    const workbookBuffer =
      formulaFile instanceof File && formulaFile.size > 0
        ? Buffer.from(await formulaFile.arrayBuffer())
        : await readFile(defaultTemplatePath);

    const result = await generateFromFormulaWorkbook({
      workbookBuffer,
      productInput: products
    });

    const stamp = timestamp();
    return NextResponse.json({
      fileName: `报价单_${stamp}.xlsx`,
      fileBase64: result.buffer.toString("base64"),
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
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "报价单生成失败。" },
      { status: 500 }
    );
  }
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
