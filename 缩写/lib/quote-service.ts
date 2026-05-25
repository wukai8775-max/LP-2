import { parseProductInput } from "./input";
import { matchProducts } from "./match";
import { readPriceWorkbook } from "./price-workbook";
import { fillQuoteTemplate } from "./quote-template";

export async function generateQuoteExcel({
  priceBuffer,
  templateBuffer,
  customerName,
  remark,
  productInput
}: {
  priceBuffer: Buffer;
  templateBuffer: Buffer;
  customerName: string;
  remark: string;
  productInput: string;
}) {
  const records = readPriceWorkbook(priceBuffer);
  if (records.length === 0) {
    throw new Error("销售价格表中没有识别到产品数据，请检查表头是否包含 Abbreviation / Product Name / Price。");
  }

  const requested = parseProductInput(productInput);
  if (requested.length === 0) {
    throw new Error("请输入需要报价的产品。");
  }

  const { matched, unmatched } = matchProducts(requested, records);
  return fillQuoteTemplate({
    templateBuffer,
    matched,
    unmatched,
    customerName,
    remark
  });
}
