export type RequestedProduct = {
  input: string;
  quantity: number;
};

export type ProductRecord = {
  abbreviation: string;
  productName: string;
  specification: string;
  price: string | number;
  discount?: string | number;
  purchasePrice?: string | number;
  sheetName: string;
  rowNumber: number;
};

export type MatchedProduct = {
  request: RequestedProduct;
  product: ProductRecord;
  matchType: string;
  score: number;
};

export type QuoteGenerationResult = {
  buffer: Buffer;
  matched: MatchedProduct[];
  unmatched: RequestedProduct[];
};
