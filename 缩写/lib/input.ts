import type { RequestedProduct } from "./types";

export function parseProductInput(input: string): RequestedProduct[] {
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
