import type { MatchedProduct, ProductRecord, RequestedProduct } from "./types";

export function normalizeText(value: unknown): string {
  return String(value ?? "")
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

export function compactText(value: unknown): string {
  return normalizeText(value).replace(/[\s._\-\/\\()（）]+/g, "");
}

export function matchProducts(requested: RequestedProduct[], records: ProductRecord[]) {
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
