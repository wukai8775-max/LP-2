(function () {
  const TEMPLATE_URL = "./assets/sales-price-template.xls";
  const TEMPLATE_NAME = "5.21 销售价格表.xls";
  const MAX_QUOTE_LINES = 10;
  const SHIPPING_TIERS = [
    { max: 10, fee: 45 },
    { max: 20, fee: 70 },
    { max: 30, fee: 95 },
    { max: 40, fee: 120 },
    { max: 50, fee: 145 },
    { max: 60, fee: 170 },
    { max: 70, fee: 195 },
    { max: 80, fee: 220 },
    { max: 90, fee: 245 },
    { max: 100, fee: 270 },
  ];
  const QUOTE_LABELS_TO_BOLD = ["abb eviatio", "total price", "shipping fee", "transaction fees (5%)", "amount"];
  const QUOTE_COLUMNS = {
    serial: 0,
    abbreviation: 1,
    productName: 2,
    specification: 3,
    quantity: 4,
    price: 5,
    discount: 6,
    amount: 7,
    purchaseUnit: 8,
    purchaseTotal: 9,
    profit: 10,
  };
  const QUOTE_HEADER_ROW = 1;
  const QUOTE_PRODUCT_START_ROW = 2;
  const PROFIT_EXCHANGE_RATE = 7;

  const state = {
    workbook: null,
    priceSheetInfo: null,
    quoteSheetName: "",
    productRows: [],
    searchIndex: new Map(),
    matched: [],
    missing: [],
    pending: [],
  };

  const el = {
    fileInput: document.getElementById("templateInput"),
    fileName: document.getElementById("fileName"),
    status: document.getElementById("statusPill"),
    sheetName: document.getElementById("sheetName"),
    productCount: document.getElementById("productCount"),
    quoteSheetName: document.getElementById("quoteSheetName"),
    input: document.getElementById("productInput"),
    preview: document.getElementById("previewButton"),
    generate: document.getElementById("generateButton"),
    results: document.getElementById("results"),
    matchedList: document.getElementById("matchedList"),
    missingList: document.getElementById("missingList"),
  };

  const headerAliases = {
    abbreviation: ["abbreviation", "abbreviatio", "abbr", "abb", "产品缩写", "缩写", "简称", "型号", "model"],
    productName: ["product name", "product", "name", "产品全称", "产品名称", "品名", "名称", "description"],
    specification: ["specification", "specificatio", "spec", "产品规格", "规格", "size"],
    price: ["price ($/box)", "price", "unit price", "销售价格", "售价", "单价", "报价", "价格"],
    purchaseCost: ["采购", "采购单价", "采购成本", "purchase", "purchase cost", "cost"],
  };
  const PRODUCT_ALIAS_GROUPS = [
    { name: "AOD-9604", abbreviations: ["5AD", "10AD"] },
    { name: "Cagrilintide", abbreviations: ["CGL5", "CGL10"] },
    { name: "Tesamorelin", abbreviations: ["TSM5", "TSM10"] },
    { name: "HGH 191AA", abbreviations: ["H10", "H24", "H36"] },
    { name: "CJC-1295+IPA", abbreviations: ["CP10"] },
    { name: "CJC-1295 With DAC", abbreviations: ["CD5"] },
    { name: "Ipamorelin", abbreviations: ["IP5", "IP10"] },
    { name: "IGF-1 LR3", abbreviations: ["IG01", "IG1"] },
    { name: "MOTS-C", abbreviations: ["MS10", "MS15", "MS20", "MS40"] },
    { name: "Sermorelin Acetate", abbreviations: ["SMO-5", "SMO-10"] },
    { name: "SS-31", abbreviations: ["2S10", "2S50"] },
    { name: "BPC-157", abbreviations: ["BC5", "BC10"] },
    { name: "BPC+TB组合", abbreviations: ["BB10"] },
    { name: "TB-500", abbreviations: ["BT5", "BT10"] },
    { name: "GLOW", abbreviations: ["BBG70"] },
    { name: "KLOW", abbreviations: ["BBKG80"] },
    { name: "GHK-Cu", abbreviations: ["CU50", "CU100"] },
    { name: "DSIP", abbreviations: ["DS5", "DS10"] },
    { name: "Epithalon", abbreviations: ["ET10", "ET50"] },
    { name: "Selank", abbreviations: ["SK10"] },
    { name: "Thymosin Alpha-1", abbreviations: ["TA5", "TA10"] },
    { name: "KPV", abbreviations: ["KPV"] },
    { name: "Glutathione", abbreviations: ["GTT"] },
    { name: "NAD+", abbreviations: ["NJ100", "NJ500", "NJ1000"] },
    { name: "HCG", abbreviations: ["HC5000"] },
    { name: "Kisspeptin-10", abbreviations: ["KS10"] },
    { name: "PT-141", abbreviations: ["P41"] },
    { name: "MT-2", abbreviations: ["ML10"] },
    { name: "Semax", abbreviations: ["XA5", "XA10"] },
    { name: "Melanotan I", abbreviations: ["MT1"] },
    { name: "SNAP-8", abbreviations: ["NP810"] },
    { name: "BAC WATER", abbreviations: ["WA3", "WA1"] },
  ];

  function normalize(value) {
    return String(value == null ? "" : value)
      .replace(/\u3000/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  }

  function compact(value) {
    return normalize(value).replace(/[\s._\-\/\\()[\]（）]+/g, "");
  }

  function aliasKey(value) {
    return normalize(value).replace(/[^a-z0-9\u4e00-\u9fff]+/g, "");
  }

  function normalizeSpec(value) {
    return normalize(value).replace(/[\s._\-\/\\()[\]（）*xX×]+/g, "");
  }

  function displayValue(cell) {
    if (!cell) return "";
    if (cell.w != null) return String(cell.w);
    if (cell.v != null) return String(cell.v);
    return "";
  }

  function abbreviationParts(value) {
    const match = compact(value).match(/^([a-z]+)(\d+(?:\.\d+)?)?$/i);
    return {
      prefix: match ? match[1] : compact(value),
      number: match && match[2] ? Number(match[2]) : null,
    };
  }

  function specificationNumber(value) {
    const match = String(value || "").match(/(\d+(?:\.\d+)?)\s*(?:mg|ml)?/i);
    return match ? Number(match[1]) : Number.POSITIVE_INFINITY;
  }

  function extractTrailingSpec(value) {
    const text = String(value || "").trim();
    const match = text.match(/^(.*?)(\d+(?:\.\d+)?\s*(?:mg|ml)\s*(?:[*xX×]\s*\d+\s*vials?)?)\s*$/i);
    if (!match || !match[1].trim()) return { input: text, specification: "" };
    return {
      input: match[1].trim(),
      specification: match[2].replace(/\s+/g, ""),
    };
  }

  function isStandaloneSpecification(line) {
    return /^\d+(?:\.\d+)?\s*(?:mg|ml|iu)\s*(?:[*xX×脳]\s*\d+\s*vials?)?\s*$/i.test(String(line || "").trim());
  }

  function isStandaloneQuantity(line) {
    const text = String(line || "").trim();
    const match = text.match(/^(?:x\s*)?(\d+(?:\.\d+)?)(?:\s*(?:盒|box|boxes|vials?))?$/i);
    if (!match) return false;
    const quantity = Number(match[1]);
    return Number.isFinite(quantity) && quantity > 0;
  }

  function parseStandaloneQuantity(line) {
    const match = String(line || "").trim().match(/^(?:x\s*)?(\d+(?:\.\d+)?)(?:\s*(?:盒|box|boxes|vials?))?$/i);
    return match ? Number(match[1]) : 1;
  }

  function normalizeStandaloneSpec(line) {
    return String(line || "").trim().replace(/\s+/g, "");
  }

  function parseBracketQuantity(line) {
    const match = String(line || "").trim().match(/^(.*?)\s*[\(（]\s*(\d+(?:\.\d+)?)\s*[\)）]\s*$/);
    if (!match || !match[1].trim()) return null;
    const quantity = Number(match[2]);
    if (!Number.isFinite(quantity) || quantity <= 0) return null;
    return { input: match[1].trim(), quantity };
  }

  function splitInlineProductItems(line) {
    const text = String(line || "").replace(/\s+/g, " ").trim();
    if (!text) return [];
    const matches = Array.from(text.matchAll(/(.+?)\s*(?:[\(（]\s*(\d+(?:\.\d+)?)\s*[\)）]|\b[xX]\s*(\d+(?:\.\d+)?))(?=\s+\S+\s*(?:[\(（]|\b[xX]\s*)|$)/g));
    if (matches.length <= 1) return [];

    const items = [];
    let consumed = "";
    for (const match of matches) {
      const name = String(match[1] || "").trim();
      const quantity = Number(match[2] || match[3]);
      if (!name || !Number.isFinite(quantity) || quantity <= 0) return [];
      items.push(`${name} x ${quantity}`);
      consumed += match[0];
    }

    return consumed.replace(/\s+/g, "") === text.replace(/\s+/g, "") ? items : [];
  }

  function hasInlineRequestInfo(line) {
    const parsed = parseRequestedLine(line);
    return Boolean(parsed.specification) || parsed.quantity !== 1 || String(line || "").includes("|");
  }

  function buildVerticalRequests(lines) {
    const requests = [];
    let current = null;

    function flushCurrent() {
      if (!current) return;
      requests.push({
        input: current.input,
        specification: current.specification || "",
        quantity: Number.isFinite(current.quantity) && current.quantity > 0 ? current.quantity : 1,
      });
      current = null;
    }

    for (const rawLine of lines) {
      const line = rawLine.trim();
      if (!line) continue;

      if (isStandaloneSpecification(line)) {
        if (current && !current.specification) {
          current.specification = normalizeStandaloneSpec(line);
        } else {
          flushCurrent();
          current = { input: line, specification: "", quantity: 1 };
        }
        continue;
      }

      if (isStandaloneQuantity(line) && current) {
        current.quantity = parseStandaloneQuantity(line);
        flushCurrent();
        continue;
      }

      flushCurrent();
      current = parseRequestedLine(line);
      if (hasInlineRequestInfo(line)) flushCurrent();
    }

    flushCurrent();
    return requests;
  }

  function setStatus(text, tone) {
    el.status.textContent = text;
    el.status.className = `status-pill ${tone || ""}`.trim();
  }

  function parseRequestedProducts() {
    return el.input.value
      .split(/[\n\r,，;；]+/g)
      .map((item) => item.trim())
      .filter(Boolean)
      .map(parseRequestedLine);
  }

  function parseRequestedProductsSmartOld() {
    const lines = el.input.value
      .split(/[\n\r,，;；]+/g)
      .map((item) => item.trim())
      .filter(Boolean);
    return buildVerticalRequests(lines);
  }

  function parseRequestedProductsSmart() {
    const baseLines = el.input.value
      .split(/[\n\r,，;；]+/g)
      .map((item) => item.trim())
      .filter(Boolean);
    const lines = [];
    for (const line of baseLines) {
      const inlineItems = splitInlineProductItems(line);
      if (inlineItems.length > 1) lines.push(...inlineItems);
      else {
        const bracketQuantity = parseBracketQuantity(line);
        lines.push(bracketQuantity ? `${bracketQuantity.input} x ${bracketQuantity.quantity}` : line);
      }
    }
    return buildVerticalRequests(lines);
  }

  function parseRequestedLine(line) {
    let normalized = line.replace(/\s+/g, " ").trim();
    let specification = "";
    let explicitSpecification = false;

    if (normalized.includes("|")) {
      const parts = normalized.split("|");
      normalized = parts.shift().trim();
      specification = parts.join("|").trim();
      explicitSpecification = Boolean(specification);
    }

    const match = normalized.match(/^(.*?)(?:\s*[*xX×]\s*|\s+)(\d+(?:\.\d+)?)$/);
    let quantity = 1;
    if (match) {
      quantity = Number(match[2]);
      normalized = match[1].trim() || normalized;
    }
    if (!explicitSpecification) {
      const extracted = extractTrailingSpec(normalized);
      normalized = extracted.input;
      specification = extracted.specification;
    }
    return {
      input: normalized,
      specification,
      quantity: Number.isFinite(quantity) && quantity > 0 ? quantity : 1,
    };
  }

  function cellText(sheet, row, col) {
    if (col == null || col < 0) return "";
    return displayValue(sheet[XLSX.utils.encode_cell({ r: row, c: col })]).trim();
  }

  function findHeaderInfo(sheet) {
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    const maxRow = Math.min(range.e.r, range.s.r + 80);

    for (let row = range.s.r; row <= maxRow; row += 1) {
      const columns = { abbreviation: -1, productName: -1, specification: -1, price: -1, purchaseCost: -1 };

      for (let col = range.s.c; col <= range.e.c; col += 1) {
        const text = normalize(displayValue(sheet[XLSX.utils.encode_cell({ r: row, c: col })]));
        if (!text) continue;

        for (const key of Object.keys(headerAliases)) {
          if (columns[key] >= 0) continue;
          if (headerAliases[key].some((alias) => text === normalize(alias) || text.includes(normalize(alias)))) {
            columns[key] = col;
          }
        }
      }

      const hasProductColumn = columns.abbreviation >= 0 || columns.productName >= 0;
      if (hasProductColumn && columns.price >= 0) {
        if (columns.productName < 0) columns.productName = columns.abbreviation;
        if (columns.abbreviation < 0) columns.abbreviation = columns.productName;
        if (columns.specification < 0) columns.specification = columns.productName;
        if (columns.purchaseCost < 0) columns.purchaseCost = columns.price + 1;
        return { headerRow: row, columns };
      }
    }

    return null;
  }

  function findPriceSheet(workbook) {
    const candidates = [];

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet || !sheet["!ref"]) return;
      const headerInfo = findHeaderInfo(sheet);
      if (!headerInfo) return;

      const range = XLSX.utils.decode_range(sheet["!ref"]);
      let filledRows = 0;
      for (let row = headerInfo.headerRow + 1; row <= range.e.r; row += 1) {
        const name = cellText(sheet, row, headerInfo.columns.productName);
        const abbr = cellText(sheet, row, headerInfo.columns.abbreviation);
        const price = cellText(sheet, row, headerInfo.columns.price);
        if ((name || abbr) && price) filledRows += 1;
      }

      candidates.push({ sheetName, sheet, filledRows, ...headerInfo });
    });

    candidates.sort((a, b) => b.filledRows - a.filledRows);
    return candidates[0] || null;
  }

  function isPriceSheetName(sheetName, priceSheetName) {
    return sheetName === priceSheetName || /采购成本|purchase cost/i.test(sheetName);
  }

  function getQuoteSheetCandidates(workbook, priceSheetName) {
    return workbook.SheetNames.filter((name) => !isPriceSheetName(name, priceSheetName));
  }

  function findQuoteSheetName(workbook, priceSheetName, productCount) {
    const candidates = getQuoteSheetCandidates(workbook, priceSheetName);
    if (productCount) {
      const numbered = candidates.find((name) => name.trim() === String(productCount));
      if (numbered) return numbered;
    }

    const preferred = candidates.find((name) => /报价工具|报价单|报价|quote/i.test(name));
    if (preferred) return preferred;

    const firstNumbered = candidates.find((name) => /^\d+$/.test(name.trim()));
    if (firstNumbered) return firstNumbered;

    return candidates[0] || "";
  }

  function findQuoteInputSheetName(workbook, priceSheetName) {
    const candidates = getQuoteSheetCandidates(workbook, priceSheetName);
    return candidates.find((name) => /报价工具|quote tool/i.test(name)) || "";
  }

  function addSearchKey(value, product) {
    const key = compact(value);
    if (!key) return;
    if (!state.searchIndex.has(key)) state.searchIndex.set(key, []);
    state.searchIndex.get(key).push(product);
  }

  function aliasGroupsForProduct(product) {
    const productNameKey = aliasKey(product.productName);
    const abbreviationKey = aliasKey(product.abbreviation);
    return PRODUCT_ALIAS_GROUPS.filter((group) => {
      const groupNameKey = aliasKey(group.name);
      const abbreviationKeys = group.abbreviations.map(aliasKey);
      return (
        abbreviationKeys.includes(abbreviationKey) ||
        (groupNameKey && productNameKey && (productNameKey.includes(groupNameKey) || groupNameKey.includes(productNameKey)))
      );
    });
  }

  function splitSearchKeys(value) {
    return String(value || "")
      .split(/[,\uFF0C;；、/\\]+/g)
      .map((item) => item.trim())
      .filter((item) => item.length >= 2);
  }

  function indexProducts() {
    const info = state.priceSheetInfo;
    const range = XLSX.utils.decode_range(info.sheet["!ref"] || "A1:A1");
    state.productRows = [];
    state.searchIndex.clear();

    for (let row = info.headerRow + 1; row <= range.e.r; row += 1) {
      const abbreviation = cellText(info.sheet, row, info.columns.abbreviation);
      const productName = cellText(info.sheet, row, info.columns.productName);
      const specification = cellText(info.sheet, row, info.columns.specification);
      const price = cellText(info.sheet, row, info.columns.price);
      const purchaseCost = cellText(info.sheet, row, info.columns.purchaseCost);

      if (!abbreviation && !productName && !specification) continue;
      if (!price && /total|合计|小计|备注|remark|note/i.test(`${abbreviation} ${productName} ${specification}`)) continue;

      const abbrInfo = abbreviationParts(abbreviation);
      const product = {
        row,
        abbreviation,
        productName,
        specification,
        price,
        purchaseCost,
        seriesPrefix: abbrInfo.prefix,
        abbreviationNumber: abbrInfo.number,
        specificationNumber: specificationNumber(specification),
      };
      state.productRows.push(product);

      for (const value of [abbreviation, productName, specification]) {
        addSearchKey(value, product);
        splitSearchKeys(value).forEach((part) => addSearchKey(part, product));
      }
      for (const group of aliasGroupsForProduct(product)) {
        addSearchKey(group.name, product);
        group.abbreviations.forEach((alias) => addSearchKey(alias, product));
      }
    }
  }

  function productLabel(product) {
    return [product.abbreviation, product.productName, product.specification].filter(Boolean).join(" / ");
  }

  function filterBySpecification(products, spec) {
    if (!spec) return products;
    const normalizedSpec = normalizeSpec(spec);
    return products.filter((item) => normalizeSpec(item.specification).includes(normalizedSpec));
  }

  function pickSmallestSpec(products) {
    return uniqueProducts(products).sort((a, b) => {
      const specDiff = a.specificationNumber - b.specificationNumber;
      if (specDiff !== 0) return specDiff;
      return (a.abbreviationNumber || Number.POSITIVE_INFINITY) - (b.abbreviationNumber || Number.POSITIVE_INFINITY);
    })[0] || null;
  }

  function sameProductFamily(products) {
    const unique = uniqueProducts(products);
    const names = new Set(unique.map((item) => compact(item.productName)).filter(Boolean));
    const prefixes = new Set(unique.map((item) => compact(item.seriesPrefix)).filter(Boolean));
    return unique.length > 0 && (names.size === 1 || prefixes.size === 1);
  }

  function editDistance(a, b) {
    if (a === b) return 0;
    if (!a) return b.length;
    if (!b) return a.length;

    const previous = Array.from({ length: b.length + 1 }, (_, index) => index);
    const current = new Array(b.length + 1);
    for (let i = 1; i <= a.length; i += 1) {
      current[0] = i;
      for (let j = 1; j <= b.length; j += 1) {
        const cost = a[i - 1] === b[j - 1] ? 0 : 1;
        current[j] = Math.min(
          previous[j] + 1,
          current[j - 1] + 1,
          previous[j - 1] + cost
        );
      }
      for (let j = 0; j <= b.length; j += 1) previous[j] = current[j];
    }
    return previous[b.length];
  }

  function typoDistance(query, candidate) {
    if (!query || !candidate) return Number.POSITIVE_INFINITY;
    if (candidate.includes(query) || query.includes(candidate)) return Math.abs(candidate.length - query.length);
    return editDistance(query, candidate);
  }

  function fuzzyCandidates(query, spec) {
    const scored = [];
    for (const product of state.productRows) {
      if (spec && !normalizeSpec(product.specification).includes(spec)) continue;
      const fields = [
        { value: product.abbreviation, priority: 0 },
        { value: product.seriesPrefix, priority: 1 },
        { value: product.productName, priority: 2 },
      ];
      aliasGroupsForProduct(product).forEach((group) => {
        group.abbreviations.forEach((alias) => fields.push({ value: alias, priority: 0 }));
        fields.push({ value: group.name, priority: 2 });
      });
      let best = null;
      for (const field of fields) {
        const value = compact(field.value);
        const distance = typoDistance(query, value);
        if (distance > 3) continue;
        const score = distance * 10 + field.priority + product.specificationNumber / 1000;
        if (!best || score < best.score) best = { product, score, distance };
      }
      if (best) scored.push(best);
    }

    return uniqueProducts(
      scored
        .sort((a, b) => a.score - b.score)
        .map((item) => item.product)
    ).slice(0, 3);
  }

  function matchSeriesCandidates(query, spec) {
    const queryParts = abbreviationParts(query);
    const candidates = state.productRows.filter((product) => {
      const name = compact(product.productName);
      const prefix = compact(product.seriesPrefix);
      const abbr = compact(product.abbreviation);
      const aliasGroups = aliasGroupsForProduct(product);
      if (prefix && query === prefix) return true;
      if (name && (name.startsWith(query) || name.includes(query))) return true;
      if (query.length >= 3 && abbr.startsWith(query)) return true;
      if (queryParts.prefix && query === queryParts.prefix && prefix === queryParts.prefix) return true;
      if (aliasGroups.some((group) => {
        const groupName = compact(group.name);
        return groupName.startsWith(query) || groupName.includes(query) || group.abbreviations.some((alias) => compact(alias).startsWith(query));
      })) return true;
      return false;
    });
    return filterBySpecification(candidates, spec);
  }

  function findProduct(request) {
    const query = compact(request.input);
    const spec = normalizeSpec(request.specification);
    const exact = uniqueProducts(state.searchIndex.get(query) || []);
    const exactBySpec = spec ? filterBySpecification(exact, spec) : exact;
    if (exactBySpec.length === 1) return { product: exactBySpec[0], reason: "" };
    if (exactBySpec.length > 1 && sameProductFamily(exactBySpec)) return { product: pickSmallestSpec(exactBySpec), reason: "" };
    if (exactBySpec.length > 1) return { product: null, reason: "匹配到多个规格，请用 |规格 补充说明" };

    const seriesCandidates = matchSeriesCandidates(query, spec);
    if (seriesCandidates.length === 1) return { product: seriesCandidates[0], reason: "" };
    if (seriesCandidates.length > 1 && sameProductFamily(seriesCandidates)) return { product: pickSmallestSpec(seriesCandidates), reason: "" };
    if (seriesCandidates.length > 1) return { product: null, reason: "匹配到多个产品系列，请输入更完整的缩写或名称" };

    let best = null;
    let bestScore = 0;
    let sameScoreCount = 0;
    for (const product of state.productRows) {
      for (const value of [product.abbreviation, product.productName, product.specification]) {
        const candidate = compact(value);
        let score = 0;
        if (candidate === query) score = 100;
        else if (candidate.includes(query)) score = 84;
        else if (query.includes(candidate) && candidate.length >= 3) score = 78;
        if (spec && !normalizeSpec(product.specification).includes(spec)) score = 0;

        if (score > bestScore) {
          best = product;
          bestScore = score;
          sameScoreCount = 1;
        } else if (score > 0 && score === bestScore) {
          sameScoreCount += 1;
        }
      }
    }

    if (bestScore >= 78 && sameScoreCount === 1) return { product: best, reason: "" };
    if (bestScore >= 78) return { product: null, reason: "匹配到多个相似产品，请输入更完整的缩写或名称" };
    const suggestions = fuzzyCandidates(query, spec);
    if (suggestions.length > 0) return { product: null, pending: true, suggestions, reason: "疑似拼写错误，请确认候选产品" };
    return { product: null, reason: "价格表中未找到" };
  }

  function uniqueProducts(products) {
    const seen = new Set();
    const result = [];
    for (const product of products) {
      if (seen.has(product.row)) continue;
      seen.add(product.row);
      result.push(product);
    }
    return result;
  }

  function analyzeMatches() {
    const requests = parseRequestedProductsSmart();
    const seen = new Set();
    state.matched = [];
    state.missing = [];
    state.pending = [];

    for (const request of requests) {
      const result = findProduct(request);
      if (result.pending) {
        state.pending.push({ request, suggestions: result.suggestions, selectedIndex: 0, reason: result.reason });
        continue;
      }
      if (!result.product) {
        state.missing.push({ ...request, reason: result.reason });
        continue;
      }

      const key = String(result.product.row);
      if (seen.has(key)) continue;
      state.matched.push({ request, product: result.product });
      seen.add(key);
    }

    renderResults();
    return state.matched.length + state.missing.length + state.pending.length;
  }

  function makeListItem(text) {
    const item = document.createElement("li");
    item.textContent = text;
    return item;
  }

  function renderResults() {
    el.results.classList.add("is-visible");
    el.matchedList.innerHTML = "";
    el.missingList.innerHTML = "";

    if (state.matched.length === 0) {
      el.matchedList.appendChild(makeListItem("没有匹配到产品"));
    } else {
      for (const item of state.matched) {
        el.matchedList.appendChild(
          makeListItem(`${item.request.input} x ${item.request.quantity} -> ${productLabel(item.product)}，${formatPrice(item.product.price)}`)
        );
      }
    }

    for (const item of state.pending) {
      const choices = item.suggestions.map((product, index) => `${index + 1}. ${productLabel(product)}`).join(" / ");
      el.matchedList.appendChild(makeListItem(`待确认：${item.request.input} x ${item.request.quantity} -> ${choices}`));
    }

    if (state.missing.length === 0) {
      el.missingList.appendChild(makeListItem("无"));
    } else {
      for (const item of state.missing) {
        el.missingList.appendChild(makeListItem(`${item.input} x ${item.quantity}（${item.reason}）`));
      }
    }
  }

  function parseMoney(value) {
    if (typeof value === "number") return value;
    const text = String(value == null ? "" : value).replace(/[^0-9.-]/g, "");
    const number = Number(text);
    return Number.isFinite(number) ? number : 0;
  }

  function getTotalQuantity() {
    return state.matched.reduce((sum, item) => sum + item.request.quantity, 0);
  }

  function getShippingFee(quantity) {
    const tier = SHIPPING_TIERS.find((item) => quantity <= item.max);
    if (tier) return tier.fee;

    const baseTier = SHIPPING_TIERS[SHIPPING_TIERS.length - 1];
    const extraTiers = Math.ceil((quantity - baseTier.max) / 10);
    return baseTier.fee + extraTiers * 25;
  }

  function confirmPendingMatches() {
    if (state.pending.length === 0) return true;

    const lines = state.pending.map((item, index) => {
      const product = item.suggestions[item.selectedIndex || 0];
      return `${index + 1}. ${item.request.input} x ${item.request.quantity} -> ${productLabel(product)}`;
    });
    const ok = confirm(`以下产品是根据拼写相近自动建议的，请确认是否按这些产品生成报价单：\n\n${lines.join("\n")}`);
    if (!ok) return false;

    const seen = new Set(state.matched.map((item) => String(item.product.row)));
    for (const item of state.pending) {
      const product = item.suggestions[item.selectedIndex || 0];
      const key = String(product.row);
      if (seen.has(key)) continue;
      state.matched.push({ request: item.request, product, confirmedFuzzy: true });
      seen.add(key);
    }
    state.pending = [];
    renderResults();
    return true;
  }

  function getQuoteTotals() {
    const totalPrice = state.matched.reduce((sum, item) => {
      return sum + parseMoney(item.product.price) * item.request.quantity;
    }, 0);
    const purchaseTotal = state.matched.reduce((sum, item) => {
      return sum + parseMoney(item.product.purchaseCost) * item.request.quantity;
    }, 0);
    const totalQuantity = getTotalQuantity();
    const shippingFee = getShippingFee(totalQuantity);
    const transactionFees = shippingFee == null ? 0 : (totalPrice + shippingFee) * 0.05;
    const amount = totalPrice + (shippingFee || 0) + transactionFees;
    const profit = totalPrice * PROFIT_EXCHANGE_RATE - purchaseTotal;

    return { totalQuantity, totalPrice, purchaseTotal, shippingFee, transactionFees, amount, profit };
  }

  function cloneWorkbook(workbook) {
    const data = XLSX.write(workbook, { bookType: "xlsx", type: "array", cellStyles: true, compression: true });
    return XLSX.read(data, { type: "array", cellFormula: true, cellStyles: true, cellNF: true, cellDates: true });
  }

  function clearQuoteInputs(sheet) {
    for (let row = 2; row <= 11; row += 1) {
      for (let col = 13; col <= 17; col += 1) {
        delete sheet[XLSX.utils.encode_cell({ r: row, c: col })];
      }
    }
  }

  function writeCell(sheet, address, value, type) {
    const existing = sheet[address] || {};
    sheet[address] = { ...existing, t: type || (typeof value === "number" ? "n" : "s"), v: value, w: undefined, f: undefined };
    if (typeof value === "number" && !sheet[address].z) sheet[address].z = "0";
  }

  function setCellStyle(sheet, address, options) {
    const cell = sheet[address];
    if (!cell) return;
    const currentStyle = cell.s || {};
    const currentFont = currentStyle.font || {};
    const currentAlignment = currentStyle.alignment || {};
    cell.s = {
      ...currentStyle,
      font: {
        ...currentFont,
        name: "Arial",
        sz: 12,
        bold: options.bold == null ? true : options.bold,
      },
      alignment: {
        ...currentAlignment,
        horizontal: options.horizontal || "center",
        vertical: options.vertical || "center",
      },
    };
  }

  function applyQuoteSheetStyle(sheet) {
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let col = range.s.c; col <= range.e.c; col += 1) {
        setCellStyle(sheet, XLSX.utils.encode_cell({ r: row, c: col }), {
          bold: true,
          horizontal: "center",
          vertical: "center",
        });
      }
    }
  }

  function findLabelCells(sheet, labels) {
    const wanted = new Set(labels.map((label) => compact(label)));
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    const found = [];

    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let col = range.s.c; col <= range.e.c; col += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: col });
        const text = compact(displayValue(sheet[address]));
        if (wanted.has(text)) found.push({ address, row, col, label: text });
      }
    }

    return found;
  }

  function findAmountCellOnRow(sheet, row, startCol) {
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    for (let col = range.e.c; col > startCol; col -= 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = sheet[address];
      if (!cell) continue;
      const value = displayValue(cell);
      if (cell.f || typeof cell.v === "number" || /\$|^\s*-?\d+(?:\.\d+)?\s*$/.test(value)) return address;
    }
    return XLSX.utils.encode_cell({ r: row, c: Math.min(startCol + 3, range.e.c) });
  }

  function setNumberCell(sheet, row, col, value, format) {
    const address = XLSX.utils.encode_cell({ r: row, c: col });
    writeCell(sheet, address, value, "n");
    sheet[address].z = format || "0";
    return address;
  }

  function setTextCell(sheet, row, col, value) {
    const address = XLSX.utils.encode_cell({ r: row, c: col });
    writeCell(sheet, address, value == null ? "" : String(value), "s");
    return address;
  }

  function addMerge(sheet, merge) {
    if (!sheet["!merges"]) sheet["!merges"] = [];
    sheet["!merges"] = sheet["!merges"].filter((item) => {
      return !(item.s.r === merge.s.r && item.e.r === merge.e.r && item.s.c === merge.s.c && item.e.c === merge.e.c);
    });
    sheet["!merges"].push(merge);
  }

  function removeRowMerges(sheet, row, startCol, endCol) {
    if (!sheet["!merges"]) return;
    sheet["!merges"] = sheet["!merges"].filter((merge) => {
      const sameRow = merge.s.r <= row && merge.e.r >= row;
      const overlaps = merge.s.c <= endCol && merge.e.c >= startCol;
      return !(sameRow && overlaps);
    });
  }

  function isNumberedQuoteSheet() {
    return /^\d+$/.test(String(state.quoteSheetName || "").trim());
  }

  function writeAmountForLabel(sheet, label, value) {
    const labelCell = findLabelCells(sheet, [label])[0];
    if (!labelCell) return false;

    const amountAddress = findAmountCellOnRow(sheet, labelCell.row, labelCell.col);
    const existing = sheet[amountAddress] || { t: "n" };
    sheet[amountAddress] = {
      ...existing,
      t: "n",
      v: value,
      w: undefined,
      f: undefined,
      z: existing.z || "$#,##0.00",
    };
    setCellStyle(sheet, labelCell.address, { bold: true });
    setCellStyle(sheet, amountAddress, { bold: true });
    return true;
  }

  function boldProductPrices(sheet) {
    for (let row = 2; row <= 11; row += 1) {
      for (let col = 0; col <= 12; col += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = sheet[address];
        if (!cell) continue;
        const value = displayValue(cell);
        if (cell.f || /\$/.test(value) || (typeof cell.v === "number" && col >= 5)) {
          setCellStyle(sheet, address, { bold: true });
        }
      }
    }
  }

  function applyFixedSummaryLayout(sheet, totals) {
    if (!isNumberedQuoteSheet()) return false;

    const lineCount = Number(state.quoteSheetName) || state.matched.length;
    const totalRow = QUOTE_PRODUCT_START_ROW + lineCount;
    const shippingRow = totalRow + 1;
    const feesRow = totalRow + 2;
    const amountRow = totalRow + 3;
    const rows = [
      { row: totalRow, label: "Total price", value: totals.totalPrice },
      { row: shippingRow, label: "Shipping fee", value: totals.shippingFee },
      { row: feesRow, label: "Transaction fees（5%）", value: totals.transactionFees },
      { row: amountRow, label: "Amount", value: totals.amount },
    ];

    addMerge(sheet, { s: { r: 0, c: 0 }, e: { r: 0, c: QUOTE_COLUMNS.profit } });
    rows.forEach((item) => {
      removeRowMerges(sheet, item.row, QUOTE_COLUMNS.abbreviation, QUOTE_COLUMNS.discount);
      addMerge(sheet, {
        s: { r: item.row, c: QUOTE_COLUMNS.abbreviation },
        e: { r: item.row, c: QUOTE_COLUMNS.discount },
      });

      const labelAddress = setTextCell(sheet, item.row, QUOTE_COLUMNS.abbreviation, item.label);
      const amountAddress = setNumberCell(sheet, item.row, QUOTE_COLUMNS.amount, item.value, "$#,##0.00");
      setCellStyle(sheet, labelAddress, { bold: true });
      setCellStyle(sheet, amountAddress, { bold: true });
    });

    setNumberCell(sheet, totalRow, QUOTE_COLUMNS.purchaseTotal, totals.purchaseTotal, "¥#,##0.00");
    setNumberCell(sheet, totalRow, QUOTE_COLUMNS.profit, totals.profit, "0.00");
    setCellStyle(sheet, XLSX.utils.encode_cell({ r: totalRow, c: QUOTE_COLUMNS.purchaseTotal }), { bold: true });
    setCellStyle(sheet, XLSX.utils.encode_cell({ r: totalRow, c: QUOTE_COLUMNS.profit }), { bold: true });
    return true;
  }

  function findOutputProductTable(sheet) {
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    const maxRow = Math.min(range.e.r, 8);
    const labels = {
      abbreviation: ["abbeviatio", "abbreviation", "abbreviatio", "abbr", "abb", "产品缩写", "缩写"],
      productName: ["productname", "product", "产品名称", "产品全称", "品名"],
      specification: ["specificatio", "specification", "spec", "规格"],
      quantity: ["q", "t", "qt", "qty", "quantity", "数量"],
      price: ["unitprice", "price", "单价", "价格"],
      amount: ["tatol", "total", "amount", "totalprice", "金额", "总价"],
    };
    let headerRow = -1;
    const columns = {};

    for (let row = range.s.r; row <= maxRow; row += 1) {
      for (let col = range.s.c; col <= range.e.c; col += 1) {
        const text = compact(displayValue(sheet[XLSX.utils.encode_cell({ r: row, c: col })]));
        if (!text) continue;
        for (const [key, aliases] of Object.entries(labels)) {
          if (columns[key] != null) continue;
          if (aliases.some((alias) => text === compact(alias) || text.includes(compact(alias)))) {
            columns[key] = col;
            headerRow = headerRow < 0 ? row : Math.min(headerRow, row);
          }
        }
      }
    }

    if (headerRow < 0) headerRow = 1;
    if (columns.abbreviation == null) columns.abbreviation = 1;
    if (columns.productName == null) columns.productName = columns.abbreviation + 1;
    if (columns.specification == null) columns.specification = columns.productName + 1;
    if (columns.quantity == null) columns.quantity = columns.specification + 1;
    if (columns.price == null) columns.price = columns.quantity + 1;
    if (columns.amount == null) columns.amount = columns.price + 1;

    return { headerRow, startRow: headerRow + 1, columns };
  }

  function clearOutputProductRows(sheet, table) {
    const visibleLineCount = Number(state.quoteSheetName) || state.matched.length;
    for (let row = table.startRow; row < table.startRow + visibleLineCount; row += 1) {
      Object.values(QUOTE_COLUMNS).forEach((col) => {
        const address = XLSX.utils.encode_cell({ r: row, c: col });
        const existing = sheet[address] || {};
        sheet[address] = { ...existing, t: "s", v: "", w: undefined, f: undefined };
      });
    }
  }

  function overwriteOutputProductRows(workbook) {
    if (!state.quoteSheetName || isPriceSheetName(state.quoteSheetName, state.priceSheetInfo.sheetName)) return;
    const sheet = workbook.Sheets[state.quoteSheetName];
    if (!sheet) return;

    const table = isNumberedQuoteSheet()
      ? { headerRow: QUOTE_HEADER_ROW, startRow: QUOTE_PRODUCT_START_ROW, columns: QUOTE_COLUMNS }
      : findOutputProductTable(sheet);
    clearOutputProductRows(sheet, table);
    state.matched.slice(0, MAX_QUOTE_LINES).forEach((item, index) => {
      const row = table.startRow + index;
      const price = parseMoney(item.product.price);
      const amount = price * item.request.quantity;
      const purchaseUnit = parseMoney(item.product.purchaseCost);
      const purchaseTotal = purchaseUnit * item.request.quantity;
      const profit = amount * PROFIT_EXCHANGE_RATE - purchaseTotal;

      setNumberCell(sheet, row, table.columns.serial, index + 1, "0");
      setTextCell(sheet, row, table.columns.abbreviation, item.product.abbreviation || item.request.input);
      setTextCell(sheet, row, table.columns.productName, item.product.productName);
      setTextCell(sheet, row, table.columns.specification, item.product.specification);
      setNumberCell(sheet, row, table.columns.quantity, item.request.quantity, "0");
      setNumberCell(sheet, row, table.columns.price, price, "$0.00");
      setTextCell(sheet, row, table.columns.discount, "0%");
      setNumberCell(sheet, row, table.columns.amount, amount, "$0.00");
      setNumberCell(sheet, row, table.columns.purchaseUnit, purchaseUnit, "0");
      setNumberCell(sheet, row, table.columns.purchaseTotal, purchaseTotal, "0");
      setNumberCell(sheet, row, table.columns.profit, profit, "0.00");
      setCellStyle(sheet, XLSX.utils.encode_cell({ r: row, c: table.columns.price }), { bold: true });
      setCellStyle(sheet, XLSX.utils.encode_cell({ r: row, c: table.columns.amount }), { bold: true });
      setCellStyle(sheet, XLSX.utils.encode_cell({ r: row, c: table.columns.abbreviation }), { bold: true });
    });
  }

  function applyQuoteOutputFormatting(workbook, totals) {
    const sheet = workbook.Sheets[state.quoteSheetName];
    if (!sheet) return;

    findLabelCells(sheet, QUOTE_LABELS_TO_BOLD).forEach((cell) => setCellStyle(sheet, cell.address, { bold: true }));
    if (!applyFixedSummaryLayout(sheet, totals)) {
      writeAmountForLabel(sheet, "Total price", totals.totalPrice);
      writeAmountForLabel(sheet, "Shipping fee", totals.shippingFee);
      writeAmountForLabel(sheet, "Transaction fees (5%)", totals.transactionFees);
      writeAmountForLabel(sheet, "Amount", totals.amount);
    }
    boldProductPrices(sheet);
    applyQuoteSheetStyle(sheet);
  }

  function fillQuoteInputs(workbook) {
    const inputSheetName = findQuoteInputSheetName(workbook, state.priceSheetInfo.sheetName);
    if (!inputSheetName || isPriceSheetName(inputSheetName, state.priceSheetInfo.sheetName)) return false;
    const sheet = workbook.Sheets[inputSheetName];
    if (!sheet) return false;

    clearQuoteInputs(sheet);
    state.matched.slice(0, MAX_QUOTE_LINES).forEach((item, index) => {
      const row = 3 + index;
      writeCell(sheet, `N${row}`, item.product.abbreviation || item.product.productName, "s");
      writeCell(sheet, `O${row}`, item.product.productName, "s");
      writeCell(sheet, `P${row}`, item.product.specification, "s");
      writeCell(sheet, `Q${row}`, item.request.quantity, "n");
      writeCell(sheet, `R${row}`, 0, "n");
    });

    if (workbook.Workbook) {
      workbook.Workbook.CalcPr = { calcMode: "auto", fullCalcOnLoad: "1", forceFullCalc: "1" };
    }
    return true;
  }

  function makeOutputWorkbook(totals) {
    const workbook = cloneWorkbook(state.workbook);
    fillQuoteInputs(workbook);
    overwriteOutputProductRows(workbook);
    applyQuoteOutputFormatting(workbook, totals);
    return workbook;
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function money(value, symbol) {
    const number = Number(value) || 0;
    return `${symbol || "$"}${number.toLocaleString("en-US", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })}`;
  }

  function plainNumber(value, digits) {
    const number = Number(value) || 0;
    return number.toLocaleString("en-US", {
      minimumFractionDigits: digits || 0,
      maximumFractionDigits: digits || 0,
    });
  }

  function makeFormattedQuoteHtml(totals) {
    const rows = state.matched.map((item, index) => {
      const price = parseMoney(item.product.price);
      const purchaseUnit = parseMoney(item.product.purchaseCost);
      const quoteTotal = price * item.request.quantity;
      const purchaseTotal = purchaseUnit * item.request.quantity;
      const profit = quoteTotal * PROFIT_EXCHANGE_RATE - purchaseTotal;

      return `
        <tr>
          <td>${index + 1}</td>
          <td class="abbr">${escapeHtml(item.product.abbreviation || item.request.input)}</td>
          <td>${escapeHtml(item.product.productName)}</td>
          <td>${escapeHtml(item.product.specification)}</td>
          <td>${item.request.quantity}</td>
          <td>${money(price)}</td>
          <td>0%</td>
          <td>${money(quoteTotal)}</td>
          <td>${plainNumber(purchaseUnit)}</td>
          <td>${plainNumber(purchaseTotal)}</td>
          <td>${plainNumber(profit, 2)}</td>
        </tr>`;
    }).join("");

    return `<!doctype html>
<html>
<head>
  <meta charset="UTF-8" />
  <style>
    table.quote {
      border-collapse: collapse;
      table-layout: fixed;
      font-family: Arial, sans-serif;
      font-size: 12pt;
      font-weight: bold;
      text-align: center;
      vertical-align: middle;
      mso-displayed-decimal-separator: ".";
      mso-displayed-thousand-separator: ",";
    }
    .quote td, .quote th {
      border: 1px solid #000;
      height: 38px;
      text-align: center;
      vertical-align: middle;
      font-family: Arial, sans-serif;
      font-size: 12pt;
      font-weight: bold;
      white-space: nowrap;
    }
    .quote .title {
      height: 38px;
      background: #10aee3;
      color: #000;
      font-size: 16pt;
      letter-spacing: 1px;
    }
    .quote th {
      background: #c00000;
      color: #fff;
      height: 36px;
    }
    .quote .abbr {
      font-size: 15pt;
    }
    .quote .summary-label {
      height: 38px;
      font-size: 14pt;
    }
    .quote .summary-value {
      font-size: 14pt;
    }
    col.c1 { width: 74px; }
    col.c2 { width: 92px; }
    col.c3 { width: 220px; }
    col.c4 { width: 125px; }
    col.c5 { width: 42px; }
    col.c6 { width: 120px; }
    col.c7 { width: 110px; }
    col.c8 { width: 125px; }
    col.c9 { width: 125px; }
    col.c10 { width: 125px; }
    col.c11 { width: 125px; }
  </style>
</head>
<body>
  <table class="quote">
    <colgroup>
      <col class="c1" /><col class="c2" /><col class="c3" /><col class="c4" /><col class="c5" />
      <col class="c6" /><col class="c7" /><col class="c8" /><col class="c9" /><col class="c10" /><col class="c11" />
    </colgroup>
    <tr><td class="title" colspan="11">订单号：20261018</td></tr>
    <tr>
      <th>列1</th>
      <th>Abb eviatio</th>
      <th>Product Name</th>
      <th>Specificatio</th>
      <th>Q<br />T</th>
      <th>Price<br />($/box)</th>
      <th>Discount</th>
      <th>TATOL</th>
      <th>采购单价</th>
      <th>采购总价</th>
      <th>利润</th>
    </tr>
    ${rows}
    <tr>
      <td></td>
      <td class="summary-label" colspan="6">Total price</td>
      <td class="summary-value">${money(totals.totalPrice)}</td>
      <td></td>
      <td class="summary-value">${money(totals.purchaseTotal, "￥")}</td>
      <td>${plainNumber(totals.profit, 2)}</td>
    </tr>
    <tr>
      <td></td>
      <td class="summary-label" colspan="6">Shipping fee</td>
      <td class="summary-value">${money(totals.shippingFee)}</td>
      <td></td><td></td><td></td>
    </tr>
    <tr>
      <td></td>
      <td class="summary-label" colspan="6">Transaction fees（5%）</td>
      <td class="summary-value">${money(totals.transactionFees)}</td>
      <td></td><td></td><td></td>
    </tr>
    <tr>
      <td></td>
      <td class="summary-label" colspan="6">Amount</td>
      <td class="summary-value">${money(totals.amount)}</td>
      <td></td><td></td><td></td>
    </tr>
  </table>
</body>
</html>`;
  }

  function escapeAttribute(value) {
    return escapeHtml(value).replace(/'/g, "&#39;");
  }

  function formulaCell(className, formula, value) {
    return `<td class="${className}">${escapeHtml(formula)}</td>`;
  }

  function uniqueDataProducts() {
    const seen = new Set();
    return state.productRows.filter((product) => {
      const abbreviation = String(product.abbreviation || "").trim();
      if (!abbreviation) return false;
      const key = abbreviation.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function dataProductRows(products) {
    return products.map((product) => {
      return `
        <tr>
          <td>${escapeHtml(product.abbreviation)}</td>
          <td>${escapeHtml(product.productName)}</td>
          <td>${escapeHtml(product.specification)}</td>
          <td class="number">${parseMoney(product.price)}</td>
          <td class="number">${parseMoney(product.purchaseCost)}</td>
        </tr>`;
    }).join("");
  }

  function makeFormattedQuoteHtml(totals) {
    const dataProducts = uniqueDataProducts();
    const dataLastRow = Math.max(dataProducts.length + 1, 2);
    const dataRange = `ProductData!$A$2:$E$${dataLastRow}`;
    const firstProductRow = 3;
    const lastProductRow = Math.max(state.matched.length + 2, firstProductRow);
    const totalRow = state.matched.length + 3;
    const shippingRow = totalRow + 1;
    const transactionRow = totalRow + 2;
    const totalPriceFormula = `=SUM(H${firstProductRow}:H${lastProductRow})`;
    const purchaseTotalFormula = `=SUM(J${firstProductRow}:J${lastProductRow})`;
    const totalProfitFormula = `=SUM(K${firstProductRow}:K${lastProductRow})`;
    const totalQuantityFormula = `SUM(E${firstProductRow}:E${lastProductRow})`;
    const shippingFormula = `=IF(${totalQuantityFormula}<=0,0,45+(CEILING(${totalQuantityFormula}/10,1)-1)*25)`;
    const transactionFormula = `=(H${totalRow}+H${shippingRow})*5%`;
    const amountFormula = `=H${totalRow}+H${shippingRow}+H${transactionRow}`;
    const productDataRows = dataProductRows(dataProducts);
    const rows = state.matched.map((item, index) => {
      const price = parseMoney(item.product.price);
      const purchaseUnit = parseMoney(item.product.purchaseCost);
      const quoteTotal = price * item.request.quantity;
      const purchaseTotal = purchaseUnit * item.request.quantity;
      const profit = quoteTotal * PROFIT_EXCHANGE_RATE - purchaseTotal;
      const rowNumber = index + 3;
      const lookup = (columnIndex, fallback) => {
        return `=IFERROR(VLOOKUP(B${rowNumber},${dataRange},${columnIndex},FALSE),${fallback})`;
      };

      return `
        <tr>
          <td>${index + 1}</td>
          <td class="abbr">${escapeHtml(item.product.abbreviation || item.request.input)}</td>
          ${formulaCell("", lookup(2, "\"\""), item.product.productName)}
          ${formulaCell("", lookup(3, "\"\""), item.product.specification)}
          <td>${item.request.quantity}</td>
          ${formulaCell("usd", lookup(4, "0"), price)}
          <td class="percent">0</td>
          ${formulaCell("usd", `=IFERROR(E${rowNumber}*F${rowNumber}*(1-G${rowNumber}),0)`, quoteTotal)}
          ${formulaCell("number", lookup(5, "0"), purchaseUnit)}
          ${formulaCell("number", `=IFERROR(E${rowNumber}*I${rowNumber},0)`, purchaseTotal)}
          ${formulaCell("number-2", `=IFERROR(H${rowNumber}*${PROFIT_EXCHANGE_RATE}-J${rowNumber},0)`, profit)}
        </tr>`;
    }).join("");

    return `<!doctype html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="UTF-8" />
  <!--[if gte mso 9]><xml>
    <x:ExcelWorkbook>
      <x:ExcelWorksheets>
        <x:ExcelWorksheet>
          <x:Name>Quote</x:Name>
          <x:WorksheetSource HRef="#Quote" />
          <x:WorksheetOptions><x:DisplayGridlines /></x:WorksheetOptions>
        </x:ExcelWorksheet>
        <x:ExcelWorksheet>
          <x:Name>ProductData</x:Name>
          <x:WorksheetSource HRef="#ProductData" />
          <x:WorksheetOptions><x:Visible>SheetHidden</x:Visible></x:WorksheetOptions>
        </x:ExcelWorksheet>
      </x:ExcelWorksheets>
      <x:Calculation>Automatic</x:Calculation>
      <x:ForceFullCalculation />
    </x:ExcelWorkbook>
  </xml><![endif]-->
  <style>
    table.quote {
      border-collapse: collapse;
      table-layout: fixed;
      font-family: Arial, sans-serif;
      font-size: 12pt;
      font-weight: bold;
      text-align: center;
      vertical-align: middle;
      mso-displayed-decimal-separator: ".";
      mso-displayed-thousand-separator: ",";
    }
    .quote td, .quote th {
      border: 1px solid #000;
      height: 38px;
      text-align: center;
      vertical-align: middle;
      font-family: Arial, sans-serif;
      font-size: 12pt;
      font-weight: bold;
      white-space: nowrap;
    }
    .quote .title {
      height: 38px;
      background: #10aee3;
      color: #000;
      font-size: 16pt;
      letter-spacing: 1px;
    }
    .quote th {
      background: #c00000;
      color: #fff;
      height: 36px;
    }
    .quote .abbr {
      font-size: 15pt;
    }
    .quote .summary-label {
      height: 38px;
      font-size: 14pt;
    }
    .quote .summary-value {
      font-size: 14pt;
    }
    .quote .usd {
      mso-number-format: "\\0022$\\0022#,##0.00";
    }
    .quote .number {
      mso-number-format: "0";
    }
    .quote .number-2 {
      mso-number-format: "0.00";
    }
    .quote .percent {
      mso-number-format: "0%";
    }
    .data td, .data th {
      font-family: Arial, sans-serif;
      font-size: 10pt;
      mso-number-format: "\\@";
    }
    .data .number {
      mso-number-format: "0.00";
    }
    col.c1 { width: 74px; }
    col.c2 { width: 92px; }
    col.c3 { width: 220px; }
    col.c4 { width: 125px; }
    col.c5 { width: 42px; }
    col.c6 { width: 120px; }
    col.c7 { width: 110px; }
    col.c8 { width: 125px; }
    col.c9 { width: 125px; }
    col.c10 { width: 125px; }
    col.c11 { width: 125px; }
  </style>
</head>
<body>
  <div id="Quote" style="mso-element:sheet">
  <table class="quote">
    <colgroup>
      <col class="c1" /><col class="c2" /><col class="c3" /><col class="c4" /><col class="c5" />
      <col class="c6" /><col class="c7" /><col class="c8" /><col class="c9" /><col class="c10" /><col class="c11" />
    </colgroup>
    <tr><td class="title" colspan="11">订单号：20261018</td></tr>
    <tr>
      <th>列1</th>
      <th>Abb eviatio</th>
      <th>Product Name</th>
      <th>Specificatio</th>
      <th>Q<br />T</th>
      <th>Price<br />($/box)</th>
      <th>Discount</th>
      <th>TATOL</th>
      <th>采购单价</th>
      <th>采购总价</th>
      <th>利润</th>
    </tr>
    ${rows}
    <tr>
      <td></td>
      <td class="summary-label" colspan="6">Total price</td>
      ${formulaCell("summary-value usd", totalPriceFormula, totals.totalPrice)}
      <td></td>
      ${formulaCell("summary-value number-2", purchaseTotalFormula, totals.purchaseTotal)}
      ${formulaCell("number-2", totalProfitFormula, totals.profit)}
    </tr>
    <tr>
      <td></td>
      <td class="summary-label" colspan="6">Shipping fee</td>
      ${formulaCell("summary-value usd", shippingFormula, totals.shippingFee)}
      <td></td><td></td><td></td>
    </tr>
    <tr>
      <td></td>
      <td class="summary-label" colspan="6">Transaction fees (5%)</td>
      ${formulaCell("summary-value usd", transactionFormula, totals.transactionFees)}
      <td></td><td></td><td></td>
    </tr>
    <tr>
      <td></td>
      <td class="summary-label" colspan="6">Amount</td>
      ${formulaCell("summary-value usd", amountFormula, totals.amount)}
      <td></td><td></td><td></td>
    </tr>
  </table>
  </div>
  <div id="ProductData" style="mso-element:sheet">
    <table class="data">
      <tr>
        <th>Abbreviation</th>
        <th>Product Name</th>
        <th>Specification</th>
        <th>Price</th>
        <th>Purchase Unit</th>
      </tr>
      ${productDataRows}
    </table>
  </div>
</body>
</html>`;
  }

  function lookupDataCells(product) {
    if (!product) {
      return '<td class="lookup-data"></td><td class="lookup-data"></td><td class="lookup-data"></td><td class="lookup-data"></td><td class="lookup-data"></td>';
    }
    return `
      <td class="lookup-data">${escapeHtml(product.abbreviation)}</td>
      <td class="lookup-data">${escapeHtml(product.productName)}</td>
      <td class="lookup-data">${escapeHtml(product.specification)}</td>
      <td class="lookup-data number">${parseMoney(product.price)}</td>
      <td class="lookup-data number">${parseMoney(product.purchaseCost)}</td>`;
  }

  function makeFormattedQuoteHtml(totals) {
    const dataProducts = uniqueDataProducts();
    const visibleRowCount = state.matched.length + 6;
    const tableRowCount = Math.max(visibleRowCount, dataProducts.length + 1);
    const dataLastRow = Math.max(dataProducts.length + 1, 2);
    const spacerCells = '<td class="spacer"></td>'.repeat(15);
    const dataRange = `$AA$2:$AE$${dataLastRow}`;
    const firstProductRow = 3;
    const lastProductRow = Math.max(state.matched.length + 2, firstProductRow);
    const totalRow = state.matched.length + 3;
    const shippingRow = totalRow + 1;
    const transactionRow = totalRow + 2;
    const totalPriceFormula = `=SUM(H${firstProductRow}:H${lastProductRow})`;
    const purchaseTotalFormula = `=SUM(J${firstProductRow}:J${lastProductRow})`;
    const totalProfitFormula = `=SUM(K${firstProductRow}:K${lastProductRow})`;
    const totalQuantityFormula = `SUM(E${firstProductRow}:E${lastProductRow})`;
    const shippingFormula = `=IF(${totalQuantityFormula}<=0,0,45+(CEILING(${totalQuantityFormula}/10,1)-1)*25)`;
    const transactionFormula = `=(H${totalRow}+H${shippingRow})*5%`;
    const amountFormula = `=H${totalRow}+H${shippingRow}+H${transactionRow}`;
    const dataForRow = (rowNumber) => `${spacerCells}${lookupDataCells(dataProducts[rowNumber - 2])}`;
    const quoteRows = state.matched.map((item, index) => {
      const price = parseMoney(item.product.price);
      const purchaseUnit = parseMoney(item.product.purchaseCost);
      const quoteTotal = price * item.request.quantity;
      const purchaseTotal = purchaseUnit * item.request.quantity;
      const profit = quoteTotal * PROFIT_EXCHANGE_RATE - purchaseTotal;
      const rowNumber = index + 3;
      const lookup = (columnIndex, fallback) => {
        return `=IFERROR(VLOOKUP(B${rowNumber},${dataRange},${columnIndex},FALSE),${fallback})`;
      };

      return `
        <tr>
          <td>${index + 1}</td>
          <td class="abbr">${escapeHtml(item.product.abbreviation || item.request.input)}</td>
          ${formulaCell("", lookup(2, "\"\""), item.product.productName)}
          ${formulaCell("", lookup(3, "\"\""), item.product.specification)}
          <td>${item.request.quantity}</td>
          ${formulaCell("usd", lookup(4, "0"), price)}
          <td class="percent">0</td>
          ${formulaCell("usd", `=IFERROR(E${rowNumber}*F${rowNumber}*(1-G${rowNumber}),0)`, quoteTotal)}
          ${formulaCell("number", lookup(5, "0"), purchaseUnit)}
          ${formulaCell("number", `=IFERROR(E${rowNumber}*I${rowNumber},0)`, purchaseTotal)}
          ${formulaCell("number-2", `=IFERROR(H${rowNumber}*${PROFIT_EXCHANGE_RATE}-J${rowNumber},0)`, profit)}
          ${dataForRow(rowNumber)}
        </tr>`;
    }).join("");
    const extraDataRows = Array.from({ length: Math.max(tableRowCount - visibleRowCount, 0) }, (_, index) => {
      const rowNumber = visibleRowCount + index + 1;
      return `<tr class="data-only"><td colspan="11"></td>${dataForRow(rowNumber)}</tr>`;
    }).join("");

    return `<!doctype html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="UTF-8" />
  <style>
    table.quote {
      border-collapse: collapse;
      table-layout: fixed;
      font-family: Arial, sans-serif;
      font-size: 12pt;
      font-weight: bold;
      text-align: center;
      vertical-align: middle;
      mso-displayed-decimal-separator: ".";
      mso-displayed-thousand-separator: ",";
    }
    .quote td, .quote th {
      border: 1px solid #000;
      height: 38px;
      text-align: center;
      vertical-align: middle;
      font-family: Arial, sans-serif;
      font-size: 12pt;
      font-weight: bold;
      white-space: nowrap;
    }
    .quote .title {
      height: 38px;
      background: #10aee3;
      color: #000;
      font-size: 16pt;
      letter-spacing: 1px;
    }
    .quote th {
      background: #c00000;
      color: #fff;
      height: 36px;
    }
    .quote .abbr {
      font-size: 15pt;
    }
    .quote .summary-label {
      height: 38px;
      font-size: 14pt;
    }
    .quote .summary-value {
      font-size: 14pt;
    }
    .quote .usd {
      mso-number-format: "\\0022$\\0022#,##0.00";
    }
    .quote .number {
      mso-number-format: "0";
    }
    .quote .number-2 {
      mso-number-format: "0.00";
    }
    .quote .percent {
      mso-number-format: "0%";
    }
    .quote .spacer {
      width: 34px;
      border: none;
      color: #fff;
      font-size: 1pt;
      overflow: hidden;
    }
    .quote .lookup-data {
      width: 34px;
      border: none;
      color: #fff;
      font-size: 1pt;
      overflow: hidden;
    }
    .quote .data-only td {
      height: 0;
      border: none;
      font-size: 1pt;
      color: #fff;
      overflow: hidden;
    }
    col.c1 { width: 74px; }
    col.c2 { width: 92px; }
    col.c3 { width: 220px; }
    col.c4 { width: 125px; }
    col.c5 { width: 42px; }
    col.c6 { width: 120px; }
    col.c7 { width: 110px; }
    col.c8 { width: 125px; }
    col.c9 { width: 125px; }
    col.c10 { width: 125px; }
    col.c11 { width: 125px; }
    col.spacer-col { width: 34px; }
    col.lookup-col { width: 34px; }
  </style>
</head>
<body>
  <table class="quote">
    <colgroup>
      <col class="c1" /><col class="c2" /><col class="c3" /><col class="c4" /><col class="c5" />
      <col class="c6" /><col class="c7" /><col class="c8" /><col class="c9" /><col class="c10" /><col class="c11" />
      <col class="spacer-col" /><col class="spacer-col" /><col class="spacer-col" /><col class="spacer-col" /><col class="spacer-col" />
      <col class="spacer-col" /><col class="spacer-col" /><col class="spacer-col" /><col class="spacer-col" /><col class="spacer-col" />
      <col class="spacer-col" /><col class="spacer-col" /><col class="spacer-col" /><col class="spacer-col" /><col class="spacer-col" />
      <col class="lookup-col" /><col class="lookup-col" /><col class="lookup-col" /><col class="lookup-col" /><col class="lookup-col" />
    </colgroup>
    <tr>
      <td class="title" colspan="11">订单号：20261018</td>
      ${spacerCells}
      <td class="lookup-data">Abbreviation</td>
      <td class="lookup-data">Product Name</td>
      <td class="lookup-data">Specification</td>
      <td class="lookup-data">Price</td>
      <td class="lookup-data">Purchase Unit</td>
    </tr>
    <tr>
      <th>列1</th>
      <th>Abb eviatio</th>
      <th>Product Name</th>
      <th>Specificatio</th>
      <th>Q<br />T</th>
      <th>Price<br />($/box)</th>
      <th>Discount</th>
      <th>TATOL</th>
      <th>采购单价</th>
      <th>采购总价</th>
      <th>利润</th>
      ${dataForRow(2)}
    </tr>
    ${quoteRows}
    <tr>
      <td></td>
      <td class="summary-label" colspan="6">Total price</td>
      ${formulaCell("summary-value usd", totalPriceFormula, totals.totalPrice)}
      <td></td>
      ${formulaCell("summary-value number-2", purchaseTotalFormula, totals.purchaseTotal)}
      ${formulaCell("number-2", totalProfitFormula, totals.profit)}
      ${dataForRow(totalRow)}
    </tr>
    <tr>
      <td></td>
      <td class="summary-label" colspan="6">Shipping fee</td>
      ${formulaCell("summary-value usd", shippingFormula, totals.shippingFee)}
      <td></td><td></td><td></td>
      ${dataForRow(shippingRow)}
    </tr>
    <tr>
      <td></td>
      <td class="summary-label" colspan="6">Transaction fees (5%)</td>
      ${formulaCell("summary-value usd", transactionFormula, totals.transactionFees)}
      <td></td><td></td><td></td>
      ${dataForRow(transactionRow)}
    </tr>
    <tr>
      <td></td>
      <td class="summary-label" colspan="6">Amount</td>
      ${formulaCell("summary-value usd", amountFormula, totals.amount)}
      <td></td><td></td><td></td>
      ${dataForRow(transactionRow + 1)}
    </tr>
    ${extraDataRows}
  </table>
</body>
</html>`;
  }

  function downloadFormattedQuoteXls(totals, stamp) {
    const html = makeFormattedQuoteHtml(totals);
    const blob = new Blob(["\ufeff", html], { type: "application/vnd.ms-excel;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `报价单_${stamp}.xls`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function quoteCellStyle(options) {
    const border = options.border === false ? undefined : {
      top: { style: "thin", color: { rgb: "000000" } },
      bottom: { style: "thin", color: { rgb: "000000" } },
      left: { style: "thin", color: { rgb: "000000" } },
      right: { style: "thin", color: { rgb: "000000" } },
    };
    return {
      font: {
        name: "Arial",
        sz: options.size || 12,
        bold: options.bold == null ? true : options.bold,
        color: { rgb: options.color || "000000" },
      },
      alignment: {
        horizontal: "center",
        vertical: "center",
        wrapText: !!options.wrap,
      },
      fill: options.fill ? { patternType: "solid", fgColor: { rgb: options.fill } } : undefined,
      border,
    };
  }

  function setXlsxCell(sheet, row, col, value, options) {
    const address = XLSX.utils.encode_cell({ r: row, c: col });
    const cell = {
      t: options && options.type ? options.type : (typeof value === "number" ? "n" : "s"),
      v: value,
    };
    if (options && options.format) cell.z = options.format;
    if (options && options.style) cell.s = options.style;
    sheet[address] = cell;
    return cell;
  }

  function setXlsxFormula(sheet, row, col, formula, cachedValue, options) {
    const cell = setXlsxCell(sheet, row, col, cachedValue, options);
    cell.f = formula;
    return cell;
  }

  function makeShippingFormula(totalQuantityFormula) {
    return `IF(${totalQuantityFormula}<=0,0,45+(CEILING(${totalQuantityFormula}/10,1)-1)*25)`;
  }

  function downloadFormulaQuoteXlsx(totals, stamp) {
    const workbook = XLSX.utils.book_new();
    const sheet = {};
    const dataProducts = uniqueDataProducts();
    const productCount = state.matched.length;
    const dataLastRow = Math.max(dataProducts.length + 1, 2);
    const dataRange = `$AA$2:$AE$${dataLastRow}`;
    const firstProductRow = 3;
    const lastProductRow = Math.max(productCount + 2, firstProductRow);
    const totalRow = productCount + 3;
    const shippingRow = totalRow + 1;
    const feesRow = totalRow + 2;
    const amountRow = totalRow + 3;
    const maxRows = Math.max(amountRow, dataProducts.length + 1);
    const maxCols = 30;
    const styles = {
      title: quoteCellStyle({ fill: "10AEE3", size: 16 }),
      header: quoteCellStyle({ fill: "C00000", color: "FFFFFF" }),
      base: quoteCellStyle({}),
      abbr: quoteCellStyle({ size: 15 }),
      summary: quoteCellStyle({ size: 14 }),
      lookup: quoteCellStyle({ size: 1, color: "FFFFFF", border: false }),
      spacer: quoteCellStyle({ size: 1, color: "FFFFFF", border: false }),
    };

    for (let row = 0; row < maxRows; row += 1) {
      for (let col = 0; col <= 10; col += 1) {
        setXlsxCell(sheet, row, col, "", { style: styles.base });
      }
      for (let col = 11; col <= maxCols; col += 1) {
        setXlsxCell(sheet, row, col, "", { style: styles.spacer });
      }
    }

    setXlsxCell(sheet, 0, 0, "订单号：20261018", { style: styles.title });
    sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 10 } }];

    ["列1", "Abb eviatio", "Product Name", "Specificatio", "Q\nT", "Price\n($/box)", "Discount", "TATOL", "采购单价", "采购总价", "利润"].forEach((label, col) => {
      setXlsxCell(sheet, 1, col, label, { style: styles.header });
    });

    state.matched.forEach((item, index) => {
      const row = index + 2;
      const excelRow = row + 1;
      const abbr = item.product.abbreviation || item.request.input;
      const price = parseMoney(item.product.price);
      const purchaseUnit = parseMoney(item.product.purchaseCost);
      const quoteTotal = price * item.request.quantity;
      const purchaseTotal = purchaseUnit * item.request.quantity;
      const profit = quoteTotal * PROFIT_EXCHANGE_RATE - purchaseTotal;
      const lookup = (columnIndex, fallback) => `IFERROR(VLOOKUP(B${excelRow},${dataRange},${columnIndex},FALSE),${fallback})`;

      setXlsxCell(sheet, row, 0, index + 1, { style: styles.base, format: "0" });
      setXlsxCell(sheet, row, 1, abbr, { style: styles.abbr });
      setXlsxFormula(sheet, row, 2, lookup(2, '""'), item.product.productName, { style: styles.base });
      setXlsxFormula(sheet, row, 3, lookup(3, '""'), item.product.specification, { style: styles.base });
      setXlsxCell(sheet, row, 4, item.request.quantity, { style: styles.base, format: "0" });
      setXlsxFormula(sheet, row, 5, lookup(4, "0"), price, { style: styles.base, format: "$#,##0.00" });
      setXlsxCell(sheet, row, 6, 0, { style: styles.base, format: "0%" });
      setXlsxFormula(sheet, row, 7, `IFERROR(E${excelRow}*F${excelRow}*(1-G${excelRow}),0)`, quoteTotal, { style: styles.base, format: "$#,##0.00" });
      setXlsxFormula(sheet, row, 8, lookup(5, "0"), purchaseUnit, { style: styles.base, format: "0" });
      setXlsxFormula(sheet, row, 9, `IFERROR(E${excelRow}*I${excelRow},0)`, purchaseTotal, { style: styles.base, format: "0" });
      setXlsxFormula(sheet, row, 10, `IFERROR(H${excelRow}*${PROFIT_EXCHANGE_RATE}-J${excelRow},0)`, profit, { style: styles.base, format: "0.00" });
    });

    [
      { row: totalRow - 1, label: "Total price", amountFormula: `SUM(H${firstProductRow}:H${lastProductRow})`, amount: totals.totalPrice, purchaseFormula: `SUM(J${firstProductRow}:J${lastProductRow})`, purchase: totals.purchaseTotal, profitFormula: `SUM(K${firstProductRow}:K${lastProductRow})`, profit: totals.profit },
      { row: shippingRow - 1, label: "Shipping fee", amountFormula: makeShippingFormula(`SUM(E${firstProductRow}:E${lastProductRow})`), amount: totals.shippingFee },
      { row: feesRow - 1, label: "Transaction fees (5%)", amountFormula: `(H${totalRow}+H${shippingRow})*5%`, amount: totals.transactionFees },
      { row: amountRow - 1, label: "Amount", amountFormula: `H${totalRow}+H${shippingRow}+H${feesRow}`, amount: totals.amount },
    ].forEach((item) => {
      sheet["!merges"].push({ s: { r: item.row, c: 1 }, e: { r: item.row, c: 6 } });
      setXlsxCell(sheet, item.row, 1, item.label, { style: styles.summary });
      setXlsxFormula(sheet, item.row, 7, item.amountFormula, item.amount, { style: styles.summary, format: "$#,##0.00" });
      if (item.purchaseFormula) setXlsxFormula(sheet, item.row, 9, item.purchaseFormula, item.purchase, { style: styles.summary, format: "0" });
      if (item.profitFormula) setXlsxFormula(sheet, item.row, 10, item.profitFormula, item.profit, { style: styles.summary, format: "0.00" });
    });

    ["Abbreviation", "Product Name", "Specification", "Price", "Purchase Unit"].forEach((label, index) => {
      setXlsxCell(sheet, 0, 26 + index, label, { style: styles.lookup });
    });
    dataProducts.forEach((product, index) => {
      const row = index + 1;
      setXlsxCell(sheet, row, 26, product.abbreviation, { style: styles.lookup });
      setXlsxCell(sheet, row, 27, product.productName, { style: styles.lookup });
      setXlsxCell(sheet, row, 28, product.specification, { style: styles.lookup });
      setXlsxCell(sheet, row, 29, parseMoney(product.price), { style: styles.lookup, format: "0.00" });
      setXlsxCell(sheet, row, 30, parseMoney(product.purchaseCost), { style: styles.lookup, format: "0.00" });
    });

    sheet["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRows - 1, c: maxCols } });
    sheet["!cols"] = [
      { wch: 8 }, { wch: 12 }, { wch: 28 }, { wch: 16 }, { wch: 5 },
      { wch: 14 }, { wch: 13 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
      ...Array.from({ length: 15 }, () => ({ wch: 3 })),
      ...Array.from({ length: 5 }, () => ({ wch: 3 })),
    ];
    sheet["!rows"] = Array.from({ length: maxRows }, (_, row) => ({ hpt: row === 0 ? 28 : 26 }));
    workbook.Workbook = { CalcPr: { calcMode: "auto", fullCalcOnLoad: "1", forceFullCalc: "1" } };
    XLSX.utils.book_append_sheet(workbook, sheet, "报价单");
    XLSX.writeFile(workbook, `报价单_${stamp}.xlsx`, { bookType: "xlsx", cellStyles: true, compression: true });
  }

  function escapeXml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function xmlCell(options) {
    const attrs = [];
    if (options.index) attrs.push(`ss:Index="${options.index}"`);
    if (options.mergeAcross) attrs.push(`ss:MergeAcross="${options.mergeAcross}"`);
    if (options.style) attrs.push(`ss:StyleID="${options.style}"`);
    if (options.formula) attrs.push(`ss:Formula="${escapeXml(options.formula)}"`);
    const type = options.type || (typeof options.value === "number" ? "Number" : "String");
    return `<Cell ${attrs.join(" ")}><Data ss:Type="${type}">${escapeXml(options.value)}</Data></Cell>`;
  }

  function xmlRow(cells, height) {
    return `<Row${height ? ` ss:Height="${height}"` : ""}>${cells.join("")}</Row>`;
  }

  function makeFormulaQuoteXml(totals) {
    const dataProducts = uniqueDataProducts();
    const productCount = state.matched.length;
    const dataLastRow = Math.max(dataProducts.length + 1, 2);
    const firstProductRow = 3;
    const lastProductRow = Math.max(productCount + 2, firstProductRow);
    const totalRow = productCount + 3;
    const shippingRow = totalRow + 1;
    const feesRow = totalRow + 2;
    const amountRow = totalRow + 3;
    const totalRows = Math.max(amountRow, dataProducts.length + 1);
    const lookupRange = `R2C27:R${dataLastRow}C31`;
    const widths = [52, 78, 180, 104, 34, 94, 84, 96, 96, 96, 96];
    const columns = [
      ...widths.map((width) => `<Column ss:Width="${width}" />`),
      ...Array.from({ length: 15 }, () => '<Column ss:Width="22" />'),
      ...Array.from({ length: 5 }, () => '<Column ss:Width="22" />'),
    ].join("");
    const rows = [];

    rows.push(xmlRow([
      xmlCell({ style: "Title", mergeAcross: 10, value: "订单号：20261018" }),
      xmlCell({ index: 27, style: "Lookup", value: "Abbreviation" }),
      xmlCell({ style: "Lookup", value: "Product Name" }),
      xmlCell({ style: "Lookup", value: "Specification" }),
      xmlCell({ style: "Lookup", value: "Price" }),
      xmlCell({ style: "Lookup", value: "Purchase Unit" }),
    ], 30));

    const firstData = dataProducts[0];
    rows.push(xmlRow([
      ...["列1", "Abb eviatio", "Product Name", "Specificatio", "Q\nT", "Price\n($/box)", "Discount", "TATOL", "采购单价", "采购总价", "利润"].map((label) => xmlCell({ style: "Header", value: label })),
      xmlCell({ index: 27, style: "Lookup", value: firstData && firstData.abbreviation }),
      xmlCell({ style: "Lookup", value: firstData && firstData.productName }),
      xmlCell({ style: "Lookup", value: firstData && firstData.specification }),
      xmlCell({ style: "LookupNumber", value: firstData ? parseMoney(firstData.price) : 0 }),
      xmlCell({ style: "LookupNumber", value: firstData ? parseMoney(firstData.purchaseCost) : 0 }),
    ], 28));

    state.matched.forEach((item, index) => {
      const excelRow = index + 3;
      const dataProduct = dataProducts[excelRow - 2];
      const price = parseMoney(item.product.price);
      const purchaseUnit = parseMoney(item.product.purchaseCost);
      const quoteTotal = price * item.request.quantity;
      const purchaseTotal = purchaseUnit * item.request.quantity;
      const profit = quoteTotal * PROFIT_EXCHANGE_RATE - purchaseTotal;
      const lookup = (column, fallback) => `=IFERROR(VLOOKUP(RC2,${lookupRange},${column},FALSE),${fallback})`;
      rows.push(xmlRow([
        xmlCell({ style: "Base", value: index + 1 }),
        xmlCell({ style: "Abbr", value: item.product.abbreviation || item.request.input }),
        xmlCell({ style: "Base", formula: lookup(2, '""'), value: item.product.productName }),
        xmlCell({ style: "Base", formula: lookup(3, '""'), value: item.product.specification }),
        xmlCell({ style: "Base", value: item.request.quantity }),
        xmlCell({ style: "Money", formula: lookup(4, "0"), value: price }),
        xmlCell({ style: "Percent", value: 0 }),
        xmlCell({ style: "Money", formula: "=IFERROR(RC5*RC6*(1-RC7),0)", value: quoteTotal }),
        xmlCell({ style: "Integer", formula: lookup(5, "0"), value: purchaseUnit }),
        xmlCell({ style: "Integer", formula: "=IFERROR(RC5*RC9,0)", value: purchaseTotal }),
        xmlCell({ style: "Number2", formula: `=IFERROR(RC8*${PROFIT_EXCHANGE_RATE}-RC10,0)`, value: profit }),
        xmlCell({ index: 27, style: "Lookup", value: dataProduct && dataProduct.abbreviation }),
        xmlCell({ style: "Lookup", value: dataProduct && dataProduct.productName }),
        xmlCell({ style: "Lookup", value: dataProduct && dataProduct.specification }),
        xmlCell({ style: "LookupNumber", value: dataProduct ? parseMoney(dataProduct.price) : 0 }),
        xmlCell({ style: "LookupNumber", value: dataProduct ? parseMoney(dataProduct.purchaseCost) : 0 }),
      ], 28));
    });

    [
      { row: totalRow, label: "Total price", amountFormula: `=SUM(R${firstProductRow}C8:R${lastProductRow}C8)`, amount: totals.totalPrice, purchaseFormula: `=SUM(R${firstProductRow}C10:R${lastProductRow}C10)`, purchase: totals.purchaseTotal, profitFormula: `=SUM(R${firstProductRow}C11:R${lastProductRow}C11)`, profit: totals.profit },
      { row: shippingRow, label: "Shipping fee", amountFormula: `=IF(SUM(R${firstProductRow}C5:R${lastProductRow}C5)<=0,0,45+(CEILING(SUM(R${firstProductRow}C5:R${lastProductRow}C5)/10,1)-1)*25)`, amount: totals.shippingFee },
      { row: feesRow, label: "Transaction fees (5%)", amountFormula: `=(R${totalRow}C8+R${shippingRow}C8)*5%`, amount: totals.transactionFees },
      { row: amountRow, label: "Amount", amountFormula: `=R${totalRow}C8+R${shippingRow}C8+R${feesRow}C8`, amount: totals.amount },
    ].forEach((item) => {
      const dataProduct = dataProducts[item.row - 2];
      rows.push(xmlRow([
        xmlCell({ style: "Base", value: "" }),
        xmlCell({ style: "Summary", mergeAcross: 5, value: item.label }),
        xmlCell({ style: "MoneySummary", formula: item.amountFormula, value: item.amount }),
        xmlCell({ style: "Base", value: "" }),
        item.purchaseFormula ? xmlCell({ style: "IntegerSummary", formula: item.purchaseFormula, value: item.purchase }) : xmlCell({ style: "Base", value: "" }),
        item.profitFormula ? xmlCell({ style: "Number2Summary", formula: item.profitFormula, value: item.profit }) : xmlCell({ style: "Base", value: "" }),
        xmlCell({ index: 27, style: "Lookup", value: dataProduct && dataProduct.abbreviation }),
        xmlCell({ style: "Lookup", value: dataProduct && dataProduct.productName }),
        xmlCell({ style: "Lookup", value: dataProduct && dataProduct.specification }),
        xmlCell({ style: "LookupNumber", value: dataProduct ? parseMoney(dataProduct.price) : 0 }),
        xmlCell({ style: "LookupNumber", value: dataProduct ? parseMoney(dataProduct.purchaseCost) : 0 }),
      ], 28));
    });

    for (let excelRow = amountRow + 1; excelRow <= totalRows; excelRow += 1) {
      const dataProduct = dataProducts[excelRow - 2];
      rows.push(xmlRow([
        xmlCell({ index: 27, style: "Lookup", value: dataProduct && dataProduct.abbreviation }),
        xmlCell({ style: "Lookup", value: dataProduct && dataProduct.productName }),
        xmlCell({ style: "Lookup", value: dataProduct && dataProduct.specification }),
        xmlCell({ style: "LookupNumber", value: dataProduct ? parseMoney(dataProduct.price) : 0 }),
        xmlCell({ style: "LookupNumber", value: dataProduct ? parseMoney(dataProduct.purchaseCost) : 0 }),
      ], 8));
    }

    return `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
<Styles>
<Style ss:ID="Base"><Font ss:FontName="Arial" ss:Size="12" ss:Bold="1"/><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
<Style ss:ID="Title" ss:Parent="Base"><Font ss:FontName="Arial" ss:Size="16" ss:Bold="1"/><Interior ss:Color="#10AEE3" ss:Pattern="Solid"/></Style>
<Style ss:ID="Header" ss:Parent="Base"><Font ss:FontName="Arial" ss:Size="12" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#C00000" ss:Pattern="Solid"/></Style>
<Style ss:ID="Abbr" ss:Parent="Base"><Font ss:FontName="Arial" ss:Size="15" ss:Bold="1"/></Style>
<Style ss:ID="Summary" ss:Parent="Base"><Font ss:FontName="Arial" ss:Size="14" ss:Bold="1"/></Style>
<Style ss:ID="Money" ss:Parent="Base"><NumberFormat ss:Format="&quot;$&quot;#,##0.00"/></Style>
<Style ss:ID="MoneySummary" ss:Parent="Summary"><NumberFormat ss:Format="&quot;$&quot;#,##0.00"/></Style>
<Style ss:ID="Integer" ss:Parent="Base"><NumberFormat ss:Format="0"/></Style>
<Style ss:ID="IntegerSummary" ss:Parent="Summary"><NumberFormat ss:Format="0"/></Style>
<Style ss:ID="Number2" ss:Parent="Base"><NumberFormat ss:Format="0.00"/></Style>
<Style ss:ID="Number2Summary" ss:Parent="Summary"><NumberFormat ss:Format="0.00"/></Style>
<Style ss:ID="Percent" ss:Parent="Base"><NumberFormat ss:Format="0%"/></Style>
<Style ss:ID="Lookup"><Font ss:FontName="Arial" ss:Size="1" ss:Color="#FFFFFF"/></Style>
<Style ss:ID="LookupNumber" ss:Parent="Lookup"><NumberFormat ss:Format="0.00"/></Style>
</Styles><Worksheet ss:Name="报价单"><Table>${columns}${rows.join("")}</Table><WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><Selected/></WorksheetOptions></Worksheet></Workbook>`;
  }

  function downloadFormulaQuoteXml(totals, stamp) {
    const xml = makeFormulaQuoteXml(totals);
    const blob = new Blob(["\ufeff", xml], { type: "application/vnd.ms-excel;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `报价单_${stamp}.xls`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  }

  function formatPrice(price) {
    const text = String(price == null ? "" : price).trim();
    if (!text) return "-";
    return text.startsWith("$") ? text : `$${text}`;
  }

  async function loadWorkbook(data, fileName) {
    const workbook = XLSX.read(data, {
      type: "array",
      cellDates: true,
      cellFormula: true,
      cellNF: true,
      cellStyles: true,
      bookVBA: true,
    });

    const priceSheetInfo = findPriceSheet(workbook);
    if (!priceSheetInfo) {
      throw new Error("没有识别到价格表，请确认表头包含 Abbreviation/Product Name/Price 或对应中文列。");
    }

    state.workbook = workbook;
    state.priceSheetInfo = priceSheetInfo;
    state.quoteSheetName = findQuoteSheetName(workbook, priceSheetInfo.sheetName);
    indexProducts();

    el.fileName.textContent = fileName;
    el.sheetName.textContent = priceSheetInfo.sheetName;
    el.productCount.textContent = String(state.productRows.length);
    el.quoteSheetName.textContent = state.quoteSheetName;
    setStatus("已读取", "ok");
  }

  async function loadTemplate() {
    try {
      setStatus("读取模板中");
      const response = await fetch(TEMPLATE_URL);
      if (!response.ok) throw new Error(`模板文件加载失败：HTTP ${response.status}`);
      const data = await response.arrayBuffer();
      await loadWorkbook(data, TEMPLATE_NAME);
    } catch (error) {
      console.error(error);
      setStatus("模板失败", "error");
      alert(error.message || "读取内置价格表失败。");
    }
  }

  function generate() {
    if (!state.workbook) {
      alert("价格表还没有加载完成，请稍等。");
      return;
    }

    const count = analyzeMatches();
    if (count === 0) {
      alert("请输入需要报价的产品。");
      return;
    }
    if (state.matched.length === 0 && state.pending.length === 0) {
      alert("没有匹配到可报价的产品，请检查输入。");
      return;
    }
    if (!confirmPendingMatches()) {
      alert("已取消生成。请确认待确认产品后再生成。");
      return;
    }
    if (state.matched.length === 0) {
      alert("没有确认可报价的产品。");
      return;
    }
    const totals = getQuoteTotals();

    const date = new Date();
    const stamp = `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, "0")}${String(date.getDate()).padStart(2, "0")}_${String(date.getHours()).padStart(2, "0")}${String(date.getMinutes()).padStart(2, "0")}`;
    downloadFormulaQuoteXml(totals, stamp);
    setStatus("已生成", "ok");
  }

  el.fileInput.addEventListener("change", async (event) => {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    try {
      const data = await file.arrayBuffer();
      await loadWorkbook(data, file.name);
      if (el.input.value.trim()) analyzeMatches();
    } catch (error) {
      console.error(error);
      setStatus("读取失败", "error");
      alert(error.message || "读取上传文件失败。");
    }
  });
  el.preview.addEventListener("click", analyzeMatches);
  el.generate.addEventListener("click", generate);
  window.quoteAppDebug = state;
  loadTemplate();
})();
