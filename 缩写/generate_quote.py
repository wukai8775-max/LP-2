"""
Excel 报价单自动生成工具

使用方法：
1. 在 products.txt 中每行填写一个产品，支持缩写或产品全称。
   可选格式：
      RT30
      Retatrutide | 30mg*10vials
      RT30, 2
      RT30, 2, 0.05
   说明：逗号第 2 项是数量，第 3 项是折扣；竖线后面是规格。

2. 安装依赖并运行：
      pip install -r requirements.txt
      python generate_quote.py

3. 程序会复制原 Excel 报价工具，生成新的 .xlsx 文件：
      报价单_YYYYMMDD_HHMMSS.xlsx

重点：
- 本脚本依赖本机 Microsoft Excel 来打开 .xls 并另存为 .xlsx。
- 脚本只写入“报价工具”sheet 的输入区 N:R，不改公式区 A:M。
- 字体、格式、颜色、边框、合并单元格、公式、数字格式、打印格式均由 Excel 原生复制保留。
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
import shutil
import sys
from typing import Iterable

try:
    import xlwings as xw
except ImportError:
    print("缺少依赖 xlwings。请先运行：pip install -r requirements.txt")
    raise


SOURCE_FILE = Path(r"C:\Users\Administrator\Desktop\5.21销售价格表_报价工具.xls")
PRODUCT_INPUT_FILE = Path(__file__).with_name("products.txt")
OUTPUT_DIR = Path(__file__).resolve().parent

PRICE_SHEET = "采购成本"
QUOTE_SHEET = "报价工具"

HEADER_ROW = 2
DATA_START_ROW = 3
DATA_END_ROW = 68

COL_ABBR = "B"
COL_PRODUCT_NAME = "C"
COL_SPEC = "D"
COL_PRICE = "E"
COL_PURCHASE_COST = "F"

INPUT_START_ROW = 3
INPUT_END_ROW = 12
INPUT_COL_ABBR = "N"
INPUT_COL_PRODUCT_NAME = "O"
INPUT_COL_SPEC = "P"
INPUT_COL_QTY = "Q"
INPUT_COL_DISCOUNT = "R"


@dataclass
class ProductRecord:
    row: int
    abbreviation: str
    product_name: str
    specification: str
    price: object
    purchase_cost: object


@dataclass
class ProductRequest:
    raw: str
    query: str
    specification: str = ""
    quantity: int = 1
    discount: float = 0.0


@dataclass
class MatchResult:
    request: ProductRequest
    product: ProductRecord | None
    reason: str = ""


def normalize(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\u3000", " ")).strip().lower()


def compact(value: object) -> str:
    return re.sub(r"[\s._\-/\\()（）]+", "", normalize(value))


def parse_product_line(line: str) -> ProductRequest | None:
    raw = line.strip()
    if not raw or raw.startswith("#"):
        return None

    parts = [part.strip() for part in re.split(r"[,，\t]", raw)]
    query_part = parts[0]
    spec = ""
    if "|" in query_part:
        query_part, spec = [part.strip() for part in query_part.split("|", 1)]

    quantity = 1
    discount = 0.0
    if len(parts) >= 2 and parts[1]:
        quantity = int(float(parts[1]))
    if len(parts) >= 3 and parts[2]:
        discount = float(parts[2].rstrip("%"))
        if discount > 1:
            discount = discount / 100

    return ProductRequest(raw=raw, query=query_part, specification=spec, quantity=quantity, discount=discount)


def load_requests(path: Path) -> list[ProductRequest]:
    if not path.exists():
        raise FileNotFoundError(f"未找到输入文件：{path}")
    requests: list[ProductRequest] = []
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        request = parse_product_line(line)
        if request:
            requests.append(request)
    return requests


def read_products(sheet) -> list[ProductRecord]:
    records: list[ProductRecord] = []
    for row in range(DATA_START_ROW, DATA_END_ROW + 1):
        abbr = str(sheet.range(f"{COL_ABBR}{row}").value or "").strip()
        name = str(sheet.range(f"{COL_PRODUCT_NAME}{row}").value or "").strip()
        spec = str(sheet.range(f"{COL_SPEC}{row}").value or "").strip()
        if not any([abbr, name, spec]):
            continue
        records.append(
            ProductRecord(
                row=row,
                abbreviation=abbr,
                product_name=name,
                specification=spec,
                price=sheet.range(f"{COL_PRICE}{row}").value,
                purchase_cost=sheet.range(f"{COL_PURCHASE_COST}{row}").value,
            )
        )
    return records


def match_one(request: ProductRequest, records: Iterable[ProductRecord]) -> MatchResult:
    query_key = compact(request.query)
    spec_key = compact(request.specification)
    records = list(records)

    exact_abbr = [item for item in records if compact(item.abbreviation) == query_key]
    if spec_key:
        exact_abbr = [item for item in exact_abbr if spec_key in compact(item.specification)]
    if len(exact_abbr) == 1:
        return MatchResult(request, exact_abbr[0])
    if len(exact_abbr) > 1:
        return MatchResult(request, None, "缩写匹配到多个规格，请在输入中加 |规格")

    exact_name = [item for item in records if compact(item.product_name) == query_key]
    if spec_key:
        exact_name = [item for item in exact_name if spec_key in compact(item.specification)]
    if len(exact_name) == 1:
        return MatchResult(request, exact_name[0])
    if len(exact_name) > 1:
        return MatchResult(request, None, "产品全称匹配到多个规格，请输入缩写或加 |规格")

    contains = [
        item
        for item in records
        if query_key and (query_key in compact(item.abbreviation) or query_key in compact(item.product_name))
    ]
    if spec_key:
        contains = [item for item in contains if spec_key in compact(item.specification)]
    if len(contains) == 1:
        return MatchResult(request, contains[0])
    if len(contains) > 1:
        return MatchResult(request, None, "模糊匹配到多个产品，请输入更完整的缩写/名称")

    return MatchResult(request, None, "未找到")


def clear_quote_inputs(sheet) -> None:
    sheet.range(f"{INPUT_COL_ABBR}{INPUT_START_ROW}:{INPUT_COL_DISCOUNT}{INPUT_END_ROW}").clear_contents()


def fill_quote_inputs(sheet, matches: list[MatchResult]) -> None:
    if len(matches) > INPUT_END_ROW - INPUT_START_ROW + 1:
        raise ValueError("报价工具最多支持 10 个产品；请拆成多次生成。")

    clear_quote_inputs(sheet)
    for offset, match in enumerate(matches):
        row = INPUT_START_ROW + offset
        request = match.request
        product = match.product
        assert product is not None

        # 只填写报价工具的输入区，让原公式继续计算产品名称、规格、价格、运费和合计。
        sheet.range(f"{INPUT_COL_ABBR}{row}").value = product.abbreviation
        sheet.range(f"{INPUT_COL_PRODUCT_NAME}{row}").value = ""
        sheet.range(f"{INPUT_COL_SPEC}{row}").value = product.specification
        sheet.range(f"{INPUT_COL_QTY}{row}").value = request.quantity
        sheet.range(f"{INPUT_COL_DISCOUNT}{row}").value = request.discount


def make_output_path() -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return OUTPUT_DIR / f"报价单_{stamp}.xlsx"


def save_as_xlsx_with_excel(source: Path, output: Path, requests: list[ProductRequest]) -> tuple[list[MatchResult], list[MatchResult]]:
    if not source.exists():
        raise FileNotFoundError(f"未找到源文件：{source}")

    temp_xls = output.with_suffix(".working.xls")
    shutil.copy2(source, temp_xls)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    book = None
    try:
        book = app.books.open(str(temp_xls), update_links=False, read_only=False)
        price_sheet = book.sheets[PRICE_SHEET]
        quote_sheet = book.sheets[QUOTE_SHEET]

        records = read_products(price_sheet)
        results = [match_one(request, records) for request in requests]
        matched = [result for result in results if result.product is not None]
        unmatched = [result for result in results if result.product is None]

        fill_quote_inputs(quote_sheet, matched)

        book.app.api.Calculation = -4105  # xlCalculationAutomatic
        book.app.api.CalculateFullRebuild()
        book.api.ForceFullCalculation = True
        book.api.RefreshAll()
        quote_sheet.activate()
        book.api.SaveAs(str(output), FileFormat=51)  # 51 = .xlsx
        return matched, unmatched
    finally:
        if book is not None:
            book.close()
        app.quit()
        if temp_xls.exists():
            temp_xls.unlink()


def main() -> int:
    requests = load_requests(PRODUCT_INPUT_FILE)
    if not requests:
        print("products.txt 中没有有效产品。")
        return 1

    output = make_output_path()
    matched, unmatched = save_as_xlsx_with_excel(SOURCE_FILE, output, requests)

    print(f"已生成报价单：{output}")
    print()
    print("已成功生成的产品：")
    if matched:
        for item in matched:
            product = item.product
            assert product is not None
            print(f"- {item.request.raw} -> {product.abbreviation} / {product.product_name} / {product.specification}")
    else:
        print("- 无")

    print()
    print("未找到的产品清单：")
    if unmatched:
        for item in unmatched:
            print(f"- {item.request.raw}（{item.reason}）")
    else:
        print("- 无")
    return 0


if __name__ == "__main__":
    sys.exit(main())
