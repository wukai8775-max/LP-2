"use client";

import { useState } from "react";

type ApiResult = {
  fileName: string;
  fileBase64: string;
  matched: Array<{
    input: string;
    quantity: number;
    abbreviation: string;
    productName: string;
    specification: string;
    price: string | number;
    sheetName: string;
    matchType: string;
  }>;
  unmatched: Array<{ input: string; quantity: number }>;
};

export default function HomePage() {
  const [formulaFile, setFormulaFile] = useState<File | null>(null);
  const [products, setProducts] = useState("");
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<ApiResult | null>(null);
  const [error, setError] = useState("");

  async function submit() {
    setError("");
    setResult(null);

    if (!products.trim()) {
      setError("请输入需要生成的产品名称。");
      return;
    }

    const formData = new FormData();
    if (formulaFile) formData.append("formulaFile", formulaFile);
    formData.append("products", products);

    setLoading(true);
    try {
      const response = await fetch("/api/quote", {
        method: "POST",
        body: formData
      });
      const payload = await response.json();

      if (!response.ok) {
        throw new Error(payload.error || "生成失败。");
      }

      setResult(payload);
      downloadExcel(payload);
    } catch (err) {
      setError(err instanceof Error ? err.message : "生成失败。");
    } finally {
      setLoading(false);
    }
  }

  function downloadExcel(payload: ApiResult) {
    const byteCharacters = atob(payload.fileBase64);
    const bytes = new Uint8Array(byteCharacters.length);
    for (let i = 0; i < byteCharacters.length; i += 1) {
      bytes[i] = byteCharacters.charCodeAt(i);
    }

    const blob = new Blob([bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = payload.fileName;
    link.click();
    URL.revokeObjectURL(url);
  }

  return (
    <main className="min-h-screen bg-[#eef3f8] p-4 text-slate-900 md:p-8">
      <section className="mx-auto max-w-6xl">
        <header className="mb-5 flex flex-col gap-3 border-b border-slate-200 pb-5 md:flex-row md:items-end md:justify-between">
          <div>
            <p className="mb-2 text-sm font-bold text-teal-700">Excel 报价工具</p>
            <h1 className="text-3xl font-black tracking-normal md:text-4xl">按原销售价格表生成报价 Excel</h1>
          </div>
          <div className="rounded border border-slate-200 bg-white px-4 py-2 text-sm font-bold text-slate-600">
            默认使用 5.21 销售价格表
          </div>
        </header>

        <div className="grid gap-4 lg:grid-cols-[0.9fr_1.1fr]">
          <section className="rounded-lg border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="mb-4 text-xl font-black">价格表模板</h2>
            <label className="grid cursor-pointer gap-2 rounded-lg border border-dashed border-slate-300 bg-slate-50 p-4 hover:border-teal-700 hover:bg-teal-50">
              <span className="text-sm font-black text-slate-800">可选：上传新的销售价格表</span>
              <span className="text-sm text-slate-600">
                {formulaFile ? formulaFile.name : "不上传时，程序会使用已内置的 5.21 销售价格表。"}
              </span>
              <input
                className="hidden"
                type="file"
                accept=".xls,.xlsx,.xlsm"
                onChange={(event) => setFormulaFile(event.target.files?.[0] || null)}
              />
            </label>

            <div className="mt-4 rounded-lg bg-slate-50 p-4 text-sm leading-7 text-slate-700">
              程序会读取原表里的产品行，匹配你输入的产品名称。匹配到的行会保留，未选择的产品行会隐藏；原工作簿里的公式、列宽、行高、合并单元格和可读取的格式会随结果文件一起导出。
            </div>
          </section>

          <section className="rounded-lg border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="mb-4 text-xl font-black">输入产品名称</h2>
            <textarea
              className="min-h-[310px] w-full resize-y rounded-lg border border-slate-200 bg-slate-50 p-4 leading-7 outline-none focus:border-teal-700 focus:ring-4 focus:ring-teal-100"
              value={products}
              onChange={(event) => setProducts(event.target.value)}
              placeholder={"每行一个产品名称，也可以写数量，例如：\nMS40*2\n产品全称 A\nWA10 x 3"}
            />
            <div className="mt-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
              <p className="text-sm text-slate-500">支持产品全称、简称、规格；表格里没有的产品会跳过并列出来。</p>
              <button
                className="h-11 rounded-lg bg-teal-700 px-6 font-black text-white hover:bg-teal-800 disabled:cursor-not-allowed disabled:bg-slate-400"
                disabled={loading}
                onClick={submit}
              >
                {loading ? "生成中..." : "生成并下载 Excel"}
              </button>
            </div>
            {error ? <p className="mt-4 rounded-lg bg-red-50 p-3 text-sm font-bold text-red-700">{error}</p> : null}
          </section>
        </div>

        {result ? (
          <section className="mt-4 grid gap-4 rounded-lg border border-slate-200 bg-white p-5 shadow-sm md:grid-cols-2">
            <ResultList
              title={`已匹配 ${result.matched.length} 个`}
              items={result.matched.map((item) => `${item.input} x ${item.quantity} -> ${item.abbreviation || item.productName || item.specification}`)}
              empty="没有匹配到产品"
            />
            <ResultList
              title={`未匹配 ${result.unmatched.length} 个`}
              items={result.unmatched.map((item) => `${item.input} x ${item.quantity}`)}
              empty="无"
            />
          </section>
        ) : null}
      </section>
    </main>
  );
}

function ResultList({ title, items, empty }: { title: string; items: string[]; empty: string }) {
  return (
    <div>
      <h3 className="mb-3 font-black">{title}</h3>
      <ul className="grid gap-2">
        {(items.length ? items : [empty]).map((item) => (
          <li key={item} className="rounded-lg bg-slate-50 px-3 py-2 text-sm text-slate-700">
            {item}
          </li>
        ))}
      </ul>
    </div>
  );
}
