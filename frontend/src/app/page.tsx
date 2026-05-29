"use client";

import { FormEvent, useEffect, useMemo, useState } from "react";
import { createRule, deleteRule, executeReplace, fetchRules, getExportUrl, ReplaceSummary, Rule } from "@/lib/api";

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <section className="rounded-xl border border-slate-300 bg-white p-5 shadow-sm space-y-3">
      <h2 className="text-base font-semibold text-slate-900">{title}</h2>
      {children}
    </section>
  );
}

export default function Home() {
  const [rules, setRules] = useState<Rule[]>([]);
  const [keyword, setKeyword] = useState("");
  const [excelColumn, setExcelColumn] = useState("");
  const [wordFile, setWordFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);

  const [startRow, setStartRow] = useState(1);
  const [endRow, setEndRow] = useState(10);
  const [fileNameColumn, setFileNameColumn] = useState("自然村名");
  const [exportMode, setExportMode] = useState<"zip" | "merge">("zip");

  const [loading, setLoading] = useState(false);
  const [executing, setExecuting] = useState(false);
  const [error, setError] = useState("");
  const [summary, setSummary] = useState<ReplaceSummary | null>(null);

  async function loadRules() {
    setLoading(true);
    try {
      setRules(await fetchRules());
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    loadRules();
  }, []);

  const canExecute = useMemo(() => {
    return !!wordFile && !!excelFile && rules.length > 0 && startRow > 0 && endRow >= startRow && fileNameColumn.trim().length > 0;
  }, [wordFile, excelFile, rules.length, startRow, endRow, fileNameColumn]);

  async function onSubmitRule(e: FormEvent) {
    e.preventDefault();
    if (!keyword.trim() || !excelColumn.trim()) return;
    setError("");
    try {
      await createRule({ keyword: keyword.trim(), excel_column: excelColumn.trim() });
      setKeyword("");
      setExcelColumn("");
      await loadRules();
    } catch (err) {
      setError(err instanceof Error ? err.message : "新增规则失败");
    }
  }

  async function onDelete(id: number) {
    setError("");
    try {
      await deleteRule(id);
      await loadRules();
    } catch (err) {
      setError(err instanceof Error ? err.message : "删除规则失败");
    }
  }

  async function onExecute(e: FormEvent) {
    e.preventDefault();
    if (!wordFile || !excelFile) return;
    setExecuting(true);
    setError("");
    try {
      const result = await executeReplace({
        wordFile,
        excelFile,
        startRow,
        endRow,
        fileNameColumn,
        exportMode,
      });
      setSummary(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : "执行失败");
    } finally {
      setExecuting(false);
    }
  }

  return (
    <main className="min-h-screen bg-slate-100 p-6 md:p-10 text-slate-900">
      <div className="mx-auto max-w-6xl space-y-5">
        <Section title="Word + Excel 批量替换（FastAPI + Next.js）">
          <p className="text-sm text-slate-600">按顺序完成上传、规则管理、执行与导出，界面全流程统一表格与居中风格。</p>
        </Section>

        <Section title="1) 文件上传">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <label className="rounded-md border border-slate-300 bg-slate-50 p-3 text-sm">
              <div className="mb-2 font-medium">Word 模板 (.docx)</div>
              <input type="file" accept=".docx" onChange={(e) => setWordFile(e.target.files?.[0] ?? null)} />
            </label>
            <label className="rounded-md border border-slate-300 bg-slate-50 p-3 text-sm">
              <div className="mb-2 font-medium">Excel 数据 (.xlsx/.xls)</div>
              <input type="file" accept=".xlsx,.xls" onChange={(e) => setExcelFile(e.target.files?.[0] ?? null)} />
            </label>
          </div>
        </Section>

        <Section title="2) 规则管理">
          <form onSubmit={onSubmitRule} className="grid grid-cols-1 md:grid-cols-3 gap-3">
            <input className="rounded-md border border-slate-300 px-3 py-2 text-sm" placeholder="模板关键字（如【姓名】）" value={keyword} onChange={(e) => setKeyword(e.target.value)} />
            <input className="rounded-md border border-slate-300 px-3 py-2 text-sm" placeholder="Excel 列名" value={excelColumn} onChange={(e) => setExcelColumn(e.target.value)} />
            <button className="rounded-md bg-slate-900 text-white px-4 py-2 text-sm">添加规则</button>
          </form>

          <div className="overflow-x-auto rounded-md border border-slate-300">
            <table className="w-full text-sm">
              <thead className="bg-slate-200">
                <tr>
                  <th className="px-3 py-2 text-center w-20">序号</th>
                  <th className="px-3 py-2 text-center">模板关键字</th>
                  <th className="px-3 py-2 text-center">Excel 列名</th>
                  <th className="px-3 py-2 text-center w-28">操作</th>
                </tr>
              </thead>
              <tbody>
                {loading ? (
                  <tr><td className="px-3 py-3 text-center" colSpan={4}>加载中...</td></tr>
                ) : rules.length === 0 ? (
                  <tr><td className="px-3 py-3 text-center" colSpan={4}>暂无规则</td></tr>
                ) : (
                  rules.map((rule, idx) => (
                    <tr key={rule.id} className="border-t border-slate-200">
                      <td className="px-3 py-2 text-center">{idx + 1}</td>
                      <td className="px-3 py-2 text-center">{rule.keyword}</td>
                      <td className="px-3 py-2 text-center">{rule.excel_column}</td>
                      <td className="px-3 py-2 text-center"><button className="rounded border border-slate-300 px-2 py-1 text-xs" onClick={() => onDelete(rule.id)}>删除</button></td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </Section>

        <Section title="3) 执行与导出">
          <form onSubmit={onExecute} className="grid grid-cols-1 md:grid-cols-5 gap-3">
            <input className="rounded-md border border-slate-300 px-3 py-2 text-sm" type="number" min={1} value={startRow} onChange={(e) => setStartRow(Number(e.target.value) || 1)} placeholder="起始行" />
            <input className="rounded-md border border-slate-300 px-3 py-2 text-sm" type="number" min={1} value={endRow} onChange={(e) => setEndRow(Number(e.target.value) || 1)} placeholder="结束行" />
            <input className="rounded-md border border-slate-300 px-3 py-2 text-sm" value={fileNameColumn} onChange={(e) => setFileNameColumn(e.target.value)} placeholder="文件名列" />
            <select className="rounded-md border border-slate-300 px-3 py-2 text-sm" value={exportMode} onChange={(e) => setExportMode(e.target.value as "zip" | "merge")}> 
              <option value="zip">ZIP 导出</option>
              <option value="merge">合并文档</option>
            </select>
            <button disabled={!canExecute || executing} className="rounded-md bg-slate-900 text-white px-4 py-2 text-sm disabled:opacity-50">
              {executing ? "执行中..." : "开始替换"}
            </button>
          </form>

          {summary ? (
            <>
              <div className="overflow-x-auto rounded-md border border-slate-300">
                <table className="w-full text-sm">
                  <thead className="bg-slate-200">
                    <tr>
                      <th className="px-3 py-2 text-center">总数</th>
                      <th className="px-3 py-2 text-center">成功</th>
                      <th className="px-3 py-2 text-center">失败</th>
                      <th className="px-3 py-2 text-center">替换次数</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr className="border-t border-slate-200">
                      <td className="px-3 py-2 text-center">{summary.total}</td>
                      <td className="px-3 py-2 text-center">{summary.success}</td>
                      <td className="px-3 py-2 text-center">{summary.failed}</td>
                      <td className="px-3 py-2 text-center">{summary.replacements}</td>
                    </tr>
                  </tbody>
                </table>
              </div>

              <div className="flex gap-3">
                <a className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm" href={getExportUrl("zip", summary.run_id)}>下载 ZIP</a>
                <a className="rounded-md border border-slate-300 bg-white px-4 py-2 text-sm" href={getExportUrl("merge", summary.run_id)}>下载合并文档</a>
              </div>
            </>
          ) : null}
        </Section>

        {error ? <section className="rounded-md border border-red-300 bg-red-50 px-3 py-2 text-sm text-red-700">{error}</section> : null}
      </div>
    </main>
  );
}
