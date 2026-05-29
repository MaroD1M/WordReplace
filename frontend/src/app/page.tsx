"use client";

import { FormEvent, useEffect, useMemo, useRef, useState } from "react";
import {
  applyTemplate,
  createTemplate,
  deleteSingleResult,
  executeReplace,
  fetchTemplateDetail,
  fetchTemplates,
  getExportUrl,
  getSingleFileExportUrl,
  previewFiles,
  PreviewData,
  ReplaceSummary,
  Rule,
  RuleTemplateListItem,
} from "@/lib/api";

const DEFAULT_CONTENT_WIDTH = 1400;
const DEFAULT_FONT_SIZE = 15;
const MAX_CONTENT_WIDTH = 1800;
const MIN_CONTENT_WIDTH = 1100;
const MAX_FONT_SIZE = 18;
const MIN_FONT_SIZE = 14;
const APP_VERSION = process.env.NEXT_PUBLIC_APP_VERSION ?? "2.0.3";

function Section({ title, hint, children }: { title: string; hint: string; children: React.ReactNode }) {
  return (
    <section className="rounded-xl border border-slate-200 bg-white/95 p-5 shadow-sm space-y-3">
      <h2 className="flex items-center gap-2 text-base font-semibold tracking-wide text-slate-900">
        <span className="h-4 w-1 rounded bg-cyan-600" />
        {title}
        <span className="ml-1 cursor-help rounded-full border border-slate-300 px-1.5 text-xs text-slate-500" title={hint}>i</span>
      </h2>
      {children}
    </section>
  );
}

export default function Home() {
  const stepAnchors = [
    { label: "上传文件", key: "step-upload" },
    { label: "自动预览", key: "step-preview" },
    { label: "规则管理", key: "step-rules" },
    { label: "执行替换", key: "step-execute" },
    { label: "导出结果", key: "step-export" },
  ] as const;

  const [rules, setRules] = useState<Rule[]>([]);
  const [keyword, setKeyword] = useState("");
  const [excelColumn, setExcelColumn] = useState("");
  const [wordFile, setWordFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);

  const [startRow, setStartRow] = useState(1);
  const [endRow, setEndRow] = useState(10);
  const [fileNameColumn, setFileNameColumn] = useState("");
  const [seqSource, setSeqSource] = useState("__auto");
  const [seqFormat] = useState<"1" | "01" | "0001" | "1." | "一">("1");
  const [namePrefix, setNamePrefix] = useState("");
  const [nameSuffix, setNameSuffix] = useState("");

  const [executing, setExecuting] = useState(false);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [error, setError] = useState("");
  const [summary, setSummary] = useState<ReplaceSummary | null>(null);
  const [preview, setPreview] = useState<PreviewData | null>(null);
  const [detailFilter, setDetailFilter] = useState<"all" | "success" | "failed">("all");
  const [editingRuleId, setEditingRuleId] = useState<number | null>(null);
  const [editingKeyword, setEditingKeyword] = useState("");
  const [editingExcelColumn, setEditingExcelColumn] = useState("");

  const [contentWidth, setContentWidth] = useState(DEFAULT_CONTENT_WIDTH);
  const [baseFontSize, setBaseFontSize] = useState(DEFAULT_FONT_SIZE);
  const [showExamplePanel, setShowExamplePanel] = useState(false);
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [previewWrap, setPreviewWrap] = useState(false);
  const [deletingItemId, setDeletingItemId] = useState<string | null>(null);
  const [stepNow, setStepNow] = useState(1);
  const [templates, setTemplates] = useState<RuleTemplateListItem[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState<number | null>(null);
  const [templateCreator, setTemplateCreator] = useState("");
  const [templateName, setTemplateName] = useState("");
  const [templateDescription, setTemplateDescription] = useState("");
  const [templateMode, setTemplateMode] = useState<"replace" | "append">("replace");
  const [templateInfo, setTemplateInfo] = useState("");
  const frameRef = useRef<number | null>(null);

  useEffect(() => {
    fetchTemplates().then(setTemplates).catch(() => void 0);
    try {
      const savedWidth = localStorage.getItem("wr_content_width");
      const savedFont = localStorage.getItem("wr_font_size");
      if (savedWidth) {
        const width = Math.min(MAX_CONTENT_WIDTH, Math.max(MIN_CONTENT_WIDTH, Number(savedWidth)));
        setContentWidth(width);
      }
      if (savedFont) {
        const font = Math.min(MAX_FONT_SIZE, Math.max(MIN_FONT_SIZE, Number(savedFont)));
        setBaseFontSize(font);
      }
    } catch {
      // ignore localStorage errors
    }
  }, []);

  useEffect(() => {
    if (!wordFile || !excelFile) {
      setPreview(null);
      return;
    }
    let cancelled = false;
    const timer = setTimeout(async () => {
      setPreviewLoading(true);
      setError("");
      try {
        const data = await previewFiles(wordFile, excelFile);
        if (cancelled) return;
        setPreview(data);
        setStartRow(1);
        setEndRow(Math.max(1, data.excel_total_rows));
        if (!fileNameColumn && data.excel_columns.length > 0) {
          setFileNameColumn(data.excel_columns[0]);
        }
      } catch (err) {
        if (!cancelled) setError(err instanceof Error ? err.message : "预览失败");
      } finally {
        if (!cancelled) setPreviewLoading(false);
      }
    }, 250);
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [wordFile, excelFile]);

  useEffect(() => {
    try {
      localStorage.setItem("wr_content_width", String(contentWidth));
      localStorage.setItem("wr_font_size", String(baseFontSize));
    } catch {
      // ignore
    }
  }, [contentWidth, baseFontSize]);

  const canExecute = useMemo(() => {
    return !!wordFile && !!excelFile && rules.length > 0 && startRow > 0 && endRow >= startRow && fileNameColumn.trim().length > 0;
  }, [wordFile, excelFile, rules.length, startRow, endRow, fileNameColumn]);
  const uiScale = baseFontSize / DEFAULT_FONT_SIZE;

  function scrollToStep(anchor: string) {
    const target = document.getElementById(anchor);
    if (!target) return;
    target.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function resetDisplaySettings() {
    setContentWidth(DEFAULT_CONTENT_WIDTH);
    setBaseFontSize(DEFAULT_FONT_SIZE);
  }

  function updateContentWidth(value: number) {
    if (frameRef.current) cancelAnimationFrame(frameRef.current);
    frameRef.current = requestAnimationFrame(() => {
      setContentWidth(value);
      frameRef.current = null;
    });
  }

  const executeHint = useMemo(() => {
    if (!wordFile || !excelFile) return "请先上传 Word 与 Excel 文件";
    if (rules.length === 0) return "请先添加至少一条规则";
    if (fileNameColumn.trim().length === 0) return "请填写文件名来源列";
    if (startRow <= 0 || endRow < startRow) return "请检查起始行与结束行";
    return "参数已完整，可开始替换";
  }, [wordFile, excelFile, rules.length, fileNameColumn, startRow, endRow]);

  function onSubmitRule(e: FormEvent) {
    e.preventDefault();
    if (!keyword.trim() || !excelColumn.trim()) return;
    const newRule: Rule = { id: Date.now(), keyword: keyword.trim(), excel_column: excelColumn.trim() };
    setRules((prev) => [...prev, newRule]);
    setError("");
    setKeyword("");
    setExcelColumn("");
  }

  function onDelete(id: number) {
    setError("");
    setRules((prev) => prev.filter((r) => r.id !== id));
  }

  function onStartEdit(rule: Rule) {
    setEditingRuleId(rule.id);
    setEditingKeyword(rule.keyword);
    setEditingExcelColumn(rule.excel_column);
  }

  function onSaveEdit() {
    if (!editingRuleId || !editingKeyword.trim() || !editingExcelColumn.trim()) return;
    setError("");
    setRules((prev) => prev.map((r) => (r.id === editingRuleId ? { ...r, keyword: editingKeyword.trim(), excel_column: editingExcelColumn.trim() } : r)));
    setEditingRuleId(null);
    setEditingKeyword("");
    setEditingExcelColumn("");
  }

  async function onExecute(e: FormEvent) {
    e.preventDefault();
    if (!wordFile || !excelFile) return;
    setExecuting(true);
    setError("");
    const useSeqPrefix = seqSource !== "__none";
    const seqColumn = seqSource.startsWith("__") ? "" : seqSource;
    try {
      const result = await executeReplace({
        wordFile,
        excelFile,
        startRow,
        endRow,
        fileNameColumn,
        rules: rules.map((r) => ({ keyword: r.keyword, excel_column: r.excel_column })),
        useSeqPrefix,
        seqFormat,
        seqColumn,
        namePrefix,
        nameSuffix,
      });
      setSummary(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : "执行失败");
    } finally {
      setExecuting(false);
    }
  }

  const filteredDetails = useMemo(() => {
    if (!summary) return [];
    if (detailFilter === "success") return summary.details.filter((d) => d.status === "成功");
    if (detailFilter === "failed") return summary.details.filter((d) => d.status === "失败");
    return summary.details;
  }, [summary, detailFilter]);

  async function copyFailedMessages() {
    if (!summary) return;
    const failed = summary.details.filter((d) => d.status === "失败");
    if (failed.length === 0) return;
    const text = failed.map((d) => `序号${d.seq} 行${d.row_number} 文件${d.file_name} 原因: ${d.message || "-"}`).join("\n");
    await navigator.clipboard.writeText(text);
  }

  async function onDeleteResult(itemId: string) {
    if (!summary) return;
    setError("");
    setDeletingItemId(itemId);
    try {
      const updated = await deleteSingleResult(summary.run_id, itemId, summary.export_token);
      setSummary({ ...summary, ...updated });
    } catch (err) {
      setError(err instanceof Error ? err.message : "删除结果失败");
    } finally {
      setDeletingItemId(null);
    }
  }

  async function onSaveTemplate() {
    if (!templateName.trim()) {
      setError("请输入模板名称");
      return;
    }
    if (rules.length === 0) {
      setError("当前没有可保存的规则");
      return;
    }
    setError("");
    try {
      await createTemplate({
        name: templateName.trim(),
        creator: templateCreator.trim() || "未填写",
        description: templateDescription.trim(),
        items: rules.map((r) => ({ keyword: r.keyword, excel_column: r.excel_column })),
      });
      setTemplateInfo("模板保存成功");
      setTemplates(await fetchTemplates());
    } catch (err) {
      setError(err instanceof Error ? err.message : "保存模板失败");
    }
  }

  async function onApplyTemplate() {
    if (!selectedTemplateId) {
      setError("请先选择模板");
      return;
    }
    if (!preview) {
      setError("请先上传并预览 Excel");
      return;
    }
    setError("");
    try {
      const result = await applyTemplate(selectedTemplateId, preview.excel_columns, templateMode);
      const detail = await fetchTemplateDetail(selectedTemplateId);
      const mapped = result.valid_items.map((it, idx) => ({
        id: Number(`${Date.now()}${idx}`),
        keyword: it.keyword,
        excel_column: it.excel_column,
      }));
      setRules((prev) => (templateMode === "replace" ? mapped : [...prev, ...mapped]));
      setTemplateInfo(`模板「${detail.name}」已应用：有效 ${result.valid_items.length} 条，无效 ${result.invalid_items.length} 条（已忽略）`);
    } catch (err) {
      setError(err instanceof Error ? err.message : "应用模板失败");
    }
  }

  useEffect(() => {
    if (!wordFile || !excelFile) setStepNow(1);
    else if (!preview) setStepNow(2);
    else if (rules.length === 0) setStepNow(3);
    else if (!summary) setStepNow(4);
    else setStepNow(5);
  }, [wordFile, excelFile, preview, rules.length, summary]);

  return (
    <main className="min-h-screen bg-[radial-gradient(ellipse_at_top,#e0f2fe_0%,#f8fafc_40%,#f8fafc_100%)] p-4 text-slate-900 md:p-8" style={{ fontSize: `${baseFontSize}px` }}>
      <div className="hidden xl:block fixed left-4 top-28 z-20 w-40">
        <div className="rounded-xl border border-slate-200 bg-white/95 p-3 shadow-sm">
          <p className="mb-2 text-xs font-semibold text-slate-500">流程进度</p>
          {stepAnchors.map((step, idx) => {
            const i = idx + 1;
            const done = i < stepNow;
            const active = i === stepNow;
            return (
              <button
                key={step.key}
                type="button"
                onClick={() => scrollToStep(step.key)}
                className="flex w-full items-center gap-2 rounded px-1 py-1 text-left hover:bg-slate-50"
              >
                <span className={`h-2.5 w-2.5 rounded-full ${done ? "bg-emerald-500" : active ? "bg-cyan-600" : "bg-slate-300"}`} />
                <span className={`text-xs ${active ? "font-semibold text-slate-800" : "text-slate-500"}`}>{step.label}</span>
              </button>
            );
          })}
        </div>
      </div>

      <div className="hidden xl:block fixed right-4 top-28 z-20 w-52">
        <div className="rounded-xl border border-slate-200 bg-white/95 p-3 shadow-sm space-y-3">
          <p className="text-xs font-semibold text-slate-500">显示设置</p>
          <label className="block text-xs">
            内容宽度：{contentWidth}px
            <input type="range" min={MIN_CONTENT_WIDTH} max={MAX_CONTENT_WIDTH} step={10} value={contentWidth} onChange={(e) => updateContentWidth(Number(e.target.value))} className="mt-1 w-full" />
          </label>
          <label className="block text-xs">
            字体大小：{baseFontSize}px
            <input type="range" min={MIN_FONT_SIZE} max={MAX_FONT_SIZE} step={1} value={baseFontSize} onChange={(e) => setBaseFontSize(Number(e.target.value))} className="mt-1 w-full" />
          </label>
          <button type="button" className="wr-btn w-full wr-btn-sm" onClick={resetDisplaySettings}>恢复默认</button>
          <div className="border-t border-slate-200 pt-2">
            <p className="mb-1 text-xs font-semibold text-slate-500">页面选项</p>
            <button type="button" className="wr-btn w-full wr-btn-sm" onClick={() => setShowAdvanced((v) => !v)}>
              {showAdvanced ? "隐藏高级项" : "显示高级项"}
            </button>
          </div>
        </div>
      </div>

      <div className="mx-auto space-y-5 ui-scale xl:px-48" style={{ maxWidth: `${contentWidth - 32}px`, ["--ui-scale" as never]: String(uiScale) }}>
        <section className="rounded-2xl border border-slate-200 bg-gradient-to-r from-slate-900 via-slate-800 to-cyan-900 p-6 text-slate-100 shadow-md">
          <div className="flex flex-col gap-4 xl:flex-row xl:items-end xl:justify-between">
            <div className="space-y-1">
              <p className="text-xs uppercase tracking-[0.18em] text-slate-300">WordReplace</p>
              <h1 className="text-2xl font-bold tracking-wide">文档批量替换助手</h1>
              <p className="text-sm text-slate-200">Word 模板 + Excel 数据，一键生成批量文档。</p>
            </div>
          </div>
        </section>

        <div id="step-upload" className="scroll-mt-24" />
        <Section title="1) 文件上传" hint="先选择 Word 和 Excel 文件，系统会自动预览。">
          <div className="rounded-xl border-2 border-dashed border-cyan-300 bg-cyan-50/70 p-4">
            <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
              <p className="text-sm font-semibold text-sky-800">请先上传 Word 模板与 Excel 数据（上传后自动预览）</p>
              <button type="button" className="wr-btn wr-btn-sm" onClick={() => setShowExamplePanel((v) => !v)}>
                {showExamplePanel ? "隐藏示例" : "示例与快速填充"}
              </button>
            </div>
            {showExamplePanel ? (
              <aside className="mb-3 wr-panel bg-white/90">
                <p className="text-sm font-medium text-slate-700">示例文件与规则（随机模拟数据）</p>
                <div className="mt-2 grid grid-cols-1 gap-2 md:grid-cols-2">
                  <div className="wr-panel bg-slate-50 p-2">
                    <p className="text-xs font-semibold text-slate-700">单页模板示例（15条）</p>
                    <div className="mt-1 flex flex-wrap gap-2 text-xs">
                      <a className="wr-btn wr-btn-sm inline-flex items-center justify-center hover:bg-slate-100" href="/examples/single_page_template.docx" download>Word</a>
                      <a className="wr-btn wr-btn-sm inline-flex items-center justify-center hover:bg-slate-100" href="/examples/single_page_data.xlsx" download>Excel</a>
                    </div>
                  </div>
                  <div className="wr-panel bg-slate-50 p-2">
                    <p className="text-xs font-semibold text-slate-700">多页模板示例（15条）</p>
                    <div className="mt-1 flex flex-wrap gap-2 text-xs">
                      <a className="wr-btn wr-btn-sm inline-flex items-center justify-center hover:bg-slate-100" href="/examples/multi_page_template.docx" download>Word</a>
                      <a className="wr-btn wr-btn-sm inline-flex items-center justify-center hover:bg-slate-100" href="/examples/multi_page_data.xlsx" download>Excel</a>
                    </div>
                  </div>
                </div>
              </aside>
            ) : null}
            <div className="grid grid-cols-1 gap-3 xl:grid-cols-2">
              <label className="wr-panel text-sm">
                <div className="mb-2 font-medium">Word 模板 (.docx)</div>
              <input className="h-10 w-full text-sm" type="file" accept=".docx" onChange={(e) => setWordFile(e.target.files?.[0] ?? null)} />
              {wordFile ? <button type="button" className="wr-btn mt-2 wr-btn-sm" onClick={() => setWordFile(null)}>清除 Word</button> : null}
              </label>
              <label className="wr-panel text-sm">
                <div className="mb-2 font-medium">Excel 数据 (.xlsx/.xls)</div>
              <input className="h-10 w-full text-sm" type="file" accept=".xlsx,.xls" onChange={(e) => setExcelFile(e.target.files?.[0] ?? null)} />
              {excelFile ? <button type="button" className="wr-btn mt-2 wr-btn-sm" onClick={() => setExcelFile(null)}>清除 Excel</button> : null}
              </label>
            </div>
          </div>

          {previewLoading ? <p className="text-sm text-slate-600">正在自动预览模板与数据...</p> : null}

          {preview ? (
            <div id="step-preview" className="scroll-mt-24">
            <div className="grid grid-cols-1 gap-3 xl:grid-cols-2">
              <div className="wr-panel bg-slate-50">
                <div className="mb-2 text-sm font-medium">Word 模板预览</div>
                <pre className="h-[360px] md:h-[420px] xl:h-[520px] overflow-auto whitespace-pre-wrap text-xs leading-5">{preview.word_text || "（模板无可显示文本）"}</pre>
              </div>
              <div className="wr-panel bg-slate-50">
                <div className="mb-2 flex flex-wrap items-center justify-between gap-2 text-sm font-medium">
                  <span>Excel 数据预览（前25行 / 共 {preview.excel_total_rows} 行）</span>
                  <label className="text-xs font-normal">
                    <input type="checkbox" checked={previewWrap} onChange={(e) => setPreviewWrap(e.target.checked)} className="mr-1" />
                    长文本换行
                  </label>
                </div>
                <div className="h-[360px] md:h-[420px] xl:h-[520px] overflow-auto rounded border border-slate-300 bg-white">
                  <table className="w-full text-xs">
                    <thead className="sticky top-0 z-10 bg-slate-100 shadow-sm">
                      <tr>
                        {preview.excel_columns.map((col) => (
                          <th key={col} className="min-w-[120px] border border-slate-300 px-2 py-1 text-center">{col}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {preview.excel_rows.map((row, idx) => (
                        <tr key={idx} className="odd:bg-slate-50/50">
                          {row.map((cell, cidx) => (
                            <td key={`${idx}-${cidx}`} className={`min-w-[120px] border border-slate-200 px-2 py-1 text-center ${previewWrap ? "whitespace-normal break-words" : "whitespace-nowrap truncate max-w-[260px]"}`} title={String(cell)}>{cell}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
            </div>
          ) : null}
        </Section>

        <div id="step-rules" className="scroll-mt-24" />
        <Section title="2) 规则管理" hint="配置模板关键字和 Excel 列名映射。">
          {showAdvanced ? (
          <div className="wr-panel bg-cyan-50/50 space-y-3">
            <div className="grid grid-cols-1 gap-3 xl:grid-cols-4">
            <input className="wr-input" placeholder="模板名称" value={templateName} onChange={(e) => setTemplateName(e.target.value)} />
            <input className="wr-input" placeholder="创建人（例如：张三）" value={templateCreator} onChange={(e) => setTemplateCreator(e.target.value)} />
            <input className="wr-input" placeholder="模板描述（可选）" value={templateDescription} onChange={(e) => setTemplateDescription(e.target.value)} />
            <button type="button" className="wr-btn" onClick={onSaveTemplate}>保存为模板</button>
            </div>
            <div className="grid grid-cols-1 gap-3 xl:grid-cols-4">
              <select className="wr-input" value={selectedTemplateId ?? ""} onChange={(e) => setSelectedTemplateId(e.target.value ? Number(e.target.value) : null)}>
                <option value="">选择模板库规则</option>
                {templates.map((t) => (
                  <option key={t.id} value={t.id}>{t.name} · {t.creator} · {t.item_count}条</option>
                ))}
              </select>
              <select className="wr-input" value={templateMode} onChange={(e) => setTemplateMode(e.target.value as "replace" | "append")}>
                <option value="replace">覆盖当前规则</option>
                <option value="append">追加到当前规则</option>
              </select>
              <button type="button" className="wr-btn" onClick={onApplyTemplate}>应用模板</button>
              <div className="flex items-center text-xs text-slate-500">{templateInfo || "模板会自动校验 Excel 列，不匹配规则将忽略"}</div>
            </div>
          </div>
          ) : null}

          <form onSubmit={onSubmitRule} className="grid grid-cols-1 gap-3 xl:grid-cols-3">
            <input className="wr-input" placeholder="模板关键字（如【姓名】）" value={keyword} onChange={(e) => setKeyword(e.target.value)} />
            <input list="excel-column-options" className="wr-input" value={excelColumn} onChange={(e) => setExcelColumn(e.target.value)} placeholder="请选择或输入 Excel 列名" disabled={!preview || preview.excel_columns.length === 0} />
            <button className="wr-btn wr-btn-primary disabled:opacity-50" disabled={!preview || preview.excel_columns.length === 0}>添加规则</button>
          </form>
          <datalist id="excel-column-options">
            {(preview?.excel_columns ?? []).map((col) => (
              <option key={col} value={col} />
            ))}
          </datalist>

          <div className="overflow-x-auto rounded-md border border-slate-300">
            <table className="w-full table-fixed text-sm">
              <thead className="bg-slate-200">
                <tr>
                  <th className="w-20 px-3 py-2 text-center">序号</th>
                  <th className="px-3 py-2 text-center">模板关键字</th>
                  <th className="px-3 py-2 text-center">Excel 列名</th>
                  <th className="w-44 px-3 py-2 text-center">操作</th>
                </tr>
              </thead>
              <tbody>
                {rules.length === 0 ? (
                  <tr><td className="px-3 py-2.5 text-center" colSpan={4}>暂无规则</td></tr>
                ) : (
                  rules.map((rule, idx) => (
                    <tr key={rule.id} className="border-t border-slate-200">
                      <td className="px-3 py-2.5 text-center">{idx + 1}</td>
                      <td className="px-3 py-2.5 text-center">
                        {editingRuleId === rule.id ? <input className="wr-input h-8 px-2 text-center" value={editingKeyword} onChange={(e) => setEditingKeyword(e.target.value)} /> : rule.keyword}
                      </td>
                      <td className="px-3 py-2.5 text-center">
                        {editingRuleId === rule.id ? <input list={`excel-column-options-${rule.id}`} className="wr-input h-8 px-2 text-center" value={editingExcelColumn} onChange={(e) => setEditingExcelColumn(e.target.value)} /> : rule.excel_column}
                        <datalist id={`excel-column-options-${rule.id}`}>{(preview?.excel_columns ?? []).map((col) => <option key={col} value={col} />)}</datalist>
                      </td>
                      <td className="px-3 py-2.5 text-center">
                        {editingRuleId === rule.id ? (
                          <div className="flex flex-nowrap items-center justify-center gap-2">
                            <button type="button" className="wr-btn wr-btn-sm" onClick={onSaveEdit}>保存</button>
                            <button type="button" className="wr-btn wr-btn-sm" onClick={() => { setEditingRuleId(null); setEditingKeyword(""); setEditingExcelColumn(""); }}>取消</button>
                          </div>
                        ) : (
                          <div className="flex flex-nowrap items-center justify-center gap-2">
                            <button type="button" className="wr-btn wr-btn-sm" onClick={() => onStartEdit(rule)}>编辑</button>
                            <button type="button" className="wr-btn wr-btn-sm" onClick={() => onDelete(rule.id)}>删除</button>
                          </div>
                        )}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </Section>

        <div id="step-execute" className="scroll-mt-24" />
        <Section title="3) 执行与导出" hint="设置行范围与命名策略，执行后下载结果。">
          <form onSubmit={onExecute} className="space-y-3">
            <div className="wr-panel">
              <div className="mb-2 text-sm font-semibold text-slate-800">基础参数</div>
              <div className="grid grid-cols-1 gap-3 xl:grid-cols-3">
                <label className="text-sm"><div className="mb-1 font-medium">起始行</div><input className="wr-input" type="number" min={1} value={startRow} onChange={(e) => setStartRow(Number(e.target.value) || 1)} /></label>
                <label className="text-sm"><div className="mb-1 font-medium">结束行</div><input className="wr-input" type="number" min={1} value={endRow} onChange={(e) => setEndRow(Number(e.target.value) || 1)} /></label>
                <label className="text-sm"><div className="mb-1 font-medium">文件名来源列</div><input list="excel-column-options-export" className="wr-input" value={fileNameColumn} onChange={(e) => setFileNameColumn(e.target.value)} /></label>
              </div>
            </div>

            <datalist id="excel-column-options-export">{(preview?.excel_columns ?? []).map((col) => <option key={col} value={col} />)}</datalist>

            <div className="wr-panel bg-slate-50">
              <div className="mb-2 text-sm font-semibold text-slate-800">文件名策略</div>
              <div className="grid grid-cols-1 gap-3 xl:grid-cols-3">
                <label className="text-sm">
                  <div className="mb-1 font-medium">序号来源</div>
                  <select className="wr-input" value={seqSource} onChange={(e) => setSeqSource(e.target.value)}>
                    <option value="__none">不添加序号</option>
                    <option value="__auto">使用替换序号</option>
                    {(preview?.excel_columns ?? []).map((col) => (
                      <option key={col} value={col}>使用 Excel 列：{col}</option>
                    ))}
                  </select>
                </label>
                <label className="text-sm"><div className="mb-1 font-medium">文件名前缀（可选）</div><input className="wr-input" value={namePrefix} onChange={(e) => setNamePrefix(e.target.value)} /></label>
                <label className="text-sm"><div className="mb-1 font-medium">文件名后缀（可选）</div><input className="wr-input" value={nameSuffix} onChange={(e) => setNameSuffix(e.target.value)} /></label>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-3">
              <button disabled={!canExecute || executing} className="wr-btn wr-btn-primary disabled:opacity-50">{executing ? "执行中..." : "开始替换"}</button>
              <span className="text-xs text-slate-500">{executeHint}</span>
            </div>
          </form>

          {summary ? (
            <>
              <div id="step-export" className="scroll-mt-24" />
              <div className="overflow-x-auto rounded-md border border-slate-300">
                <table className="w-full text-sm">
                  <thead className="bg-slate-200"><tr><th className="px-3 py-2 text-center">总数</th><th className="px-3 py-2 text-center">成功</th><th className="px-3 py-2 text-center">失败</th><th className="px-3 py-2 text-center">替换次数</th></tr></thead>
                  <tbody><tr className="border-t border-slate-200"><td className="px-3 py-2.5 text-center">{summary.total}</td><td className="px-3 py-2.5 text-center">{summary.success}</td><td className="px-3 py-2.5 text-center">{summary.failed}</td><td className="px-3 py-2.5 text-center">{summary.replacements}</td></tr></tbody>
                </table>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <span className="text-sm text-slate-600">明细筛选：</span>
                {showAdvanced ? (
                  <>
                    <button type="button" className="wr-btn wr-btn-sm" onClick={() => setDetailFilter("all")}>全部</button>
                    <button type="button" className="wr-btn wr-btn-sm" onClick={() => setDetailFilter("success")}>仅成功</button>
                    <button type="button" className="wr-btn wr-btn-sm" onClick={() => setDetailFilter("failed")}>仅失败</button>
                    <button type="button" className="wr-btn wr-btn-sm" onClick={copyFailedMessages}>复制失败原因</button>
                  </>
                ) : <span className="text-xs text-slate-500">开启“显示高级项”后可筛选明细与复制失败原因。</span>}
              </div>

              <div className="max-h-[420px] overflow-auto rounded-md border border-slate-300">
                <table className="w-full text-sm">
                  <thead className="sticky top-0 z-10 bg-slate-200">
                    <tr>
                      <th className="w-16 px-3 py-2 text-center">序号</th>
                      <th className="w-24 px-3 py-2 text-center">Excel行号</th>
                      <th className="px-3 py-2 text-center">生成文件名</th>
                      <th className="w-20 px-3 py-2 text-center">状态</th>
                      <th className="w-24 px-3 py-2 text-center">替换次数</th>
                      <th className="w-28 px-3 py-2 text-center">操作</th>
                      <th className="px-3 py-2 text-center">备注</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredDetails.map((item) => (
                      <tr key={`${item.item_id}`} className="border-t border-slate-200">
                        <td className="px-3 py-2.5 text-center">{item.seq}</td>
                        <td className="px-3 py-2.5 text-center">{item.row_number}</td>
                        <td className="px-3 py-2.5 text-center">{item.file_name}</td>
                        <td className="px-3 py-2.5 text-center"><span className={item.status === "成功" ? "rounded-full bg-emerald-50 px-2 py-0.5 text-emerald-700" : "rounded-full bg-red-50 px-2 py-0.5 text-red-700"}>{item.status}</span></td>
                        <td className="px-3 py-2.5 text-center">{item.replace_count}</td>
                        <td className="px-3 py-2.5 text-center">
                          <div className="flex flex-nowrap items-center justify-center gap-2">
                            <a className="wr-btn wr-btn-sm inline-flex items-center justify-center" href={getSingleFileExportUrl(summary.run_id, item.item_id, summary.export_token)}>下载</a>
                            <button type="button" disabled={deletingItemId === item.item_id} className="wr-btn wr-btn-sm disabled:opacity-50" onClick={() => onDeleteResult(item.item_id)}>
                              {deletingItemId === item.item_id ? "删除中" : "删除"}
                            </button>
                          </div>
                        </td>
                        <td className="px-3 py-2.5 text-center">{item.message || "-"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="flex gap-3">
                <a className="wr-btn inline-flex items-center" href={getExportUrl("zip", summary.run_id, summary.export_token)}>下载 ZIP</a>
                <a className="wr-btn inline-flex items-center" href={getExportUrl("merge", summary.run_id, summary.export_token)}>下载合并文档</a>
              </div>
            </>
          ) : null}
        </Section>

        {error ? <section className="rounded-md border border-red-300 bg-red-50 px-3 py-2 text-sm text-red-700">{error}</section> : null}

        <footer className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-slate-200 bg-slate-900 px-4 py-3 text-sm text-slate-300 shadow-sm">
          <div className="flex items-center gap-2">
            <span className="font-medium text-white">文档批量替换助手</span>
            <span className="rounded bg-slate-700 px-2 py-0.5 text-xs text-slate-100">v{APP_VERSION}</span>
          </div>
          <a
            href="https://github.com/MaroD1M/WordReplace"
            target="_blank"
            rel="noreferrer"
            className="inline-flex h-9 items-center gap-2 rounded-md border border-slate-600 bg-slate-800/80 px-3 text-sm text-slate-300 hover:bg-slate-700"
            aria-label="打开项目 GitHub 仓库"
          >
            <svg viewBox="0 0 24 24" aria-hidden="true" className="h-4 w-4 fill-current">
              <path d="M12 .5a12 12 0 0 0-3.79 23.38c.6.1.82-.26.82-.58v-2.03c-3.34.73-4.04-1.61-4.04-1.61-.55-1.4-1.34-1.78-1.34-1.78-1.1-.75.08-.74.08-.74 1.21.09 1.85 1.25 1.85 1.25 1.08 1.86 2.83 1.32 3.52 1.01.11-.79.42-1.32.76-1.63-2.67-.31-5.47-1.34-5.47-5.96 0-1.32.47-2.4 1.24-3.24-.13-.31-.54-1.57.12-3.27 0 0 1.01-.32 3.3 1.24a11.5 11.5 0 0 1 6 0c2.28-1.56 3.29-1.24 3.29-1.24.67 1.7.26 2.96.13 3.27.77.84 1.24 1.92 1.24 3.24 0 4.63-2.8 5.65-5.48 5.95.43.37.82 1.11.82 2.24v3.32c0 .32.22.69.83.57A12 12 0 0 0 12 .5z" />
            </svg>
            GitHub
          </a>
        </footer>
      </div>
    </main>
  );
}
