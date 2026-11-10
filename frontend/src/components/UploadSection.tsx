"use client";

import { useState } from "react";

import type { PreviewData } from "@/lib/api";
import Section from "@/components/Section";

type UploadSectionProps = {
  wordFile: File | null;
  setWordFile: (file: File | null) => void;
  excelFile: File | null;
  setExcelFile: (file: File | null) => void;
  preview: PreviewData | null;
  previewLoading: boolean;
  previewWrap: boolean;
  setPreviewWrap: (value: boolean) => void;
  sheetNames: string[];
  selectedSheet: string;
  setSelectedSheet: (value: string) => void;
  maxUploadSizeMb: number;
  onError: (message: string) => void;
};

function validateFile(file: File, extensions: string[], label: string, maxUploadSizeMb: number): string | null {
  const name = file.name.toLowerCase();
  if (!extensions.some((ext) => name.endsWith(ext))) {
    return `${label} 文件类型不支持`;
  }
  if (file.size > maxUploadSizeMb * 1024 * 1024) {
    return `${label} 超过 ${maxUploadSizeMb}MB 限制`;
  }
  return null;
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export default function UploadSection({
  wordFile,
  setWordFile,
  excelFile,
  setExcelFile,
  preview,
  previewLoading,
  previewWrap,
  setPreviewWrap,
  sheetNames,
  selectedSheet,
  setSelectedSheet,
  maxUploadSizeMb,
  onError,
}: UploadSectionProps) {
  const [showExamplePanel, setShowExamplePanel] = useState(false);
  const [wordDragActive, setWordDragActive] = useState(false);
  const [excelDragActive, setExcelDragActive] = useState(false);

  function handleWordChange(file: File | null) {
    if (!file) {
      setWordFile(null);
      return;
    }
    const message = validateFile(file, [".docx"], "Word", maxUploadSizeMb);
    if (message) {
      onError(message);
      return;
    }
    onError("");
    setWordFile(file);
  }

  function handleExcelChange(file: File | null) {
    if (!file) {
      setExcelFile(null);
      return;
    }
    const message = validateFile(file, [".xlsx", ".xls"], "Excel", maxUploadSizeMb);
    if (message) {
      onError(message);
      return;
    }
    onError("");
    setExcelFile(file);
  }

  return (
    <>
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
                    <a
                      className="wr-btn wr-btn-sm inline-flex items-center justify-center hover:bg-slate-100"
                      href="/examples/single_page_template.docx"
                      download
                    >
                      Word
                    </a>
                    <a
                      className="wr-btn wr-btn-sm inline-flex items-center justify-center hover:bg-slate-100"
                      href="/examples/single_page_data.xlsx"
                      download
                    >
                      Excel
                    </a>
                  </div>
                </div>
                <div className="wr-panel bg-slate-50 p-2">
                  <p className="text-xs font-semibold text-slate-700">多页模板示例（15条）</p>
                  <div className="mt-1 flex flex-wrap gap-2 text-xs">
                    <a
                      className="wr-btn wr-btn-sm inline-flex items-center justify-center hover:bg-slate-100"
                      href="/examples/multi_page_template.docx"
                      download
                    >
                      Word
                    </a>
                    <a
                      className="wr-btn wr-btn-sm inline-flex items-center justify-center hover:bg-slate-100"
                      href="/examples/multi_page_data.xlsx"
                      download
                    >
                      Excel
                    </a>
                  </div>
                </div>
              </div>
            </aside>
          ) : null}
          <div className="grid grid-cols-1 gap-3 xl:grid-cols-2">
            <div
              className={`wr-panel min-w-0 text-sm transition ${
                wordDragActive ? "bg-cyan-50/60 ring-2 ring-cyan-400" : ""
              }`}
              onDragOver={(event) => {
                event.preventDefault();
                setWordDragActive(true);
              }}
              onDragLeave={(event) => {
                event.preventDefault();
                setWordDragActive(false);
              }}
              onDrop={(event) => {
                event.preventDefault();
                setWordDragActive(false);
                handleWordChange(event.dataTransfer.files?.[0] ?? null);
              }}
            >
              <label className="block">
                <div className="mb-2 flex flex-wrap items-center justify-between gap-1">
                  <span className="font-medium">Word 模板 (.docx)</span>
                  <span className="text-xs font-normal text-slate-500">单文件上限 {maxUploadSizeMb} MB</span>
                </div>
                <input
                  className="h-10 w-full text-sm"
                  type="file"
                  accept=".docx"
                  onChange={(event) => {
                    handleWordChange(event.target.files?.[0] ?? null);
                    event.target.value = "";
                  }}
                />
                <div
                  className={`mt-1 truncate text-xs ${wordFile ? "text-emerald-700" : "text-slate-400"}`}
                  title={wordFile?.name}
                >
                  {wordFile ? `已选择：${wordFile.name}（${formatSize(wordFile.size)}）` : "未选择文件"}
                </div>
              </label>
              {wordFile ? (
                <button
                  type="button"
                  className="wr-btn mt-2 wr-btn-sm"
                  onClick={() => {
                    onError("");
                    setWordFile(null);
                  }}
                >
                  清除 Word
                </button>
              ) : null}
            </div>
            <div
              className={`wr-panel min-w-0 text-sm transition ${
                excelDragActive ? "bg-cyan-50/60 ring-2 ring-cyan-400" : ""
              }`}
              onDragOver={(event) => {
                event.preventDefault();
                setExcelDragActive(true);
              }}
              onDragLeave={(event) => {
                event.preventDefault();
                setExcelDragActive(false);
              }}
              onDrop={(event) => {
                event.preventDefault();
                setExcelDragActive(false);
                handleExcelChange(event.dataTransfer.files?.[0] ?? null);
              }}
            >
              <label className="block">
                <div className="mb-2 flex flex-wrap items-center justify-between gap-1">
                  <span className="font-medium">Excel 数据 (.xlsx/.xls)</span>
                  <span className="text-xs font-normal text-slate-500">单文件上限 {maxUploadSizeMb} MB</span>
                </div>
                <input
                  className="h-10 w-full text-sm"
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(event) => {
                    handleExcelChange(event.target.files?.[0] ?? null);
                    event.target.value = "";
                  }}
                />
                <div
                  className={`mt-1 truncate text-xs ${excelFile ? "text-emerald-700" : "text-slate-400"}`}
                  title={excelFile?.name}
                >
                  {excelFile ? `已选择：${excelFile.name}（${formatSize(excelFile.size)}）` : "未选择文件"}
                </div>
              </label>
              {excelFile ? (
                <button
                  type="button"
                  className="wr-btn mt-2 wr-btn-sm"
                  onClick={() => {
                    onError("");
                    setExcelFile(null);
                  }}
                >
                  清除 Excel
                </button>
              ) : null}
            </div>
          </div>
        </div>

        {previewLoading ? <p className="text-sm text-slate-600">正在自动预览模板与数据...</p> : null}

        {preview ? (
          <div id="step-preview" className="scroll-mt-24">
            <div className="grid grid-cols-1 gap-3 xl:grid-cols-2">
              <div className="wr-panel min-w-0 bg-slate-50">
                <div className="mb-2 text-sm font-medium">Word 模板预览</div>
                <pre className="h-[360px] md:h-[420px] xl:h-[520px] overflow-auto whitespace-pre-wrap text-xs leading-5">
                  {preview.word_text || "（模板无可显示文本）"}
                </pre>
              </div>
              <div className="wr-panel min-w-0 bg-slate-50">
                <div className="mb-2 flex flex-wrap items-center justify-between gap-2 text-sm font-medium">
                  <span>Excel 数据预览（前25行 / 共 {preview.excel_total_rows} 行）</span>
                  <div className="flex flex-wrap items-center gap-3 text-xs font-normal">
                    {sheetNames.length > 1 ? (
                      <label className="flex items-center gap-1">
                        工作表：
                        <select
                          className="rounded border border-slate-300 bg-white px-1 py-0.5"
                          value={selectedSheet}
                          onChange={(event) => setSelectedSheet(event.target.value)}
                        >
                          {sheetNames.map((sheet) => (
                            <option key={sheet} value={sheet}>
                              {sheet}
                            </option>
                          ))}
                        </select>
                      </label>
                    ) : null}
                    <label>
                      <input
                        type="checkbox"
                        checked={previewWrap}
                        onChange={(event) => setPreviewWrap(event.target.checked)}
                        className="mr-1"
                      />
                      长文本换行
                    </label>
                  </div>
                </div>
                <div className="h-[360px] md:h-[420px] xl:h-[520px] overflow-auto rounded border border-slate-300 bg-white">
                  <table className="w-full text-xs" aria-label="Excel 数据预览">
                    <thead className="sticky top-0 z-10 bg-slate-100 shadow-sm">
                      <tr>
                        {preview.excel_columns.map((col) => (
                          <th key={col} scope="col" className="min-w-[120px] border border-slate-300 px-2 py-1 text-center">
                            {col}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {preview.excel_rows.map((row, idx) => (
                        <tr key={idx} className="odd:bg-slate-50/50">
                          {row.map((cell, cidx) => (
                            <td
                              key={`${idx}-${cidx}`}
                              className={`min-w-[120px] border border-slate-200 px-2 py-1 text-center ${
                                previewWrap
                                  ? "whitespace-normal break-words"
                                  : "whitespace-nowrap truncate max-w-[260px]"
                              }`}
                              title={String(cell)}
                            >
                              {cell}
                            </td>
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
    </>
  );
}
