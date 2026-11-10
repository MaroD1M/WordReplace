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
};

export default function UploadSection({
  wordFile,
  setWordFile,
  excelFile,
  setExcelFile,
  preview,
  previewLoading,
  previewWrap,
  setPreviewWrap,
}: UploadSectionProps) {
  const [showExamplePanel, setShowExamplePanel] = useState(false);

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
            <label className="wr-panel min-w-0 text-sm">
              <div className="mb-2 font-medium">Word 模板 (.docx)</div>
              <input
                className="h-10 w-full text-sm"
                type="file"
                accept=".docx"
                onChange={(event) => setWordFile(event.target.files?.[0] ?? null)}
              />
              {wordFile ? (
                <button type="button" className="wr-btn mt-2 wr-btn-sm" onClick={() => setWordFile(null)}>
                  清除 Word
                </button>
              ) : null}
            </label>
            <label className="wr-panel min-w-0 text-sm">
              <div className="mb-2 font-medium">Excel 数据 (.xlsx/.xls)</div>
              <input
                className="h-10 w-full text-sm"
                type="file"
                accept=".xlsx,.xls"
                onChange={(event) => setExcelFile(event.target.files?.[0] ?? null)}
              />
              {excelFile ? (
                <button type="button" className="wr-btn mt-2 wr-btn-sm" onClick={() => setExcelFile(null)}>
                  清除 Excel
                </button>
              ) : null}
            </label>
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
                  <label className="text-xs font-normal">
                    <input
                      type="checkbox"
                      checked={previewWrap}
                      onChange={(event) => setPreviewWrap(event.target.checked)}
                      className="mr-1"
                    />
                    长文本换行
                  </label>
                </div>
                <div className="h-[360px] md:h-[420px] xl:h-[520px] overflow-auto rounded border border-slate-300 bg-white">
                  <table className="w-full text-xs">
                    <thead className="sticky top-0 z-10 bg-slate-100 shadow-sm">
                      <tr>
                        {preview.excel_columns.map((col) => (
                          <th key={col} className="min-w-[120px] border border-slate-300 px-2 py-1 text-center">
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
