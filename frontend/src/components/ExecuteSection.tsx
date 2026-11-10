"use client";

import { FormEvent } from "react";

import type { PreviewData, ReplaceSummary, SeqFormat } from "@/lib/api";
import { SEQ_FORMATS } from "@/lib/constants";
import ColumnSelect from "@/components/ColumnSelect";
import ResultsSection from "@/components/ResultsSection";
import Section from "@/components/Section";

type ExecuteSectionProps = {
  preview: PreviewData | null;
  startRow: number;
  setStartRow: (value: number) => void;
  endRow: number;
  setEndRow: (value: number) => void;
  fileNameColumn: string;
  setFileNameColumn: (value: string) => void;
  seqSource: string;
  setSeqSource: (value: string) => void;
  seqFormat: SeqFormat;
  setSeqFormat: (value: SeqFormat) => void;
  namePrefix: string;
  setNamePrefix: (value: string) => void;
  nameSuffix: string;
  setNameSuffix: (value: string) => void;
  executing: boolean;
  canExecute: boolean;
  executeHint: string;
  onSubmit: (event: FormEvent) => void;
  summary: ReplaceSummary | null;
  setSummary: React.Dispatch<React.SetStateAction<ReplaceSummary | null>>;
  onError: (message: string) => void;
  showAdvanced: boolean;
};

export default function ExecuteSection({
  preview,
  startRow,
  setStartRow,
  endRow,
  setEndRow,
  fileNameColumn,
  setFileNameColumn,
  seqSource,
  setSeqSource,
  seqFormat,
  setSeqFormat,
  namePrefix,
  setNamePrefix,
  nameSuffix,
  setNameSuffix,
  executing,
  canExecute,
  executeHint,
  onSubmit,
  summary,
  setSummary,
  onError,
  showAdvanced,
}: ExecuteSectionProps) {
  const excelColumns = preview?.excel_columns ?? [];

  return (
    <>
      <div id="step-execute" className="scroll-mt-24" />
      <Section title="3) 执行与导出" hint="设置行范围与命名策略，执行后下载结果。">
        <form onSubmit={onSubmit} className="space-y-3">
          <div className="wr-panel">
            <div className="mb-2 text-sm font-semibold text-slate-800">基础参数</div>
            <div className="grid grid-cols-1 gap-3 xl:grid-cols-[1fr_1fr_1.4fr]">
              <label className="min-w-0 text-sm">
                <div className="mb-1 font-medium">起始行</div>
                <input
                  className="wr-input"
                  type="number"
                  min={1}
                  value={startRow === 0 ? "" : startRow}
                  onChange={(event) => setStartRow(event.target.value === "" ? 0 : Number(event.target.value))}
                  onBlur={() => {
                    if (startRow < 1) setStartRow(1);
                  }}
                />
              </label>
              <label className="min-w-0 text-sm">
                <div className="mb-1 font-medium">结束行</div>
                <input
                  className="wr-input"
                  type="number"
                  min={1}
                  value={endRow === 0 ? "" : endRow}
                  onChange={(event) => setEndRow(event.target.value === "" ? 0 : Number(event.target.value))}
                  onBlur={() => {
                    if (endRow < 1) setEndRow(1);
                  }}
                />
              </label>
              <label className="min-w-0 text-sm">
                <div className="mb-1 font-medium">文件名来源列</div>
                <ColumnSelect
                  value={fileNameColumn}
                  onChange={setFileNameColumn}
                  options={excelColumns}
                  placeholder="请选择文件名来源列"
                  ariaLabel="选择文件名来源列"
                />
              </label>
            </div>
          </div>

          <div className="wr-panel bg-slate-50">
            <div className="mb-2 text-sm font-semibold text-slate-800">文件名策略</div>
            <div className="grid grid-cols-1 gap-3 xl:grid-cols-4">
              <label className="min-w-0 text-sm">
                <div className="mb-1 font-medium">序号来源</div>
                <select className="wr-input" value={seqSource} onChange={(event) => setSeqSource(event.target.value)}>
                  <option value="__none">不添加序号</option>
                  <option value="__auto">使用替换序号</option>
                  {excelColumns.map((col) => (
                    <option key={col} value={col}>
                      使用 Excel 列：{col}
                    </option>
                  ))}
                </select>
              </label>
              <label className="min-w-0 text-sm">
                <div className="mb-1 font-medium">序号格式</div>
                <select
                  className="wr-input"
                  value={seqFormat}
                  onChange={(event) => setSeqFormat(event.target.value as SeqFormat)}
                >
                  {SEQ_FORMATS.map((format) => (
                    <option key={format.value} value={format.value}>
                      {format.label}
                    </option>
                  ))}
                </select>
              </label>
              <label className="min-w-0 text-sm">
                <div className="mb-1 font-medium">文件名前缀（可选）</div>
                <input className="wr-input" value={namePrefix} onChange={(event) => setNamePrefix(event.target.value)} />
              </label>
              <label className="min-w-0 text-sm">
                <div className="mb-1 font-medium">文件名后缀（可选）</div>
                <input className="wr-input" value={nameSuffix} onChange={(event) => setNameSuffix(event.target.value)} />
              </label>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            <button disabled={!canExecute || executing} className="wr-btn wr-btn-primary disabled:opacity-50">
              {executing ? "执行中..." : "开始替换"}
            </button>
            <span className="text-xs text-slate-500">{executeHint}</span>
          </div>
        </form>

        {summary ? (
          <ResultsSection summary={summary} setSummary={setSummary} onError={onError} showAdvanced={showAdvanced} />
        ) : null}
      </Section>
    </>
  );
}
