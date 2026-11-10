"use client";

import { FormEvent, useEffect, useMemo, useState } from "react";

import { executeReplace, type ReplaceSummary, type Rule, type SeqFormat } from "@/lib/api";
import DisplaySettings from "@/components/DisplaySettings";
import ExecuteSection from "@/components/ExecuteSection";
import Footer from "@/components/Footer";
import ProgressSidebar from "@/components/ProgressSidebar";
import RulesSection from "@/components/RulesSection";
import UploadSection from "@/components/UploadSection";
import { useAppConfig } from "@/hooks/useAppConfig";
import { useDisplaySettings } from "@/hooks/useDisplaySettings";
import { usePreview } from "@/hooks/usePreview";

export default function Home() {
  const [rules, setRules] = useState<Rule[]>([]);
  const [wordFile, setWordFile] = useState<File | null>(null);
  const [excelFile, setExcelFile] = useState<File | null>(null);

  const [startRow, setStartRow] = useState(1);
  const [endRow, setEndRow] = useState(10);
  const [fileNameColumn, setFileNameColumn] = useState("");
  const [seqSource, setSeqSource] = useState("__auto");
  const [seqFormat, setSeqFormat] = useState<SeqFormat>("1");
  const [namePrefix, setNamePrefix] = useState("");
  const [nameSuffix, setNameSuffix] = useState("");
  const [selectedSheet, setSelectedSheet] = useState("");
  const [keepPlaceholderOnEmpty, setKeepPlaceholderOnEmpty] = useState(false);

  const [executing, setExecuting] = useState(false);
  const [error, setError] = useState("");
  const [summary, setSummary] = useState<ReplaceSummary | null>(null);
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [previewWrap, setPreviewWrap] = useState(false);

  const { preview, previewLoading, previewError } = usePreview(wordFile, excelFile, selectedSheet);
  const { config, configError } = useAppConfig();
  const {
    contentWidth,
    baseFontSize,
    setBaseFontSize,
    resetDisplaySettings,
    updateContentWidth,
    uiScale,
  } = useDisplaySettings();

  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect -- a new Excel file resets the sheet choice
    setSelectedSheet("");
  }, [excelFile]);

  // Re-apply row range and validate the filename/sequence columns whenever a
  // new preview arrives (e.g. after switching the Excel file or sheet).
  useEffect(() => {
    if (!preview) return;
    // eslint-disable-next-line react-hooks/set-state-in-effect -- reset controls when a new preview arrives
    setSelectedSheet((current) => current || preview.selected_sheet);
    setStartRow(1);
    setEndRow(Math.max(1, preview.excel_total_rows));
    setFileNameColumn((current) =>
      current && preview.excel_columns.includes(current) ? current : (preview.excel_columns[0] ?? ""),
    );
    setSeqSource((current) =>
      current.startsWith("__") || preview.excel_columns.includes(current) ? current : "__auto",
    );
  }, [preview]);

  // Any change to inputs invalidates previous results.
  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect -- results belong to the previous inputs
    setSummary(null);
  }, [
    wordFile,
    excelFile,
    rules,
    startRow,
    endRow,
    fileNameColumn,
    seqSource,
    seqFormat,
    namePrefix,
    nameSuffix,
    selectedSheet,
    keepPlaceholderOnEmpty,
  ]);

  const missingColumns = useMemo(() => {
    if (!preview) return [];
    const columns = new Set(preview.excel_columns);
    const missing = new Set<string>();
    rules.forEach((rule) => {
      if (rule.excel_column && !columns.has(rule.excel_column)) missing.add(rule.excel_column);
    });
    if (fileNameColumn && !columns.has(fileNameColumn)) missing.add(fileNameColumn);
    return [...missing];
  }, [preview, rules, fileNameColumn]);

  const stepNow = useMemo(() => {
    if (!wordFile || !excelFile) return 1;
    if (!preview) return 2;
    if (rules.length === 0) return 3;
    if (!summary) return 4;
    return 5;
  }, [wordFile, excelFile, preview, rules.length, summary]);

  const canExecute = useMemo(() => {
    return (
      !!wordFile &&
      !!excelFile &&
      rules.length > 0 &&
      startRow > 0 &&
      endRow >= startRow &&
      fileNameColumn.trim().length > 0 &&
      missingColumns.length === 0
    );
  }, [wordFile, excelFile, rules.length, startRow, endRow, fileNameColumn, missingColumns]);

  const executeHint = useMemo(() => {
    if (!wordFile || !excelFile) return "请先上传 Word 与 Excel 文件";
    if (rules.length === 0) return "请先添加至少一条规则";
    if (missingColumns.length > 0) return `以下 Excel 列不存在：${missingColumns.join("、")}`;
    if (fileNameColumn.trim().length === 0) return "请填写文件名来源列";
    if (startRow <= 0 || endRow < startRow) return "请检查起始行与结束行";
    return "参数已完整，可开始替换";
  }, [wordFile, excelFile, rules.length, missingColumns, fileNameColumn, startRow, endRow]);

  function scrollToStep(anchor: string) {
    const target = document.getElementById(anchor);
    if (!target) return;
    target.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function onExecute(event: FormEvent) {
    event.preventDefault();
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
        rules: rules.map((rule) => ({ keyword: rule.keyword, excel_column: rule.excel_column })),
        useSeqPrefix,
        seqFormat,
        seqColumn,
        namePrefix,
        nameSuffix,
        sheetName: selectedSheet,
        keepPlaceholderOnEmpty,
      });
      setSummary(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : "执行失败");
    } finally {
      setExecuting(false);
    }
  }

  const displayError = error || previewError || configError;

  return (
    <main
      className="min-h-screen bg-[radial-gradient(ellipse_at_top,#e0f2fe_0%,#f8fafc_40%,#f8fafc_100%)] p-4 text-slate-900 md:p-8"
      style={{ fontSize: `${baseFontSize}px` }}
    >
      <ProgressSidebar stepNow={stepNow} onScrollToStep={scrollToStep} />
      <DisplaySettings
        contentWidth={contentWidth}
        baseFontSize={baseFontSize}
        setBaseFontSize={setBaseFontSize}
        updateContentWidth={updateContentWidth}
        resetDisplaySettings={resetDisplaySettings}
        showAdvanced={showAdvanced}
        onToggleAdvanced={() => setShowAdvanced((value) => !value)}
      />

      <div
        className="mx-auto space-y-5 ui-scale 2xl:px-48"
        style={{ maxWidth: `${contentWidth - 32}px`, ["--ui-scale" as never]: String(uiScale) }}
      >
        <section className="rounded-2xl border border-slate-200 bg-gradient-to-r from-slate-900 via-slate-800 to-cyan-900 p-6 text-slate-100 shadow-md">
          <div className="flex flex-col gap-4 xl:flex-row xl:items-end xl:justify-between">
            <div className="space-y-1">
              <p className="text-xs uppercase tracking-[0.18em] text-slate-300">WordReplace</p>
              <h1 className="text-2xl font-bold tracking-wide">文档批量替换助手</h1>
              <p className="text-sm text-slate-200">Word 模板 + Excel 数据，一键生成批量文档。</p>
            </div>
          </div>
        </section>

        <UploadSection
          wordFile={wordFile}
          setWordFile={setWordFile}
          excelFile={excelFile}
          setExcelFile={setExcelFile}
          preview={preview}
          previewLoading={previewLoading}
          previewWrap={previewWrap}
          setPreviewWrap={setPreviewWrap}
          sheetNames={preview?.sheet_names ?? []}
          selectedSheet={selectedSheet}
          setSelectedSheet={setSelectedSheet}
          maxUploadSizeMb={config.max_upload_size_mb}
          onError={setError}
        />

        <RulesSection
          preview={preview}
          rules={rules}
          setRules={setRules}
          onError={setError}
          showAdvanced={showAdvanced}
        />

        <ExecuteSection
          preview={preview}
          startRow={startRow}
          setStartRow={setStartRow}
          endRow={endRow}
          setEndRow={setEndRow}
          fileNameColumn={fileNameColumn}
          setFileNameColumn={setFileNameColumn}
          seqSource={seqSource}
          setSeqSource={setSeqSource}
          seqFormat={seqFormat}
          setSeqFormat={setSeqFormat}
          namePrefix={namePrefix}
          setNamePrefix={setNamePrefix}
          nameSuffix={nameSuffix}
          setNameSuffix={setNameSuffix}
          keepPlaceholderOnEmpty={keepPlaceholderOnEmpty}
          setKeepPlaceholderOnEmpty={setKeepPlaceholderOnEmpty}
          executing={executing}
          canExecute={canExecute}
          executeHint={executeHint}
          onSubmit={onExecute}
          summary={summary}
          setSummary={setSummary}
          onError={setError}
          showAdvanced={showAdvanced}
        />

        {displayError ? (
          <section
            role="alert"
            aria-live="polite"
            className="rounded-md border border-red-300 bg-red-50 px-3 py-2 text-sm text-red-700"
          >
            {displayError}
          </section>
        ) : null}

        <Footer />
      </div>
    </main>
  );
}
