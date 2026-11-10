"use client";

import { useEffect, useMemo, useState } from "react";

import {
  deleteSingleResult,
  getExportUrl,
  getSingleFileExportUrl,
  parseApiError,
  type ReplaceSummary,
} from "@/lib/api";
import { RESULT_PAGE_SIZE } from "@/lib/constants";

type ResultsSectionProps = {
  summary: ReplaceSummary;
  setSummary: React.Dispatch<React.SetStateAction<ReplaceSummary | null>>;
  onError: (message: string) => void;
  showAdvanced: boolean;
};

type DetailFilter = "all" | "success" | "failed";

async function copyText(text: string): Promise<void> {
  try {
    await navigator.clipboard.writeText(text);
    return;
  } catch {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    textarea.remove();
  }
}

async function downloadFile(url: string, filename: string, onError: (message: string) => void) {
  try {
    const response = await fetch(url);
    if (!response.ok) {
      onError(await parseApiError(response, "下载失败"));
      return;
    }
    const blob = await response.blob();
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = objectUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(objectUrl);
  } catch (err) {
    onError(err instanceof Error ? err.message : "下载失败");
  }
}

export default function ResultsSection({ summary, setSummary, onError, showAdvanced }: ResultsSectionProps) {
  const [detailFilter, setDetailFilter] = useState<DetailFilter>("all");
  const [deletingItemId, setDeletingItemId] = useState<string | null>(null);
  const [page, setPage] = useState(0);

  useEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect -- a new run replaces the previous result set
    setPage(0);
  }, [summary.run_id]);

  const filteredDetails = useMemo(() => {
    if (detailFilter === "success") return summary.details.filter((item) => item.status === "成功");
    if (detailFilter === "failed") return summary.details.filter((item) => item.status === "失败");
    return summary.details;
  }, [summary, detailFilter]);

  const totalPages = Math.max(1, Math.ceil(filteredDetails.length / RESULT_PAGE_SIZE));
  const currentPage = Math.min(page, totalPages - 1);
  const pageItems = filteredDetails.slice(
    currentPage * RESULT_PAGE_SIZE,
    currentPage * RESULT_PAGE_SIZE + RESULT_PAGE_SIZE,
  );

  async function copyFailedMessages() {
    const failed = summary.details.filter((item) => item.status === "失败");
    if (failed.length === 0) return;
    const text = failed
      .map((item) => `序号${item.seq} 行${item.row_number} 文件${item.file_name} 原因: ${item.message || "-"}`)
      .join("\n");
    try {
      await copyText(text);
    } catch {
      onError("复制失败，请手动选择文本复制");
    }
  }

  async function onDeleteResult(itemId: string) {
    onError("");
    setDeletingItemId(itemId);
    try {
      const updated = await deleteSingleResult(summary.run_id, itemId, summary.export_token);
      setSummary((prev) => (prev ? { ...prev, ...updated } : prev));
    } catch (err) {
      onError(err instanceof Error ? err.message : "删除结果失败");
    } finally {
      setDeletingItemId(null);
    }
  }

  return (
    <>
      <div id="step-export" className="scroll-mt-24" />
      <div className="overflow-x-auto rounded-md border border-slate-300">
        <table className="w-full text-sm" aria-label="执行结果">
          <thead className="bg-slate-200">
            <tr>
              <th scope="col" className="px-3 py-2 text-center">总数</th>
              <th scope="col" className="px-3 py-2 text-center">成功</th>
              <th scope="col" className="px-3 py-2 text-center">失败</th>
              <th scope="col" className="px-3 py-2 text-center">替换次数</th>
            </tr>
          </thead>
          <tbody>
            <tr className="border-t border-slate-200">
              <td className="px-3 py-2.5 text-center">{summary.total}</td>
              <td className="px-3 py-2.5 text-center">{summary.success}</td>
              <td className="px-3 py-2.5 text-center">{summary.failed}</td>
              <td className="px-3 py-2.5 text-center">{summary.replacements}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <div className="flex flex-wrap items-center gap-2">
        <span className="text-sm text-slate-600">明细筛选：</span>
        {showAdvanced ? (
          <>
            <button type="button" className="wr-btn wr-btn-sm" onClick={() => setDetailFilter("all")}>
              全部
            </button>
            <button type="button" className="wr-btn wr-btn-sm" onClick={() => setDetailFilter("success")}>
              仅成功
            </button>
            <button type="button" className="wr-btn wr-btn-sm" onClick={() => setDetailFilter("failed")}>
              仅失败
            </button>
            <button type="button" className="wr-btn wr-btn-sm" onClick={copyFailedMessages}>
              复制失败原因
            </button>
          </>
        ) : (
          <span className="text-xs text-slate-500">开启“显示高级项”后可筛选明细与复制失败原因。</span>
        )}
      </div>

      <div className="flex flex-wrap items-center justify-between gap-2 text-xs text-slate-600">
        <span>
          共 {filteredDetails.length} 条，第 {currentPage + 1} / {totalPages} 页
        </span>
        <div className="flex gap-2">
          <button
            type="button"
            className="wr-btn wr-btn-sm disabled:opacity-50"
            disabled={currentPage === 0}
            onClick={() => setPage(Math.max(0, currentPage - 1))}
          >
            上一页
          </button>
          <button
            type="button"
            className="wr-btn wr-btn-sm disabled:opacity-50"
            disabled={currentPage >= totalPages - 1}
            onClick={() => setPage(Math.min(totalPages - 1, currentPage + 1))}
          >
            下一页
          </button>
        </div>
      </div>

      <div className="max-h-[420px] overflow-auto rounded-md border border-slate-300">
        <table className="w-full text-sm" aria-label="执行结果明细">
          <thead className="sticky top-0 z-10 bg-slate-200">
            <tr>
              <th scope="col" className="w-16 px-3 py-2 text-center">序号</th>
              <th scope="col" className="w-24 px-3 py-2 text-center">Excel行号</th>
              <th scope="col" className="px-3 py-2 text-center">生成文件名</th>
              <th scope="col" className="w-20 px-3 py-2 text-center">状态</th>
              <th scope="col" className="w-24 px-3 py-2 text-center">替换次数</th>
              <th scope="col" className="w-28 px-3 py-2 text-center">操作</th>
              <th scope="col" className="px-3 py-2 text-center">备注</th>
            </tr>
          </thead>
          <tbody>
            {pageItems.map((item) => (
              <tr key={item.item_id} className="border-t border-slate-200">
                <td className="px-3 py-2.5 text-center">{item.seq}</td>
                <td className="px-3 py-2.5 text-center">{item.row_number}</td>
                <td className="px-3 py-2.5 text-center">{item.file_name}</td>
                <td className="px-3 py-2.5 text-center">
                  <span
                    className={
                      item.status === "成功"
                        ? "rounded-full bg-emerald-50 px-2 py-0.5 text-emerald-700"
                        : "rounded-full bg-red-50 px-2 py-0.5 text-red-700"
                    }
                  >
                    {item.status}
                  </span>
                </td>
                <td className="px-3 py-2.5 text-center">{item.replace_count}</td>
                <td className="px-3 py-2.5 text-center">
                  <div className="flex flex-nowrap items-center justify-center gap-2">
                    <button
                      type="button"
                      className="wr-btn wr-btn-sm"
                      onClick={() =>
                        downloadFile(
                          getSingleFileExportUrl(summary.run_id, item.item_id, summary.export_token),
                          item.file_name,
                          onError,
                        )
                      }
                    >
                      下载
                    </button>
                    <button
                      type="button"
                      disabled={deletingItemId === item.item_id}
                      className="wr-btn wr-btn-sm disabled:opacity-50"
                      onClick={() => onDeleteResult(item.item_id)}
                    >
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

      <div className="flex flex-wrap gap-3">
        <button
          type="button"
          className="wr-btn inline-flex items-center disabled:opacity-50"
          disabled={summary.total === 0}
          onClick={() =>
            downloadFile(getExportUrl("zip", summary.run_id, summary.export_token), "replace_results.zip", onError)
          }
        >
          下载 ZIP
        </button>
        <button
          type="button"
          className="wr-btn inline-flex items-center disabled:opacity-50"
          disabled={summary.total === 0}
          onClick={() =>
            downloadFile(getExportUrl("merge", summary.run_id, summary.export_token), "replace_merged.docx", onError)
          }
        >
          下载合并文档
        </button>
      </div>
    </>
  );
}
