"use client";

import { useMemo, useState } from "react";

import {
  deleteSingleResult,
  getExportUrl,
  getSingleFileExportUrl,
  type ReplaceSummary,
} from "@/lib/api";

type ResultsSectionProps = {
  summary: ReplaceSummary;
  setSummary: React.Dispatch<React.SetStateAction<ReplaceSummary | null>>;
  onError: (message: string) => void;
  showAdvanced: boolean;
};

type DetailFilter = "all" | "success" | "failed";

export default function ResultsSection({ summary, setSummary, onError, showAdvanced }: ResultsSectionProps) {
  const [detailFilter, setDetailFilter] = useState<DetailFilter>("all");
  const [deletingItemId, setDeletingItemId] = useState<string | null>(null);

  const filteredDetails = useMemo(() => {
    if (detailFilter === "success") return summary.details.filter((item) => item.status === "成功");
    if (detailFilter === "failed") return summary.details.filter((item) => item.status === "失败");
    return summary.details;
  }, [summary, detailFilter]);

  async function copyFailedMessages() {
    const failed = summary.details.filter((item) => item.status === "失败");
    if (failed.length === 0) return;
    const text = failed
      .map((item) => `序号${item.seq} 行${item.row_number} 文件${item.file_name} 原因: ${item.message || "-"}`)
      .join("\n");
    await navigator.clipboard.writeText(text);
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
                    <a
                      className="wr-btn wr-btn-sm inline-flex items-center justify-center"
                      href={getSingleFileExportUrl(summary.run_id, item.item_id, summary.export_token)}
                    >
                      下载
                    </a>
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

      <div className="flex gap-3">
        <a className="wr-btn inline-flex items-center" href={getExportUrl("zip", summary.run_id, summary.export_token)}>
          下载 ZIP
        </a>
        <a
          className="wr-btn inline-flex items-center"
          href={getExportUrl("merge", summary.run_id, summary.export_token)}
        >
          下载合并文档
        </a>
      </div>
    </>
  );
}
