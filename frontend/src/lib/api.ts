const API_BASE = process.env.NEXT_PUBLIC_API_BASE ?? "";

export type Rule = {
  id: number;
  keyword: string;
  excel_column: string;
};

export type ReplaceSummary = {
  run_id: string;
  total: number;
  success: number;
  failed: number;
  replacements: number;
};

export async function fetchRules(): Promise<Rule[]> {
  const res = await fetch(`${API_BASE}/rules`, { cache: "no-store" });
  if (!res.ok) throw new Error("加载规则失败");
  return res.json();
}

export async function createRule(payload: { keyword: string; excel_column: string }): Promise<Rule> {
  const res = await fetch(`${API_BASE}/rules`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) throw new Error("创建规则失败");
  return res.json();
}

export async function deleteRule(ruleId: number): Promise<void> {
  const res = await fetch(`${API_BASE}/rules/${ruleId}`, { method: "DELETE" });
  if (!res.ok) throw new Error("删除规则失败");
}

export async function executeReplace(payload: {
  wordFile: File;
  excelFile: File;
  startRow: number;
  endRow: number;
  fileNameColumn: string;
  exportMode: "zip" | "merge";
}): Promise<ReplaceSummary> {
  const form = new FormData();
  form.append("word_file", payload.wordFile);
  form.append("excel_file", payload.excelFile);
  form.append("start_row", String(payload.startRow));
  form.append("end_row", String(payload.endRow));
  form.append("file_name_column", payload.fileNameColumn);
  form.append("export_mode", payload.exportMode);

  const res = await fetch(`${API_BASE}/replace/execute`, { method: "POST", body: form });
  if (!res.ok) {
    const msg = await res.text();
    throw new Error(msg || "执行替换失败");
  }
  return res.json();
}

export function getExportUrl(kind: "zip" | "merge", runId: string): string {
  return `${API_BASE}/export/${kind}/${runId}`;
}
