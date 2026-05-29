const API_BASE = process.env.NEXT_PUBLIC_API_BASE ?? "";

export type Rule = {
  id: number;
  keyword: string;
  excel_column: string;
};

export type RuleTemplateListItem = {
  id: number;
  name: string;
  creator: string;
  description: string;
  is_active: boolean;
  created_at_bj: string;
  updated_at_bj: string;
  item_count: number;
};

export type RuleTemplateDetail = {
  id: number;
  name: string;
  creator: string;
  description: string;
  is_active: boolean;
  created_at: string;
  updated_at: string;
  items: Array<{
    id: number;
    keyword: string;
    excel_column: string;
    order_index: number;
    is_valid: boolean;
  }>;
};

export type ReplaceSummary = {
  run_id: string;
  export_token: string;
  total: number;
  success: number;
  failed: number;
  replacements: number;
  details: {
    item_id: string;
    seq: number;
    row_number: number;
    file_name: string;
    status: string;
    replace_count: number;
    message: string;
  }[];
};

export type PreviewData = {
  word_text: string;
  excel_columns: string[];
  excel_rows: string[][];
  excel_total_rows: number;
};

export async function fetchRules(): Promise<Rule[]> {
  const res = await fetch(`${API_BASE}/rules`, { cache: "no-store" });
  if (!res.ok) throw new Error("加载规则失败");
  return res.json();
}

export async function fetchTemplates(): Promise<RuleTemplateListItem[]> {
  const res = await fetch(`${API_BASE}/rule-templates`, { cache: "no-store" });
  if (!res.ok) throw new Error("加载模板库失败");
  return res.json();
}

export async function fetchTemplateDetail(templateId: number): Promise<RuleTemplateDetail> {
  const res = await fetch(`${API_BASE}/rule-templates/${templateId}`, { cache: "no-store" });
  if (!res.ok) throw new Error("加载模板详情失败");
  return res.json();
}

export async function createTemplate(payload: {
  name: string;
  creator: string;
  description: string;
  items: Array<{ keyword: string; excel_column: string }>;
}): Promise<RuleTemplateDetail> {
  const res = await fetch(`${API_BASE}/rule-templates`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) throw new Error("创建模板失败");
  return res.json();
}

export async function applyTemplate(templateId: number, excelColumns: string[], mode: "replace" | "append"): Promise<{
  valid_items: Array<{ keyword: string; excel_column: string }>;
  invalid_items: Array<{ keyword: string; excel_column: string }>;
}> {
  const res = await fetch(`${API_BASE}/rule-templates/${templateId}/apply`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ excel_columns: excelColumns, mode }),
  });
  if (!res.ok) throw new Error("应用模板失败");
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

export async function updateRule(ruleId: number, payload: { keyword: string; excel_column: string }): Promise<Rule> {
  const res = await fetch(`${API_BASE}/rules/${ruleId}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) throw new Error("更新规则失败");
  return res.json();
}

export async function previewFiles(wordFile: File, excelFile: File): Promise<PreviewData> {
  const form = new FormData();
  form.append("word_file", wordFile);
  form.append("excel_file", excelFile);
  const res = await fetch(`${API_BASE}/preview`, { method: "POST", body: form });
  if (!res.ok) {
    const msg = await res.text();
    throw new Error(msg || "预览失败");
  }
  return res.json();
}

export async function executeReplace(payload: {
  wordFile: File;
  excelFile: File;
  startRow: number;
  endRow: number;
  fileNameColumn: string;
  rules: { keyword: string; excel_column: string }[];
  useSeqPrefix: boolean;
  seqFormat: "1" | "01" | "0001" | "1." | "一";
  seqColumn: string;
  namePrefix: string;
  nameSuffix: string;
}): Promise<ReplaceSummary> {
  const form = new FormData();
  form.append("word_file", payload.wordFile);
  form.append("excel_file", payload.excelFile);
  form.append("start_row", String(payload.startRow));
  form.append("end_row", String(payload.endRow));
  form.append("file_name_column", payload.fileNameColumn);
  form.append("export_mode", "zip");
  form.append("rules_json", JSON.stringify(payload.rules));
  form.append("use_seq_prefix", String(payload.useSeqPrefix));
  form.append("seq_format", payload.seqFormat);
  form.append("seq_column", payload.seqColumn);
  form.append("name_prefix", payload.namePrefix);
  form.append("name_suffix", payload.nameSuffix);

  const res = await fetch(`${API_BASE}/replace/execute`, { method: "POST", body: form });
  if (!res.ok) {
    const msg = await res.text();
    throw new Error(msg || "执行替换失败");
  }
  return res.json();
}

export function getExportUrl(kind: "zip" | "merge", runId: string, token: string): string {
  const q = new URLSearchParams({ token }).toString();
  return `${API_BASE}/export/${kind}/${runId}?${q}`;
}

export function getSingleFileExportUrl(runId: string, itemId: string, token: string): string {
  const q = new URLSearchParams({ token }).toString();
  return `${API_BASE}/export/file/${runId}/${itemId}?${q}`;
}

export async function deleteSingleResult(runId: string, itemId: string, token: string): Promise<Pick<ReplaceSummary, "total" | "success" | "failed" | "replacements" | "details"> & { deleted: boolean }> {
  const q = new URLSearchParams({ token }).toString();
  const res = await fetch(`${API_BASE}/export/result/${runId}/${itemId}?${q}`, { method: "DELETE" });
  if (!res.ok) {
    const msg = await res.text();
    throw new Error(msg || "删除结果失败");
  }
  return res.json();
}
