const API_BASE = process.env.NEXT_PUBLIC_API_BASE ?? "";

export type SeqFormat = "1" | "01" | "0001" | "1." | "一";

export type AppConfig = {
  app_version: string;
  max_upload_size_mb: number;
  max_excel_rows: number;
  max_replace_rules: number;
  max_preview_rows: number;
  run_cache_ttl_seconds: number;
};

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
  sheet_names: string[];
  selected_sheet: string;
};

/**
 * FastAPI returns errors as {"detail": "..."} (or a list of validation
 * objects). Parse that instead of showing raw JSON to the user.
 */
export async function parseApiError(res: Response, fallback: string): Promise<string> {
  try {
    const text = await res.text();
    if (!text) return fallback;
    try {
      const data = JSON.parse(text) as { detail?: unknown };
      if (typeof data.detail === "string") return data.detail;
      if (Array.isArray(data.detail)) {
        return data.detail
          .map((item) => (typeof item === "object" && item && "msg" in item ? String(item.msg) : JSON.stringify(item)))
          .join("；");
      }
    } catch {
      // not JSON, fall through to raw text
    }
    return text;
  } catch {
    return fallback;
  }
}

export async function fetchAppConfig(): Promise<AppConfig> {
  const res = await fetch(`${API_BASE}/config`, { cache: "no-store" });
  if (!res.ok) throw new Error(await parseApiError(res, "加载应用配置失败"));
  return res.json();
}

export async function fetchTemplates(): Promise<RuleTemplateListItem[]> {
  const res = await fetch(`${API_BASE}/rule-templates`, { cache: "no-store" });
  if (!res.ok) throw new Error(await parseApiError(res, "加载模板库失败"));
  return res.json();
}

export async function fetchTemplateDetail(templateId: number): Promise<RuleTemplateDetail> {
  const res = await fetch(`${API_BASE}/rule-templates/${templateId}`, { cache: "no-store" });
  if (!res.ok) throw new Error(await parseApiError(res, "加载模板详情失败"));
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
  if (!res.ok) throw new Error(await parseApiError(res, "创建模板失败"));
  return res.json();
}

export async function applyTemplate(
  templateId: number,
  excelColumns: string[],
  mode: "replace" | "append",
): Promise<{
  valid_items: Array<{ keyword: string; excel_column: string }>;
  invalid_items: Array<{ keyword: string; excel_column: string }>;
}> {
  const res = await fetch(`${API_BASE}/rule-templates/${templateId}/apply`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ excel_columns: excelColumns, mode }),
  });
  if (!res.ok) throw new Error(await parseApiError(res, "应用模板失败"));
  return res.json();
}

export async function previewFiles(
  wordFile: File,
  excelFile: File,
  signal?: AbortSignal,
  sheetName = "",
): Promise<PreviewData> {
  const form = new FormData();
  form.append("word_file", wordFile);
  form.append("excel_file", excelFile);
  form.append("sheet_name", sheetName);
  const res = await fetch(`${API_BASE}/preview`, { method: "POST", body: form, signal });
  if (!res.ok) throw new Error(await parseApiError(res, "预览失败"));
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
  seqFormat: SeqFormat;
  seqColumn: string;
  namePrefix: string;
  nameSuffix: string;
  sheetName: string;
  keepPlaceholderOnEmpty: boolean;
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
  form.append("sheet_name", payload.sheetName);
  form.append("keep_placeholder_on_empty", String(payload.keepPlaceholderOnEmpty));

  const res = await fetch(`${API_BASE}/replace/execute`, { method: "POST", body: form });
  if (!res.ok) throw new Error(await parseApiError(res, "执行替换失败"));
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

export async function deleteSingleResult(
  runId: string,
  itemId: string,
  token: string,
): Promise<Pick<ReplaceSummary, "total" | "success" | "failed" | "replacements" | "details"> & { deleted: boolean }> {
  const q = new URLSearchParams({ token }).toString();
  const res = await fetch(`${API_BASE}/export/result/${runId}/${itemId}?${q}`, { method: "DELETE" });
  if (!res.ok) throw new Error(await parseApiError(res, "删除结果失败"));
  return res.json();
}
