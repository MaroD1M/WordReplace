import type { SeqFormat } from "@/lib/api";

export const DEFAULT_CONTENT_WIDTH = 1400;
export const DEFAULT_FONT_SIZE = 15;
export const MAX_CONTENT_WIDTH = 1800;
export const MIN_CONTENT_WIDTH = 1100;
export const MAX_FONT_SIZE = 18;
export const MIN_FONT_SIZE = 14;

export const APP_VERSION = process.env.NEXT_PUBLIC_APP_VERSION ?? "2.0.5";

export const SEQ_FORMATS: Array<{ value: SeqFormat; label: string }> = [
  { value: "1", label: "1, 2, 3" },
  { value: "01", label: "01, 02, 03" },
  { value: "0001", label: "0001, 0002" },
  { value: "1.", label: "1., 2., 3." },
  { value: "一", label: "一, 二, 三" },
];

export const STEP_ANCHORS = [
  { label: "上传文件", key: "step-upload" },
  { label: "自动预览", key: "step-preview" },
  { label: "规则管理", key: "step-rules" },
  { label: "执行替换", key: "step-execute" },
  { label: "导出结果", key: "step-export" },
] as const;
