"use client";

import { useEffect, useState } from "react";

import { previewFiles, type PreviewData } from "@/lib/api";

function fileSignature(wordFile: File | null, excelFile: File | null, sheetName: string): string {
  if (!wordFile || !excelFile) return "";
  const describe = (file: File) => `${file.name}:${file.size}:${file.lastModified}`;
  return `${describe(wordFile)}|${describe(excelFile)}|${sheetName}`;
}

export function usePreview(wordFile: File | null, excelFile: File | null, sheetName = "") {
  const [preview, setPreview] = useState<PreviewData | null>(null);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [previewError, setPreviewError] = useState("");
  const [loadedSignature, setLoadedSignature] = useState("");

  const signature = fileSignature(wordFile, excelFile, sheetName);
  const hasFiles = signature.length > 0;

  useEffect(() => {
    if (!wordFile || !excelFile) return;

    const controller = new AbortController();
    let active = true;

    const timer = setTimeout(async () => {
      setPreviewLoading(true);
      setPreviewError("");
      try {
        const data = await previewFiles(wordFile, excelFile, controller.signal, sheetName);
        if (!active) return;
        setPreview(data);
        setLoadedSignature(fileSignature(wordFile, excelFile, sheetName));
      } catch (err) {
        if (!active || controller.signal.aborted) return;
        setPreviewError(err instanceof Error ? err.message : "预览失败");
      } finally {
        if (active) setPreviewLoading(false);
      }
    }, 250);

    return () => {
      active = false;
      clearTimeout(timer);
      controller.abort();
    };
  }, [wordFile, excelFile, sheetName]);

  return {
    preview: hasFiles && loadedSignature === signature ? preview : null,
    previewLoading: hasFiles ? previewLoading : false,
    previewError: hasFiles ? previewError : "",
  };
}
