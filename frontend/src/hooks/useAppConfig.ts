"use client";

import { useEffect, useState } from "react";

import { fetchAppConfig, type AppConfig } from "@/lib/api";
import { APP_VERSION, MAX_UPLOAD_SIZE_MB } from "@/lib/constants";

export const FALLBACK_APP_CONFIG: AppConfig = {
  app_version: APP_VERSION,
  max_upload_size_mb: MAX_UPLOAD_SIZE_MB,
  max_excel_rows: 20000,
  max_replace_rules: 500,
  max_preview_rows: 25,
  run_cache_ttl_seconds: 1800,
};

export function useAppConfig() {
  const [config, setConfig] = useState<AppConfig>(FALLBACK_APP_CONFIG);
  const [configError, setConfigError] = useState("");

  useEffect(() => {
    let active = true;
    fetchAppConfig()
      .then((data) => {
        if (!active) return;
        setConfig({ ...FALLBACK_APP_CONFIG, ...data });
        setConfigError("");
      })
      .catch((err) => {
        if (!active) return;
        setConfigError(err instanceof Error ? err.message : "加载应用配置失败");
      });
    return () => {
      active = false;
    };
  }, []);

  return { config, configError };
}
