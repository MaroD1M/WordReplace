"use client";

import { useEffect, useRef, useState } from "react";

import {
  DEFAULT_CONTENT_WIDTH,
  DEFAULT_FONT_SIZE,
  MAX_CONTENT_WIDTH,
  MAX_FONT_SIZE,
  MIN_CONTENT_WIDTH,
  MIN_FONT_SIZE,
} from "@/lib/constants";

export function useDisplaySettings() {
  const [contentWidth, setContentWidth] = useState(DEFAULT_CONTENT_WIDTH);
  const [baseFontSize, setBaseFontSize] = useState(DEFAULT_FONT_SIZE);
  const frameRef = useRef<number | null>(null);

  useEffect(() => {
    try {
      const savedWidth = localStorage.getItem("wr_content_width");
      const savedFont = localStorage.getItem("wr_font_size");
      if (savedWidth) {
        const width = Math.min(MAX_CONTENT_WIDTH, Math.max(MIN_CONTENT_WIDTH, Number(savedWidth)));
        // eslint-disable-next-line react-hooks/set-state-in-effect -- hydrate persisted settings after mount
        if (!Number.isNaN(width)) setContentWidth(width);
      }
      if (savedFont) {
        const font = Math.min(MAX_FONT_SIZE, Math.max(MIN_FONT_SIZE, Number(savedFont)));
        if (!Number.isNaN(font)) setBaseFontSize(font);
      }
    } catch {
      // ignore localStorage errors
    }
  }, []);

  useEffect(() => {
    try {
      localStorage.setItem("wr_content_width", String(contentWidth));
      localStorage.setItem("wr_font_size", String(baseFontSize));
    } catch {
      // ignore localStorage errors
    }
  }, [contentWidth, baseFontSize]);

  function resetDisplaySettings() {
    setContentWidth(DEFAULT_CONTENT_WIDTH);
    setBaseFontSize(DEFAULT_FONT_SIZE);
  }

  function updateContentWidth(value: number) {
    if (frameRef.current) cancelAnimationFrame(frameRef.current);
    frameRef.current = requestAnimationFrame(() => {
      setContentWidth(value);
      frameRef.current = null;
    });
  }

  return {
    contentWidth,
    baseFontSize,
    setBaseFontSize,
    resetDisplaySettings,
    updateContentWidth,
    uiScale: baseFontSize / DEFAULT_FONT_SIZE,
  };
}
