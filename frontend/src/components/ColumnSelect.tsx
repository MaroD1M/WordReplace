"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";

type ColumnSelectProps = {
  value: string;
  onChange: (value: string) => void;
  options: string[];
  placeholder?: string;
  disabled?: boolean;
  inputClassName?: string;
  ariaLabel?: string;
};

type PopupRect = { left: number; top: number; width: number };

/**
 * Searchable column picker.
 *
 * Native <datalist> filters suggestions by the input's current value, so a
 * prefilled value makes the dropdown look like it only contains one option.
 * This component always lists every column on focus and renders its popup in
 * a portal, so it is never clipped by scroll containers or overlapping cards.
 */
export default function ColumnSelect({
  value,
  onChange,
  options,
  placeholder,
  disabled,
  inputClassName,
  ariaLabel,
}: ColumnSelectProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLDivElement>(null);
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [highlight, setHighlight] = useState(0);
  const [rect, setRect] = useState<PopupRect | null>(null);

  const filtered = useMemo(() => {
    const keyword = query.trim().toLowerCase();
    if (!keyword) return options;
    const matched = options.filter((option) => option.toLowerCase().includes(keyword));
    return matched.length > 0 ? matched : options;
  }, [options, query]);

  useEffect(() => {
    if (!open) return;

    const updateRect = () => {
      const el = inputRef.current;
      if (!el) return;
      const box = el.getBoundingClientRect();
      const width = Math.min(Math.max(box.width, 280), window.innerWidth - 16);
      const left = Math.max(8, Math.min(box.left, window.innerWidth - width - 8));
      setRect({ left, top: box.bottom + 4, width });
    };

    updateRect();
    window.addEventListener("scroll", updateRect, true);
    window.addEventListener("resize", updateRect);

    const onPointerDown = (event: MouseEvent) => {
      const target = event.target as Node;
      if (inputRef.current?.contains(target) || listRef.current?.contains(target)) return;
      setOpen(false);
    };
    document.addEventListener("mousedown", onPointerDown);

    return () => {
      window.removeEventListener("scroll", updateRect, true);
      window.removeEventListener("resize", updateRect);
      document.removeEventListener("mousedown", onPointerDown);
    };
  }, [open]);

  function commit(option: string) {
    onChange(option);
    setQuery("");
    setOpen(false);
    inputRef.current?.focus();
  }

  function handleKeyDown(event: React.KeyboardEvent<HTMLInputElement>) {
    if (event.key === "Escape") {
      setOpen(false);
      return;
    }
    if (!open && (event.key === "ArrowDown" || event.key === "Enter")) {
      setOpen(true);
      setQuery("");
      return;
    }
    if (event.key === "ArrowDown") {
      event.preventDefault();
      setHighlight((current) => Math.min(current + 1, filtered.length - 1));
    } else if (event.key === "ArrowUp") {
      event.preventDefault();
      setHighlight((current) => Math.max(current - 1, 0));
    } else if (event.key === "Enter" && open && filtered[highlight]) {
      event.preventDefault();
      commit(filtered[highlight]);
    }
  }

  return (
    <>
      <input
        ref={inputRef}
        className={inputClassName ?? "wr-input"}
        value={value}
        placeholder={placeholder}
        disabled={disabled}
        aria-label={ariaLabel}
        title={value}
        autoComplete="off"
        onFocus={() => {
          setQuery("");
          setHighlight(0);
          setOpen(true);
        }}
        onChange={(event) => {
          onChange(event.target.value);
          setQuery(event.target.value);
          setHighlight(0);
          setOpen(true);
        }}
        onKeyDown={handleKeyDown}
      />
      {open && rect && typeof document !== "undefined"
        ? createPortal(
            <div
              ref={listRef}
              role="listbox"
              style={{ position: "fixed", left: rect.left, top: rect.top, width: rect.width }}
              className="z-[100] max-h-60 overflow-auto rounded-md border border-slate-300 bg-white py-1 text-sm shadow-lg"
            >
              {filtered.length === 0 ? (
                <div className="px-3 py-2 text-slate-400">暂无可选列</div>
              ) : (
                filtered.map((option, index) => (
                  <button
                    key={option}
                    type="button"
                    role="option"
                    aria-selected={option === value}
                    title={option}
                    onMouseDown={(event) => event.preventDefault()}
                    onMouseEnter={() => setHighlight(index)}
                    onClick={() => commit(option)}
                    className={`block w-full truncate px-3 py-1.5 text-left ${
                      index === highlight ? "bg-cyan-50 text-cyan-800" : "text-slate-700 hover:bg-slate-50"
                    }`}
                  >
                    {option}
                  </button>
                ))
              )}
            </div>,
            document.body,
          )
        : null}
    </>
  );
}
