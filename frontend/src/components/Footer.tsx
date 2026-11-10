import { APP_VERSION } from "@/lib/constants";

export default function Footer() {
  return (
    <footer className="flex flex-wrap items-center justify-between gap-3 rounded-xl border border-slate-200 bg-slate-900 px-4 py-3 text-sm text-slate-300 shadow-sm">
      <div className="flex items-center gap-2">
        <span className="font-medium text-white">文档批量替换助手</span>
        <span className="rounded bg-slate-700 px-2 py-0.5 text-xs text-slate-100">v{APP_VERSION}</span>
      </div>
      <a
        href="https://github.com/MaroD1M/WordReplace"
        target="_blank"
        rel="noreferrer"
        className="inline-flex h-9 items-center gap-2 rounded-md border border-slate-600 bg-slate-800/80 px-3 text-sm text-slate-300 hover:bg-slate-700"
        aria-label="打开项目 GitHub 仓库"
      >
        <svg viewBox="0 0 24 24" aria-hidden="true" className="h-4 w-4 fill-current">
          <path d="M12 .5a12 12 0 0 0-3.79 23.38c.6.1.82-.26.82-.58v-2.03c-3.34.73-4.04-1.61-4.04-1.61-.55-1.4-1.34-1.78-1.34-1.78-1.1-.75.08-.74.08-.74 1.21.09 1.85 1.25 1.85 1.25 1.08 1.86 2.83 1.32 3.52 1.01.11-.79.42-1.32.76-1.63-2.67-.31-5.47-1.34-5.47-5.96 0-1.32.47-2.4 1.24-3.24-.13-.31-.54-1.57.12-3.27 0 0 1.01-.32 3.3 1.24a11.5 11.5 0 0 1 6 0c2.28-1.56 3.29-1.24 3.29-1.24.67 1.7.26 2.96.13 3.27.77.84 1.24 1.92 1.24 3.24 0 4.63-2.8 5.65-5.48 5.95.43.37.82 1.11.82 2.24v3.32c0 .32.22.69.83.57A12 12 0 0 0 12 .5z" />
        </svg>
        GitHub
      </a>
    </footer>
  );
}
