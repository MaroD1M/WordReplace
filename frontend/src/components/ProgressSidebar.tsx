import { STEP_ANCHORS } from "@/lib/constants";

export default function ProgressSidebar({
  stepNow,
  onScrollToStep,
}: {
  stepNow: number;
  onScrollToStep: (anchor: string) => void;
}) {
  return (
    <div className="hidden xl:block fixed left-4 top-28 z-20 w-40">
      <div className="rounded-xl border border-slate-200 bg-white/95 p-3 shadow-sm">
        <p className="mb-2 text-xs font-semibold text-slate-500">流程进度</p>
        {STEP_ANCHORS.map((step, idx) => {
          const i = idx + 1;
          const done = i < stepNow;
          const active = i === stepNow;
          return (
            <button
              key={step.key}
              type="button"
              onClick={() => onScrollToStep(step.key)}
              className="flex w-full items-center gap-2 rounded px-1 py-1 text-left hover:bg-slate-50"
            >
              <span
                className={`h-2.5 w-2.5 rounded-full ${
                  done ? "bg-emerald-500" : active ? "bg-cyan-600" : "bg-slate-300"
                }`}
              />
              <span className={`text-xs ${active ? "font-semibold text-slate-800" : "text-slate-500"}`}>
                {step.label}
              </span>
            </button>
          );
        })}
      </div>
    </div>
  );
}
