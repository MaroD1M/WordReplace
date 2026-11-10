export default function Section({
  title,
  hint,
  children,
}: {
  title: string;
  hint: string;
  children: React.ReactNode;
}) {
  return (
    <section className="rounded-xl border border-slate-200 bg-white/95 p-5 shadow-sm space-y-3">
      <h2 className="flex items-center gap-2 text-base font-semibold tracking-wide text-slate-900">
        <span className="h-4 w-1 rounded bg-cyan-600" />
        {title}
        <span
          className="ml-1 cursor-help rounded-full border border-slate-300 px-1.5 text-xs text-slate-500"
          title={hint}
        >
          i
        </span>
      </h2>
      {children}
    </section>
  );
}
