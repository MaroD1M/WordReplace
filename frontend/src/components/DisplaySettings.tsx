import {
  MAX_CONTENT_WIDTH,
  MAX_FONT_SIZE,
  MIN_CONTENT_WIDTH,
  MIN_FONT_SIZE,
} from "@/lib/constants";

type DisplaySettingsProps = {
  contentWidth: number;
  baseFontSize: number;
  setBaseFontSize: (value: number) => void;
  updateContentWidth: (value: number) => void;
  resetDisplaySettings: () => void;
  showAdvanced: boolean;
  onToggleAdvanced: () => void;
};

export default function DisplaySettings({
  contentWidth,
  baseFontSize,
  setBaseFontSize,
  updateContentWidth,
  resetDisplaySettings,
  showAdvanced,
  onToggleAdvanced,
}: DisplaySettingsProps) {
  return (
    <div className="hidden 2xl:block fixed right-4 top-28 z-20 w-52">
      <div className="rounded-xl border border-slate-200 bg-white/95 p-3 shadow-sm space-y-3">
        <p className="text-xs font-semibold text-slate-500">显示设置</p>
        <label className="block text-xs">
          内容宽度：{contentWidth}px
          <input
            type="range"
            min={MIN_CONTENT_WIDTH}
            max={MAX_CONTENT_WIDTH}
            step={10}
            value={contentWidth}
            onChange={(event) => updateContentWidth(Number(event.target.value))}
            className="mt-1 w-full"
          />
        </label>
        <label className="block text-xs">
          字体大小：{baseFontSize}px
          <input
            type="range"
            min={MIN_FONT_SIZE}
            max={MAX_FONT_SIZE}
            step={1}
            value={baseFontSize}
            onChange={(event) => setBaseFontSize(Number(event.target.value))}
            className="mt-1 w-full"
          />
        </label>
        <button type="button" className="wr-btn w-full wr-btn-sm" onClick={resetDisplaySettings}>
          恢复默认
        </button>
        <div className="border-t border-slate-200 pt-2">
          <p className="mb-1 text-xs font-semibold text-slate-500">页面选项</p>
          <button type="button" className="wr-btn w-full wr-btn-sm" onClick={onToggleAdvanced}>
            {showAdvanced ? "隐藏高级项" : "显示高级项"}
          </button>
        </div>
      </div>
    </div>
  );
}
