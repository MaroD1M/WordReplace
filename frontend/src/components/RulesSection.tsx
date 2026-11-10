"use client";

import { FormEvent, useEffect, useState } from "react";

import {
  applyTemplate,
  createTemplate,
  fetchTemplateDetail,
  fetchTemplates,
  type PreviewData,
  type Rule,
  type RuleTemplateListItem,
} from "@/lib/api";
import ColumnSelect from "@/components/ColumnSelect";
import Section from "@/components/Section";

type RulesSectionProps = {
  preview: PreviewData | null;
  rules: Rule[];
  setRules: React.Dispatch<React.SetStateAction<Rule[]>>;
  onError: (message: string) => void;
  showAdvanced: boolean;
};

export default function RulesSection({ preview, rules, setRules, onError, showAdvanced }: RulesSectionProps) {
  const [keyword, setKeyword] = useState("");
  const [excelColumn, setExcelColumn] = useState("");
  const [editingRuleId, setEditingRuleId] = useState<number | null>(null);
  const [editingKeyword, setEditingKeyword] = useState("");
  const [editingExcelColumn, setEditingExcelColumn] = useState("");

  const [templates, setTemplates] = useState<RuleTemplateListItem[]>([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState<number | null>(null);
  const [templateCreator, setTemplateCreator] = useState("");
  const [templateName, setTemplateName] = useState("");
  const [templateDescription, setTemplateDescription] = useState("");
  const [templateMode, setTemplateMode] = useState<"replace" | "append">("replace");
  const [templateInfo, setTemplateInfo] = useState("");

  const excelColumns = preview?.excel_columns ?? [];
  const columnsReady = excelColumns.length > 0;

  useEffect(() => {
    fetchTemplates()
      .then(setTemplates)
      .catch((err) => onError(err instanceof Error ? err.message : "加载模板库失败"));
  }, [onError]);

  function onSubmitRule(event: FormEvent) {
    event.preventDefault();
    if (!keyword.trim() || !excelColumn.trim()) return;
    const newRule: Rule = { id: Date.now(), keyword: keyword.trim(), excel_column: excelColumn.trim() };
    setRules((prev) => [...prev, newRule]);
    onError("");
    setKeyword("");
    setExcelColumn("");
  }

  function onDelete(id: number) {
    onError("");
    setRules((prev) => prev.filter((rule) => rule.id !== id));
  }

  function onStartEdit(rule: Rule) {
    setEditingRuleId(rule.id);
    setEditingKeyword(rule.keyword);
    setEditingExcelColumn(rule.excel_column);
  }

  function onSaveEdit() {
    if (!editingRuleId || !editingKeyword.trim() || !editingExcelColumn.trim()) return;
    onError("");
    setRules((prev) =>
      prev.map((rule) =>
        rule.id === editingRuleId
          ? { ...rule, keyword: editingKeyword.trim(), excel_column: editingExcelColumn.trim() }
          : rule,
      ),
    );
    setEditingRuleId(null);
    setEditingKeyword("");
    setEditingExcelColumn("");
  }

  async function onSaveTemplate() {
    if (!templateName.trim()) {
      onError("请输入模板名称");
      return;
    }
    if (rules.length === 0) {
      onError("当前没有可保存的规则");
      return;
    }
    onError("");
    try {
      await createTemplate({
        name: templateName.trim(),
        creator: templateCreator.trim() || "未填写",
        description: templateDescription.trim(),
        items: rules.map((rule) => ({ keyword: rule.keyword, excel_column: rule.excel_column })),
      });
      setTemplateInfo("模板保存成功");
      setTemplates(await fetchTemplates());
    } catch (err) {
      onError(err instanceof Error ? err.message : "保存模板失败");
    }
  }

  async function onApplyTemplate() {
    if (!selectedTemplateId) {
      onError("请先选择模板");
      return;
    }
    if (!preview) {
      onError("请先上传并预览 Excel");
      return;
    }
    onError("");
    try {
      const result = await applyTemplate(selectedTemplateId, preview.excel_columns, templateMode);
      const detail = await fetchTemplateDetail(selectedTemplateId);
      const mapped = result.valid_items.map((item, idx) => ({
        id: Number(`${Date.now()}${idx}`),
        keyword: item.keyword,
        excel_column: item.excel_column,
      }));
      setRules((prev) => (templateMode === "replace" ? mapped : [...prev, ...mapped]));
      setTemplateInfo(
        `模板「${detail.name}」已应用：有效 ${result.valid_items.length} 条，无效 ${result.invalid_items.length} 条（已忽略）`,
      );
    } catch (err) {
      onError(err instanceof Error ? err.message : "应用模板失败");
    }
  }

  return (
    <>
      <div id="step-rules" className="scroll-mt-24" />
      <Section title="2) 规则管理" hint="配置模板关键字和 Excel 列名映射。">
        {showAdvanced ? (
          <div className="wr-panel bg-cyan-50/50 space-y-3">
            <div className="grid grid-cols-1 gap-3 xl:grid-cols-4">
              <input
                className="wr-input"
                placeholder="模板名称"
                value={templateName}
                onChange={(event) => setTemplateName(event.target.value)}
              />
              <input
                className="wr-input"
                placeholder="创建人（例如：张三）"
                value={templateCreator}
                onChange={(event) => setTemplateCreator(event.target.value)}
              />
              <input
                className="wr-input"
                placeholder="模板描述（可选）"
                value={templateDescription}
                onChange={(event) => setTemplateDescription(event.target.value)}
              />
              <button type="button" className="wr-btn" onClick={onSaveTemplate}>
                保存为模板
              </button>
            </div>
            <div className="grid grid-cols-1 gap-3 xl:grid-cols-4">
              <select
                className="wr-input"
                value={selectedTemplateId ?? ""}
                onChange={(event) => setSelectedTemplateId(event.target.value ? Number(event.target.value) : null)}
              >
                <option value="">选择模板库规则</option>
                {templates.map((template) => (
                  <option key={template.id} value={template.id}>
                    {template.name} · {template.creator} · {template.item_count}条
                  </option>
                ))}
              </select>
              <select
                className="wr-input"
                value={templateMode}
                onChange={(event) => setTemplateMode(event.target.value as "replace" | "append")}
              >
                <option value="replace">覆盖当前规则</option>
                <option value="append">追加到当前规则</option>
              </select>
              <button type="button" className="wr-btn" onClick={onApplyTemplate}>
                应用模板
              </button>
              <div className="flex items-center text-xs text-slate-500">
                {templateInfo || "模板会自动校验 Excel 列，不匹配规则将忽略"}
              </div>
            </div>
          </div>
        ) : null}

        <form onSubmit={onSubmitRule} className="grid grid-cols-1 gap-3 xl:grid-cols-3">
          <input
            className="wr-input"
            placeholder="模板关键字（如【姓名】）"
            value={keyword}
            onChange={(event) => setKeyword(event.target.value)}
          />
          <ColumnSelect
            value={excelColumn}
            onChange={setExcelColumn}
            options={excelColumns}
            placeholder="请选择或输入 Excel 列名"
            disabled={!columnsReady}
            ariaLabel="选择规则使用的 Excel 列"
          />
          <button className="wr-btn wr-btn-primary disabled:opacity-50" disabled={!columnsReady}>
            添加规则
          </button>
        </form>

        <div className="overflow-x-auto rounded-md border border-slate-300">
          <table className="w-full table-fixed text-sm" aria-label="替换规则">
            <thead className="bg-slate-200">
              <tr>
                <th scope="col" className="w-20 px-3 py-2 text-center">序号</th>
                <th scope="col" className="px-3 py-2 text-center">模板关键字</th>
                <th scope="col" className="px-3 py-2 text-center">Excel 列名</th>
                <th scope="col" className="w-44 px-3 py-2 text-center">操作</th>
              </tr>
            </thead>
            <tbody>
              {rules.length === 0 ? (
                <tr>
                  <td className="px-3 py-2.5 text-center" colSpan={4}>
                    暂无规则
                  </td>
                </tr>
              ) : (
                rules.map((rule, idx) => (
                  <tr key={rule.id} className="border-t border-slate-200">
                    <td className="px-3 py-2.5 text-center">{idx + 1}</td>
                    <td className="px-3 py-2.5 text-center">
                      {editingRuleId === rule.id ? (
                        <input
                          className="wr-input h-8 px-2 text-center"
                          value={editingKeyword}
                          onChange={(event) => setEditingKeyword(event.target.value)}
                        />
                      ) : (
                        rule.keyword
                      )}
                    </td>
                    <td className="px-3 py-2.5 text-center">
                      {editingRuleId === rule.id ? (
                        <ColumnSelect
                          value={editingExcelColumn}
                          onChange={setEditingExcelColumn}
                          options={excelColumns}
                          inputClassName="wr-input h-8 px-2 text-center"
                          ariaLabel="编辑规则使用的 Excel 列"
                        />
                      ) : (
                        rule.excel_column
                      )}
                    </td>
                    <td className="px-3 py-2.5 text-center">
                      {editingRuleId === rule.id ? (
                        <div className="flex flex-nowrap items-center justify-center gap-2">
                          <button type="button" className="wr-btn wr-btn-sm" onClick={onSaveEdit}>
                            保存
                          </button>
                          <button
                            type="button"
                            className="wr-btn wr-btn-sm"
                            onClick={() => {
                              setEditingRuleId(null);
                              setEditingKeyword("");
                              setEditingExcelColumn("");
                            }}
                          >
                            取消
                          </button>
                        </div>
                      ) : (
                        <div className="flex flex-nowrap items-center justify-center gap-2">
                          <button type="button" className="wr-btn wr-btn-sm" onClick={() => onStartEdit(rule)}>
                            编辑
                          </button>
                          <button type="button" className="wr-btn wr-btn-sm" onClick={() => onDelete(rule.id)}>
                            删除
                          </button>
                        </div>
                      )}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </Section>
    </>
  );
}
