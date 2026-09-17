import { useEffect, useMemo, useRef, useState } from "react";
import {
  Alert,
  Button,
  Card,
  Col,
  Row,
  Segmented,
  Space,
  Tag,
  Typography,
} from "antd";
import { DownloadOutlined } from "@ant-design/icons";
import {
  exportExcel,
  StylePresets,
  type CellStyle,
  type SheetConfig,
} from "@marcusok/excel-exporter";
import {
  buildPreviewColumns,
  cellStyleToCss,
  formatByNumFormat,
  previewCellValue,
} from "./style-preview.js";
import {
  PRESET_CHEATSHEET,
  STYLE_SCENARIOS,
  type ScenarioKey,
  type StyleScenario,
} from "./style-scenarios.js";
import { formatBytes } from "./metrics.js";

/**
 * 预览表格与代码块的样式：作用域类前缀避免与 play.css 全局样式互相干扰。
 * th/td 显式 reset UA 默认的粗体与居中，让单元格观感完全由 cellStyleToCss
 * 的内联样式驱动；nowrap 是 Excel 单元格的默认行为，wrapText 样式会覆盖它。
 * 深色主题下预览仍保持白底——模拟的是 Excel 画布，刻意不受站点主题影响。
 */
const PREVIEW_CSS = `
.xlsx-scroll { overflow-x: auto; }
.xlsx-preview {
  border-collapse: collapse;
  font-family: Calibri, "Microsoft YaHei", system-ui, sans-serif;
  font-size: 11pt;
  color: #111;
  background: #fff;
}
.xlsx-preview th, .xlsx-preview td {
  padding: 2px 8px;
  font-weight: inherit;
  text-align: left;
  white-space: nowrap;
}
.demo-code {
  margin: 0;
  padding: 12px 16px;
  max-height: 460px;
  overflow: auto;
  font-size: 12px;
  line-height: 1.7;
  border-radius: 6px;
  background: #1e1e2e;
  color: #cdd6f4;
}
.preset-item { display: flex; flex-direction: column; gap: 8px; height: 100%; }
/* 这里刻意不用 flex 做居中：flex 容器会把文本包成内容宽的匿名 flex 项，
   内联的 textAlign / verticalAlign 双双失效，预设卡就演示不出对齐了。
   块级 + line-height 居中，text-align 才能正常生效（与预览表一致）。 */
.preset-cell {
  min-height: 28px;
  line-height: 28px;
  padding: 4px 12px;
  border-radius: 4px;
}
.preset-desc { margin: 0; font-size: 12px; }
`;

/** 速查卡示例值：有 numFormat 的预设按格式码渲染，其余原样展示。 */
function presetSampleText(sample: string | number, numFormat?: string): string {
  if (numFormat !== undefined) return formatByNumFormat(sample, numFormat);
  return String(sample);
}

/**
 * 浏览器内预览表：把 SheetConfig 解析为预览列（生效样式 + numFormat 渲染），
 * 用 HTML table 模拟 Excel 观感。样式解析复用导出引擎的合并语义（见
 * style-preview.ts），因此预览即导出所得。
 */
function SheetPreviewTable({ sheet }: { sheet: SheetConfig }) {
  const columns = useMemo(() => buildPreviewColumns(sheet), [sheet]);
  return (
    <div className="xlsx-scroll">
      <table className="xlsx-preview">
        <colgroup>
          {columns.map((c) => (
            <col
              key={c.id}
              style={c.widthPx !== undefined ? { width: c.widthPx } : undefined}
            />
          ))}
        </colgroup>
        <thead>
          <tr>
            {columns.map((c) => (
              <th key={c.id} style={cellStyleToCss(c.headerStyle)}>
                {c.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {sheet.data.map((row, i) => (
            <tr key={typeof row.id === "number" ? row.id : i}>
              {columns.map((c) => (
                <td key={c.id} style={cellStyleToCss(c.dataStyle)}>
                  {previewCellValue(sheet, c, row, i)}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/** 预设速查卡内容。 */
function PresetCheatCards() {
  return (
    <Card title="内置预设速查（StylePresets）">
      <Row gutter={[12, 12]}>
        {PRESET_CHEATSHEET.map((p) => {
          // StylePresets 各项是 satisfies 收窄后的字面量对象，统一按 CellStyle
          // 读取才能访问可选字段 numFormat（只读访问，不改预设本身）
          const style: CellStyle = StylePresets[p.name];
          // 占位轮廓只给自身没有边框的预设：bordered 有自己的真实边框，
          // 两者叠加会在卡片上画出双线
          const hasBorder =
            style.border !== undefined && Object.keys(style.border).length > 0;
          return (
            <Col key={p.name} xs={24} sm={12} xl={6}>
              <div className="preset-item">
                <div
                  className="preset-cell"
                  style={{
                    ...cellStyleToCss(style),
                    ...(hasBorder
                      ? {}
                      : { boxShadow: "inset 0 0 0 1px #d9d9d9" }),
                  }}
                >
                  {presetSampleText(p.sample, style.numFormat)}
                </div>
                <div>
                  <Typography.Text code strong>
                    StylePresets.{p.name}
                  </Typography.Text>
                  <Typography.Paragraph
                    type="secondary"
                    className="preset-desc"
                  >
                    {p.desc}
                  </Typography.Paragraph>
                </div>
              </div>
            </Col>
          );
        })}
      </Row>
    </Card>
  );
}

/** 场景配方卡：切换 / 说明 / 代码 / 预览 / 要点 / 导出。 */
function ScenarioPanel() {
  const [key, setKey] = useState<ScenarioKey>("combo");
  const scenario = useMemo(
    () => STYLE_SCENARIOS.find((s) => s.key === key)!,
    [key],
  );
  const [exporting, setExporting] = useState(false);
  const [result, setResult] = useState<{
    ok: boolean;
    degraded: boolean;
    message: string;
  } | null>(null);
  // 卸载防御：导航离开后异步导出不再触碰组件状态（与 basic-export 一致）
  const cancelledRef = useRef(false);

  useEffect(() => {
    cancelledRef.current = false;
    return () => {
      cancelledRef.current = true;
    };
  }, []);

  const runExport = async (target: StyleScenario): Promise<void> => {
    if (exporting) return;
    cancelledRef.current = false;
    setExporting(true);
    setResult(null);
    try {
      const res = await exportExcel({
        filename: target.filename,
        // 8 行 + auto 模式必走 main（Workbook 路径），样式必然生效
        sheets: [target.sheet],
      });
      if (cancelledRef.current) return;
      if (res.success) {
        // success 且带软错误 = 样式被降级剥离，按警告展示（正常不会发生）
        setResult(
          res.error
            ? {
                ok: true,
                degraded: true,
                message: `已导出 ${target.filename}，但发生降级：${res.error.message}`,
              }
            : {
                ok: true,
                degraded: false,
                message: `已导出 ${target.filename} · 模式 ${res.mode} · ${formatBytes(res.blob?.size ?? null)}，用 Excel / WPS 打开与预览比对`,
              },
        );
      } else {
        setResult({
          ok: false,
          degraded: false,
          message: `导出失败：${res.error?.message ?? "未知错误"}`,
        });
      }
    } catch (err) {
      if (cancelledRef.current) return;
      setResult({
        ok: false,
        degraded: false,
        message: `导出失败：${(err as Error).message}`,
      });
    } finally {
      if (!cancelledRef.current) setExporting(false);
    }
  };

  return (
    <Card title="场景配方">
      <Space orientation="vertical" size={16} style={{ width: "100%" }}>
        <Segmented<ScenarioKey>
          value={key}
          onChange={(next) => {
            setKey(next);
            setResult(null);
          }}
          options={STYLE_SCENARIOS.map((s) => ({
            value: s.key,
            label: s.title,
          }))}
        />
        <Typography.Paragraph style={{ marginBottom: 0 }}>
          {scenario.intro}
        </Typography.Paragraph>

        <Row gutter={[16, 16]}>
          <Col xs={24} xl={12}>
            <Card size="small" title="代码">
              <pre className="demo-code">{scenario.code}</pre>
            </Card>
          </Col>
          <Col xs={24} xl={12}>
            <Card
              size="small"
              title="浏览器预览（模拟 Excel）"
              extra={scenario.tags.map((t) => (
                <Tag key={t} color="blue">
                  {t}
                </Tag>
              ))}
            >
              <SheetPreviewTable sheet={scenario.sheet} />
            </Card>
          </Col>
        </Row>

        <ul style={{ margin: 0, paddingLeft: 20 }}>
          {scenario.notes.map((n, i) => (
            <li key={i}>
              <Typography.Text
                type={n.tone === "warning" ? "warning" : undefined}
              >
                {n.text}
              </Typography.Text>
            </li>
          ))}
        </ul>

        <Space wrap>
          <Button
            type="primary"
            icon={<DownloadOutlined />}
            loading={exporting}
            onClick={() => void runExport(scenario)}
          >
            导出 {scenario.filename} 验证真实文件
          </Button>
          {result && (
            <Alert
              style={{ flex: 1, minWidth: 320 }}
              type={
                !result.ok ? "error" : result.degraded ? "warning" : "success"
              }
              showIcon
              title={result.message}
            />
          )}
        </Space>
      </Space>
    </Card>
  );
}

export default function StylesDemo() {
  return (
    <Space orientation="vertical" size={16} style={{ width: "100%" }}>
      <style>{PREVIEW_CSS}</style>
      <Alert
        type="warning"
        showIcon
        title="样式仅在 50,000 行以下的 Workbook 路径渲染"
        description="auto 模式下浏览器 <20,000 行走主线程、20,000–49,999 行走 Worker + Workbook，两者样式均生效；流式路径（≥50,000 行或降级导出）会剥离全部样式并输出 console 告警。本页数据固定 8 行，恒走 Workbook 路径。"
      />
      <PresetCheatCards />
      <ScenarioPanel />
    </Space>
  );
}
