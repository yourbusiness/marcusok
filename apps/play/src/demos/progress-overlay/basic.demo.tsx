import { useEffect, useRef, useState } from "react";
import { Alert, Button, Card, Checkbox, Flex, Space, Typography } from "antd";
import { PlayCircleOutlined } from "@ant-design/icons";
import {
  showProgressOverlay,
  type ProgressOverlayHandle,
  type ProgressOverlayOptions,
} from "@marcusok/progress-overlay";

/**
 * progress-overlay 独立演示：不经任何业务包，直接驱动通用遮罩。
 *
 * 模拟任务分两段——先 1.5s 无进度（展示不确定态 spinner 与 hint），再
 * setProgress 逐格推进（展示确定态进度条与百分比）。两段切换正是
 * excel-exporter 里 workbook 路由（全程不确定）与 stream 路由（中间进度）
 * 的观感对照。
 */

type ThemeChoice = NonNullable<ProgressOverlayOptions["theme"]>;

export default function ProgressOverlayDemo() {
  const [theme, setTheme] = useState<ThemeChoice>("auto");
  const [customText, setCustomText] = useState(false);
  const [delayed, setDelayed] = useState(true);
  const [running, setRunning] = useState(false);
  const [lastRun, setLastRun] = useState<string | null>(null);
  const runningRef = useRef(false);

  useEffect(() => {
    return () => {
      // 组件卸载兜底：极端情况下任务句柄尚未关闭时不留挂遮罩。
      runningRef.current = true;
    };
  }, []);

  const runTask = async (): Promise<void> => {
    if (runningRef.current) return;
    runningRef.current = true;
    setRunning(true);
    setLastRun(null);
    const startedAt = performance.now();

    const overlay: ProgressOverlayHandle = showProgressOverlay({
      delayMs: delayed ? 200 : 0,
      theme,
      text: customText
        ? {
            title: "正在同步数据",
            initial: "连接服务…",
            phases: { pulling: "拉取远端数据…", saving: "写入本地…" },
            hint: "演示任务共 3.5 秒，可提前感受 delayMs 门控",
          }
        : undefined,
    });

    try {
      // 第一段：无中间进度（spinner + hint）。真实场景对应 workbook 构建。
      await sleep(1500);
      overlay.setPhase(customText ? "pulling" : "building");
      // 第二段：流式进度（百分比条）。真实场景对应 stream 每 1000 行上报。
      for (const p of [0.12, 0.3, 0.5, 0.72, 0.9]) {
        await sleep(280);
        overlay.setProgress(p);
      }
      overlay.setPhase(customText ? "saving" : "downloading");
      await sleep(200);
      overlay.setProgress(1);
      setLastRun(
        `任务完成，总耗时 ${Math.round(performance.now() - startedAt)}ms`,
      );
    } finally {
      overlay.close();
      runningRef.current = false;
      setRunning(false);
    }
  };

  return (
    <Space orientation="vertical" size={16} style={{ width: "100%" }}>
      <Typography.Paragraph type="secondary" style={{ marginBottom: 0 }}>
        通用全屏进度遮罩的独立演示：毛玻璃面板 + 旋转圆环（不确定态）/ 百分比
        条（确定态）。任务先停 1.5 秒无进度（spinner），随后流式推进到 100%
        （进度条）。excel-exporter 2.8.0 起的 <code>overlay</code> 选项由本包
        驱动，默认开启。
      </Typography.Paragraph>

      <Card>
        <Flex wrap gap={24} align="flex-end">
          <Space orientation="vertical" size={6}>
            <Typography.Text type="secondary">主题</Typography.Text>
            <Space>
              {(["auto", "light", "dark"] as const).map((t) => (
                <Button
                  key={t}
                  size="small"
                  type={theme === t ? "primary" : "default"}
                  onClick={() => setTheme(t)}
                >
                  {t}
                </Button>
              ))}
            </Space>
          </Space>
          <Space orientation="vertical" size={6}>
            <Typography.Text type="secondary">选项</Typography.Text>
            <Space orientation="vertical" size={4}>
              <Checkbox
                checked={customText}
                onChange={(e) => setCustomText(e.target.checked)}
              >
                自定义文案（phases 表）
              </Checkbox>
              <Checkbox
                checked={delayed}
                onChange={(e) => setDelayed(e.target.checked)}
              >
                delayMs 200（关掉则同步挂载）
              </Checkbox>
            </Space>
          </Space>
          <Space>
            <Button
              type="primary"
              size="large"
              icon={<PlayCircleOutlined />}
              loading={running}
              onClick={() => void runTask()}
            >
              模拟任务（3.5s）
            </Button>
          </Space>
        </Flex>
      </Card>

      {lastRun && <Alert type="success" showIcon message={lastRun} closable />}
    </Space>
  );
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
