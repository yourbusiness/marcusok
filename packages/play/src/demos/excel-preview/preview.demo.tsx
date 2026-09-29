import { useCallback, useEffect, useRef, useState } from "react";
import {
  Alert,
  Button,
  Card,
  Checkbox,
  Flex,
  Space,
  Tag,
  Typography,
  Upload,
} from "antd";
import { InboxOutlined, ReloadOutlined } from "@ant-design/icons";
import {
  createPreview,
  type PreviewError,
  type PreviewInstance,
  type PreviewParsedInfo,
} from "@marcusok/excel-preview";

const { Text } = Typography;

const ERROR_CODE_TEXT: Record<PreviewError["code"], string> = {
  PASSWORD_PROTECTED: "文件已加密（Agile AES-256）",
  LEGACY_FORMAT: "旧版 .xls（BIFF8）不支持，请另存为 .xlsx",
  UNSUPPORTED: "无法识别的文件格式",
  CORRUPT: "不是有效的 xlsx/zip 文件",
  WASM: "当前环境不支持 WebAssembly",
  UNKNOWN: "未知错误",
};

export default function ExcelPreviewDemo() {
  const containerRef = useRef<HTMLDivElement>(null);
  const instanceRef = useRef<PreviewInstance | null>(null);
  const bytesRef = useRef<Uint8Array | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [hasFile, setHasFile] = useState(false);
  const [info, setInfo] = useState<PreviewParsedInfo | null>(null);
  const [error, setError] = useState<PreviewError | null>(null);
  const [showHeaders, setShowHeaders] = useState(true);
  const [showTabs, setShowTabs] = useState(true);
  const [renderSeq, setRenderSeq] = useState(0);

  // 挂载统一由下方 effect 驱动（单一数据流）：onFile / 开关 / 重建按钮都
  // 只改状态，不在事件处理器里直接 mount——否则 fileName 变化会触发 effect
  // 二次 mount（同一文件解析两遍），而 renderSeq 变化又不触发（按钮清空
  // 预览后无人重建，实测缺陷）。
  const mount = useCallback(
    (bytes: Uint8Array) => {
      if (!containerRef.current) return;
      instanceRef.current?.destroy();
      instanceRef.current = null;
      containerRef.current.textContent = "";
      setInfo(null);
      setError(null);
      instanceRef.current = createPreview(containerRef.current, {
        source: bytes,
        showHeaders,
        showTabs,
        onParsed: setInfo,
        onError: setError,
      });
    },
    [showHeaders, showTabs],
  );

  const onFile = useCallback(async (file: File) => {
    const buf = await file.arrayBuffer();
    bytesRef.current = new Uint8Array(buf);
    setHasFile(true);
    setFileName(file.name); // 触发下方 effect 挂载
    return false;
  }, []);

  // 新文件（fileName）/ 选项变化（mount 身份随 showHeaders/showTabs 变）/
  // 手动重建（renderSeq）→ 以缓存字节重建（选项需重建实例才生效，演示层
  // 面可接受；真实业务可按需保留实例切开关）
  useEffect(() => {
    if (bytesRef.current && fileName) {
      mount(bytesRef.current);
    }
  }, [mount, fileName, renderSeq]);

  // 卸载销毁
  useEffect(() => {
    return () => {
      instanceRef.current?.destroy();
      instanceRef.current = null;
    };
  }, []);

  return (
    <Flex vertical gap={16}>
      <Card size="small">
        <Flex justify="space-between" align="center" wrap gap={12}>
          <Upload.Dragger
            accept=".xlsx,.xlsm,.csv"
            showUploadList={false}
            beforeUpload={(file) => {
              void onFile(file);
              return false;
            }}
            style={{ maxWidth: 420, padding: "8px 0" }}
          >
            <p className="ant-upload-drag-icon">
              <InboxOutlined />
            </p>
            <p className="ant-upload-text">点击或拖入 .xlsx / .xlsm / .csv</p>
            <p className="ant-upload-hint">
              解析在 Worker 内完成，大文件不冻结
              UI；加密文件错误码会提示需要密码
            </p>
          </Upload.Dragger>
          <Space direction="vertical" size={8}>
            <Checkbox
              checked={showHeaders}
              onChange={(e) => setShowHeaders(e.target.checked)}
            >
              行列表头
            </Checkbox>
            <Checkbox
              checked={showTabs}
              onChange={(e) => setShowTabs(e.target.checked)}
            >
              sheet 页签
            </Checkbox>
            <Button
              icon={<ReloadOutlined />}
              disabled={!hasFile}
              onClick={() => setRenderSeq((n) => n + 1)}
            >
              重建渲染
            </Button>
          </Space>
        </Flex>
      </Card>

      {error && (
        <Alert
          type="error"
          showIcon
          message={`${ERROR_CODE_TEXT[error.code] ?? error.code}`}
          description={error.message}
        />
      )}

      {info && (
        <Space size={8} wrap>
          <Tag color="blue">{fileName}</Tag>
          <Tag>{info.sheetCount} sheets</Tag>
          <Tag>
            {info.rowCount.toLocaleString()} 行 × {info.colCount} 列
          </Tag>
          <Tag color="green">解析 {info.duration.parse}ms</Tag>
          <Tag color="green">渲染 {info.duration.render}ms</Tag>
          <Tag>合计 {info.duration.total}ms</Tag>
        </Space>
      )}

      <Card
        size="small"
        styles={{ body: { padding: 0 } }}
        title={<Text type="secondary">预览（虚拟滚动：只渲染视口内格子）</Text>}
      >
        <div
          key={renderSeq}
          ref={containerRef}
          style={{ height: 560, position: "relative" }}
        />
      </Card>
    </Flex>
  );
}
