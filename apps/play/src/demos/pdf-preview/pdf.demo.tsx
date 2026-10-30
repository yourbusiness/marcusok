import { useCallback, useEffect, useRef, useState } from "react";
import {
  Alert,
  Button,
  Card,
  Flex,
  Segmented,
  Space,
  Switch,
  Tag,
  Typography,
  Upload,
} from "antd";
import {
  InboxOutlined,
  FilePdfOutlined,
  ReloadOutlined,
} from "@ant-design/icons";
import {
  createPdfPreview,
  type PdfPreviewError,
  type PdfPreviewInstance,
  type PdfRenderedInfo,
  type PdfScale,
} from "@marcusok/pdf-preview";

const { Text } = Typography;

const ERROR_CODE_TEXT: Record<PdfPreviewError["code"], string> = {
  UNSUPPORTED: "不是 PDF（头部 1024 字节内无 %PDF- 魔数）",
  PASSWORD_PROTECTED: "文件已加密（首版仅识别并报错，不解密）",
  CORRUPT: "文件结构损坏或解析失败",
  UNKNOWN: "未知错误",
};

type ScaleMode = "fit-width" | "fit-page" | "1" | "1.5";
const SCALE_MODES: ScaleMode[] = ["fit-width", "fit-page", "1", "1.5"];
const scaleOf = (m: ScaleMode): PdfScale =>
  m === "fit-width" || m === "fit-page" ? m : Number(m);

/** 合成 200 页纯文本文档（pdf-lib 浏览器端生成，懒渲染滚动/释放的行为验证）。 */
async function bigPdf(pages = 200): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts, rgb } = await import("pdf-lib");
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (let i = 1; i <= pages; i++) {
    const page = doc.addPage([612, 792]);
    page.drawText(`Synthetic page ${i} / ${pages}`, {
      x: 36,
      y: 738,
      size: 12,
      font,
      color: rgb(0, 0, 0),
    });
    // 文本必须停留在 WinAnsi 可编码范围：StandardFonts.Helvetica 画中文会
    // 直接抛错（"WinAnsi cannot encode"，包内 fixtures.ts 同口径限制），
    // 中文 PDF 的 cmap 链路请上传真实文档验证
    page.drawText(
      "Lazy rendering benchmark line — pipeline behavior verification, text layer selection.",
      { x: 36, y: 710, size: 10, font, color: rgb(0.2, 0.2, 0.2) },
    );
  }
  return doc.save();
}

/** 双实例演示文档：两页不同尺寸（验证逐页 viewport 占位与居中）。 */
async function dualPdf(label: string): Promise<Uint8Array> {
  const { PDFDocument, StandardFonts } = await import("pdf-lib");
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  // 第一页 US Letter，第二页 A4 横向——尺寸不同，验证占位/居中按页计算
  const p1 = doc.addPage([612, 792]);
  p1.drawText(`${label} — page 1 (US Letter)`, {
    x: 36,
    y: 740,
    size: 13,
    font,
  });
  const p2 = doc.addPage([842, 595]);
  p2.drawText(`${label} — page 2 (A4 landscape)`, {
    x: 36,
    y: 545,
    size: 13,
    font,
  });
  return doc.save();
}

export default function PdfPreviewDemo() {
  const containerRef = useRef<HTMLDivElement>(null);
  const instanceRef = useRef<PdfPreviewInstance | null>(null);
  const bytesRef = useRef<Uint8Array | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [hasFile, setHasFile] = useState(false);
  const [info, setInfo] = useState<PdfRenderedInfo | null>(null);
  const [error, setError] = useState<PdfPreviewError | null>(null);
  const [scaleMode, setScaleMode] = useState<ScaleMode>("fit-width");
  const [textLayer, setTextLayer] = useState(true);
  const [renderSeq, setRenderSeq] = useState(0);
  // 快速连续选文件/点合成按钮时，只允许最后一次选择生效（异步读取会乱序完成）
  const loadSeqRef = useRef(0);

  // 双实例区（懒渲染管线互相独立）
  const dualARef = useRef<HTMLDivElement>(null);
  const dualBRef = useRef<HTMLDivElement>(null);
  const dualInstA = useRef<PdfPreviewInstance | null>(null);
  const dualInstB = useRef<PdfPreviewInstance | null>(null);

  // scale/textLayer 变更即重建（首版无热重排 API，重 create 即可）
  useEffect(() => {
    if (!bytesRef.current || !containerRef.current) return;
    instanceRef.current?.destroy();
    instanceRef.current = null;
    containerRef.current.textContent = "";
    setInfo(null);
    setError(null);
    instanceRef.current = createPdfPreview(containerRef.current, {
      source: bytesRef.current,
      scale: scaleOf(scaleMode),
      textLayer,
      onRendered: setInfo,
      onError: setError,
    });
  }, [scaleMode, textLayer, renderSeq]);

  const onFile = useCallback(async (file: File) => {
    const seq = ++loadSeqRef.current;
    const buf = await file.arrayBuffer();
    if (seq !== loadSeqRef.current) return false; // 期间又选了别的文件
    bytesRef.current = new Uint8Array(buf);
    setHasFile(true);
    setFileName(file.name);
    setRenderSeq((n) => n + 1);
    return false;
  }, []);

  useEffect(() => {
    return () => {
      // 卸载后仍在读取中的选择无需作废：其结果只会写 ref/state，不会再建实例
      instanceRef.current?.destroy();
      dualInstA.current?.destroy();
      dualInstB.current?.destroy();
    };
  }, []);

  const loadBigDoc = useCallback(async () => {
    const seq = ++loadSeqRef.current;
    const bytes = await bigPdf();
    if (seq !== loadSeqRef.current) return; // 期间选了真实文件
    bytesRef.current = bytes;
    setHasFile(true);
    setFileName("synthetic-200-pages.pdf");
    setRenderSeq((n) => n + 1);
  }, []);

  const mountDual = useCallback(async () => {
    dualInstA.current?.destroy();
    dualInstB.current?.destroy();
    dualInstA.current = null;
    dualInstB.current = null;
    if (dualARef.current && dualBRef.current) {
      // 同页双实例：懒渲染管线与 IntersectionObserver 各自独立
      const [a, b] = await Promise.all([
        dualPdf("Instance A"),
        dualPdf("Instance B"),
      ]);
      dualInstA.current = createPdfPreview(dualARef.current, { source: a });
      dualInstB.current = createPdfPreview(dualBRef.current, { source: b });
    }
  }, []);

  return (
    <Flex vertical gap={16}>
      <Card size="small">
        <Flex justify="space-between" align="center" wrap gap={12}>
          <Upload.Dragger
            accept=".pdf"
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
            <p className="ant-upload-text">点击或拖入 .pdf</p>
            <p className="ant-upload-hint">
              解析在 Web Worker；加密文档给出明确错误码；中文 PDF 走 cmaps
              链路（建议上传真实中文文档验证）
            </p>
          </Upload.Dragger>
          <Space direction="vertical" size={8}>
            <Space size={8} wrap>
              <Button
                icon={<FilePdfOutlined />}
                onClick={() => void loadBigDoc()}
              >
                合成 200 页大文档
              </Button>
              <Button
                icon={<ReloadOutlined />}
                disabled={!hasFile}
                onClick={() => setRenderSeq((n) => n + 1)}
              >
                重建渲染
              </Button>
            </Space>
            <Space size={8} wrap>
              <Text type="secondary">缩放</Text>
              <Segmented
                size="small"
                options={SCALE_MODES}
                value={scaleMode}
                onChange={(v) => setScaleMode(v)}
              />
              <Text type="secondary">文本层</Text>
              <Switch
                size="small"
                checked={textLayer}
                onChange={setTextLayer}
              />
            </Space>
          </Space>
        </Flex>
      </Card>

      {error && (
        <Alert
          type="error"
          showIcon
          title={ERROR_CODE_TEXT[error.code] ?? error.code}
          description={error.message}
        />
      )}

      {info && (
        <Space size={8} wrap>
          <Tag color="blue">{fileName}</Tag>
          <Tag>{info.pageCount} 页</Tag>
          {/* 懒渲染口径：首页可见耗时 + 从 create 起的端到端耗时 */}
          <Tag color="green">首页出图 {info.firstPageMs}ms</Tag>
          <Tag color="cyan">端到端 {info.durationMs}ms</Tag>
        </Space>
      )}

      <Card
        size="small"
        styles={{ body: { padding: 0, overflow: "auto" } }}
        title={
          <Text type="secondary">
            预览（连续滚动懒渲染：进入视口渲染、远离视口释放
            canvas；文本层可选中复制）
          </Text>
        }
      >
        <div
          key={renderSeq}
          ref={containerRef}
          style={{ minHeight: 560, background: "#f5f5f5" }}
        />
      </Card>

      <Card
        size="small"
        title={
          <Flex justify="space-between" align="center">
            <Text type="secondary">
              双实例演示（独立懒渲染管线；第二页为 A4 横向验证逐页尺寸占位）
            </Text>
            <Button size="small" onClick={() => void mountDual()}>
              渲染双实例
            </Button>
          </Flex>
        }
        styles={{ body: { padding: 12 } }}
      >
        <Flex gap={12} wrap>
          <div
            ref={dualARef}
            style={{
              flex: 1,
              minWidth: 320,
              maxHeight: 420,
              overflow: "auto",
              border: "1px dashed #d9d9d9",
              borderRadius: 8,
              padding: 8,
              background: "#fafafa",
            }}
          />
          <div
            ref={dualBRef}
            style={{
              flex: 1,
              minWidth: 320,
              maxHeight: 420,
              overflow: "auto",
              border: "1px dashed #d9d9d9",
              borderRadius: 8,
              padding: 8,
              background: "#fafafa",
            }}
          />
        </Flex>
      </Card>
    </Flex>
  );
}
