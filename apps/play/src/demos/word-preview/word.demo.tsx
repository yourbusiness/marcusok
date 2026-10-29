import { useCallback, useEffect, useRef, useState } from "react";
import {
  Alert,
  Button,
  Card,
  Flex,
  Space,
  Tag,
  Typography,
  Upload,
} from "antd";
import {
  InboxOutlined,
  FileWordOutlined,
  ReloadOutlined,
} from "@ant-design/icons";
import { strToU8, zipSync } from "fflate";
import {
  createWordPreview,
  type WordPreviewError,
  type WordPreviewInstance,
  type WordRenderedInfo,
} from "@marcusok/word-preview";

const { Text } = Typography;

const ERROR_CODE_TEXT: Record<WordPreviewError["code"], string> = {
  UNSUPPORTED: "不是 Word 文档（非 zip 容器或缺 Word 必需部件）",
  // .doc 老格式与加密 docx 同为 OLE2 魔数，首版两者均不支持（合并表述）
  LEGACY_FORMAT: ".doc 老格式或加密的 .docx 不支持，请另存为 .docx",
  PASSWORD_PROTECTED: "文件已加密（zip 加密形态）",
  CORRUPT: "文件结构损坏或内容解析失败",
  UNKNOWN: "未知错误",
};

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const u = (s: string) => strToU8(HEAD + s);

/** 最小四文件 docx（与包内测试同配方：缺根关系文件会崩出模糊错）。 */
function buildDocx(body: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": u(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    ),
    "_rels/.rels": u(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": u(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`,
    ),
    "word/_rels/document.xml.rels": u(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
    ),
  });
}

/** 合成大文档：1000 段落 + 100×5 表格（纯文本，浏览器实测完整渲染耗时）。 */
function bigDocx(): Uint8Array {
  const paras = Array.from(
    { length: 1000 },
    (_, i) =>
      `<w:p><w:r><w:t>Paragraph ${i} — 合成大文档性能测量 benchmark text for word preview rendering pipeline.</w:t></w:r></w:p>`,
  ).join("");
  const rows = Array.from(
    { length: 100 },
    (_, r) =>
      `<w:tr>${Array.from(
        { length: 5 },
        (_, c) =>
          `<w:tc><w:tcPr/><w:p><w:r><w:t>cell ${r}-${c}</w:t></w:r></w:p></w:tc>`,
      ).join("")}</w:tr>`,
  ).join("");
  return buildDocx(
    paras +
      `<w:tbl><w:tblPr/><w:tblGrid>${"<w:gridCol/>".repeat(5)}</w:tblGrid>${rows}</w:tbl>`,
  );
}

/** 双实例演示文档：不同文本 + lastRenderedPageBreak 分页（浏览器复核真实分页符）。 */
function dualInstanceDocx(text: string): Uint8Array {
  return buildDocx(
    `<w:p><w:r><w:t>${text} — page 1</w:t></w:r></w:p>` +
      `<w:p><w:r><w:lastRenderedPageBreak/><w:t>${text} — page 2</w:t></w:r></w:p>`,
  );
}

export default function WordPreviewDemo() {
  const containerRef = useRef<HTMLDivElement>(null);
  const instanceRef = useRef<WordPreviewInstance | null>(null);
  const bytesRef = useRef<Uint8Array | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [hasFile, setHasFile] = useState(false);
  const [info, setInfo] = useState<WordRenderedInfo | null>(null);
  const [error, setError] = useState<WordPreviewError | null>(null);
  const [renderSeq, setRenderSeq] = useState(0);
  // 快速连续选文件时，只允许最后一次选择生效（异步读取会乱序完成）
  const loadSeqRef = useRef(0);

  // 双实例区（隔离演示）：两个独立容器 + 实例
  const dualARef = useRef<HTMLDivElement>(null);
  const dualBRef = useRef<HTMLDivElement>(null);
  const dualInstA = useRef<WordPreviewInstance | null>(null);
  const dualInstB = useRef<WordPreviewInstance | null>(null);

  const mount = useCallback((bytes: Uint8Array) => {
    if (!containerRef.current) return;
    instanceRef.current?.destroy();
    instanceRef.current = null;
    containerRef.current.textContent = "";
    setInfo(null);
    setError(null);
    instanceRef.current = createWordPreview(containerRef.current, {
      source: bytes,
      onRendered: setInfo,
      onError: setError,
    });
  }, []);

  const onFile = useCallback(async (file: File) => {
    const seq = ++loadSeqRef.current;
    const buf = await file.arrayBuffer();
    if (seq !== loadSeqRef.current) return false; // 期间又选了别的文件
    bytesRef.current = new Uint8Array(buf);
    setHasFile(true);
    setFileName(file.name);
    return false;
  }, []);

  useEffect(() => {
    if (bytesRef.current && fileName) {
      mount(bytesRef.current);
    }
  }, [mount, fileName, renderSeq]);

  // 卸载销毁（含双实例）
  useEffect(() => {
    return () => {
      instanceRef.current?.destroy();
      dualInstA.current?.destroy();
      dualInstB.current?.destroy();
    };
  }, []);

  const loadBigDoc = useCallback(() => {
    loadSeqRef.current++; // 作废仍在读取中的文件选择
    bytesRef.current = bigDocx();
    setHasFile(true);
    setFileName("synthetic-1000-paras.docx");
  }, []);

  const mountDual = useCallback(() => {
    dualInstA.current?.destroy();
    dualInstB.current?.destroy();
    dualInstA.current = null;
    dualInstB.current = null;
    if (dualARef.current && dualBRef.current) {
      // 同页双实例：本包为每实例生成唯一 className 隔离全局 CSS
      // （上游固定 "docx" 类名时后渲染者会覆盖前者的编号/主题样式）
      dualInstA.current = createWordPreview(dualARef.current, {
        source: dualInstanceDocx("Instance A"),
      });
      dualInstB.current = createWordPreview(dualBRef.current, {
        source: dualInstanceDocx("Instance B"),
      });
    }
  }, []);

  return (
    <Flex vertical gap={16}>
      <Card size="small">
        <Flex justify="space-between" align="center" wrap gap={12}>
          <Upload.Dragger
            accept=".docx,.docm"
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
            <p className="ant-upload-text">点击或拖入 .docx / .docm</p>
            <p className="ant-upload-hint">
              渲染在主线程完成（引擎限制）；.doc
              老格式与加密文档会给出明确错误码
            </p>
          </Upload.Dragger>
          <Space direction="vertical" size={8}>
            <Button icon={<FileWordOutlined />} onClick={loadBigDoc}>
              合成大文档（1000 段落 + 100×5 表格）
            </Button>
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
          message={ERROR_CODE_TEXT[error.code] ?? error.code}
          description={error.message}
        />
      )}

      {info && (
        <Space size={8} wrap>
          <Tag color="blue">{fileName}</Tag>
          <Tag>{info.pageCount} 页（section）</Tag>
          {/* durationMs 含图片/内嵌字体加载的完整渲染耗时（引擎 resolve 前 await 全部资源） */}
          <Tag color="green">完整渲染 {info.durationMs}ms</Tag>
        </Space>
      )}

      <Card
        size="small"
        styles={{ body: { padding: 0, overflow: "auto" } }}
        title={
          <Text type="secondary">
            预览（视觉保真渲染：分页/页眉页脚/脚注随文档呈现）
          </Text>
        }
      >
        <div key={renderSeq} ref={containerRef} style={{ minHeight: 560 }} />
      </Card>

      <Card
        size="small"
        title={
          <Flex justify="space-between" align="center">
            <Text type="secondary">
              双实例隔离演示（唯一 className 隔离全局 CSS）
            </Text>
            <Button size="small" onClick={mountDual}>
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
            }}
          />
        </Flex>
        <Text type="secondary" style={{ fontSize: 12 }}>
          两个实例各渲染一份两页文档；样式互相独立（Inspect 查看两个容器内的
          section 类名 docx-N 不同）
        </Text>
      </Card>
    </Flex>
  );
}
