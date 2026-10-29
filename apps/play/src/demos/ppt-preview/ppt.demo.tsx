import { useCallback, useEffect, useRef, useState } from "react";
import {
  Alert,
  Button,
  Card,
  Flex,
  Input,
  Segmented,
  Space,
  Tag,
  Typography,
  Upload,
} from "antd";
import {
  InboxOutlined,
  FileImageOutlined,
  LeftOutlined,
  RightOutlined,
  LockOutlined,
} from "@ant-design/icons";
import { strToU8, zipSync } from "fflate";
import {
  createPptPreview,
  type PptPreviewError,
  type PptPreviewInstance,
  type PptRenderedInfo,
} from "@marcusok/ppt-preview";

const { Text } = Typography;

const ERROR_CODE_TEXT: Record<PptPreviewError["code"], string> = {
  UNSUPPORTED: "不是 .pptx（或超资源预算 / 密码无法解决的不支持项）",
  LEGACY_FORMAT: "老式二进制 .ppt，请另存为 .pptx",
  PASSWORD_PROTECTED: "文件已加密（缺密码或密码错误）",
  CORRUPT: "文件结构损坏",
  UNKNOWN: "未知错误（Worker / WASM / 资产加载等）",
};

type WidthMode = "720" | "960" | "1280";
const WIDTH_MODES: WidthMode[] = ["720", "960", "1280"];

/**
 * 合成多页 pptx（浏览器端 fflate 打 zip，与包内测试 fixture 同配方——
 * 最小四类部件即可）。中文文本用于验证 CJK 字形回退。
 */
function synthPptx(slides: number, label: string): Uint8Array {
  const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const u = (s: string) => strToU8(HEAD + s);
  const slide = (text: string, bg: string) =>
    `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bg}"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" sz="4400" b="1"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
  const BGS = ["1F4E79", "7030A0", "C00000", "1E6B36", "8C5A1E"];
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": u(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>${Array.from(
        { length: slides },
        (_, i) =>
          `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
      ).join("")}</Types>`,
    ),
    "_rels/.rels": u(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    ),
    "ppt/presentation.xml": u(
      `<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst>${Array.from(
        { length: slides },
        (_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`,
      ).join(
        "",
      )}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
    ),
    "ppt/_rels/presentation.xml.rels": u(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${Array.from(
        { length: slides },
        (_, i) =>
          `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`,
      ).join("")}</Relationships>`,
    ),
  };
  for (let i = 0; i < slides; i++) {
    files[`ppt/slides/slide${i + 1}.xml`] = u(
      slide(`${label} 第 ${i + 1} / ${slides} 页`, BGS[i % BGS.length]!),
    );
  }
  return zipSync(files);
}

export default function PptPreviewDemo() {
  const containerRef = useRef<HTMLDivElement>(null);
  const instanceRef = useRef<PptPreviewInstance | null>(null);
  const bytesRef = useRef<Uint8Array | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [password, setPassword] = useState("");
  const [info, setInfo] = useState<PptRenderedInfo | null>(null);
  const [error, setError] = useState<PptPreviewError | null>(null);
  const [widthMode, setWidthMode] = useState<WidthMode>("960");
  const [textSel, setTextSel] = useState(true);
  const [renderSeq, setRenderSeq] = useState(0);
  const [navIndex, setNavIndex] = useState<number | null>(null);
  // 快速连续选文件/点合成按钮时，只允许最后一次选择生效（异步读取会乱序完成）
  const loadSeqRef = useRef(0);

  // 双实例区（懒渲染管线与引擎 Worker 互相独立）
  const dualARef = useRef<HTMLDivElement>(null);
  const dualBRef = useRef<HTMLDivElement>(null);
  const dualInstA = useRef<PptPreviewInstance | null>(null);
  const dualInstB = useRef<PptPreviewInstance | null>(null);

  const loadBytes = useCallback((bytes: Uint8Array, name: string) => {
    bytesRef.current = bytes;
    setFileName(name);
    setError(null);
    setInfo(null);
    setNavIndex(null);
    setRenderSeq((s) => s + 1);
  }, []);

  // 参数（宽度/文本层/密码）或 source 变更即重建：首版无热重排 API，
  // destroy + 重 create 是文档倡导的路径
  useEffect(() => {
    if (!bytesRef.current || !containerRef.current) return;
    instanceRef.current?.destroy();
    instanceRef.current = null;
    containerRef.current.textContent = "";
    setInfo(null);
    setError(null);
    instanceRef.current = createPptPreview(containerRef.current, {
      source: bytesRef.current,
      slideWidth: Number(widthMode),
      enableTextSelection: textSel,
      ...(password.trim() !== "" ? { password: password.trim() } : {}),
      onRendered: setInfo,
      onError: setError,
    });
    return () => {
      instanceRef.current?.destroy();
      instanceRef.current = null;
    };
  }, [renderSeq, widthMode, textSel, password]);

  // 双实例：同源两份字节（source 复用安全性由包内防御性复制保证）
  useEffect(() => {
    if (!bytesRef.current || !dualARef.current || !dualBRef.current) return;
    dualInstA.current?.destroy();
    dualInstB.current?.destroy();
    dualARef.current.textContent = "";
    dualBRef.current.textContent = "";
    dualInstA.current = createPptPreview(dualARef.current, {
      source: bytesRef.current,
      slideWidth: 480,
    });
    dualInstB.current = createPptPreview(dualBRef.current, {
      source: bytesRef.current,
      slideWidth: 480,
      enableTextSelection: false,
    });
    return () => {
      dualInstA.current?.destroy();
      dualInstB.current?.destroy();
      dualInstA.current = null;
      dualInstB.current = null;
    };
  }, [renderSeq]);

  // render 内不读 ref（react-hooks/refs）：页数快照在实例创建后写入 state
  const [total, setTotal] = useState(0);
  useEffect(() => {
    const t = setTimeout(
      () => setTotal(instanceRef.current?.slideCount ?? 0),
      300,
    );
    return () => clearTimeout(t);
  }, [renderSeq, widthMode, textSel, password]);

  return (
    <Space direction="vertical" size="middle" style={{ width: "100%" }}>
      <Card size="small">
        <Flex wrap="wrap" gap="middle" align="center">
          <Upload
            accept=".pptx"
            showUploadList={false}
            beforeUpload={(file) => {
              const seq = ++loadSeqRef.current;
              void file.arrayBuffer().then((buf) => {
                if (seq !== loadSeqRef.current) return; // 期间又选了别的来源
                loadBytes(new Uint8Array(buf), file.name);
              });
              return false;
            }}
          >
            <Button icon={<InboxOutlined />}>上传 / 拖入 .pptx</Button>
          </Upload>
          <Button
            icon={<FileImageOutlined />}
            onClick={() => {
              loadSeqRef.current++;
              loadBytes(synthPptx(30, "合成样本"), "合成 30 页样本.pptx");
            }}
          >
            合成 30 页样本
          </Button>
          <Button
            onClick={() => {
              loadSeqRef.current++;
              loadBytes(synthPptx(120, "滚动压力"), "合成 120 页样本.pptx");
            }}
          >
            合成 120 页样本（滚动压力）
          </Button>
          <Input
            style={{ width: 200 }}
            prefix={<LockOutlined />}
            placeholder="加密文件密码"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            allowClear
          />
          <Segmented
            options={WIDTH_MODES}
            value={widthMode}
            onChange={(v) => setWidthMode(v)}
          />
          <Button
            icon={textSel ? "Aa" : <span />}
            onClick={() => setTextSel((s) => !s)}
          >
            文本层：{textSel ? "开" : "关"}
          </Button>
        </Flex>
      </Card>

      {error && (
        <Alert
          type={error.code === "UNKNOWN" ? "warning" : "error"}
          showIcon
          message={
            <Space size="small">
              <Tag color="red">{error.code}</Tag>
              <Text>{ERROR_CODE_TEXT[error.code]}</Text>
            </Space>
          }
          description={error.message}
          closable
          onClose={() => setError(null)}
        />
      )}
      {info && (
        <Alert
          type="success"
          showIcon
          message={
            <Space size="small" wrap>
              <Tag color="green">onRendered</Tag>
              <Text>共 {info.slideCount} 页</Text>
              <Text>首张页渲染 {info.firstSlideMs}ms</Text>
              <Text>端到端 {info.durationMs}ms</Text>
              <Text type="secondary">
                （端到端含全 deck
                preflight——非渐进加载会预检全部页，懒渲染只延迟绘制）
              </Text>
            </Space>
          }
        />
      )}

      <Card
        size="small"
        title={fileName ?? "未选择文件"}
        extra={
          <Space size="small">
            <Button
              size="small"
              icon={<LeftOutlined />}
              onClick={() => instanceRef.current?.prevSlide()}
            >
              上一页
            </Button>
            <Input
              size="small"
              style={{ width: 90 }}
              placeholder="0-based 页码"
              value={navIndex ?? ""}
              onChange={(e) => {
                const v = e.target.value.trim();
                setNavIndex(v === "" ? null : Number(v));
              }}
              onPressEnter={() => {
                if (navIndex === null) return;
                try {
                  instanceRef.current?.goToSlide(Math.trunc(navIndex));
                } catch (err) {
                  setError({
                    code: "UNSUPPORTED",
                    message: `goToSlide: ${err instanceof Error ? err.message : String(err)}`,
                  });
                }
              }}
            />
            <Button
              size="small"
              onClick={() => instanceRef.current?.nextSlide()}
            >
              下一页
              <RightOutlined />
            </Button>
            <Text type="secondary">slideCount = {total}</Text>
          </Space>
        }
      >
        <div
          ref={containerRef}
          style={{
            height: 520,
            overflowY: "auto",
            background: "#f5f5f5",
            borderRadius: 8,
          }}
        />
      </Card>

      <Card size="small" title="双实例（同源字节，独立 Worker / 懒渲染）">
        <Flex gap="middle">
          <div style={{ flex: 1, minWidth: 0 }}>
            <Text type="secondary">A：文本层开</Text>
            <div
              ref={dualARef}
              style={{
                height: 320,
                overflowY: "auto",
                background: "#f5f5f5",
                borderRadius: 8,
                marginTop: 4,
              }}
            />
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <Text type="secondary">B：文本层关（纯 canvas）</Text>
            <div
              ref={dualBRef}
              style={{
                height: 320,
                overflowY: "auto",
                background: "#f5f5f5",
                borderRadius: 8,
                marginTop: 4,
              }}
            />
          </div>
        </Flex>
      </Card>

      <Alert
        type="info"
        showIcon
        message="验收提示"
        description={
          <ul style={{ margin: 0, paddingLeft: 18 }}>
            <li>
              加密
              .pptx（PowerPoint「打开密码」）在密码框输入后自动重建解密；错误密码报
              PASSWORD_PROTECTED。
            </li>
            <li>
              滚动主预览区验证懒渲染：视口外页的 canvas 会释放（DevTools
              Elements 观察），快速滚动回落不白屏。
            </li>
            <li>
              上传真实中文文档可验证 CJK 字形回退；含图表/SmartArt 的文档按
              limits 说明降级。
            </li>
            <li>
              goToSlide 输入 0-based 页码回车；越界抛 RangeError
              的行为在控制台可见（demo 内转为 Alert）。
            </li>
          </ul>
        }
      />
    </Space>
  );
}
