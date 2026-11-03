/**
 * createWordPreview 的实现：对 docx-preview `renderAsync` 的薄封装。
 *
 * 封装职责（偏离上游默认的三处 + 安全加固，均见包 README / limits）：
 * 1. sniff 前置（见 ../sniff.ts）——上游对"zip 非 docx"只报模糊空引用错；
 * 2. 每实例唯一 `className`——上游默认固定 `"docx"`，注入的默认样式/文档
 *    样式/编号变量全部是文档级全局 CSS，同页双实例会互相覆盖（真实
 *    Chromium 已复现）；本包传 `docx-<seq>` 隔离；
 * 3. `useBase64URL: true`——上游 0 处 revokeObjectURL，blob URL 必然泄漏；
 * 4. `renderAltChunks: false`——上游默认 true 且 iframe srcdoc 无 sandbox；
 * 5. 渲染后安全后处理：超链接 scheme 白名单、（可选）iframe sandbox、
 *    （可选）styleNonce 换节点重建。
 */
import { renderAsync } from "docx-preview";
import type { Options } from "docx-preview";
import { sniffWordDocument } from "../sniff";
import type {
  WordPreviewInstance,
  WordPreviewOptions,
  WordSource,
} from "../types";

/** 实例序号：内部唯一 className 的来源（docx-1、docx-2…）。 */
let instanceSeq = 0;

/**
 * 跨 realm 判定：`instanceof` 在不同 iframe/Node vm 的构造器间恒为 false，
 * 会让跨 realm 的 Uint8Array 落到 `.arrayBuffer()`（TypedArray 没有该
 * 方法）而报 "source.arrayBuffer is not a function"。ArrayBuffer.isView
 * 按内部槽判定，跨 realm 有效；ArrayBuffer 自身用 toString 标签兜底。
 */
function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === "[object ArrayBuffer]";
}

async function toBytes(source: WordSource): Promise<Uint8Array> {
  if (ArrayBuffer.isView(source)) {
    return new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  }
  if (isArrayBuffer(source)) return new Uint8Array(source);
  const buf = await source.arrayBuffer();
  return new Uint8Array(buf);
}

/**
 * 用户回调异常隔离：回调是调用方代码（常见如 React 组件已卸载），其抛错
 * 不得被 boot 的 catch 误报成渲染错误，也不得产生未处理拒绝。与
 * excel-preview 的 guardCallback 同口径：仅 warn。
 *
 * 异步回调是第二形态：TS 允许把 `async () => …` 赋给返回 void 的回调类型，
 * 其 rejected promise 不会被同步 catch 捕到——必须显式 .catch，否则就是
 * unhandled rejection。
 */
function guardCallback<A extends unknown[]>(
  cb: ((...args: A) => void) | undefined,
  name: string,
): ((...args: A) => void) | undefined {
  if (!cb) return cb;
  const warn = (e: unknown): void => {
    console.warn(`[word-preview] user ${name} callback threw; ignoring`, e);
  };
  return (...args: A) => {
    try {
      const result = cb(...args) as unknown;
      // 按 thenable 判断而非 `instanceof Promise`：iframe 等跨 realm 的
      // Promise 不满足后者，会让异步回调的 rejection 漏成 unhandled。
      if (
        result !== null &&
        (typeof result === "object" || typeof result === "function") &&
        typeof (result as { then?: unknown }).then === "function"
      ) {
        void (result as PromiseLike<unknown>).then(undefined, warn);
      }
    } catch (e) {
      warn(e);
    }
  };
}

/** 超链接放行协议（相对形态经 base 解析后落在其中）。 */
const SAFE_HREF_PROTOCOLS = new Set(["http:", "https:", "mailto:", "tel:"]);

/** 由本包接管、禁止透传上游的 renderOptions 键（见创建处的选项映射）。 */
const TAKEN_OVER_OPTION_KEYS = [
  "className",
  "useBase64URL",
  "renderAltChunks",
  "h",
  "experimental",
  "debug",
  "trimXmlDeclaration",
] as const;

/**
 * scheme 白名单判定。用 `URL` 解析而非 `startsWith`：WHATWG URL 规范在
 * 解析时剔除前后 C0 控制字符与空白、把协议名归一为小写——`" JavaScript:x"`
 * 一类的绕过形态都会被正确解析为 `javascript:` 并拒绝。无 scheme 的相对
 * 形态（`//host`、路径）经 baseURI 解析后继承 http(s)，放行；`#anchor`
 * 走函数开头的短路（不经协议判定）。
 */
function isSafeHref(raw: string): boolean {
  // `#` 前缀的锚点永远指向当前文档（纯片段，无导航/协议语义），先于
  // 协议白名单无条件放行：否则在 file:// 或自定义协议页面（Electron、
  // 本地打开的 HTML 报告）上片段会继承页面协议（file:）而被误剥，
  // 与 README/limits "保留锚点形态" 的承诺不符。
  if (raw.startsWith("#")) return true;
  try {
    const url = new URL(raw, document.baseURI);
    return SAFE_HREF_PROTOCOLS.has(url.protocol);
  } catch {
    return false; // URL 解析失败（非法字符等）不放行
  }
}

/**
 * 上游把“解析不到的图片/字体”写成字符串 `"null"`（`img.src = null` /
 * `url(null)`），浏览器会把它当相对 URL 请求到 `/null`（真实 Chromium
 * 实测 404），且坏图标可见。这里清掉这些占位，避免无谓请求与误导 UI：
 * - img[src]、SVG image[href]/[xlink:href] 指向 "null" 时摘除属性；
 * - <style> 中 @font-face 的 url(null) 整块删除，其余 url(null) 换成
 *   无网络请求的空 data URL。
 * 注意必须在上游资源任务全部 await 完成后执行（本函数即在 renderAsync
 * resolve 之后调用）；不改变合法图片/字体。
 */
function scrubBrokenResources(wrapper: HTMLElement): void {
  for (const el of wrapper.querySelectorAll("img[src]")) {
    if (el.getAttribute("src") === "null") el.removeAttribute("src");
  }
  for (const el of wrapper.querySelectorAll("image")) {
    for (const attr of ["href", "xlink:href"] as const) {
      if (el.getAttribute(attr) === "null") el.removeAttribute(attr);
    }
  }
  const brokenUrl = /url\(\s*(?:'null'|"null"|null)\s*\)/gi;
  for (const style of wrapper.querySelectorAll("style")) {
    const css = style.textContent;
    if (!css || !brokenUrl.test(css)) continue;
    brokenUrl.lastIndex = 0;
    style.textContent = css
      .replace(
        /@font-face\s*\{[^{}]*url\(\s*(?:'null'|"null"|null)\s*\)[^{}]*\}/gi,
        "",
      )
      .replace(brokenUrl, 'url("data:,")');
  }
}

/**
 * 渲染后安全后处理（与渲染同通道，上游 resolve 时图片/字体等异步任务已
 * 全部 await 完成，这里无需二次等待）：
 * - 超链接：上游 `renderHyperlink` 把外链关系 target **原始字符串直赋**
 *   href、无任何 scheme 校验（源码实证），恶意 docx 可写 `javascript:` /
 *   `data:` target、点击触发。非白名单一律摘除 href、保留链接文本。
 * - altChunk（仅 allowAltChunks 路径）：上游 iframe 无 sandbox，补
 *   `sandbox=""`（禁止脚本）。注意这是渲染后属性赋值，与 srcdoc 首帧
 *   加载的先后不保证——尽力而为，limits 文档明示残余风险。
 * - 解析不到的图片/字体资源：清掉 "null" 占位（见 scrubBrokenResources）。
 */
function postProcessRender(
  wrapper: HTMLElement,
  sandboxIframes: boolean,
): void {
  for (const a of wrapper.querySelectorAll("a[href]")) {
    if (!isSafeHref(a.getAttribute("href") ?? "")) {
      a.removeAttribute("href");
    }
  }
  if (sandboxIframes) {
    for (const f of wrapper.querySelectorAll("iframe")) {
      f.setAttribute("sandbox", "");
    }
  }
  scrubBrokenResources(wrapper);
}

/**
 * 空引用类上游错误的文案归一（V8 / Firefox / Safari 三种形态）。
 * 典型来源是文档引用了关系部件但包里缺失对应文件（如超链接/图片存在、
 * word/_rels/document.xml.rels 缺失）；不解析原始文案会直接把引擎内部
 * 错误暴露给调用方。
 */
const NULL_PROPERTY_ERROR_RE =
  /cannot read properties of (?:null|undefined) \(reading '[^']+'\)|can't access property ["'][^"']+["'],?\s*\S+ is (?:null|undefined)|null is not an object \(evaluating '[^']+'\)/i;

/** 导出供测试直接覆盖三种浏览器文案（不进入包公开 API）。 */
export function describeRenderFailure(message: string): string {
  return NULL_PROPERTY_ERROR_RE.test(message)
    ? "the document references a relationship part that is missing from the package (for example word/_rels/document.xml.rels). The file is incomplete or corrupt."
    : message;
}

/**
 * altChunk 的前置 sandbox：docx-preview 的 `renderAltChunk` 先 `h({tagName:
 * "iframe"})`，**随后**才在异步任务里设置 `srcdoc`。渲染完成后补
 * `sandbox` 属性已经太晚（实测 Chromium 下脚本必然先执行）——必须在
 * iframe 节点创建的那一刻就带 sandbox。
 *
 * 上游 `h` 未导出，这里按 0.4.1 的实现镜像（pinned 版本，升级时必须复核
 * docx-preview 的 h 语义）。仅 allowAltChunks 路径启用，避免影响默认路径。
 */
function createHardenedH(): NonNullable<Options["h"]> {
  const h: NonNullable<Options["h"]> = (elemOrText) => {
    if (typeof elemOrText === "string") {
      return document.createTextNode(elemOrText);
    }
    if (elemOrText instanceof Node) return elemOrText;
    const { tagName, className, style, children, ns, ...props } = elemOrText;
    if (tagName === "#fragment") {
      const fragment = document.createDocumentFragment();
      for (const child of children ?? []) fragment.appendChild(h(child));
      return fragment;
    }
    if (tagName === "#comment") {
      return document.createComment(
        typeof children?.[0] === "string" ? children[0] : "",
      );
    }
    const el = ns
      ? document.createElementNS(ns, tagName)
      : document.createElement(tagName);
    // 关键：在返回给上游之前设置，早于 srcdoc 赋值/节点挂载
    if (tagName.toLowerCase() === "iframe") {
      el.setAttribute("sandbox", "");
    }
    if (className) el.setAttribute("class", className);
    if (typeof style === "string") el.setAttribute("style", style);
    else if (style) Object.assign((el as HTMLElement).style, style);
    for (const [key, value] of Object.entries(props)) {
      if (value !== undefined)
        (el as unknown as Record<string, unknown>)[key] = value;
    }
    for (const child of children ?? []) el.appendChild(h(child));
    return el;
  };
  return h;
}

/**
 * styleNonce 的"换节点"实现：CSP 对 `<style>` 的拦截发生在节点**插入时**，
 * 事后给已被拒的 style 补 nonce 属性不会重新生效——必须删旧节点、新建带
 * nonce 的 `<style>` 塞同样 CSS（excel-preview styleNonce 的同款口径）。
 *
 * 实现借 renderAsync 第三参 styleContainer：把一个游离容器传给上游，全部
 * `<style>` 节点会被路由进去（缺省并入 bodyContainer）。游离容器不进
 * 文档流，重建的带 nonce 节点必须挂回渲染容器内，否则整页丢样式。
 */
function rebuildStylesWithNonce(
  detached: HTMLElement,
  wrapper: HTMLElement,
  nonce: string,
): void {
  for (const style of detached.querySelectorAll("style")) {
    const el = document.createElement("style");
    el.setAttribute("nonce", nonce);
    el.textContent = style.textContent;
    wrapper.appendChild(el);
  }
}

/**
 * 清理上游向全局注册表写入的本实例条目：renderComments 路径会向
 * `CSS.highlights` 注册 `<className>-comments`（Highlight 持有 Range，
 * 随注册表常驻；上游全库 0 处 delete）——不清理的话 React 反复
 * create/destroy 会无界累积。CSS.highlights 是较新 API（happy-dom 等
 * 环境无实现），特性检测兜底；delete 失败不阻断卸载主流程。
 */
function cleanupGlobalRegistrations(className: string): void {
  if (typeof CSS === "undefined" || !CSS.highlights) return;
  try {
    CSS.highlights.delete(`${className}-comments`);
  } catch (e) {
    console.warn(
      "[word-preview] comment highlight cleanup failed; ignoring",
      e,
    );
  }
}

export function createWordPreview(
  container: HTMLElement,
  options: WordPreviewOptions,
): WordPreviewInstance {
  const onRendered = guardCallback(options.onRendered, "onRendered");
  const onError = guardCallback(options.onError, "onError");
  let destroyed = false;

  // 每实例唯一 className：隔离上游按固定 "docx" 类名注入的全局 CSS（默认
  // 样式、文档样式 .docx_<styleId>、编号 p.docx-num-…、主题变量）。上游
  // 生成选择器/变量时全部以该类名为前缀（rootSelector、processStyleName、
  // 编号 --docx-N-* 变量，源码实证），换名即整体换前缀。
  const className = `docx-${++instanceSeq}`;

  // 私有 wrapper（bodyContainer）：上游 renderAsync 在 await 解析+渲染
  // **之后**才清空容器并 append 节点——destroy 清空用户容器兜不住迟到
  // 结果（会重新写入）。改为同步摘除 wrapper：迟到的渲染只写入已脱离
  // 文档的节点。用户容器只作挂载点。
  const wrapper = document.createElement("div");
  wrapper.className = className;
  container.appendChild(wrapper);

  const boot = async () => {
    const t0 = performance.now();
    // 读输入源失败（如 Blob 底层数据不可用）与文档内容无关，归 UNKNOWN
    // 而非 CORRUPT——CORRUPT 专指"是 Word 文档但内容/结构坏了"（sniff 之后
    // 的失败），api 文档按此口径描述 UNKNOWN
    let bytes: Uint8Array;
    try {
      bytes = await toBytes(options.source);
    } catch (err) {
      if (destroyed) return;
      wrapper.remove();
      onError?.({
        code: "UNKNOWN",
        message: `Failed to read the source: ${err instanceof Error ? err.message : String(err)}`,
        cause: err,
      });
      return;
    }
    try {
      if (destroyed) return;

      const sniffed = await sniffWordDocument(bytes);
      if (destroyed) return;
      if (!sniffed.ok) {
        // 失败即无副作用：摘除私有 wrapper 再报错（excel-preview 的错误
        // 路径不碰容器，本包因防 destroy 竞态提前挂了 wrapper，须在终止
        // 路径显式回收，否则错误重试会在用户容器累积空 div）
        wrapper.remove();
        onError?.(sniffed.error);
        return;
      }

      // 选项映射：白名单透传 + 本包接管的默认加固（useBase64URL /
      // ignoreLastRenderedPageBreak / renderAltChunks / className）。
      // WordRenderOptions 类型层面没有这些接管键，但 JS 调用方可强传——
      // 必须显式删除，否则 h/experimental/debug 等仍会直达上游（h 可接管
      // 整棵渲染树的元素创建、experimental 会打开上游定时器），与对外
      // "由本包接管、传入无效" 的承诺不符。
      const ro = options.renderOptions ?? {};
      const upstream: Partial<Options> = { ...ro };
      for (const key of TAKEN_OVER_OPTION_KEYS) delete upstream[key];
      Object.assign(upstream, {
        className,
        useBase64URL: true,
        ignoreLastRenderedPageBreak: ro.ignoreLastRenderedPageBreak ?? false,
        renderAltChunks: options.allowAltChunks === true,
      });
      if (upstream.renderAltChunks) {
        upstream.h = createHardenedH();
      }

      // 有 nonce 时用游离容器收拢上游全部 <style>（渲染后统一重建）；
      // 无 nonce 不传第三参（上游并入 bodyContainer，勿多走一次重建）
      const styleDetached = options.styleNonce
        ? document.createElement("div")
        : undefined;

      // sniff 已确认必需部件齐全，之后 renderAsync 抛错即"内容损坏/
      // 解析失败"，统一归 CORRUPT（保留 cause 供诊断）
      await renderAsync(bytes, wrapper, styleDetached, upstream);
      if (destroyed) {
        // 上游可能在 destroy 的首次清理之后才注册注释高亮；迟到完成时
        // 必须再删一次，避免 Highlight 持有的 Range 永久留在全局注册表。
        cleanupGlobalRegistrations(className);
        // 迟到的渲染结果全部写进了已脱离文档的 wrapper；destroy 的第一轮
        // 清空发生在这些节点存在之前，必须再清一次，否则实例闭包会一直
        // 持有整份渲染 DOM（含 base64 图片/字体）。
        wrapper.replaceChildren();
        return; // wrapper 已摘除，丢弃且不回调
      }

      if (styleDetached) {
        rebuildStylesWithNonce(styleDetached, wrapper, options.styleNonce!);
      }
      postProcessRender(wrapper, upstream.renderAltChunks === true);

      // pageCount 按容器内 section 元素计数（上游每"页"渲染为一个
      // <section class="docx-N">，源码 createPageElement）
      const pageCount = wrapper.querySelectorAll(`section.${className}`).length;
      onRendered?.({
        pageCount,
        durationMs: Math.round(performance.now() - t0),
      });
    } catch (err) {
      if (destroyed) {
        // reject 同样可能发生在迟到注册之后；销毁态虽不回调，仍须收尾。
        cleanupGlobalRegistrations(className);
        wrapper.replaceChildren();
        return;
      }
      // 同 sniff 失败路径：失败即无副作用，摘除 wrapper 再报错
      // （destroy 分支已自行 remove，此处必是未 destroy 的失败）。
      // renderAsync 抛错前可能已注册注释高亮（renderComments 半途失败），
      // 全局注册表一并清理
      wrapper.remove();
      cleanupGlobalRegistrations(className);
      const message = err instanceof Error ? err.message : String(err);
      // 缺关系部件时上游只抛浏览器各异的空引用 TypeError——这正是本包
      // 声称要消除的晦涩错误。统一改写为可操作说明，保留 cause 供诊断。
      const friendly = describeRenderFailure(message);
      onError?.({
        code: "CORRUPT",
        message: `Failed to render document: ${friendly}`,
        cause: err,
      });
    }
  };
  void boot();

  return {
    destroy() {
      if (destroyed) return;
      destroyed = true;
      // 同步摘除私有 wrapper：用户容器立即干净；在途的 renderAsync 完成
      // 后只会往已脱离文档的 wrapper 写入，boot 的 destroyed 检查丢弃结果。
      // destroy 后所有回调静默（取消不是错误）。上游注册进全局
      // CSS.highlights 的注释高亮一并删除（见 cleanupGlobalRegistrations）。
      // 先清空再摘除：wrapper 仍被本 destroy 闭包引用，只 remove() 会把
      // 整份渲染 DOM（base64 图片/字体）留在闭包里直到实例被回收。
      wrapper.replaceChildren();
      wrapper.remove();
      cleanupGlobalRegistrations(className);
    },
  };
}
