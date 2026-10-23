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

async function toBytes(source: WordSource): Promise<Uint8Array> {
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  const buf = await source.arrayBuffer();
  return new Uint8Array(buf);
}

/**
 * 用户回调异常隔离：回调是调用方代码（常见如 React 组件已卸载），其抛错
 * 不得被 boot 的 catch 误报成渲染错误，也不得产生未处理拒绝。与
 * excel-preview 的 guardCallback 同口径：仅 warn。
 */
function guardCallback<A extends unknown[]>(
  cb: ((...args: A) => void) | undefined,
  name: string,
): ((...args: A) => void) | undefined {
  if (!cb) return cb;
  return (...args: A) => {
    try {
      cb(...args);
    } catch (e) {
      console.warn(`[word-preview] user ${name} callback threw; ignoring`, e);
    }
  };
}

/** 超链接放行协议（相对形态经 base 解析后落在其中）。 */
const SAFE_HREF_PROTOCOLS = new Set(["http:", "https:", "mailto:", "tel:"]);

/**
 * scheme 白名单判定。用 `URL` 解析而非 `startsWith`：WHATWG URL 规范在
 * 解析时剔除前后 C0 控制字符与空白、把协议名归一为小写——`" JavaScript:x"`
 * 一类的绕过形态都会被正确解析为 `javascript:` 并拒绝。无 scheme 的相对
 * 形态（含 `#anchor`、`//host`）经 baseURI 解析后继承 http(s)，放行。
 */
function isSafeHref(raw: string): boolean {
  try {
    const url = new URL(raw, document.baseURI);
    return SAFE_HREF_PROTOCOLS.has(url.protocol);
  } catch {
    return false; // URL 解析失败（非法字符等）不放行
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
      // WordRenderOptions 类型层面就没有 className 等接管键；运行期强传
      // 同名键也会被下面的显式赋值覆盖（不在白名单的 experimental 等键
      // 类型未含，JS 调用方强传会原样透传上游——薄封装不做二次过滤）。
      const ro = options.renderOptions ?? {};
      const upstream: Partial<Options> = {
        ...ro,
        className,
        useBase64URL: true,
        ignoreLastRenderedPageBreak: ro.ignoreLastRenderedPageBreak ?? false,
        renderAltChunks: options.allowAltChunks === true,
      };

      // 有 nonce 时用游离容器收拢上游全部 <style>（渲染后统一重建）；
      // 无 nonce 不传第三参（上游并入 bodyContainer，勿多走一次重建）
      const styleDetached = options.styleNonce
        ? document.createElement("div")
        : undefined;

      // sniff 已确认必需部件齐全，之后 renderAsync 抛错即"内容损坏/
      // 解析失败"，统一归 CORRUPT（保留 cause 供诊断）
      await renderAsync(bytes, wrapper, styleDetached, upstream);
      if (destroyed) return; // 迟到结果：wrapper 已摘除，丢弃且不回调

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
      if (destroyed) return;
      // 同 sniff 失败路径：失败即无副作用，摘除 wrapper 再报错
      // （destroy 分支已自行 remove，此处必是未 destroy 的失败）
      wrapper.remove();
      const message = err instanceof Error ? err.message : String(err);
      onError?.({
        code: "CORRUPT",
        message: `Failed to render document: ${message}`,
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
      // destroy 后所有回调静默（取消不是错误）。
      wrapper.remove();
    },
  };
}
