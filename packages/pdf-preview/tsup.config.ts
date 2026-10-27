import { defineConfig } from "tsup";

// 单配置：pdfjs-dist（现代构建）整体打进本包 dist，pdfjs-dist 只在
// devDependencies——npm 消费方零传递依赖（word-preview 打 docx-preview 的
// 同款决策；上游无 dependencies/peerDependencies，仅 Node 端的 optional
// @napi-rs/canvas，不影响消费方）。
//
// platform 保持 browser：pdfjs-dist 的 browser 字段把 fs/url/http/https/
// canvas 映射为 false，esbuild 据此剔除 Node 分支（word-preview 打 jszip
// 的同类机制）。
//
// 不 minify（仓库惯例）。注意：esbuild 的 legalComments 只保留带
// @license/@preserve///! 标记的注释——pdf.js 许可头用的是 @licstart/@licend
// 标记，**不会**被自动保留（实测 dist 无 Mozilla 头），故以 banner 显式
// 补回同一份许可声明；构建后 grep "Mozilla Foundation" 验证（见
// THIRD_PARTY_NOTICES.md）。
//
// worker/cmaps/standard_fonts 不经 esbuild——由 onSuccess 的 copy-assets.mjs
// 从 pdfjs-dist 原样拷入 dist（资产不带哈希、路径稳定，self-host 与
// bundler copy 插件都好引用）。
export default defineConfig({
  entry: { index: "src/index.ts" },
  format: ["esm"],
  dts: { resolve: true },
  splitting: true,
  treeshake: true,
  clean: true,
  sourcemap: true,
  target: "es2022",
  noExternal: ["pdfjs-dist"],
  onSuccess: "node scripts/copy-assets.mjs",
  platform: "browser",
  // pdf.js（Apache-2.0）的许可声明：esbuild 不认 @licstart 标记，以 banner
  // 原样保留（与 pdf.mjs 文件头逐字一致）
  banner: {
    js: `/**
 * @licstart The following is the entire license notice for the
 * JavaScript code in this page
 *
 * Copyright 2024 Mozilla Foundation
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * @licend The above is the entire license notice for the
 * JavaScript code in this page.
 */`,
  },
});
