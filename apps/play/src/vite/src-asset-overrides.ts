/**
 * src 联调模式下的资产字面量改写（纯函数，供 vite.config.ts 的插件胶水使用）。
 *
 * 背景：play 把 @marcusok/* 主入口别名到 src/，但包内的
 * `new URL("./<file>", import.meta.url)` 字面量相对 src 解析——资产不在 src
 * 而在 dist/，dev 服务与生产构建都解析不到。
 *
 * 改写方式：把字面量**整体**替换为 `new URL(<id>, import.meta.url)`，其中
 * `<id>` 是注入的 `import <id> from "<pkg>/dist/<file>?url"` 资源 URL 变量：
 *  - dev：marcusok-resolver（同 config）把 `<pkg>/dist/...?url` 解析到包内
 *    磁盘文件，vite 以 ?url 语义直接服务该路径（server.fs.allow 已含
 *    packages/），Worker/fetch 同源可加载；
 *  - build：?url 是 vite 官方资产机制，emit 哈希副本并内联 URL 常量——
 *    行为确定，不依赖 asset-import-meta-url 插件对 `../` 相对形态的解析
 *    （vite 8 / rolldown 实测该形态不稳定：同仓两个worker 一个被资产化、
 *    一个警告 "doesn't exist at build time" 后保留运行时相对解析、部署
 *    后 404）。此前的相对路径改写（`../dist/` 固定前缀）还有层级 bug：
 *    子目录模块（excel-preview 的 src/worker/）被改到不存在的 src/dist/，
 *    preview 的 parse worker 在 play 里从未真正加载过，一直静默走主线程
 *    fallback。
 *
 * **必须包回 `new URL(...)` 而不能裸用 `<id>`**：被替换的字面量是 URL
 * **对象**表达式，消费方形态包括 `.href` 取属性（pdf-preview 的
 * workerSrc）、`instanceof URL` 分支与传给 fetch/Worker——?url 导入在
 * dev/build 下都是**字符串**，裸替换会让 `.href` 变成 undefined（pdf-preview
 * 的 worker 曾因此在 play 的 dev/build 双模式里静默失效，报 "Invalid
 * workerSrc"）。首参是标识符而非字符串字面量，bundler 的
 * asset-import-meta-url 模式匹配不会二次介入，?url 已完成资产定位。
 */

/** 标识符安全化：@marcusok/excel-preview → _marcusok_excel-preview。 */
function safeIdent(pkgName: string): string {
  return pkgName.replace(/[^a-zA-Z0-9_$]/g, "_");
}

/**
 * 把代码中每个 `new URL("./<file>", import.meta.url)` 字面量替换为模块
 * 顶部注入的 `?url` 导入变量（见头注释的 dev/build 行为）。
 * 无任何命中返回 null（插件协议：null = 不改写）。
 */
export function rewriteAssetLiterals(
  code: string,
  pkgName: string,
  files: string[],
): string | null {
  let out = code;
  const imports: string[] = [];
  files.forEach((file, i) => {
    const id = `__marcusok_asset_${safeIdent(pkgName)}_${i}`;
    const literal = `new URL("./${file}", import.meta.url)`;
    // 包回 new URL 保持 URL 对象语义（见头注释）；base 沿用 import.meta.url
    const replacement = `new URL(${id}, import.meta.url)`;
    if (!out.includes(literal)) return;
    imports.push(`import ${id} from "${pkgName}/dist/${file}?url";`);
    out = out.split(literal).join(replacement);
  });
  if (imports.length === 0) return null;
  return `${imports.join("\n")}\n${out}`;
}
