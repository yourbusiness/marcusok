// Keep the en (root) and zh/ page trees in sync.
// Rules:
// - Site-level pages (guide/, index.md, play.md) must always be mirrored.
// - A package under packages/<dir>/ is "bilingual" only when zh/packages/<dir>
//   exists; en-only packages are exempt from the mirror requirement. Remove the
//   zh/ directory to mark a package en-only, create it to require full mirrors.
// - Every zh/ page must have an en original (extra zh pages fail the check).
// Runs as `pnpm test` for the docs app (turbo test in CI).
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { dirname, join, relative, sep } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const zhRoot = join(root, "zh");
const SKIP_DIRS = new Set(["zh", ".vitepress", "node_modules", "public"]);

function collectMd(dir, baseDir) {
  const out = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const abs = join(dir, entry.name);
    if (entry.isDirectory()) {
      if (SKIP_DIRS.has(entry.name)) continue;
      out.push(...collectMd(abs, baseDir));
    } else if (
      entry.isFile() &&
      entry.name.endsWith(".md") &&
      !(dir === baseDir && entry.name === "README.md")
    ) {
      out.push(relative(baseDir, abs).replaceAll(sep, "/"));
    }
  }
  return out;
}

const en = collectMd(root, root).sort();
const zh = collectMd(zhRoot, zhRoot).sort();

const PACKAGES_PREFIX = "packages/";

// Package pages are required to be mirrored only when the package has a zh/
// mirror directory; site-level pages are always required.
const required = en.filter((f) => {
  if (!f.startsWith(PACKAGES_PREFIX)) return true;
  const pkgDir = f.slice(PACKAGES_PREFIX.length).split("/")[0];
  return existsSync(join(zhRoot, PACKAGES_PREFIX, pkgDir));
});

const missingInZh = required.filter((f) => !zh.includes(f));
const extraInZh = zh.filter((f) => !en.includes(f));
const skippedEnOnly = en.length - required.length;

// ---- content sanity: a mirror that is a byte-identical copy of the en page,
// or contains almost no Chinese outside code blocks, is almost certainly an
// untranslated placeholder rather than a deliberate translation. ----
const CJK_RE = /[㐀-䶿一-鿿]/g; // CJK ext-A + unified ideographs
const CJK_RATIO_FLOOR = 0.1;

function stripCode(content) {
  return content
    .replace(/```[\s\S]*?```/g, "") // fenced code blocks
    .replace(/`[^`\n]*`/g, "") // inline code
    .replace(/<[^>]+>/g, ""); // inline HTML
}

function cjkRatio(content) {
  const text = stripCode(content).replace(/\s+/g, "");
  if (text.length === 0) return 0;
  return (text.match(CJK_RE) ?? []).length / text.length;
}

const contentErrors = [];
for (const f of required) {
  // Skip pages with no zh mirror: the file-level report below already lists
  // them, and readFileSync here would throw ENOENT before any report prints.
  if (missingInZh.includes(f)) continue;
  const enContent = readFileSync(join(root, f), "utf8");
  const zhContent = readFileSync(join(zhRoot, f), "utf8");
  if (zhContent === enContent) {
    contentErrors.push(
      `zh/ mirror is byte-identical to the en original (untranslated?): ${f}`,
    );
    continue;
  }
  const ratio = cjkRatio(zhContent);
  if (ratio < CJK_RATIO_FLOOR) {
    contentErrors.push(
      `zh/ mirror contains only ${(ratio * 100).toFixed(0)}% CJK outside code blocks (untranslated?): ${f}`,
    );
  }
}

if (missingInZh.length || extraInZh.length || contentErrors.length) {
  for (const f of missingInZh) {
    console.error(`[check-i18n] missing zh/ mirror: ${f}`);
  }
  for (const f of extraInZh) {
    console.error(`[check-i18n] unexpected zh page (no en original): ${f}`);
  }
  for (const msg of contentErrors) {
    console.error(`[check-i18n] ${msg}`);
  }
  console.error(
    "[check-i18n] en/zh page trees are out of sync or contain untranslated mirrors.",
  );
  process.exit(1);
}

console.log(
  `[check-i18n] OK — ${required.length} required en pages and ${zh.length} zh pages in sync` +
    (skippedEnOnly > 0
      ? ` (${skippedEnOnly} pages in en-only packages skipped).`
      : "."),
);
