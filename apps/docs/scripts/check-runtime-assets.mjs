// Verify the runtimeAssets copy mechanism (file form + dir form) against a
// synthetic fixture tree — the pdf-preview integration added directory
// assets (169 cmaps), and a regression here would only surface as a silent
// CJK degradation in the deployed site (cmap fetch 404), never as a build
// failure. Runs as part of `pnpm test` for the docs app.
//
// Node >= 22.18 / 24 loads ../.vitepress/runtime-assets.ts via built-in type
// stripping (the file is erasable-TS only; keep it that way).
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import {
  copyRuntimeAssets,
  stableAssetNamesFor,
} from "../.vitepress/runtime-assets.ts";

function fail(msg) {
  console.error(`[check-runtime-assets] FAIL: ${msg}`);
  process.exit(1);
}

/** Build a fake installed-package tree: <root>/<name>/package.json + files. */
function fakePkg(root, name, files) {
  const dir = join(root, name);
  mkdirSync(dir, { recursive: true });
  writeFileSync(join(dir, "package.json"), JSON.stringify({ name }));
  for (const [rel, content] of Object.entries(files)) {
    const f = join(dir, rel);
    mkdirSync(dirname(f), { recursive: true });
    writeFileSync(f, content ?? "x");
  }
  return dir;
}

const tmp = mkdtempSync(join(tmpdir(), "runtime-assets-"));
try {
  // Fixture: a package whose entry sits in build/, with assets both at the
  // package root (cmaps/, standard_fonts/) and next to the entry (a wasm).
  // resolveFrom accepts absolute paths (require.resolve semantics), so the
  // fixture tree needs no node_modules wiring.
  const engineDir = fakePkg(tmp, "fake-engine", {
    "build/entry.mjs": "export {}",
    "build/asset.wasm": "wasm-bytes",
    "cmaps/Adobe-Japan1-0.bcmap": "bcmap-bytes",
    "cmaps/LICENSE": "adobe-license",
    "standard_fonts/LiberationSans-Regular.ttf": "ttf-bytes",
  });
  const entryAbs = join(engineDir, "build", "entry.mjs");

  const pkgs = [
    {
      npmName: "@marcusok/fake-preview",
      runtimeAssets: [
        // file form: file relative to the resolved entry's directory
        { resolveFrom: entryAbs, file: "asset.wasm", to: "engine.wasm" },
        // dir form: dir relative to the package root of resolveFrom
        { resolveFrom: engineDir, dir: "cmaps", to: "cmaps" },
        { resolveFrom: engineDir, dir: "standard_fonts", to: "fonts" },
      ],
    },
  ];

  const publicDir = join(tmp, "public");
  copyRuntimeAssets(pkgs, publicDir);

  // file form landed
  if (!existsSync(join(publicDir, "engine.wasm")))
    fail("file asset was not copied");
  // dir form copies recursively, including nested LICENSE files
  if (!existsSync(join(publicDir, "cmaps", "Adobe-Japan1-0.bcmap")))
    fail("dir asset did not copy nested files");
  if (!existsSync(join(publicDir, "cmaps", "LICENSE")))
    fail("dir asset dropped the LICENSE file (must ship with the data)");
  if (!existsSync(join(publicDir, "fonts", "LiberationSans-Regular.ttf")))
    fail("dir asset did not honor the custom destination dir");

  // `through`: resolve the engine from *another package's* dependency tree
  // (the pdf-preview integration resolves pdfjs-dist through the package).
  // The through package is addressed by absolute path so the temp fixture
  // needs no workspace wiring.
  fakePkg(join(tmp, "node_modules"), "fake-through", {});
  fakePkg(join(tmp, "node_modules"), "fake-engine", {
    "cmaps/Through.bcmap": "bcmap-bytes",
  });
  const throughDir = join(tmp, "node_modules", "fake-through");
  copyRuntimeAssets(
    [
      {
        npmName: "@marcusok/fake-preview-through",
        runtimeAssets: [
          {
            resolveFrom: "fake-engine",
            through: throughDir,
            dir: "cmaps",
            to: "cmaps-through",
          },
        ],
      },
    ],
    join(tmp, "public-through"),
  );
  if (
    !existsSync(join(tmp, "public-through", "cmaps-through", "Through.bcmap"))
  ) {
    fail("through-resolved dir asset was not copied");
  }

  // missing source fails loudly (not a silent 404 later)
  let threw = null;
  try {
    copyRuntimeAssets(
      [
        {
          npmName: "@marcusok/fake-preview",
          runtimeAssets: [
            { resolveFrom: entryAbs, file: "nope.wasm", to: "x" },
          ],
        },
      ],
      join(tmp, "public2"),
    );
  } catch {
    threw = true;
  }
  if (!threw) fail("missing source asset did not throw");

  // stableAssetNamesFor: file form in, dir form out
  const names = stableAssetNamesFor(pkgs);
  if (!names.has("engine.wasm"))
    fail("file asset basename missing from stableAssetNamesFor");
  if (names.has("cmaps") || names.has("fonts"))
    fail(
      "dir asset leaked into stableAssetNamesFor (would mis-key Rollup's assetFileNames)",
    );

  console.log(
    "[check-runtime-assets] ok (file form, dir form, through resolution, LICENSE passthrough, loud failure, stable-name exclusion)",
  );
} finally {
  rmSync(tmp, { recursive: true, force: true });
}
