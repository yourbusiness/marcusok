# Format Fidelity

Cell display text goes through a compensation layer that fixes a set of engine defects — each one verified against real behavior and locked by unit tests. This page documents what is restored exactly and what is approximated.

## Builtin number formats: Excel's actual table

ECMA-376 defines builtin id 14 as `mm-dd-yy`, but Excel renders it as `m/d/yyyy` (en-US). The engine returns the ECMA string, and its `loadFormatTable` override cannot win (the builtin table always takes lookup precedence — a structural property of the engine). The preview therefore carries its own **Excel-behavior table** for builtin ids 0–22 and 37–49 in en-US behavior (14 → `m/d/yyyy`, 18 → `h:mm AM/PM`, 46 → `[h]:mm:ss`, …). The locale-dependent ids in those ranges (14–22 among them) are where Excel's rendering drifts from ECMA-376's strings; the ranges are covered **wholesale** — many ids there, `46` included, are locale-invariant — because the engine's builtin table always wins lookup precedence, so re-resolving every id through the preview's own table is the only way any override can take effect. Ids are resolved to format _strings_ before formatting. The East-Asian groups 27–36 and 50–58 (only ever written by East-Asian-locale Excel/WPS — zh-CN files store `2025年1月21日` as `numFmtId=31`) are covered with zh-CN behavior strings; the engine's `getBuiltinFormat` returns `undefined` for them, so without the table those dates rendered as raw serials.

## Literal compensation in numeric sections

The engine drops literal characters inside numeric sections: `¥#,##0.00` renders as `1,234.50` (currency gone), `#,##0.00;-#,##0.00` renders `1,234.50` for -1234.5 (sign gone), `"neg"0` renders `5` (quoted text gone). The layer extracts safe edge literals (quoted runs, escapes, currency symbols, `-`, `(`, `)`, spaces), hands the numeric core to the engine with the **absolute value**, and re-joins — which also restores Excel's exact sign semantics (a negative section without `-` shows no sign, exactly like Excel).

Restored: currency prefixes/suffixes, units (`0.0" kg"`), text prefixes (`"Total: "@`), negative signs, accounting parentheses. Thousands-scaling commas (`#,##0,` → value ÷ 1000 per trailing comma, `0.0,,"M"` → `1.5M`) are divided out before the engine formats the value. Display rounding follows Excel's decimal semantics (15 significant digits, half-up — `2.675` with `0.00` shows `2.68`, where the engine's binary `toFixed` would show `2.67`). `[DBNum1/2/3]` digit-substitution modifiers are stripped so the numeric section still renders (full Chinese-numeral rendering is out of scope), and section-level `General` (`[Red]General`, WPS's `G/通用格式`) renders the value instead of the literal text.

Approximated: literals interleaved _between_ digits (phone-style `000-0000`) are still dropped; `_` width-padding and `*` fill are intentionally dropped.

## Dates and times: fully self-implemented

Two engine defects made a full rewrite necessary (the engine renders `mm` after `h` as **month** — `0.5` with `h:mm` shows `12:12` instead of `12:00` — and `dddd` as the day-of-month number):

- `mm` minute/month disambiguation by adjacency (after `h`, before `s` → minutes)
- Weekday names (`ddd`/`dddd`), month names (`mmm`/`mmmm`), 12-hour + `AM/PM`
- Elapsed durations `[h]:mm:ss`, `[m]:ss`, `[s]` (all broken in the engine)
- The bracket-less `mm:ss` (which the engine misreads as month:second) renders as time-of-day minutes:seconds — elapsed minutes is the _bracketed_ `[m]:ss` semantics in Excel, and the two are kept distinct
- Fractional seconds (`ss.0`, builtin 47) round to the displayed precision
- date1904 workbooks: serials are shifted +1462 days before formatting
- Quoted literals (`yyyy"年"m"月"d"日"`) and the 1900 ghost-day (serial 60) are handled

All time math uses UTC components (the industry-standard approach that avoids DST gaps), so a file renders identically in every timezone.

## General

`General` is self-implemented: 15 significant digits with trailing zeros trimmed (the engine emits 16-digit float tails), uppercase `TRUE`/`FALSE` for booleans, raw strings for text.

## Theme and indexed colors

The engine's read path drops `theme`/`tint`/`indexed` colors to `null` — that would lose the default font color (theme 1) and every themed fill on real-world files. The overlay layer unzips only `styles.xml` + `theme1.xml` (a few KB each), recovers the raw color specs, resolves them (theme palette with the 0/1・2/3 index swap, ECMA tint on HSL luminance, the indexed palette) and merges by array index. Explicit RGB colors take precedence; a failed overlay degrades gracefully to RGB-only.

## Self-closed style elements (real-Excel id alignment)

Real Excel files write the default border as a self-closed `<border/>` (self-closed `<font/>`/`<fill/>` occur occasionally too). The engine's parser skips self-closed elements, so its `fonts`/`fills`/`borders` arrays come out one item short and every `cellXfs` id reference shifts: `borderId=0` picks up the file's border #1, and the trailing reference dangles (verified with hand-built OOXML — the engine's _own_ writer emits expanded forms, which is why round-trips through the exporter never trip this). The overlay layer therefore rebuilds all three collections from `styles.xml` in document order — self-closed entries included as empty items — keeping ids aligned with `cellXfs` exactly; the engine arrays serve only as the fallback when `styles.xml` itself cannot be read.

## Colors from format codes

`[Red]` / `[Blue]` / `[Color 3]` section colors are extracted and applied as inline cell colors (the engine's plain `formatCell` mangles sections containing them entirely).
