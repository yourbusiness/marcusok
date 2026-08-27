---
"@marcusok/excel-exporter": patch
---

修复多路径输入校验与进度契约问题：

- stream/SheetJS 路径遇 `NaN`/`Infinity` 不再写出非法 `<v>NaN</v>` XML（此前 Excel 会判定文件损坏），改为写入可见字符串 `"NaN"`/`"Infinity"`；workbook 路径行为不变（空单元格）。
- 新增 `merges` 统一校验（整数、`row`/`col` ≥ 0、`rowspan`/`colspan` ≥ 1、不超出数据区、互不重叠），三条路径一致以 `{ success: false, error }` 失败并指明问题项，替代原先直接生成损坏文件的行为。
- 重复 sheet 名三条路径统一报错（此前 stream 路径产出损坏文件、workbook 路径意外降级丢样式、SheetJS 路径静默改名）。
- worker 失败降级 SheetJS 的路径修复 `onProgress(1)` 双发，恢复「trailing 1 恰好一次」契约（types.ts）。
- stream 模式特性跳过警告改为递归扫描列树，多级表头下嵌套的 `width`/`style`/`headerStyle` 不再被静默丢弃，兑现 README「dropped with a warning」承诺。
- 新增 17 个测试（输入校验 + worker 进度契约），总计 84 个用例。
