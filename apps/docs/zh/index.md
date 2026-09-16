---
layout: home

hero:
  name: "MarcusOK"
  text: "后台业务前端公共库"
  tagline: 声明式、可组合的 TypeScript 库集合，10 万行 Excel 导出不到 1 秒。
  image:
    src: /logo.svg
    alt: MarcusOK logo
  actions:
    - theme: brand
      text: 快速开始
      link: /zh/guide/01-getting-started
    - theme: alt
      text: 在线演示
      link: /zh/play
---

<InstallCommand />

## 生态亮点

<PackageHighlights />

<StatsBlock />

## 包生态

<PackageCards />

## 性能参考

Fast stream 在 10 万行数据下约 0.8s 完成导出，而同一引擎的 Workbook 整表序列化路径在约 5.5 万行后耗时陡增（10 万行 17.5s）。

<ClientOnly>
  <BenchmarkChart />
</ClientOnly>
