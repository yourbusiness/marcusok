---
layout: home

hero:
  name: "MarcusOK"
  text: "Frontend libraries for admin products"
  tagline: Declarative, composable TypeScript libraries. Export 100k rows to Excel in under a second.
  image:
    src: /logo.svg
    alt: MarcusOK logo
  actions:
    - theme: brand
      text: Get Started
      link: /guide/01-getting-started
    - theme: alt
      text: Play
      link: /play
---

<InstallCommand />

## Ecosystem Highlights

<PackageHighlights />

<StatsBlock />

## Packages

<PackageCards />

## Performance Snapshot

Fast stream completes a 100k-row export in ~0.8s, while the same engine's Workbook path slows sharply beyond ~55k rows (17.5s at 100k).

<ClientOnly>
  <BenchmarkChart />
</ClientOnly>
