# Fixture 16 — Bundesliga match schedule (`16_bundesliga`)

## Source

`data.xlsx` was **not** authored for this project. It was obtained from the
public internet to exercise grepxcel against a real-world, third-party Excel
file rather than a hand-crafted fixture.

It was found by searching **Bing** for Excel files with the query:

```
filetype:xlsx
```

The first result led, via a Bing redirect URL, to **bulibox.de**:

- Search/redirect link (as clicked):
  <https://www.bing.com/ck/a?!&&p=334ca830711da017a0261040109db036cf3ce8521399b80d60eaf0facace2603JmltdHM9MTc4MDE4NTYwMA&ptn=3&ver=2&hsh=4&fclid=2ab9f301-d69b-6a6a-1a05-e5f7d7876b58&psq=filetype%3axlsx&u=a1aHR0cDovL3d3dy5idWxpYm94LmRlL2Rvd25sb2Fkcy9zcGllbHBsYW4lMjAxLmxpZ2ElMjAyMDI1LTIwMjYueGxz>

- Decoded download URL (the `u=` parameter, base64-decoded):
  <http://www.bulibox.de/downloads/spielplan%201.liga%202025-2026.xls>

- Site: <http://www.bulibox.de> › downloads

## Contents

A German football (Bundesliga) 1. Liga 2025/2026 season match schedule
("Spielplan"). The workbook has two sheets:

- **`Spielplan`** (320 rows × 10 columns) — the match schedule table, with a
  banner row, a title/header block, and per-matchday fixtures (date, time,
  home/away teams, results).
- **`para`** (18 rows × 3 columns) — a small parameter/lookup sheet used by the
  original spreadsheet's formulas.

## Sanitisation

Because this is an unknown file downloaded from the public internet, it was
**not** committed as-is. The original download was uploaded to
<https://metadefender.com/> (OPSWAT MetaDefender), which scans and
**content-disarms-and-reconstructs (CDR)** the file. The sanitised version
produced by MetaDefender is what is stored here as `data.xlsx`.

## Why it's here

This is a messy, real-world file (merged cells, banner rows, mixed
date/time formats, formula-driven cells) — useful for testing how the engine
and the `draft` command cope with input that was never designed with grepxcel
in mind.

## Licensing / attribution

The file is third-party content from bulibox.de, included here only as test
data. It is not covered by this repository's licence. If redistribution is a
concern, replace it with an equivalently-structured synthetic fixture.
