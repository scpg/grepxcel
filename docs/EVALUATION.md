# Drafter Evaluation Notes

How we measure the quality of `grepxcel draft` — the command that analyses an
Excel file and asks an LLM to produce a starter pattern file.

This document records the **methodology** and the **findings** so the knowledge
travels with the repository rather than living in one developer's head. It is
intentionally candid about what works, what doesn't, and where the measurements
themselves are misleading. The numbers below come from a representative subset
of fixtures, not an exhaustive benchmark — treat them as direction, not gospel.

> The evaluation *harness* itself (model downloads, batch runner, caching) is
> exploratory tooling kept out of the shipped package. This file captures the
> conclusions; reproducing the exact numbers requires re-running that tooling
> against the `tests/fixtures/` set.

---

## Why evaluate at all

`extract` and `docs` are deterministic. `draft` is not: it hands an Excel
structure to a language model and hopes for a usable pattern. To choose a
default local model, to know when a cloud model is worth the cost, and to catch
regressions when we change the analysis prompt, we need a repeatable score.

---

## Methodology — execution-based scoring

We never compare pattern *text* (there are many valid ways to write the same
pattern). Instead we compare what the patterns **produce**:

1. For a fixture that has a hand-written reference `pattern.xlsx`, run the model
   to draft a candidate pattern from the fixture's `data.xlsx`.
2. Run `extract` twice on the same data — once with the reference pattern, once
   with the candidate.
3. Compare the two JSON outputs.

This rewards a draft that extracts the *right data*, regardless of how the
pattern is phrased.

### Two metrics, and why the strict one lies

**Key-score** — flatten both JSON trees to `dotted.path → value` leaves and
count exact matches. Strict and easy, but it conflates two very different
failures:

- the model extracted the **wrong data** (a real failure), and
- the model extracted the **right data under different field names** (cosmetic).

Because every leaf key is prefixed by its group (`item.name`, `item.sku`, …), a
model that calls the group `products` instead of `item`, or the first column
`product` instead of `name`, scores **0%** on a table it extracted *perfectly*.

**Value-recall** — ignore keys entirely; treat each side's leaf values as a
multiset and measure how many of the reference's values appear in the candidate.
This separates "missed the data" from "named it differently." (It measures
recall, not precision: a model that over-extracts can still score 100%.)

The gap between the two is the single most important thing this evaluation
revealed. Example, `02_product_catalog` with `qwen-coder-7b`:

| Metric | Score | Meaning |
|---|---|---|
| Key-score | 0% | every key mismatched |
| Value-recall | 100% (32/32) | every value extracted correctly |

The model did the job; it just chose `table_0.product` where the reference said
`item.name`. **Read key-score and value-recall together.** A low key-score with
high value-recall is a naming difference, not a defect.

---

## Findings (representative subset: fixtures 01, 02, 03, 05, 09)

### Local models, current code

| Fixture | Layout | qwen-coder-7b | gemma-4-e4b |
|---|---|---|---|
| 01_simple_invoice | key-value | 100% | 100% |
| 02_product_catalog | table | 100% | 0% |
| 03_purchase_order | mixed | 33% | 100% |
| 05_expense_report | tables | 10% | 0% |
| 09_sales_by_region | table | 100% | 100% |
| **value-recall avg** | | **~69%** | **~60%** |

(Values are value-recall. The corresponding key-scores averaged ~20% and ~18%
respectively — a stark illustration of how much the strict metric understates
extraction quality.)

Takeaways:

- **qwen-coder-7b** (the shipped default) and **gemma-4-e4b** are roughly
  comparable and far more capable than key-score suggests. They differ on *which*
  fixtures they nail rather than on overall strength.
- **Gemma 4 is a generational leap over Gemma 2.** The older `gemma-2-9b` scored
  ~0% across the board in earlier runs; `gemma-4-e4b` is now a credible local
  option at a similar (~5 GB) footprint.
- Genuine failures remain (e.g. `gemma-4-e4b` on `02`/`05`) — these are real,
  not naming artefacts (value-recall is also 0%).

### Cloud models

For the hardest real-world tables, capable cloud models extract end-to-end where
local 7B-class models stall. `claude-opus-4-8` handled a messy 9-column
Bundesliga schedule (merged-cell banners removed, an empty `split` column) and
extracted all 306 rows; the local models produced structurally close but
non-extracting patterns. Indicative cost: **~$0.02–0.04 per draft** for Opus.

Rough cloud ordering (quality, and cost per 1M tokens): `claude-opus-4-8`
(\$5/\$25) > `claude-sonnet-4-6` (\$3/\$15) ≈ `claude-haiku-4-5` (\$1/\$5).
Gemini backends exist in code but require GCP billing to be enabled.

---

## Known limitations of local models

1. **lbl: / var: confusion in tables.** Smaller models often define every table
   column as `lbl:` (anchors) and create no `var:` fields, then reference the
   `lbl:` names in the `DATA:` row — which never reach the output. The analysis
   prompt now tags each column explicitly (`HEADER (→lbl:) → DATA (→var:)`), which
   helps, but 7B-class models still hit a ceiling on wide, multi-column tables.
2. **Field-naming divergence.** Even when extraction is correct, models rarely
   guess the reference's exact group/field names — which is *fine in practice*
   (you rename to taste) but tanks the strict key-score.
3. **Multi-section footers** (free-text blocks below a table) can confuse smaller
   models into emitting stray instructions.

These argue for a future **semantic / LLM-as-judge** metric that scores "do these
two patterns extract the same information?" rather than exact key equality.

---

## Practical guidance

- **Default local model:** `qwen-coder-7b` — strong on key-value layouts, decent
  on tables, runs offline. `gemma-4-e4b` is a fair alternative worth trying.
- **Hard or unfamiliar tables:** use `--backend claude` (Opus for the toughest
  cases). It is the most reliable path to a working pattern on messy input, at a
  few cents per draft.
- **Always expect to rename fields** in the drafted pattern. The model gets the
  structure and values; the names are yours to set.
- A draft is a *starting point*, not a finished pattern — open it, fix regexes
  and names, then run `extract`.
