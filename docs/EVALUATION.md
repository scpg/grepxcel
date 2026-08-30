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

1. For a fixture that has a hand-written reference pattern (`*_pattern-manual.xlsx`), run the model
   to draft a candidate pattern from the fixture's `*_data.xlsx`.
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

## Findings (full set — all 17 fixtures, value-recall)

| Rank | Model | Backend | Value-recall | Fixtures ≥99% | $ |
|---|---|---|---|---|---|
| 1 | `mistral-ai/codestral-2501` | github | **86%** | 12/17 | free |
| 1 | `openai/gpt-4.1` | github | **86%** | 12/17 | free |
| 1 | `openai/gpt-4o` | github | **86%** | 12/17 | free |
| 4 | `meta/llama-3.3-70b-instruct` | github | 63–81% | 8–12/17 | free |
| 5 | **gemma-4-e4b** *(default local)* | local | 51% | 8/17 | free |
| 6 | qwen-coder-7b *(previous default)* | local | 42% | 5/17 | free |
| — | `deepseek/deepseek-v3-0324` | github | *rate-limited, partial* | — | free |

The whole benchmark cost **$0.00** (GitHub Models is free with a subscription;
local runs on your GPU). Figures are from two full runs; LLM sampling
(temperature 0.1) gives a few points of run-to-run variance — see the llama note.

Takeaways:

- **The free GitHub Models backend wins by a wide margin.** The large GitHub
  models clearly beat the local models (42–51%), at no dollar cost. For the best
  quality, use `grepxcel draft --backend github` (`--github-model openai/gpt-4.1`,
  `openai/gpt-4o`, or `mistral-ai/codestral-2501`).
- **Three-way tie at the top (86%):** `codestral`, `gpt-4.1`, and `gpt-4o` are
  indistinguishable — pick any. `codestral` (a code model) tying the GPTs fits,
  since drafting a structured pattern is a code-like task.
- **`llama-3.3-70b` is the least consistent** of the GitHub models (≈81% one run,
  ≈63% another) — run-to-run sampling variance plus occasional throttling. Still
  well ahead of the local models, but less reliable than the top three.
- **`gemma-4-e4b` is the best *local* model** (the shipped default), ahead of the
  previous default `qwen-coder-7b` by ~9 points, at a similar ~5 GB footprint.
  **Gemma 4 is a generational leap over Gemma 2** (`gemma-2-9b` scored ~0%). Its
  remaining failures are drafted patterns that over-specify cells (e.g. expecting
  a cell past the end of the sheet) — a local-model quality limit, not a crash.
- **`deepseek-v3` is not usable for bulk via GitHub** — its tight quota window
  throttles rapid runs (`Too many requests`); confirmed across two runs.
- An earlier 5-fixture subset over-flattered the local models (qwen ~69%) because
  it excluded the hard fixtures; the full set above is the honest picture.

For the hardest real-world tables, the capable cloud/GitHub models extract
end-to-end where local models stall — e.g. the messy 9-column Bundesliga
schedule (empty `split` column) is extracted in full by GPT-4o and Claude Opus,
while local models produce structurally close but non-extracting patterns.

Anthropic `claude` backend (metered) ordering and cost per 1M tokens:
`claude-opus-4-8` (\$5/\$25) > `claude-sonnet-4-6` (\$3/\$15) ≈ `claude-haiku-4-5`
(\$1/\$5); ~$0.003–0.04 per draft. Gemini exists in code but needs GCP billing.

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

- **Best quality, free:** `--backend github --github-model openai/gpt-4.1`
  (or `openai/gpt-4o`). Free with a GitHub subscription, and the strongest
  results in this eval. Needs `GITHUB_TOKEN` with `Models: read`.
- **Fully offline / no account:** the default `local` backend (`gemma-4-e4b`) —
  best of the local models, runs on your machine, nothing leaves it.
- **Metered alternative:** `--backend claude` (Opus for the toughest cases) — a
  few cents per draft; use it if you prefer Anthropic or have no GitHub access.
- **Avoid for bulk:** `deepseek-v3` via GitHub (rate-limited).
- **Always expect to rename fields** in the drafted pattern. The model gets the
  structure and values; the names are yours to set.
- A draft is a *starting point*, not a finished pattern — open it, fix regexes
  and names, then run `extract`.
