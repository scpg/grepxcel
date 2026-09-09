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

> **Note (2026-08-30):** GitHub Models retired its free-tier endpoint (HTTP 410
> retirement brownout). The github backend is now disabled in grepxcel; the
> figures below are historical and cannot be reproduced without a replacement
> endpoint. The local backend is now the primary free option.

| Rank | Model | Backend | Value-recall | Fixtures ≥99% | $ |
|---|---|---|---|---|---|
| 1 | `mistral-ai/codestral-2501` | github *(retired)* | **86%** | 12/17 | was free |
| 1 | `openai/gpt-4.1` | github *(retired)* | **86%** | 12/17 | was free |
| 1 | `openai/gpt-4o` | github *(retired)* | **86%** | 12/17 | was free |
| 4 | `meta/llama-3.3-70b-instruct` | github *(retired)* | 63–81% | 8–12/17 | was free |
| 5 | **gemma-4-e4b** *(previous local default)* | local | 51% | 8/17 | free |
| 6 | qwen-coder-7b *(older local)* | local | 42% | 5/17 | free |
| — | **Qwen3-8B** *(current local default)* | local | *eval pending* | — | free |
| — | `deepseek/deepseek-v3-0324` | github *(retired)* | *rate-limited, partial* | — | was free |

The original benchmark cost **$0.00** (GitHub Models was free with a subscription;
local runs on your GPU). Figures are from two full runs; LLM sampling
(temperature 0.1) gives a few points of run-to-run variance.

Takeaways from the original eval:

- **`gemma-4-e4b` was the best *local* model** (now replaced by Qwen3-8B as the
  default). `gemma-4-e4b` was ahead of `qwen-coder-7b` by ~9 points, at similar
  weight size. **Gemma 4 was a generational leap over Gemma 2** (`gemma-2-9b`
  scored ~0%). Remaining failures were over-specified cell sequences and table
  syntax errors (e.g. `table:*` placed before `START:`).
- **GitHub Models free tier is gone.** As of 2026 GitHub retired the free-tier
  Models endpoint. The `github` backend is disabled in grepxcel (flip
  `_GITHUB_MODELS_ENABLED` in drafter.py to re-enable for a future replacement).
- **`claude` backend is the primary quality option** for non-local runs.
  Default model upgraded to `claude-sonnet-5` (Claude 5, high quality at $3/$15 per 1M tokens).
  Use `ClaudeBackend(model='claude-opus-5')` for the absolute best quality.
  A draft costs ~$0.003–0.01 per fixture at Sonnet 5 pricing.

**2026-08-30 model upgrade**: switched default local model from `gemma-4-E4B-it`
(51%, 4B params, 4.6 GB Q4_K_M) to `Qwen3-8B` (8B params, 4.7 GB Q4_K_M).
Qwen3-8B has not yet been benchmarked on the grepxcel fixture set; re-run the
evaluation harness to update the table.  Reasons for the upgrade:
- Twice the parameter count in the same VRAM footprint
- Qwen3 (2025) is a substantially newer generation than Gemma 4
- Existing gemma-4 failures included `table:*` syntax errors and 4096-token
  context overflows (now fixed to n_ctx=8192); a stronger model should help both
- GitHub Models retirement removes the strongest free alternative, raising the
  importance of local quality

For the hardest real-world tables, the cloud models extract end-to-end where
local models stall — e.g. the messy 9-column Bundesliga schedule is extracted in
full by GPT-4o and Claude Opus, while local models produce structurally close but
non-extracting patterns.

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

- **Fully offline / no account:** the default `local` backend (`Qwen3-8B`) —
  best available local model, runs entirely on your machine, nothing leaves it.
  On an 8 GB GPU all layers are offloaded; a draft takes ~10–15 s.
- **Best quality (metered):** `--backend claude` (Haiku for speed/cost, Sonnet
  or Opus for the toughest fixtures) — a few cents per draft.
- **GitHub Models (retired 2026):** the `github` backend is disabled.
  Implementation is preserved; flip `_GITHUB_MODELS_ENABLED` in drafter.py if
  GitHub provides a new endpoint.
- **Always expect to rename fields** in the drafted pattern. The model gets the
  structure and values; the names are yours to set.
- A draft is a *starting point*, not a finished pattern — open it, fix regexes
  and names, then run `extract`.
