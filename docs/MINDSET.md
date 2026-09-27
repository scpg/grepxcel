# The mindset behind grepxcel

> This document is the *why*, not the *how*. For usage see the
> [README](../README.md) and [pattern-file.md](pattern-file.md).

`grepxcel` exists to fill one specific gap:

**the gap between "Excel as a human document" and "Excel as a reliable data source."**

Spreadsheets are where humans *enter* data — and that isn't going to change, because
Excel is a genuinely great tool for that. But the same flexibility that makes it
great for people makes it hostile as a data source: labels sit next to values,
tables start at arbitrary positions, cells get merged for looks, and the same report
arrives in a slightly different shape every month. Over time it only gets messier.

`grepxcel` is **not** trying to remove Excel. `grepxcel` tries to build a **trustworthy airlock around it**: the human surface on one side, clean structured JSON on the other, and a deliberate, auditable gate in between. The honest promise is not "replace Excel" — it is *"make data extracted from Excel safe to depend on programmatically."*

---

## Who this is for (and who it isn't)

The core audience is **developers and IT-savvy people** who are willing to invest some
effort up front in a pattern, in exchange for processing **hundreds or thousands of
files** reliably and feeding the result into a system that can actually use the data.

Writing a good pattern is an engineering task — regex, structure, types — and `grepxcel` is designed to reward that investment with extraction that is robust, auditable, and maintainable. That focused scope is intentional.

That said, **the person who authors the pattern and the person who uses it daily don't have to be the same person.** Non-technical users are explicitly accounted for: the web wizard lets anyone build or inspect a pattern by clicking cells in a browser, without writing a line of config. AI-assisted drafting and worked examples lower the bar further. A developer can own the pattern; the people who run it every day never need to touch the internals.
The engineering rigour stays; the barrier to *using* it comes down.

---

## The core principle: deterministic, or honestly broken

This is the heart of the tool.

An LLM can read a messy spreadsheet today, and that is useful — but at ~80% accuracy
it "needs to be rechecked," and rechecking does not scale to thousands of files.
At that volume nobody wants 80% *silent* accuracy. They want:

- **reproducible** results on the files that conform to the pattern, and
- a **loud, specific flag** on the files that don't — *which file, which cell, which anchor, and why* — so a human reviews only the small fraction that actually deviated.

A validated pattern gives exactly that. Same file + same pattern → same JSON, every
time. And when reality stops matching the pattern, grepxcel **fails loudly** instead of
guessing confidently. *Correct, or honestly broken — that's the goal.* Silent errors
are the worst outcome, so every mismatch grepxcel can detect is surfaced loudly. The
quality of extraction depends heavily on the pattern — a well-written pattern catches
more problems; a loose one may let unexpected values through.

This is why the AI-assisted `draft` command is a **bootstrap**, not the runtime: use
the LLM to get from a messy file to a *candidate* pattern in minutes, then a human
refines it once, and from then on the **deterministic engine** runs in production.

---

## DevOps discipline applied to spreadsheet ingestion

The one-line framing of the whole project:

> **Treat Excel as a reliable data source — with the DevOps discipline of
> observability, fail-fast, and proper logging.**

That isn't a slogan; it's a design compass. Each principle maps to something concrete:

| Principle | In grepxcel |
|---|---|
| **Config as code** | The pattern file is version-controlled, diffable, reviewable in a PR — infrastructure-as-code for extraction. |
| **Fail fast** | The parser rejects impossible/ambiguous patterns up front (e.g. a field used as both a value and a parent), not halfway through a run. |
| **Linting** | `validate-pattern` statically checks a pattern before you ever touch data. |
| **Health checks** | `doctor` is a readiness probe for the environment (deps, keys, model, proxy/TLS). |
| **Observability** | Per-field `-v/-vv` trace; structured, JSON-serialisable log records meant to be consumed by machines, not just read. |
| **Proper logging** | Severity levels, a file sink kept plain while the console is coloured — *logs for machines, output for humans.* |
| **Idempotency** | Deterministic: same input always yields the same output. |
| **Scriptable** | Non-zero exit codes on failure, so it composes inside a pipeline. |

And it tells us what to build **next**, by following the same compass to the parts
that aren't done yet:

- **Fleet-level observability** — not just "this file," but the *batch*: "3,988 of
  4,000 conformed (99.7%); here are the 12 that didn't and exactly why."
- **Conformance as an SLO** — track that rate over time; a drop is **drift detection**
  (someone changed the template upstream) and the basis for an **alert**.
- **Schema contracts** — validate the extracted JSON against the *consuming* system's
  expected schema, turning extraction into a producer/consumer contract test.
- **Lineage** — every output record can say which pattern version produced it.
- **A reject bin** — non-conforming files routed for human review rather than silently dropped.

---

## What success actually depends on

It is tempting to think the clever extraction engine is what wins adoption. It isn't.
People will drop a tool that extracts brilliantly but leaves them guessing why 12 files
failed — and keep a tool that is merely *good* at extraction but tells them precisely
what went wrong and is cheap to maintain.

So the things that decide whether this is *used in production* are the unglamorous ones:
the **exception report**, the **maintenance story** when layouts drift, the docs, the
error messages, the install, the release hygiene. We choose to build that half — CI
across Python versions, signed commits, security triage, a real release pipeline —
precisely because that half is what turns "a clever tool" into "a thing teams depend on."

A caution to ourselves: the DevOps framing is only worth anything if the *substance* is
real. "Observability" must mean an exception report that is genuinely excellent and logs
a machine can ingest — not the word on a README over a pile of print statements. The
audience this tool is for will smell the difference immediately.

---

## The ethos

Code can be generated now. **Judgment and direction can't.** Knowing where the tool
should go, what "done right" means, and insisting on the discipline to get there — that
is the rarer and more valuable half of building something.

I know the direction. I know quality when I see it. The AI handles the volume — the
typing, the boilerplate, the first draft. What it doesn't replace is the decision of
*what to build*, the refusal to ship something half-right, and the willingness to go
back and fix it properly when it isn't. That stays mine.
