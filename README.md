# AI Vibe-Coding Guide

**English** | [简体中文](README.zh.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Doc Format](https://img.shields.io/badge/Doc-DOCX%20%2B%20PDF-2E5496.svg)](docs/)
[![Pages](https://img.shields.io/badge/Pages-32-blue.svg)](docs/AI-Vibe-Coding-Guide.pdf)
[![Tables](https://img.shields.io/badge/Tables-34-2A7B7B.svg)](docs/AI-Vibe-Coding-Guide.pdf)
[![Bilingual](https://img.shields.io/badge/Docs-中文%20%7C%20English-success.svg)](#)
[![Built with](https://img.shields.io/badge/Built%20with-python--docx-3776AB.svg)](https://python-docx.readthedocs.io/)

> A practical, opinionated guide that upgrades ad-hoc AI "vibe coding" into a **verifiable three-stage pipeline**: plan the work with a Code-Architect prompt, implement it with an agentic coding CLI, then verify the result against a fixed engineering standard (P1–P11).

---

## Why this guide exists

Vibe coding is fast but has two structural defects: it is **not reproducible** (the same sentence produces a different implementation on a different day) and it is **not verifiable** (the AI says "done", but nobody ever defined what *done* means). This guide fixes both by inserting a planning stage in front and a verification stage behind.

## The three-stage pipeline

| Stage | Tool | Role | Output | Done when |
|---|---|---|---|---|
| ① **Plan** | Chatbox (Code-Architect persona) | Translate a vague request into a structured task package | Subtask list + executable Prompt suite | Every subtask has inputs, outputs and acceptance criteria |
| ② **Implement** | Claude Code (`--effort high`) / DeepSeek Harness (`dsh`) | Write code, run tests, commit | Runnable code + Conventional Commits | Local tests pass, working tree clean |
| ③ **Verify** | Hermes Agent + P1–P11 standard | Audit the artifact against engineering specs | Audit report + evidence (measured numbers) | All 11 specs rated, gates can actually block a merge |

## What's inside

| Chapter | Contents |
|---|---|
| **1. Overview** | From vibe coding to an engineering pipeline; five core principles; toolchain positioning |
| **2. Stage ① — Prompt planning with Chatbox** | The complete reusable *Code-Architect* system prompt; five-step working method; the 7-element task-card template; the five-dimension decomposition framework (data / logic / interface / presentation / infrastructure); a worked end-to-end example (8 subtasks, dependency DAG, two full production-grade prompts); six common failure modes |
| **3. Stage ② — Implementation** | Wiring Claude Code to a third-party Anthropic-compatible endpoint; effort levels (`low`/`medium`/`high`/`xhigh`/`max`) and when to use each; the five-step single-task loop; parallel orchestration with git worktrees; real CLI flag reference; what DeepSeek Harness (`dsh`) is and when to prefer it; traceable commit conventions |
| **4. Stage ③ — Verification (P1–P11)** | Eleven engineering specs, each with Role / Context / Task / Constraints / Input / Output / Acceptance Criteria **plus a "Hermes execution layer"** describing exactly how to run and verify it, with copy-pasteable commands |
| **5. Risks** | Six structural risks (over-specification, stale tool versions, test anti-patterns, multi-language drift, open-source vs internal, gate over-tightening) + six AI-coding-specific risks (hallucinated APIs, silent failure, excessive permissions, secret leakage, context drift, unreviewed merges) + a 10-item pre-release checklist |
| **Appendix** | Prompt 7-element cheat card; Hermes verification command reference; recommended project layout; glossary |

The P1–P11 standard covered in Chapter 4:

| ID | Spec | ID | Spec |
|---|---|---|---|
| P1 | Coding standard (naming / style / structure) | P7 | Test strategy & the test pyramid |
| P2 | Git commit & branching workflow | P8 | Unit-test standard |
| P3 | Lint / format toolchain automation | P9 | Integration-test standard |
| P4 | Documentation & comment standard | P10 | E2E-test standard |
| P5 | Code Review standard | P11 | CI/CD pipeline & quality gates |
| P6 | Open-source governance & release | | |

## Highlights

- **Evidence over claims.** Every verification step is expressed as a command you can run or a number you can check — never as "the AI said it works".
- **Tool-accurate.** All CLI flags and commands were captured from a real environment on 2026-09-11 (Claude Code v2.1.126, `deepseek-ai/deepseek-harness`). Version numbers are deliberately *not* hard-coded into the guidance, so the document does not rot.
- **Honest about limits.** DeepSeek Harness is explicitly flagged as a *developer preview* with breaking changes; the doc says so instead of pretending otherwise.
- **Verification is falsifiable.** The guide insists that a CI quality gate is only real once you have deliberately pushed a failing change and watched it block — a gate that exists only in YAML is called a *fake gate*.

## Quick start

The guide is delivered as a Word document with a matching PDF:

| File | Purpose |
|---|---|
| [`docs/AI-Vibe-Coding-Guide.docx`](docs/AI-Vibe-Coding-Guide.docx) | Editable Word source — fork it and adapt the standard to your team |
| [`docs/AI-Vibe-Coding-Guide.pdf`](docs/AI-Vibe-Coding-Guide.pdf) | Print/share version, 32 pages, table of contents with page numbers |

```
# Just open the Word document and read it.
# To adapt it to your own team standard, edit tools/build_guide.py and rebuild:

pip install python-docx
python tools/build_guide.py docs/AI-Vibe-Coding-Guide.docx
```

## Preview

<!-- Images are referenced through an absolute CDN URL on purpose: repo-relative image
     paths are served from raw.githubusercontent.com, which is DNS-blocked in mainland
     China. An absolute URL makes GitHub proxy the image through camo.githubusercontent.com,
     which stays reachable. -->

![Cover](https://cdn.jsdelivr.net/gh/Misaka4396/AI-Vibe-Coding-Guide@main/assets/preview-1-cover.png)

![Table of contents](https://cdn.jsdelivr.net/gh/Misaka4396/AI-Vibe-Coding-Guide@main/assets/preview-2-toc.png)

![P1/P2 specification pages](https://cdn.jsdelivr.net/gh/Misaka4396/AI-Vibe-Coding-Guide@main/assets/preview-3-specs.png)

## Repository layout

```
AI-Vibe-Coding-Guide/
├─ docs/
│   ├─ AI-Vibe-Coding-Guide.docx      # editable Word deliverable (32 pages)
│   └─ AI-Vibe-Coding-Guide.pdf       # rendered PDF (TOC updated)
├─ assets/                            # preview images used in this README
├─ tools/
│   └─ build_guide.py                 # regenerates the DOCX from the content source
├─ README.md                          # this file (English)
├─ README.zh.md                       # 简体中文版
├─ CHANGELOG.md                       # Keep a Changelog
└─ LICENSE                            # MIT
```

## Verification

Document build verified on 2026-09-11 with real tooling, not by inspection alone:

| Check | Method | Result |
|---|---|---|
| Opens without repair | Microsoft Word via COM automation (`Documents.Open`) | PASS — opened, 1714 paragraphs / 34 tables |
| Renders to PDF | Word `ExportAsFixedFormat` | PASS — 32 pages, 1.30 MB |
| Table of contents resolved | `TablesOfContents(1).Update()` then re-paginate | PASS — page numbers present (30 → 32 pages after TOC fill) |
| Text layer is real text | `pymupdf` per-page `get_text()` + keyword search | PASS — 11 hits for "代码架构师", 19 for "P11", 58 for "Hermes" |
| No missing-glyph boxes | scan extracted text for U+FFFD / □ | PASS — 0 occurrences |
| No leftover markup | scan extracted text for literal `**` / backticks | PASS — only intentional glob patterns (`tests\unit\**\*.ts`) remain |

## License

[MIT](LICENSE) © 2026 Misaka4396
