---
name: indian-it-comps
description: Build an FY trading-comparables workbook (EV/Revenue, EV/EBITDA, P/E with peer statistics) for Indian listed IT services companies from consolidated Annual Report PDFs in input_pdfs/, writing output/trading_comps.xlsx with live formulas and a page-cited audit log.
disable-model-invocation: true
---

# Indian IT Services Trading Comps

Run from the workspace root. `S` = `.claude/skills/indian-it-comps/scripts`. Every command is
`cd <S> && python <script> ...`. All arithmetic is in Python or in Excel formulas — never compute,
convert or adjust a number in your own text.

## Context budget rules (strict)
- **Never** read a PDF, a `work/*/pages/*.txt` file, `extracted.json`, or the workbook yourself.
  Page text is read only by extractor subagents. You read only script summaries.
- **Never** open or edit `tests/golden/*` (ground truth, human-owned) or `config/policy.yaml`
  (accounting policy, user decisions only). Never show golden values to a subagent.

## Steps

1. **Prep (deterministic):** `cd <S> && python pipeline.py prep`
   - One line per peer, then `TO_EXTRACT: <co_ids>`.
   - `MISSING_PDF` → list for the user at the end; continue with the rest.
   - `INDEX_PROBLEMS` / `slice problems` → do not extract that company; report it.

2. **Extract (parallel LLM):** for every co_id in `TO_EXTRACT`, launch one `general-purpose`
   subagent, **all in a single message** so they run concurrently. Prompt (fill `<co_id>`):

   > You are the extractor subagent for an M&A trading-comps pipeline. Read and follow exactly:
   > .claude/skills/indian-it-comps/reference/extractor_brief.md
   > All paths are relative to the workspace root (your starting directory).
   > Substitutions: <co_id> = `<co_id>`, <root> = `.` (the workspace root).
   > Read only the files the brief lists. Do not open the PDF, the tests/ folder, other companies'
   > folders, or any existing extracted.json — extract from scratch and overwrite it.
   > Validate with: cd .claude/skills/indian-it-comps/scripts && python validate.py <co_id>
   > Shell is Git Bash on Windows; Python is `python`. Reply with the one-line status from the brief
   > plus up to 5 short lines. No numbers dump.

   Wait for every subagent. Keep only its status line and flags.

3. **Finish (deterministic):** `cd <S> && python pipeline.py finish`
   - Validates every extraction, fetches market data, builds and verifies the workbook.
   - `VALIDATION FAIL <co>` → re-launch that company's extractor **once**, adding the failure
     lines to the prompt; then re-run finish. If it fails again, leave it out and report it.
   - `MARKET BLOCKED <co>` → yfinance share count is unexplained vs AR + `config/share_bridge.yaml`.
     Do **not** guess a bridge entry. Report the gap; the user supplies the issuance/buyback with source.
   - `MARKET DATA ERROR` → retry finish once (Yahoo is flaky), then report.
   - `VERIFY` must say `PASS`. If not, report the diff lines verbatim — do not hand-edit the workbook.

4. **Report to the user** (short):
   - Table per company: status, EV/Revenue, EV/EBITDA, P/E (adjusted) — copy from the `BUILD:` lines.
   - Peer statistics come from the workbook; say how many companies are in them and which are
     excluded (`EXCLUDED FROM STATS`) and why.
   - Golden scores (if shown), missing PDFs, blocked or failed companies, and any extractor flag
     that changes the valuation: post-balance-sheet deals, share issuances, restatements, listed subsidiaries.
   - Path: `output/trading_comps.xlsx`.

## Adding a peer
Add it to `config/peers.yaml` (co_id, name, `nse` ticker, `pdf` filename, `listed_subsidiaries`) and
drop `input_pdfs/<pdf>`. A listed subsidiary's NCI must be valued at market (policy) — flag it to the
user if the pipeline has no market-value input for it yet.

## Reference (load only if a step needs it)
- `reference/accounting_rules.md`, `reference/extractor_brief.md`, `reference/extracted.schema.json` — extractor contract
- `config/policy.yaml` — locked accounting policy; `config/share_bridge.yaml` — documented share issuances
- `scripts/score_golden.py` — developer regression against `tests/golden/` (run automatically by finish)
