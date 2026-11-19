# Extractor subagent brief

You extract one company's FY figures from pre-sliced Annual Report pages into
`work/<co_id>/extracted.json`. Placeholders: `<co_id>`, `<root>` (workspace root).

## Inputs (read only these)
1. `<root>/.claude/skills/indian-it-comps/reference/accounting_rules.md` — classification rules. Read first.
2. `<root>/work/<co_id>/manifest.json` — page list (`pages[].uid`, `why`), `targets`, `reporting_unit`.
3. `<root>/work/<co_id>/pages/<uid>.txt` — one printed page each. Read the pages whose `why` you need.
4. `<root>/.claude/skills/indian-it-comps/reference/extracted.schema.json` — output contract.

Do **not** open the PDF, other companies' folders, or any other file. Do not search the web.

## Procedure
1. Read the face pages (`face:bs`, `face:pl`) and fill `pl`, `bs`, `expense_lines`.
2. Read note pages by `why` tag and fill `cash_components`, `restricted_carveouts`,
   `investment_components`, `debt_like_items`, `esop_plans`, `shares`.
   `debt_like_scan` pages are hints — they may or may not contain a balance.
3. Write `<root>/work/<co_id>/extracted.json` (UTF-8, valid JSON, matching the schema).
4. Run: `python <root>/.claude/skills/indian-it-comps/scripts/validate.py <co_id>`
5. If it reports FAIL lines, fix **only** the failing fields by re-reading the cited page, rewrite the
   file, and re-run. Maximum 3 fix rounds. Never change a correct value to force a check to pass —
   if a check fails because the report itself does not reconcile, add a flag explaining why.

## Return value
Reply with **one line only**:
`<co_id>: <PASS|FAIL> checks=<passed>/<total> flags=<n> rounds=<k>`
followed by up to 5 short lines naming any remaining failures or important flags.
Do not paste extracted numbers or page text into your reply.
