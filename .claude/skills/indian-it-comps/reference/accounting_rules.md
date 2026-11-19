# Extraction & classification rules (read by the extractor subagent)

You **record and classify**. You never compute EBITDA, EV, multiples, unit conversions or
adjusted figures — Python does that. Every number you write must be visible on the cited page.

## General
- Use the **current-year column** (FY ending on `period_end`), never the prior year.
- Values **as printed**, in the report's unit (`unit_as_printed`). Brackets `(1,234)` → `-1234`.
  `-` / `—` / `Nil` → `0`. Indian grouping `1,48,814` → `148814`.
- `uid` and `printed_page` come from the `### SOURCE` header of the page file you read the value on.
- `label` = the line label as printed (trimmed). `note` = the note reference if shown.
- Consolidated only. If a page header says "Standalone", it is the wrong page — flag it.

## P&L
- `expense_lines`: every line between *Total income* and *Total expenses*, including finance costs,
  D&A and any separate FX line. Credits (e.g. inventory increase printed in brackets) negative.
- `exceptional_charge`: positive if it reduces profit, negative if it is a gain. null if none.
- `exceptional_tax_effect`: only if the exceptional note states the tax on it; else null.
- `fx_line_charge`: a separate FX gain/loss line on the face (loss positive). null if none.
- `fx_in_other_income`: the foreign-exchange gain/(loss) line **inside the Other Income note**,
  signed as it contributes to Other Income (gain positive, loss negative, e.g. `(646.59)` → `-646.59`).
  null if the Other Income note has no FX line. Never put the face FX line here.
- `fx_in_other_income_extra`: any *additional* FX lines inside the Other Income note (e.g. a separate
  forward-contract / hedge gain or loss line), same sign convention. Omit or [] if there is only one.
- `share_of_equity_investees`: share of profit/(loss) of associates / JVs (equity-accounted investees)
  printed on the face between Total expenses and Profit before exceptional items and tax. Signed as it
  contributes to profit (loss negative). null if the face has no such line.
- `eps_basic` / `eps_diluted`: total (continuing + discontinued). `eps_diluted_continuing`: the
  continuing-operations diluted EPS if printed separately, else null.
- `profit_continuing_ops` / `profit_discontinued_ops`: null if the P&L has no such split.
- `discontinued_attrib_owners`: only if a note splits discontinued profit between owners/NCI.
- `pat_nci`: null if there is no NCI line (do not write 0 unless printed).

## Balance sheet
- Lease liabilities are separate from borrowings. If a borrowings note *includes* leases, flag it.
- `finance_lease_receivable`: list every face line (current and non-current).
- `equity_method_investments`: investments in associates / JVs on the face; null if none.

## Cash components (`cash_components`, `restricted_carveouts`)
List **every line** of the cash & cash equivalents note and the other bank balances note.
`restricted: true` for: earmarked / unpaid / unclaimed dividend accounts, margin money,
deposits under lien or held against guarantees, escrow, balances not freely usable.
If a restriction is disclosed only in narrative ("deposits include ₹X held against bank
guarantees"), keep the line `restricted: false` and add the ₹X to `restricted_carveouts`.

## Investment components (`investment_components`)
List **every line** of the current and non-current investment notes.
- `treasury`: mutual funds (liquid/debt/index), bonds, NCDs, CPs, T-bills, term deposits with banks
  or financial institutions, government securities, units of InvITs / REITs.
- `strategic_equity`: equity stakes in other companies (unquoted or quoted) held for business reasons.
- `restricted`: pledged / lien-marked investments.
- `other`: anything else — explain in `comment`.
Exclude investments in associates/JVs here (those go to `bs.equity_method_investments`).

## Debt-like items (`debt_like_items`)
From other financial liabilities notes, business-combination notes and fair-value notes:
contingent consideration / earn-outs, deferred consideration payable, put-option or redemption
liabilities over NCI (e.g. "financial liability for future acquisition"). Use the current-year
carrying amount. If a note splits it into current and non-current columns, record **two items**
(one per column, label suffixed " - current" / " - non-current") — never add them yourself.
If the same liability appears in several notes, record it **once** from the balance-sheet-linked
note (other financial liabilities) and mention the others in `comment`.

## ESOPs (`esop_plans`)
One entry per plan / instrument: options outstanding at year end and weighted-average exercise
price (INR per share). RSUs at nil/face-value exercise price: `instrument: rsu`, `waep_inr` as printed.

## Shares
- `shares_outstanding_year_end`: number of equity shares issued and fully paid at year end
  (count, not amount) from the share capital note.
- `treasury_shares`: shares held by an ESOP trust and deducted from equity; null if none.
- `weighted_avg_diluted_shares`: from the EPS note, if printed.

## Flags
Add a short flag string for anything unusual: restatements, units that differ between pages,
truncated notes (table cut off at a page edge), pages that look standalone, a value that appears
on the face but not in the note (or vice versa), anything you could not find.
