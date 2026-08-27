# TCA deck builder — KIC & NPS

One file: `tca_deck.py`. Copy it to the target machine with the order exports,
run it. Nothing else to install or configure.

```bash
python tca_deck.py --client KIC --data KIC/symboldetailstable_full_KIC.xlsx
python tca_deck.py --client NPS --data NPS/symboldetailstable_full_NPS.xlsx
```

Output lands in `<client>/output`:

| File | What it is |
|---|---|
| `<CLIENT>_TCA_deck.pptx` | the deck |
| `charts/*.png` | every exhibit standalone at 200dpi — **lift these into the corporate template** |
| `tables.xlsx` | every underlying table, one sheet each |
| `dark_regression.txt` | full regression output (KIC only) |
| `run_log.txt` | sanity report + reference check |

## Run this first

Before building anything, on the target machine:

```bash
python tca_deck.py --client NPS --data NPS/symboldetailstable_full_NPS.xlsx --probe
```

`--probe` reads the file and stops. It prints the sheet names, the first four
unheadered rows, every column with dtype and a sample value, which columns
mapped, the algo and market spellings with dark coverage for each, a sign check
on all four benchmarks, and the reference check. **Send me that output** — it's
everything needed to confirm the mapping before a single slide is built.

## Requirements

Python 3.7 or newer.

```bash
pip install pandas numpy matplotlib python-pptx openpyxl
pip install statsmodels     # optional: enables the dark regression
```

Without `statsmodels` everything else still runs.

## Test without data

```bash
python tca_deck.py --client KIC --sample
```

Generates a synthetic file and runs end to end. Output is stamped
`SAMPLE DATA` on every slide and named `..._SAMPLE_DATA.pptx`.

## The two clients

Everything client-specific lives in the `CLIENTS` dict at the top of the script.

|  | KIC | NPS |
|---|---|---|
| Published | 2,074 orders · $505m · **+6.2 bps** | 4,052 orders · $1,095m · **−1.4 bps** |
| Algos reviewed | VWAP, IIS, PROG | **VWAP, IIS only** (PROG/SDMA excluded) |
| Dark | 13.4% — the centrepiece | 1.0% — **section switched off** |
| Slides | 11 | 8 |

NPS runs four algos, but PROG and SDMA are a rounding error, so the review is
scoped to the two that carry the flow. The script reports what it excluded and
what share of value that left. The reference check runs on the **full** book
before the filter, so a scoped review still reconciles.

For NPS the dark section is off entirely — four slides about 1% of the book is
not advice. That slide becomes **"How Your Algos Execute"** instead: IIS puts
89% into the auction, VWAP posts 68%, and that difference is the conversation.

## Sign convention

**Positive = saving, negative = cost**, matching the published post-trade report.
Blue = saving, red = cost, everywhere. If an export ever uses the opposite sign,
set `POSITIVE_IS_SAVING = False` — the reference check will tell you.

## Report-sourced tables

Two things are **not** in the order-level export and are transcribed from each
client's report into `CLIENTS`:

- **venue segmentation** (auction / post / take / dark) — the order file carries
  the dark share only
- **industry breakdown** — there is no sector column at all

Update both when the period changes. Everything else is computed from the file.

## If the columns don't match

Headers match case-insensitively, ignoring spaces, dots and underscores, so
`Dark %`, `dark%` and `DARK_%` all resolve. On failure the script exits with the
headers actually present. Two things to check:

- `HEADER_ROW` — defaults to `1` (row 2), because the export puts a merged
  `Symbol Details` banner on row 1. Override with `--header-row 0`.
- `COLUMNS` at the top of the script.

Only `Value Exec` and `ClientAlgo` are strictly required, plus **one** benchmark
column and **one** dark column. `Dark %` is rebuilt from `Dark Value` or
`Dark Qty` if absent.

## Dark markets

```python
"dark_markets": ["Hong Kong", "Japan", "Australia"]
```

Every dark comparison is restricted to these. An order in Taiwan never had the
option of a dark fill, so including it in the "no dark" group compares *markets*
rather than *venues*. Markets in the list with no executions are never mentioned.

## Deck structure

Both decks: cover · executive summary · where you trade · what drives the number ·
**by industry** · order size · venue mix · recommendations.

KIC adds three dark slides: did it help · or were those orders just easier ·
where you can get it.

Every sentence is generated from the numbers. Nothing is hardcoded — if the data
says dark didn't help, the deck says dark didn't help. Recommendations scale to
the size of the prize: a market worth −0.2 bps is not called "the biggest drag".

## How the dark analysis avoids the obvious trap

Dark fills land on liquid, tight-spread names that were easy anyway, so a naive
dark-vs-lit split credits dark with a difficulty effect it didn't cause. So:

1. only dark-capable algos (IIS does no dark — including it would make "no dark"
   mean "IIS");
2. only dark-enabled markets;
3. "no dark fill" vs "any dark fill" — a real control group;
4. the same split repeated **inside** bands of %ADV, spread, order size and
   market;
5. a notional-weighted regression with size / spread / participation / duration
   controls and market fixed effects, reported alongside the equal-weighted fit.

It stays associational and the deck says so: dark fills happen when there's
contra flow, and contra flow may itself mean the order was easier.

## Warnings it will raise

- **a self-referential benchmark.** IIS executes in the closing auction, so
  measured against the close its slippage is exactly zero on every order — that
  is arithmetic, not performance, and the benchmark cannot rank it. The script
  tests which case it is by comparing each order's execution price against the
  benchmark price: if they match, the benchmark is tautological and the algo
  should be judged on arrival instead; if they differ, the field really is
  unpopulated
- buys and sells with opposite signs (export not side-adjusted)
- percentage fields arriving as 0–1 fractions
- weighted and equal-weighted regression fits disagreeing in sign
- any algo with fewer than 100 orders — asterisked in every table

## Options

```
--client KIC|NPS   which client config (default KIC)
--data PATH        the order export
--probe            inspect the file and stop
--sample           synthetic data
--out DIR          output directory (default <client>/output)
--header-row N     0-indexed header row
--sheet NAME|N     worksheet
--no-deck          tables and charts only
```
