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

## The short deck (4 slides)

```bash
python tca_deck_v2.py --client KIC
python tca_deck_v2.py --client NPS --slides 5
```

`tca_deck_v2.py` is the four-slide version for a short meeting:

1. **The number, and what we see** — cover and summary merged
2. **Where the money is made and lost** — by country
3. **What explains it** — chosen from the data, see below
4. **What we would change**

Slide 3 is not fixed. If one country-and-side cell is worth more than half the
whole result, that cell *is* the story and the buy/sell chart leads — which is
what happens for NPS, where Japan buys alone exceed the entire shortfall.
Otherwise the venue mix leads, because that is what decides the fix — which is
what happens for KIC, where India crossing the spread on 100% of its volume is
the actionable finding.

Output: `<client>/output/<CLIENT>_TCA_short.pptx`. `--slides 5` adds a slide on
algorithm and order size.

Everything else — data, analysis, wording, speaker notes — is identical to
`tca_deck.py`. Only the slide selection differs.

## The client deck

```bash
python tca_deck.py --client KIC --simple
python tca_deck.py --client NPS --simple
```

Eight slides, built **entirely from the published report figures** in `CLIENTS`
— no order file is read, so nothing on any slide can be synthetic. Cover and
KPIs · summary · by algorithm · by market · by industry · by order size ·
where volume executes · what we would change.

Output: `<client>/output/<CLIENT>_TCA.pptx`, plus every chart as a standalone
PNG and `tables.xlsx`.

The longer deck (`--data`) adds the order-level dark analysis on top of this.

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

## Where each number comes from

**The published report is the truth for every aggregate.** The order file is
used only for what the report cannot express.

| Exhibit | Source |
|---|---|
| Headline KPIs, algo, country, market cap, order size, side, industry, venue | **report** (`CLIENTS`) |
| Dark: no-dark vs any-dark, controlled splits, availability by market, regression | order file |
| Worst orders by money | order file |
| Benchmark matrix (all four benchmarks) | order file |

Two of the report tables *cannot* be derived at all — there is no sector column,
and the order file carries only the dark share, not the auction/post/take split.
The rest could be derived, but aren't: taking them from the report means every
exhibit agrees with what the client already holds, and the deck is correct even
before anyone has run the order file.

Every report table is typed in by hand, so each is checked on load: weights must
sum to 100 and contributions must sum to the published headline. A transcription
slip is reported at the top of the run, not discovered in a client meeting.

The recomputation is not wasted — it becomes a **per-algo reconciliation**.
Every run compares orders, notional and bps for each algo against the report and
flags any gap. A total that reconciles while one algo is wrong — two errors
cancelling — no longer slips through. Both views land in `tables.xlsx` as
`by_algo` and `by_algo_computed`, with the differences in `algo_reconciliation`.

Update the report tables in `CLIENTS` when the period changes.

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
