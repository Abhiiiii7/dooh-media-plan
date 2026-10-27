# Programmatic DOOH Media Plan & Reporting Model

A sample digital out-of-home (DOOH) media plan for a fictional grocery retailer, **Brightside Grocers**, across New York, Los Angeles, and Chicago. It covers the full account-analyst loop: brief → venue selection → budget allocation → pacing and delivery tracking → client recap.

> **All data is simulated.** Venues, publishers, and the brand are fictional. I built this to practice the planning and reporting workflow used in programmatic DOOH, not to report on a real campaign.

**Tools:** Excel (live-formula model), Power BI (pacing dashboard), Python (data simulation and plan logic)

---

## The brief

| | |
|---|---|
| Brand | Brightside Grocers (fictional) |
| Objective | Drive store visits for a spring produce launch |
| Audience | Adults 25–54, primary grocery shoppers |
| Flight | Mar 2 – Mar 29, 2026 (28 days) |
| Budget | $75,000 — New York 45%, Los Angeles 35%, Chicago 20% |

## How venues were chosen

Each of the 131 venues (6 venue types, 6 fictional publishers) gets a **Venue Score**:

```
Venue Score = (Audience Index / 100) × Daypart Fit ÷ (CPM / metro average CPM)
```

- **Audience Index:** how much more likely the venue's audience is to be in the target (100 = average).
- **Daypart Fit:** how well the venue's traffic matches when the brand wants to be seen. Grocery trips skew to evenings and weekends, so those dayparts are weighted higher (weights are inputs on the Brief sheet).
- **CPM vs metro average:** cheaper than the local average scores higher.

Budget went to the highest-scoring venues in each metro until that metro's budget was used. Each venue was capped by its sellable impressions (40% of weekly) and by 15% of the metro budget, so the plan stays spread out.

**Result:** 49 of 131 venues, $74,900 of $75,000 placed, about 7.1M planned impressions.

| Venue type | Share of spend | Why |
|---|---|---|
| Mall | 32% | High audience index, strong weekend traffic |
| Grocery | 30% | Highest audience index (avg 148), right at the point of purchase |
| Gas station | 27% | Lowest CPMs, decent index |
| Transit | 6% | Big reach but average audience fit |
| Gym | 4% | Small venues, limited inventory |
| Office lobby | 0% | Lowest index, highest CPM, no weekend traffic |

## Pacing and delivery (as of day 14 of 28)

The model simulates two weeks of delivery from publishers, then tracks each publisher and metro against plan:

- **Pace Index** = delivered to date ÷ expected to date (1.00x = on plan)
- **Behind** below 0.90x, **Ahead** above 1.10x (thresholds are inputs)
- **Projected final** = delivered to date ÷ % of flight elapsed

**What it flags at day 14:**
- Overall pace 0.92x: 3.26M of 3.55M expected impressions delivered.
- **PumpPlay TV (gas stations) is at 0.71x.** At this rate it finishes about 640K impressions short.
- **New York is behind at 0.89x**, because it carries the most PumpPlay spend.
- Every other publisher is between 0.98x and 1.04x.

That's the point of the tracker: catch the problem at the halfway mark, while there's still time to ask the publisher for make-goods or move budget to MallVision, which is delivering ahead.

## What's in the repo

```
dooh-media-plan/
├── src/build_plan.py          # simulates inventory + delivery, builds the plan and the Excel model
├── data/
│   ├── venues.csv             # 131 venues: metro, publisher, type, impressions, CPM, index, daypart mix, score
│   ├── media_plan.csv         # 49 selected venues with planned spend and impressions
│   └── delivery_daily.csv     # daily delivered impressions and spend, days 1–14
├── output/DOOH_Media_Plan.xlsx
└── powerbi_guide.md           # step-by-step build of the pacing dashboard
```

### The Excel model (`output/DOOH_Media_Plan.xlsx`)

Every number is a live formula. Change the budget, a daypart weight, or the as-of day, and everything updates.

| Sheet | What it does |
|---|---|
| Brief | Campaign inputs: budget, metro split, daypart weights, assumptions, as-of day |
| Inventory | All 131 venues with Daypart Fit, Venue Score, and rank in metro (formulas) |
| Media Plan | Selected venues, planned spend (input), CPM lookups, planned and target impressions, metro totals vs budget |
| Delivery Data | Raw daily delivery from publishers |
| Pacing | By publisher and by metro: expected vs delivered, pace index, projected shortfall, status |
| Recap Template | KPIs that fill in automatically, plus a client story section (what worked, what didn't, what changed, next flight) |

Blue text = inputs, black = formulas, green = pulled from another sheet.

## Run it

```bash
pip install pandas numpy openpyxl
python src/build_plan.py
```

The seed is fixed (`42`), so you get the same data every run.
