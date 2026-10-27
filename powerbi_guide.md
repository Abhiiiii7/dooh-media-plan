# Power BI Build Guide — DOOH Pacing & Delivery Dashboard

Build time is about 60–90 minutes. When you're done, your numbers should match the Excel Pacing sheet exactly (see the check at the end). That match is your proof the dashboard is right.

## 1. Load the data

**Get Data → Text/CSV**, load all three files from `data/`:
- `venues.csv` → rename table to **Venues**
- `media_plan.csv` → **Plan**
- `delivery_daily.csv` → **Delivery**

In **Power Query**, on Delivery:
1. Set `date` to type **Date**.
2. Add Column → Custom Column, name it `flight_day`:
   ```
   Duration.Days([date] - #date(2026, 3, 2)) + 1
   ```
   Set the type to **Whole Number**.

## 2. Model

**Model view.** Create two relationships, both many-to-one with single direction:
- `Plan[venue_id]` → `Venues[venue_id]`
- `Delivery[venue_id]` → `Venues[venue_id]`

Always slice by **Venues[publisher]**, **Venues[metro]**, and **Venues[venue_type]**, never by the columns of the same name on Plan or Delivery. Venues is the shared dimension, so plan and delivery line up.

**What-if parameter:** Modeling → New parameter → Numeric range. Name it `As Of Day`, set min 1, max 28, increment 1, default 14. This adds a slider.

## 3. Measures

Create a blank table called `_Measures` to hold these.

```DAX
Flight Days = 28

As Of = SELECTEDVALUE ( 'As Of Day'[As Of Day], 14 )

Pct Flight Elapsed = DIVIDE ( [As Of], [Flight Days] )

Planned Impressions = SUM ( Plan[planned_impressions] )

Planned Spend = SUM ( Plan[planned_spend] )

Delivered To Date =
CALCULATE (
    SUM ( Delivery[delivered_impressions] ),
    Delivery[flight_day] <= [As Of]
)

Expected To Date = [Planned Impressions] * [Pct Flight Elapsed]

Pace Index = DIVIDE ( [Delivered To Date], [Expected To Date] )

Pct Of Plan Delivered = DIVIDE ( [Delivered To Date], [Planned Impressions] )

Projected Final = DIVIDE ( [Delivered To Date], [Pct Flight Elapsed] )

Projected Shortfall = MAX ( 0, [Planned Impressions] - [Projected Final] )

Delivered Spend =
CALCULATE (
    SUM ( Delivery[delivered_spend] ),
    Delivery[flight_day] <= [As Of]
)

eCPM = DIVIDE ( [Delivered Spend], [Delivered To Date] ) * 1000

Pace Status =
SWITCH (
    TRUE (),
    [Planned Impressions] = 0, "Not in plan",
    [Pace Index] < 0.9, "Behind",
    [Pace Index] > 1.1, "Ahead",
    "On track"
)
```

For the delivery-curve chart:

```DAX
Cumulative Delivered =
VAR d = MAX ( Delivery[flight_day] )
RETURN
    CALCULATE (
        SUM ( Delivery[delivered_impressions] ),
        Delivery[flight_day] <= d,
        REMOVEFILTERS ( Delivery[flight_day], Delivery[date] )
    )

Cumulative Planned =
[Planned Impressions] * DIVIDE ( MAX ( Delivery[flight_day] ), [Flight Days] )
```

Format the measures: impressions `#,##0`, Pace Index `0.00"x"`, percents `0.0%`, eCPM and spend as currency.

## 4. Pages

### Page 1 — Pacing Overview
- **Slicers:** As Of Day (slider), Metro, Venue Type
- **KPI cards:** Planned Impressions, Delivered To Date, Pace Index, eCPM, Projected Shortfall
- **Line chart:** X = `Delivery[flight_day]`, Y = Cumulative Delivered and Cumulative Planned. The gap between the lines is the story.
- **Table by publisher:** Publisher, Planned Impressions, Expected To Date, Delivered To Date, Pace Index, Projected Shortfall, Pace Status.
  - Conditional formatting on Pace Index: background rules < 0.90 red, 0.90–1.10 neutral, > 1.10 blue.

### Page 2 — Publisher × Metro
- **Matrix:** rows = Publisher, columns = Metro, values = Pace Index. This shows where a publisher's problem sits.
- **Bar chart:** Projected Shortfall by Publisher, sorted descending.
- **Bar chart:** Delivered To Date by Venue Type.

### Page 3 — Recap
- Cards for the final KPIs.
- A text box for the client story: what worked, what under-delivered, what changed, and the recommendation for next time. Copy these from the Recap Template sheet in Excel.

## 5. Check against Excel (As Of Day = 14)

| Check | Should read |
|---|---|
| Planned Impressions (total) | 7,104,341 |
| Delivered To Date (total) | 3,256,109 |
| Pace Index (total) | 0.92x |
| PumpPlay TV Pace Index | 0.71x — **Behind** |
| New York Pace Index | 0.89x — **Behind** |
| LobbyLink | Not in plan |

If a number is off, the usual causes are a relationship pointing the wrong way, or a slicer that uses `Plan[publisher]` instead of `Venues[publisher]`.

## 6. Publish

Save as `DOOH_Pacing_Dashboard.pbix` in the repo's `output/` folder, and add 2–3 screenshots to the README. Interviewers look at screenshots before they open files.
