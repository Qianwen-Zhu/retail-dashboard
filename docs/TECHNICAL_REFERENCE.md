# Retail Business Dashboard — Build Documentation

Source of truth for the **Retail Business Dashboard** weekly deliverable. Excel template lives at `inputs/Dashboard_Template.xlsx`; the data-feeding query is `sql/weekly_retail_actuals.sql`; weekly outputs land in `outputs/YYYY-MM-DD/`.

This doc covers: what each section measures, how it's computed, what data caveats exist, which are hard limitations vs. fixable, and a runnable checklist of the fixable items.

---

## 1. Architecture overview

### 1.1 Output schema (one row per `week_start × dashboard_channel`)

| Column | Notes |
|---|---|
| `week_start` | Monday of the Mon-Sun week |
| `dashboard_channel` | One of: `Amazon`, `Amazon 3P`, `Walmart`, `Home Depot`, `Other Omni`, `TikTok Shop`, `DTC US`, `DTC CA`, `International`, plus virtual rollup rows `Total Retail`, `All Channels` |
| `sell_in_gross_rev_{actual,target,cam_actual,noncam_actual}` | $ from `BUDGET_HW_WITH_ACTUALS_WEEKLY` |
| `sell_through_units_{actual,target,cam_actual,noncam_actual,cam_target,noncam_target}` | Device-multiplied units from `TEST_FORECAST_SELL_THROUGH_PRODUCT_LINE` |
| `new_customers_{actual,cam_actual,noncam_actual}` | **Definition varies by row** — see §3.3 |
| `promo_spend_actual` | $ from inlined 2026 promo CTEs |
| `ads_onsite_actual` / `ads_dsp_actual` / `ads_total_actual` | Only Amazon + TikTok Shop populated; others NULL |

### 1.2 Data sources

**Sell-in / Sell-through / Forecast**

| Source | Used for | Native week | Notes |
|---|---|---|---|
| `DATA_MART.FINANCE.BUDGET_HW_WITH_ACTUALS_WEEKLY` | Sell-in $ + target | ISO Mon-Sun (END_DATE = Sun); year-boundary rows split | User defaults to GROSS_REV |
| `DATA_MART.FINANCE.TEST_FORECAST_SELL_THROUGH_PRODUCT_LINE` | Sell-through units + target + LY | Mon-Sun (WEEK_START = Mon) | PRODUCT_LINE × CHANNEL grain; units pre-multiplied by `DEVICE_UNIT_MULTIPLIER` |

**Promo — campaign tracker + per-channel sell-through volumes**

| Source | Used for | Native week | Notes |
|---|---|---|---|
| `DATA_MART.FINANCE.ALL_PROMOS_BY_CAMPAIGN_26` (and `_25`) | Promo tracker — campaigns × discount × CHANNEL rec ID filter | n/a (campaign window) | 26 = list-string format (`CONTAINS()` filters); 25 = plain string (`=` / `IN`) |
| `SRC_RAW.WMT_SCINTILLA.OMNI_SALES` | Walmart 1P promo sell-through (in-store + online) | Sat-Fri | Channels `'Walmart 1P Store'`, `'Walmart 1P Ecomm'`. **Reliable from WM Sat 2026-04-18 only** |
| `SRC_CLEAN.SUPPLYCHAIN_SPS.ITEM_HOMEDEPOT` + `SRC_CLEAN.SUPPLYCHAIN_SPS.ACTIVITY_HOMEDEPOT` | HD promo sell-through (in-store + online) | daily (`period_ending_date`) | Split via `SPS_LOCATION_ID` for in-store vs online; one specific ID = online |
| `DATA_MART.SUPPLYCHAIN_SPS.DATA_BESTBUYDOTCOM` | BestBuy promo sell-through | Weekly (Sun-Sat, `PERIOD_START_DATE` = Sun) | `.com` only — BB in-store not tracked |
| `ecomm_edm_prod.edm_tiktok_staging.tiktokorderlineitems` | TTS promo (order-level `seller_discount` / `exchange_currency_rate`) | daily (`order_created_at`) | Filter `_DATON_SOURCEVERSION_INTEGRATION_ID = 'tiktok_us'` (exclude MX) |
| `DASHBOARD.SUPPLYCHAIN_VENDOR_CENTRAL.VENDOR_SALES_REPORT_MANUFACTURING` + `DATA_MART.REFERENCE.AMAZON_1P_MAPPING` | Amazon promo sell-through (1P / Vendor Central only) | daily (`START_DATE`) | 1P-volume basis only; 3P units NOT in promo calc |
| `DATA_MART.FINANCE.UPC_SKU_MAPPING` | UPC → SKU lookup helper | n/a (mapping) | Used by HD + Walmart promo CTEs |

**Ads**

| Source | Used for | Native week | Notes |
|---|---|---|---|
| `SRC_CLEAN.SALES_AMAZON_ADS.US_AMAZON_ADS_SPONSORD_PRODUCTS_REPORT_CLEAN` (+1P variant) | Amazon SP on-site ads | daily | 1P starts Feb 2026; 3P starts Dec 2025 |
| `DATA_MART.FINANCE.AMAZON_DSP_ADS_TEMP` | Amazon DSP ads | Sun-Sat (week_start = Sun) | **Manual upload**; values shifted `+1 day` to Mon |
| `ecomm_daton.daton_dbt_staging.tiktok_ads_us_test_gmv_max_campaign_daily` | TTS GMV Max ads | daily (`stat_time_day`) | Excludes `'LIVE GMV Max'`; dedup latest `daton_batch_runtime` |

**New customers**

| Source | Used for | Native week | Notes |
|---|---|---|---|
| `SRC_CLEAN.SALES_AMAZON_ADS.SB_NTB_1P_US` + `SD_NTB_1P_US` | Amazon NTB (Ads-attributed new customers) — 1P only, SB + SD summed | daily | Column = `ntb_purchases` (same for both). SB earliest 2026-03-27, SD earliest 2026-04-17 (small); nothing before 2026-03-27. SP NTB not yet available. |
| `RUDDERSTACK_EVENTS.USER_ATTRIBUTE.USER_ATTRIBUTES` | New customer first-binding event | n/a | `FIRSTDEVICE_FIRSTBINDINGTIME`; `D2C_HARDWARE_ORDERS` + `FIRST_D2C_HW_ORDER_CREATETIME` for Shopify-attribution flag |
| `DEVICE_DATA.CLEAN.GE_USER_INFO` + `DEVICE_DATA.ARCHIVE.GE_RELATION_USER_DEVICE` + `DATA_MART.REFERENCE.PRODUCT_MAPPING` | New customer's first-bound device classification | n/a | Cam / Non-Cam split (Unknown lumped into NonCam) |

### 1.3 Week convention

- **Mon-Sun, `week_start = Monday`** throughout.
- Output rows include only weeks whose Sunday is strictly before `CURRENT_DATE` (`per_channel` WHERE filter). Today = Sun 5/31 → most recent week shown = Mon 5/18. Today = Mon 6/1 → most recent = Mon 5/25.
- NULL `week_start` rows from future-week budget entries with NULL END_DATE in `BUDGET_HW_WITH_ACTUALS_WEEKLY` get dropped by this same filter (NULL < CURRENT_DATE is NULL).
- **Year-boundary fix (2026-05-31)**: `BUDGET_HW_WITH_ACTUALS_WEEKLY` splits the cross-year ISO week into two partial rows (e.g., 2025-W53: Mon-Wed only; 2026-W01: Thu-Sun only). Using `DATE_TRUNC('week', END_DATE)` (Snowflake default Mon-start) maps both to the same Mon-Sun `week_start = 2025-12-29`, summed by GROUP BY.

### 1.4 Channel mapping (sources → dashboard_channel)

| Dashboard | Sell-in SUB_CHANNEL | Sell-through CHANNEL | Promo raw key | Notes |
|---|---|---|---|---|
| Amazon | `Amazon 1P` | `Amazon 1P` | `amazon` (1P-volume) | 1P only — 3P moved to own row 2026-06-01 |
| Amazon 3P | `Amazon 3P` | `Amazon 3P` | — (none) | Separate row; contributes to All Channels only, NOT Total Retail. No promo / ads / NTB. |
| Walmart | `Walmart` | `Walmart 1P` + `Walmart 3P` (3P added 2026-06-09 → sell-through now 1P+3P, matches sell-in) | `walmart` (1P Scintilla/manual + 3P order-lines) | Sell-through 1P+3P; **promo now 1P+3P** (3P added 2026-06-22, rides Ecomm campaigns); 1P pre-cutover from manual backfill |
| Home Depot | `Home Depot`, `Home Depot Canada` | `Home Depot`, `Home Depot US Online` | `hd-instore`, `hd-online` (US only) | HD CA in sell-in only (~$20-50K/wk); sell-through and promo for CA unavailable |
| Other Omni | `Best Buy`, `Costco`, `Other Retail` | `Best Buy`, `Costco.com` | `bestbuy` only | Costco / Other Retail promo (no campaign tracker rows) |
| TikTok Shop | `TikTok Shop`, `TikTok Shop Mexico` | `TikTok-US`, `TikTok-MX` | `tts` (both US + MX) | Ads (GMV Max) US-only — no MX source |
| DTC US | `Wyze.com US` | `DTC-US` | — (not pulled) | — |
| DTC CA | `Wyze.com CA` | `DTC-CANADA` | — | — |
| International | `International` | — (no rows) | — | — |
| (excluded everywhere) | `Shared Cost`, `DTC Service` | — | — | Finance buckets, not real retail channels |

### 1.5 Virtual rollup rows

- **`Total Retail` = Amazon (1P) + Walmart + Home Depot (US + CA) + Other Omni + TikTok Shop (US + MX)** (5 sections; per user 2026-06-01 — sums "the sections below" verbatim)
- **`All Channels` = everything in per_channel** (Total Retail + Amazon 3P + DTC US + DTC CA + International). Includes Shopify D2C, International, Mexico, Amazon 3P — literally everything.
- **New customers** in rollups DOES NOT use the per-channel SUM — overridden by separate-source LEFT JOINs (see §3.3)

---

## 2. The SQL pipeline

`sql/weekly_retail_actuals.sql` is the local read-only wrapper for
`DATA_MART.RETAIL_DASHBOARD.WEEKLY_RETAIL_ACTUALS`. The view definition owns
the following pipeline sections:

1. **Sell-in** — `BUDGET_HW_WITH_ACTUALS_WEEKLY` filtered to `END_DATE >= '2025-12-29'`, `DATE_TRUNC('week', END_DATE)` for week_start, channel mapping CASE, Cam/NonCam split via `PRODUCT_GROUP`.
2. **Sell-through** — `TEST_FORECAST_SELL_THROUGH_PRODUCT_LINE` filtered to 2026 weeks, channel mapping, Cam/NonCam from `PRODUCT_GROUP`.
3. **Promo (2026 only, inlined)** — five sub-blocks (HD / Walmart / BB / TTS / Amazon 1P), each builds `(week_start, sku, channel, promo_spend_usd)`. Walmart is special: pulls 1P units from Scintilla `OMNI_SALES` raw (same logic as `walmart_omni_base` in `wyze_sell_through.sql`), subchannels `'Walmart 1P Store'` + `'Walmart 1P Ecomm'`, promos expanded by CHANNEL rec IDs (`recyJpz5ObkkqwN3Z` → Store, `recb89H7psdHrEXoA` → Ecomm) so combined / online-only / store-only campaigns multiply only their applicable subchannel's units. **3P promo NOT added here** (deferred — sell-through includes 3P but promo stays 1P until confirmed).
4. **Ads** — Amazon SP daily (1P + 3P UNION) → weekly; Amazon DSP TEMP shifted +1 day to Mon; TTS GMV Max daily → weekly. Pivoted by dashboard_channel into `ads_onsite_actual` / `ads_dsp_actual` / `ads_total_actual`.
5. **New customers** — 4 CTEs chained (`nc_new_users` → `nc_user_nid` → `nc_first_bound_device` → `nc_classified`) mirroring `new_user_by_channel.sql`. Then 3 outputs: `new_users_total_weekly` (R5 All Channels), `retail_new_customers_weekly` (R18 non-Shopify with Cam/NonCam pivot, Unknown lumped into NonCam), `amazon_ntb_weekly` (R39 SB NTB orders, 1P only).
6. **`per_channel` join** — `sell_in_metrics` FULL OUTER JOIN `sell_through_metrics` FULL OUTER JOIN `promo_metrics` FULL OUTER JOIN `ads_metrics` LEFT JOIN `amazon_ntb_weekly`. `new_customers_actual` populated only for Amazon row at this level. WHERE filter: `DATEADD('day', 6, week_start) < CURRENT_DATE`.
7. **Rollup CTEs** — `total_retail` (filters per_channel to retail channels + LEFT JOIN `retail_new_customers_weekly` for `new_customers_actual` override) and `all_channels` (sums everything + LEFT JOIN `new_users_total_weekly`). Both use the same nested-subquery pattern so the new_customers SUM is replaced, not just added.
8. **Final SELECT** — UNION ALL per_channel + total_retail + all_channels, ordered by week_start DESC then dashboard_channel.

---

## 3. Section-by-section explanation

### 3.1 Sell-in (R3 / R8 / R27 / R71 / R100 / R129 / R153)

**What**: weekly gross revenue at the wholesale (shipment-to-retailer) layer.
**Source**: `BUDGET_HW_WITH_ACTUALS_WEEKLY.GROSS_REV` (user-confirmed default; not `NET_REV` or `PRODUCT_PROFIT`).
**Cam/NonCam**: `PRODUCT_GROUP = 'Cameras'` vs `'Non-Cameras'` (Service product_group excluded by NULL filter on dashboard_channel).
**Target**: `BUDGET_GROSS_REV` from same table — populated for ALL 2026 weeks (we get year-end forecast values).
**Week convention**: ISO Mon-Sun. Year-boundary handled by `DATE_TRUNC('week', END_DATE)` (see §1.3).
**Display**: dashboard values in $k (divide by 1000 in Excel).

### 3.2 Sell-through units (R4 / R12 / R32 / R76 / R105 / R134 / R158)

**What**: weekly units sold by the retailer to end customers (device-level, multi-pack expanded).
**Source**: `TEST_FORECAST_SELL_THROUGH_PRODUCT_LINE.GROSS_SALES_UNITS`. The view is built on top of `WYZE_SELL_THROUGH_DATA` and **pre-multiplies SKU units by `DEVICE_UNIT_MULTIPLIER`** before aggregating to PRODUCT_LINE × CHANNEL × WEEK_START grain.
**Cam/NonCam**: `PRODUCT_GROUP = 'Cameras'` vs `'Non-Cameras'`.
**Target**: `TARGET_UNITS` from the same view (forecast joined in via FULL OUTER JOIN). Only matches at PRODUCT_LINE × CHANNEL × week level; non-matched sell-through rows have NULL target.
**Cam/NonCam target**: derived similarly via PRODUCT_GROUP.

### 3.3 New customers (R5 / R18 / R39 only)

**Three sources for three rows — NOT cross-row summable.**

| Row | Source | Definition |
|---|---|---|
| **R5 All Channels** | `USER_ATTRIBUTES` filtered to first-binding events | `COUNT(DISTINCT USER_ID)` per week (company-wide, no channel) |
| **R18 Total Retail** | Same source + acquisition_channel classification | `COUNT(DISTINCT USER_ID)` where `acquisition_channel = 'Other'` (i.e., not Shopify D2C); Cam/NonCam from first-bound device → `PRODUCT_MAPPING.PRODUCT_FAMILY`; Unknown bucket lumped into NonCam so Cam + NonCam = Total |
| **R39 Amazon** | `SB_NTB_1P_US` + `SD_NTB_1P_US` | `SUM(ntb_purchases)` (SB + SD NTB orders, 1P only; Amazon 14-day click/view attribution). SP NTB not yet available. |

**Implementation**: `nc_classified` CTE (mirrors `new_user_by_channel.sql` chain) builds per-user (week, channel, cam_flag). Aggregates feed into rollup CTEs which OVERRIDE the SUM(per_channel) — important because Amazon's NTB orders aren't summable with Total Retail's USER count.

**Target (2026-06-29)**: New-Customers target is a flat **52%** (constant, unit %). `% of Target` rows are **percentage-points** = actual share − 52%. Total Retail: R22 = `R21 − R24` (R21 = NC ÷ sell-through units), R23 (YTD) = `ΣNC ÷ Σunits − 52%`. Amazon (R44–48) shows actual (SB+SD NTB) + 52% target only — its `As % of sales units` / `% of Target` / `% of Target - YTD` rows stay **blank** (NTB orders vs sell-through units = mixed denominator). Earlier "52% × sell-through units" target (2026-06-22) is superseded. Row numbers per section live in `build_outputs.py` constants (template is 212 rows as of 2026-06-29).

### 3.4 Promo per channel (R44 / R83 / R112 / R141 / R165)

**Amazon promo (R44)**: based on Amazon 1P (Vendor Central) units × campaign discount.
- ⚠ **Known limitation**: Amazon 3P (Seller Central) sell-through volumes NOT in promo calc. Promo from 3P campaigns × 3P units is missing. Same limitation that exists in CAC tracker. Reproduces existing behavior.

**Walmart promo (R83)**: 1P (Store + Ecomm, **subchannel-aware**) + **3P** (added 2026-06-22). 1P units from Scintilla `OMNI_SALES` (≥ cutover) / manual backfill (< cutover, bucketed Ecomm); 3P units from `SRC_CLEAN.WALMART_3P.*` order-lines, riding the Ecomm (`recb89...`) campaigns (no 3P-specific rec ID). Toggle 3P off by removing `wmt_3p_weekly` / `wmt_3p_promo_26` from the `wmt_promo_spend` unions.
- Promos with CHANNEL containing `recyJpz5ObkkqwN3Z` apply to Store units; `recb89H7psdHrEXoA` applies to Ecomm. Combined promos (both rec IDs) expand to two rows.
- 2025 promos (table `_25`, plain-string CHANNEL) historically only used `recb89H7psdHrEXoA` → hardcoded to Ecomm.
- WM-suffix CROSS JOIN preserved (some Walmart SKUs sell under both bare and `WYZECOPBWMT`-style suffix).

**HD promo (R112)**: standard ALL_PROMOS_BY_CAMPAIGN_26 with `recFlcYMqRgq507TO` (in-store) or `recSQ7fMin1DdMliW` (online) rec IDs. Sell-through from raw `SPS.ACTIVITY_HOMEDEPOT` split by `SPS_LOCATION_ID`.

**Other Omni promo (R141)**: only BB tracked. Costco / Other Retail / HD Canada have no campaign tracker rows.

**TikTok Shop promo (R165)**: order-line `seller_discount` from `tiktokorderlineitems`, both US and MX (no `_DATON_SOURCEVERSION_INTEGRATION_ID` filter as of 2026-06-01).

### 3.5 Amazon Ads (R49 / R54 / R59)

- **R49 Ads — On-site**: SP 1P (`US_1P_AMAZON_ADS_SPONSORD_PRODUCTS_REPORT_CLEAN`) + SP 3P (`US_AMAZON_ADS_SPONSORD_PRODUCTS_REPORT_CLEAN`), unioned and aggregated to week. Amazon SP 1P starts Feb 2026; 3P starts Dec 2025.
- **R54 Ads — DSP**: `AMAZON_DSP_ADS_TEMP` (manual upload, Sun-Sat → Mon shift +1 day). Coverage: ~13 LY weeks (2025-03 to 2025-05) + CY from late 2025 onwards.
- **R59 Ads (On-site + DSP)**: Excel formula `=E49+E54` (per column). Recalculates automatically when R49 or R54 change.

### 3.6 TikTok Shop Ads (R170)

- TikTok US GMV Max from `tiktok_ads_us_test_gmv_max_campaign_daily`, excluding LIVE campaigns, deduplicated on `(campaign_id, stat_time_day)` via `daton_batch_runtime DESC`.

### 3.7 Home Depot Ads (R127) / Walmart / Other Omni Ads

- **Home Depot Ads — wired (2026-06-03).** `hd_ads_weekly` reads `DATA_MART.FINANCE.HD_ADS_DAILY` (Orange Access PLA + Banner onsite + Google PMAX offsite; manual upload, no API) → feeds the `Home Depot` ads row.
- **Walmart Ads — wired (2026-06-22).** `walmart_ads_weekly` reads `DATA_MART.FINANCE.WALMART_ADS_DAILY` (`SUM(ad_spend)` by `report_date`, 1P+3P combined) → daily → Mon-Sun → feeds the `Walmart` ads (detail) row. **Other Omni Ads — still no weekly feed** (FY26 plan estimate only). (Detail Ads rows show actual + Budgeted only since 2026-06-29; `% of budget` lives on Promo + Total Marketing, finance-based.)

### 3.8 Other Marketing (R69 / R101 / R133 / R159 / R191)

- **No *weekly* actual feed.** Plan supplemental has the FY26 budget. Monthly Finance Actuals now populate the TikTok Shop Other Marketing row (see §3.9).
- Includes things like: Amazon MDF, retail in-store displays, retail marketing packages, TTS commissions & samples, affiliate spend.

### 3.9 Finance Actuals (Monthly) — real GL spend (wired 2026-06-09, source → NetSuite CSV 2026-06-18)

- **Source:** `inputs/finance/Marketing_Spend_2026YTD.csv` — a **long** table pulled through `DATA_MART.RETAIL_DASHBOARD.MARKETING_SPEND_YTD`, whose underlying source is `DATA_MART.FINANCE.NETSUITE_CPAM_DETAILS` (Wyze Labs US / Sub2). The local wrapper is `sql/finance_actuals.sql`. Refresh ~monthly: the root launcher queries the view and **replaces the CSV only when a newer month exists** (no date in the filename). Replaces the old `Wyze_Marketing_Spend_2026YTD.xlsx` (wide `Channel × Category` tab), now unused.
- `load_finance_actuals()` reads the CSV → `{(channel, category): {(year,month): $}}`; `apply_finance_actuals()` writes the inline **Finance Actuals (Monthly)** row (2nd row of each metric block) at the LAST Mon-Sun week of each month (`last_mon_of_month_col`, same placement as Budgeted), `$ → $k`. Pre-filled `% of budget` formulas (finance ÷ budget, same column) auto-compute. HTML gets matching `finance`-class rows.
- **Categories** (mapping + sign handled in the query — see `finance_actuals.sql`): `Promos` → Promo row, `Advertising` → Ads row, `Other Mktg` → Other Marketing row.
  - `Promos` = accts 4094 + 4095 + 4098 (contra-revenue, sign-flipped to positive in the query).
  - `Advertising` = acct 6601.
  - `Other Mktg` = accts 6102 + 6107 + 6108 + 6602 + 6604 + 6605 + 6610.
  - **Discounts EXCLUDED** (4092 Early Pay + 4093 Retail Terms — dropped in the query, never folded into Promo).
  - Channels missing a category (Amazon/Walmart/Other Retail have no dashboard Advertising row) just leave that finance row blank — the loader still loads those keys, but `FINANCE_ROW_MAP` has no row for them so they're not written.
- **Channel mapping** (`FINANCE_ROW_MAP` / `FINANCE_HTML_MAP` — unchanged; the query emits exactly these 5 channel keys) — 5 sections populate:
  - `Amazon 1P` → Amazon (Promo R45, Other Mktg R69)
  - `Walmart` → Walmart (Promo R89, Other Mktg R101)
  - `Home Depot` → Home Depot (Promo R121, Advertising R127, Other Mktg R133) — **includes Home Depot Canada** (folded in by the query)
  - `Other Retail` → Other Omni (Promo R153, Other Mktg → "Total marketing" R159) — = Best Buy + Costco + Other Retail
  - `TTS` → TikTok Shop (Promo R179, Advertising R185, Other Mktg R191) — **includes TikTok Shop Mexico** (folded in by the query)
  - Excluded by the query: `Amazon 3P`, `DTC` (Wyze.com US/CA), `International`, DTC paid-media platforms (Apple / Google / Roku — already in CAC via Northbeam), and all `*Service` buckets.
- **Caveat:** HD ad spend is mostly booked under Other Mktg (Channel Marketing 6102), not Advertising 6601 — so HD's finance **Advertising is tiny** (Apr $1.1K) and won't match the weekly `HD_ADS_DAILY` actual (different row). Expected.

---

## 4. Excel + HTML output

### 4.1 Excel

Loaded from `inputs/Dashboard_Template.xlsx` template (**208 rows** × 87 cols including 53 weeks of date headers — user added a `% of Target - YTD` row to every sell-in / sell-through / new-customers block 2026-06-18). All row-number constants live in `build_outputs.py` (`CH_ROWS`, `TR_ROWS`, `YTD_PCT_ROWS`, `TARGET_YTD_PCT_ROWS`, `METRIC_BLOCKS`, `PLAN_TARGETS`, `FINANCE_ROW_MAP`); if the template layout changes, update those. Per regeneration:

- **E1** is hardcoded `2026-01-04` (Sunday ending Mon-Sun week 2025-12-29); F1.. are `=prev+7`.
- **Numbers are integer everywhere** (2026-06-18): `#,##0` for $k / # / # ppl, `0%` for percent. No decimals (`apply_number_formats`).
- **Per-channel actuals + targets**: hardcoded values from CSV (`/1000` for $ columns).
- **`% of Target` rows** (single-week): Excel formulas `=IFERROR(actual/target,"")` — auto-update on edit.
- **`% of Target - YTD` rows** (NEW): Excel formulas `=IFERROR(SUM($E$actual:actual)/SUM($E$target:target),"")` — cumulative actual ÷ cumulative target, week-by-week (`TARGET_YTD_PCT_ROWS`). NC rows compute blank until a New-Customers target source exists.
- **Amazon Ads (On-site + DSP)** row: Excel formula `=E{ads_on}+E{ads_dsp}` per column.
- **Total Retail SUM rows**: Excel formulas summing the 5 retail channels (Amazon 1P + Walmart + HD + Other Omni + TikTok Shop) — `sum_formula_5ch`.
- **Total Retail New Customer rows**: hardcoded (separate source, not summable from per-channel).
- **All Channels rows (R3-R5)**: hardcoded (includes DTC / International which have no section rows).
- **`% of Total - YTD` rows** (`YTD_PCT_ROWS`): cumulative channel ÷ All Channels (denominator R3 sell-in / R4 sell-through).

### 4.2 HTML

Wide-format dashboard at `outputs/YYYY-MM-DD/Retail Business Dashboard_YYYY-MM-DD.html`:
- Sticky left columns (Metric label + Unit), sticky top header (week columns)
- Most recent column highlighted yellow
- `% of Target` cells color-coded: green ≥100%, yellow 95-100%, red <95%
- Target rows shown in light purple background
- Section dividers with bold header
- Top: data anomaly callout (red box) + general notes (yellow box)

### 4.3 Generator scripts

Excel + HTML are generated by ad-hoc Python in the conversation history (uses `openpyxl` for Excel). Not currently in repo as standalone script — should be extracted if regenerating becomes routine (see §6 checklist).

---

## 5. Hard limitations (cannot fix without major redesign)

| # | Limitation | Why hard |
|---|---|---|
| H1 | **Amazon NTB is orders, not unique customers** | Amazon's NTB API doesn't dedupe at user level. One NTB customer who places 2 orders within the 14-day window counts as 2. No workaround. |
| H2 | **NTB v3 API 60-day rolling backfill window** | Amazon API restriction. Historical NTB before 2026-03-27 cannot be retrieved. Forward-only from now on. |
| H3 | **Total Retail New Customer count contaminated by TTS + International** | `acquisition_channel = 'Other'` is "non-Shopify D2C" — there's no Wyze-side field that distinguishes retail-channel-purchased devices from TTS/Intl-purchased devices. Would require a device_serial → retailer mapping table that doesn't exist today. User accepted this as dashboard treats TTS as retail-like. |
| H4 | **Binding lags purchase + 2nd-hand devices counted** | Customer activates Wyze device whenever they unbox; could be days or weeks after purchase. Also any new user whose first device happens to be 2nd-hand (acquired/gifted) is counted as a new customer. No way to filter these out from `USER_ATTRIBUTES`. |
| H5 | **Cam/NonCam "Unknown" lumped into NonCam** | Users whose `first_binding_ts` doesn't match anything in `GE_RELATION_USER_DEVICE` within ±60min get Unknown product_type. Lumping into NonCam keeps `Cam + NonCam = Total Retail New Customers`. The alternative (drop Unknown) would make per-row Cam+NonCam < Total. Either way, ~real classification info is lost for these users. |
| ~~H6~~ | **RESOLVED 2026-06-22 — Walmart 3P promo now in.** | Sell-through already included Walmart 3P; promo now does too (`wmt_3p_weekly` + `wmt_3p_promo_26`, riding Ecomm campaigns). 3P coverage = whatever the `WALMART_3P` order-line feed holds; the 1P manual backfill does not include 3P. |
| H7 | **Other Retail sell-through not available** | `BUDGET_HW_WITH_ACTUALS_WEEKLY` has a `Other Retail` SUB_CHANNEL with sell-in $; but `WYZE_SELL_THROUGH_DATA` has no corresponding channel — Other Retail sell-through is invisible. Other Omni sell-through under-reports vs. sell-in. (No source available — accepted.) |
| H8 | **Amazon promo on 1P-volume only — BY DESIGN** | Amazon promo = 1P sell-through × discount; 3P volumes intentionally excluded (per user 2026-06-22: not counting Amazon 3P promo is correct). Not a gap. |
| ~~H9~~ | **RESOLVED 2026-06-22 — Walmart 1P pre-cutover via manual backfill.** | Scintilla unreliable before WM Sat 2026-04-25; weeks `[2025-12-29, cutover)` now come from `WALMART1P_UNITS_20250615_20260613`. Only remaining gap: pre-`2025-12-29` weeks (floor; extend if 2025-H2 wanted). |
| ~~H10~~ | **RESOLVED 2026-06-30 — future-week sell-in budget recovered.** | `BUDGET_HW_WITH_ACTUALS_WEEKLY` future weeks have YEAR + WEEK_NUMBER + BUDGET_GROSS_REV but NULL START/END_DATE (the view borrows dates from actuals). `sell_in_raw` now derives week_start from WEEK_NUMBER for those rows (`DATEADD((WEEK_NUMBER-1)*7,'2025-12-29')`), so the full-year sell-in TARGET shows. Validated the END_DATE→WEEK_NUMBER mapping is a clean +7/week series. (Cross-year first 3 days 2025-12-29..31 = 2025-W53 still excluded: 2025 has no BUDGET_GROSS_REV.) |
| H11 | **SP / DSP NTB not available** | Amazon NTB now covers **SB + SD** (1P). **SP** (Sponsored Products) NTB still has no correct data; **DSP** doesn't provide NTB metrics for Wyze. So Amazon NTB undercounts by the SP (+DSP) portion until those land. |
| H12 | **Week boundary mismatches across sources (accepted)** | Walmart Scintilla WM-Sat-Fri week mapped to Mon-Sun via `DATE_TRUNC('week', WM_Sat + 2)` — 5/7 days overlap. Amazon DSP TEMP Sun-Sat shifted +1 day → 6/7 days overlap with Mon-Sun. These bleed boundaries between adjacent weeks slightly. Same accepted treatment as CAC tracker HW loss. |

---

## 6. Fixable items — checklist

Things that can be improved later when source data / time available. Loose grouping by what's needed.

### 6.1 Needs new data feed (blocked on engineering / finance)

- [x] **Finance monthly marketing actuals** — DONE 2026-06-09; source migrated to NetSuite CSV 2026-06-18 (see §3.9). Now `inputs/finance/Marketing_Spend_2026YTD.csv` (long `PERIOD, CATEGORY, CHANNEL, SPEND` from `DATA_MART.RETAIL_DASHBOARD.MARKETING_SPEND_YTD`), wired via `load_finance_actuals()` / `apply_finance_actuals()`. **5 sections populate** (Amazon, Walmart, HD, Other Omni, TikTok Shop).
- [x] **Walmart Ads actuals** — DONE 2026-06-22. `walmart_ads_weekly` ← `WALMART_ADS_DAILY` (Walmart Connect, 1P+3P). Wired into dashboard `ads_metrics` (`'Walmart'`) + CAC `ads_cost.sql` / `v_ads_cost_weekly.sql`.
- [x] **Home Depot Ads actuals** — DONE 2026-06-03. `hd_ads_weekly` ← `HD_ADS_DAILY` (Orange Access manual upload). Feeds the Home Depot ads row.
- [ ] **Other Omni Ads actuals** — replace plan estimate (BB / Costco / Other Retail ad spend).
- [ ] **Other Marketing actuals (all channels)** — Amazon MDF, retail in-store displays, retail marketing packages, TTS commissions & samples, affiliate, content/video. Multiple sources, possibly need finance team to provide.
- [ ] **Amazon DSP daily pipeline** — replace `AMAZON_DSP_ADS_TEMP` (manual upload, currently ~13 weeks LY + recent CY only). Once daily feed lands, drop the `+1 day` shift and use `DATE_TRUNC('week', date)` like Amazon SP.
- [x] **Amazon SD NTB** — DONE 2026-06-22. `SD_NTB_1P_US` unioned with `SB_NTB_1P_US` in `amazon_ntb_weekly` (same `ntb_purchases` column). SD earliest 2026-04-17, small magnitude.
- [ ] **Amazon SP NTB** — no correct data yet; add a third UNION branch in `amazon_ntb_weekly` when it lands.
- [ ] **Amazon DSP NTB** — if Amazon ever provides DSP NTB metrics for Wyze, add to `amazon_ntb_weekly`.
- [x] **Walmart 3P sell-through** — DONE 2026-06-09. `WYZE_SELL_THROUGH_DATA` gained a `'Walmart 3P'` channel; `TEST_FORECAST` reads it; dashboard `sell_through_raw` folds `'Walmart 3P'` → `'Walmart'`. Sell-through now 1P+3P.
- [x] **Walmart 3P promo** — DONE 2026-06-22. `wmt_3p_weekly` (order-line units) + `wmt_3p_promo_26` (Ecomm campaigns re-labelled 'Walmart 3P') added to `wmt_promo_spend`. Dashboard Walmart promo now 1P+3P (matches CAC).
- [ ] **Other Omni Ads actuals** / **HD Canada sell-through+promo** / **Amazon DSP daily pipeline** — no source available; accepted for now (Other Omni & HD-CA per user 2026-06-22 "no way / not important"; DSP stays on manual `AMAZON_DSP_ADS_TEMP`).

### 6.2 Computable in SQL today (just not done yet)

- [ ] **Forecast/Target for New Customers** (R21 Total Retail target, R42 Amazon NTB target) — if marketing team produces a target, add as a target table and join. Today: blank.
- [x] **Finance Monthly Actuals data integration** — DONE 2026-06-09; source migrated to NetSuite CSV 2026-06-18 (see §3.9). The `% of budget` rows (4th/5th in each block) are pre-filled as `=IFERROR(finance/budget,"")` and auto-compute for all 5 populated sections.
- [ ] **Negative sell-in handling** — HD 2026-03-30 = -$1.13M flagged as anomaly. Decision: leave as-is (correct finance representation), or clip at zero (cleaner visualization), or split returns into a separate "Returns / Credits" row.

### 6.3 Dashboard polish / process

- [ ] **Promo per channel sister view** — `V_PROMO_SPEND_BY_CHANNEL_WEEKLY` to avoid the 600-line inline duplicated from `promo_spend_all_channels.sql`. Today inline is fine (only one consumer); extract to a view once a 2nd consumer needs per-channel promo.
- [ ] **Auto-trigger / scheduled regeneration** — schedule weekly run of `build_outputs.py` via cron / GitHub Actions / Snowflake task so user doesn't have to manually invoke. Deferred until data sources stabilize and validation work completes.
- [ ] **Devise consistent file naming** — `Retail Business Dashboard_YYYY-MM-DD-HHMM.csv` (SQL output) vs `Retail Business Dashboard_YYYY-MM-DD.xlsx/.html` (dashboard outputs). OK for now but worth aligning.
- [ ] **Add UI for stakeholders** — currently HTML opens in browser. Could be hosted (S3 / Streamlit / etc.) for shareable URL instead of file attachment.
- [ ] **Add filter / drilldown** — interactivity to filter weeks / show per-product-line breakdown. Currently the HTML is a static snapshot.

### 6.4 Verification / data quality TODOs

- [ ] **Investigate HD 2026-03-30 = -$1,129,900** — confirm with finance whether this is a real return/credit or a data quality issue in `BUDGET_HW_WITH_ACTUALS_WEEKLY`.
- [ ] **Investigate HD 2026-01-26 = $1.53M spike** — confirm post-holiday restock vs. anomaly.
- [ ] **Investigate Walmart promo near-zero for 2026-04-20 onwards** — Scintilla data coverage starts 4/20 but promo is still ~$0 throughout April. Check if promo tracker has Walmart 2026 campaigns matching to SKU-channel-week joining correctly.
- [ ] **Verify Scintilla `OMNI_SALES` data quality post-4/18** — spot-check actual sell-through against Walmart Vendor Central directly for a few weeks to confirm Scintilla numbers are reliable.

---

## 7. Operating the dashboard (refresh checklist)

### TL;DR (weekly refresh)

```bash
cd "Retail Business Dashboard"
.venv/bin/python run_retail_dashboard.py
```

The root launcher checks monthly Finance data, exports the weekly query, and
builds both outputs. It uses a temporary weekly CSV that is deleted after the
build and stores finished files under `outputs/YYYY-MM-DD/`.

### Detailed steps

To regenerate the dashboard for a new snapshot:

1. **(Pre-run)** Complete the upstream Amazon DSP, Walmart Ads, and Home Depot
   Ads updates. These sources feed the weekly query.
2. **(Local)** Run the root launcher:
   ```bash
   cd "Retail Business Dashboard"
   .venv/bin/python run_retail_dashboard.py
   ```
   Use `--snapshot-date YYYY-MM-DD` only when the output date should differ
   from today. Use `--finance-check-only` to skip the weekly dashboard.
3. **(Sanity check)** Open the HTML and verify:
   - Most recent week column highlighted is the correct Mon-Sun week
   - Each section has the expected rows (no missing channels)
   - Total Retail values ≈ sum of (Amazon 1P + Walmart + HD + Other Omni + TikTok Shop) per row; All Channels also includes Amazon 3P + DTC US + DTC CA + International
   - All Channels values ≈ Total Retail + DTC US + DTC CA + International
   - No NULL `week_start` rows
   - Any obvious anomalies (negative numbers, 10x spikes) flagged in the anomaly callout
4. **(Share)** Upload the dated Excel from `outputs/YYYY-MM-DD/`.

For a manual fallback, save the weekly Snowflake result anywhere and run:

```bash
.venv/bin/python scripts/build_outputs.py \
  "/path/to/Retail Business Dashboard_YYYY-MM-DD-HHMM.csv"
```

---

## 8. File locations

| Path | Purpose |
|---|---|
| `sql/weekly_retail_actuals.sql` | Read-only wrapper for `DATA_MART.RETAIL_DASHBOARD.WEEKLY_RETAIL_ACTUALS` |
| `sql/finance_actuals.sql` | Read-only wrapper for `DATA_MART.RETAIL_DASHBOARD.MARKETING_SPEND_YTD` → Finance CSV |
| `sql/grant_retail_dashboard_read_role.sql` | Minimal role grants for the two views |
| `scripts/build_outputs.py` | Build script: CSV → Excel + HTML in one command |
| `docs/TECHNICAL_REFERENCE.md` | This doc |
| `inputs/finance/Marketing_Spend_2026YTD.csv` | Finance Actuals (Monthly) source — replace in place ~monthly |
| `inputs/Dashboard_Template.xlsx` | Fixed Excel output template |
| `inputs/FY26_Marketing_Plan.xlsx` | Monthly marketing-plan input |
| `inputs/FY26_Promo_Plan.xlsx` | Monthly promotion-plan input |
| `outputs/YYYY-MM-DD/Retail Business Dashboard_YYYY-MM-DD.xlsx` | Snapshot Excel (per-channel hardcoded values + Total Retail formulas) |
| `outputs/YYYY-MM-DD/Retail Business Dashboard_YYYY-MM-DD.html` | Snapshot HTML dashboard |

---

## 9. Change log

| Date | Change |
|---|---|
| 2026-05-29 | Initial build: sell-in + sell-through + targets, all channels |
| 2026-05-29 | Added promo per channel (Amazon, Walmart, HD, Other Omni-BB, TTS) inline from `promo_spend_all_channels.sql` |
| 2026-05-29 | Added Amazon Ads (SP on-site + DSP TEMP) and TTS Ads (GMV Max) |
| 2026-05-29 | Added new_customers (3 sources for 3 rows: All Channels / Total Retail / Amazon SB NTB) |
| 2026-05-29 | Walmart promo migrated from SPS to Scintilla with subchannel split (Store vs Ecomm) |
| 2026-05-31 | Year-boundary fix: `DATE_TRUNC('week', END_DATE)` to merge 2025-W53 + 2026-W01 sell-in partial rows |
| 2026-05-31 | Removed LY columns from output (dashboard doesn't show LY rows) |
| 2026-05-31 | Added future-week filter: `DATEADD('day', 6, week_start) < CURRENT_DATE` |
| 2026-05-31 | Total Retail now includes TTS US (5 channels); Excel rebuilt with SUM formulas for Total Retail rows |
| 2026-05-31 | Amazon R59 (Ads On-site + DSP) changed from hardcoded to Excel formula `=E49+E54` |
| 2026-05-31 | DSP TEMP data refreshed (35 weeks coverage); Excel + HTML rebuilt |
| 2026-05-31 | Added `% of Total Retail — YTD` rows for sell-in + sell-through under each retail channel section (Amazon/Walmart/HD/Other Omni/TTS). 210 Excel formulas + HTML rows. |
| 2026-05-31 | Pre-filled `% of budget — this week` and `% of budget — YTD` rows across all 16 (channel × marketing-bucket) groups (Promo / Ads on-site / Ads DSP / Ads total / Other Marketing per channel). 672 Excel formulas — `IFERROR(actual/budget,"")` for this-week, `IFERROR(SUM(...)/SUM(...),"")` for YTD. They evaluate to blank now (no Budgeted source) and auto-populate when budget data lands. |
| 2026-05-31 | Added `Finance Monthly Actuals` placeholder sheet (second tab in the workbook) — Metric × 12 months + YTD layout, ~14 placeholder rows for total Promo / Ads sub-buckets / Other Marketing / Total Marketing Spend, plus a Variance section. Awaits finance monthly data source. |
| 2026-05-31 | Extracted ad-hoc Excel+HTML generation into `scripts/build_outputs.py`. Single command: `python3 build_outputs.py <csv>`. Snapshot date auto-parsed from CSV filename. All formula passes (per-channel %, Total Retail SUMs, % of Total Retail YTD, % of budget, Amazon Ads on+DSP) + Finance Monthly Actuals sheet all baked in. Removed Amazon promo 3P TODO (user decision to keep status quo — uncertain whether campaigns target 3P). |
| 2026-06-01 | Walmart promo CTE refactor: read `SRC_RAW.WMT_SCINTILLA.OMNI_SALES` raw (same logic as `wyze_sell_through.sql` `walmart_omni_base`) instead of `WYZE_SELL_THROUGH_DATA` view. Consistent with HD/BB/Amazon promo (raw sources) and avoids scanning the heavy unified view. Applied to all 3 files: `weekly_retail_actuals.sql`, `promo/by_channel/Walmart_promo.sql`, `promo/promo_spend_all_channels.sql`. |
| 2026-06-01 | Excel format fixes ported to `build_outputs.py`: (a) `D4='#'` override (template's wrong `$k` for sell-through units); (b) `% of Total - YTD` denominator switched from Total Retail (R8/R12) to All Channels (R3/R4); (c) Amazon R40 (`As % of sales units (ads attributed)`) removed — formula mixed NTB orders vs sell-through units denominators; (d) new `apply_number_formats()` pass: `0.0%` for any `%` row, `#,##0.0` for Ads/Promo/Other Marketing $k rows (1-decimal granularity), `#,##0` for Sell-in $k + Sell-through # + New Customer # ppl (integer). |
| 2026-06-01 | Section redefinition per user: (a) Amazon = 1P only (sell-in / sell-through / promo / ads on-site / NTB all drop 3P); Amazon 3P emitted as its own dashboard_channel row in CSV, contributes to All Channels but not Total Retail. (b) Home Depot section now includes HD Canada in sell-in (sell-through / promo for HD CA unavailable — accepted gap). (c) TikTok Shop section now includes TikTok Shop Mexico in sell-in / sell-through / promo (no MX GMV Max source for ads). (d) `'TTS US'` channel renamed to `'TikTok Shop'`; Excel template label `A152` overridden from "TikTok Shop US" → "TikTok Shop". (e) `Total Retail` is exactly the 5 retail sections summed; `All Channels` is literally everything (Total Retail + Amazon 3P + DTC US + DTC CA + International). |
| 2026-06-01 | Marketing/Promo plan targets integrated. `inputs/FY26_Promo_Plan.xlsx` (`Summary`) and `inputs/FY26_Marketing_Plan.xlsx` (`Marketing_Tracker`) loaded at runtime; mapped to 16 dashboard rows (5 promo + 11 marketing). Monthly values placed at LAST Mon-Sun week of each month (Mon 1/19→Jan, Mon 2/16→Feb, Mon 3/23→Mar, Mon 4/20→Apr, Mon 5/25→May, ...). Other Marketing rows un-grayed. |
| 2026-06-01 | Metric block layout repurposed: each 4-row block was `actual / Budgeted / % of budget this week / % of budget YTD`, now `actual / Finance Actuals (Monthly) / Budgeted (Monthly) / blank`. % of budget formulas removed (granularity mismatch — weekly actual vs monthly budget is meaningless; will reintroduce as monthly-vs-monthly once finance monthly actuals data arrives, populated in the Finance Actuals row). `Finance Monthly Actuals` second sheet deleted — finance data lives inline now. |
| 2026-06-01 | User manually restructured template (178 → 194 rows): each metric block is now 5 rows (weekly actual / Finance Actuals (Monthly) / Budgeted (Monthly) / `As % of budget - this week` / `As % of budget - YTD`). Script's row constants (`CH_ROWS`, `TR_ROWS`, `YTD_PCT_ROWS`, `METRIC_BLOCKS`, `PLAN_TARGETS`, `OTHER_MARKETING_GRAY_ROWS`) all updated to new positions. In-script `repurpose_metric_blocks()` removed (no longer needed). TikTok Shop section header override moved from `A152` → `A165`. |
| 2026-06-01 | `% of budget - this week` rows relabeled to `% of budget - this month` (col B override) since both Finance Actuals and Budgeted are monthly. Formulas pre-populated as `=IFERROR(finance/budget, "")` for "this month" and `=IFERROR(SUM(...)/SUM(...),"")` for YTD — both evaluate to blank until Finance Actuals row is populated, then auto-compute. |
| 2026-06-01 | Plan loader now keeps `$0` values (was filtering out as "missing"). FY26 plan has intentional Q1 zeros for retail promo (Amazon 1P / HD / Walmart / BB are $0 for Jan/Feb, real $ starts Mar). Cells now show `$0.0` explicitly instead of being blank. |
| 2026-06-01 | HTML output now includes `Budgeted (Monthly)` rows per channel section (light-purple-tinted, placed at last Mon-Sun week of each month — Excel and HTML kept in sync). Plan loaders moved to `main()`, passed to both `build_excel` and `build_html` to avoid duplicate file reads. |
| 2026-06-02 | **CAC promo Walmart week-convention bug fix.** When migrating WMT from SPS to Scintilla, I reused the Mon-Sun mapping (`DATE_TRUNC('week', WM_Sat + 2)` → Monday) from the retail dashboard. But CAC promo pipeline is Sun-Sat (week_start = Sunday), so Walmart was 1 day off all other channels. Changed CAC-side Walmart week_start to `WM_Sat + 1 day` (lands on Sunday directly, 6/7 days overlap with WM week). Affected: `sql/promo/by_channel/Walmart_promo.sql`, `sql/promo/promo_spend_all_channels.sql`. Retail dashboard SQL (`sql/weekly_retail_actuals.sql`) unchanged — that one is correctly Mon-Sun. `V_PROMO_SPEND_WEEKLY` view DDL needs to be re-deployed. |
| 2026-06-03 | **HD ads wired in.** User loaded HD Orange Access ads (PLA + Banner onsite + Google PMAX offsite) into `DATA_MART.FINANCE.HD_ADS_DAILY` via local Python MERGE script (no HD API). Manual download from HD Orange Access UI weekly. Google PMAX HD-attributed, not Wyze DTC (no Northbeam overlap). Wired into: `ads_cost.sql` + `v_ads_cost_weekly.sql` (CAC tracker — both now have 6 sources: northbeam, tiktok, SP 1P/3P, DSP, HD). `weekly_retail_actuals.sql` adds `hd_ads_weekly` CTE → unions into `ads_metrics` as `'Home Depot'` channel. `V_ADS_COST_WEEKLY` view DDL needs re-deployment. |
| 2026-06-03 | HD plan mapping refined: R128 HD Ads Budgeted (Monthly) now pulls from plan "Marketing package" bucket (description includes "+ media marketing"; $64.6K/mo May). R134 HD Other Marketing Budgeted drops Marketing package, keeps only "In-Store Display + Events + Brand Advocate". Caveat: Marketing package also contains a ~$11K/mo non-ads reporting/measurement fee (HD data portal etc.) — actual HD ads ≈ 83% of plan in May 2026. HTML adds an HD "Ads" weekly row + budget row to the Home Depot section (via new `has_hd_ads` flag in build_channel_section). Excel R126 HD Ads populated from CSV; R128 from plan. |
| 2026-06-09 | **Walmart 3P sell-through wired.** `sell_through_raw` CASE folds `'Walmart 3P'` → `'Walmart'` (source: `TEST_FORECAST` ← `WYZE_SELL_THROUGH_DATA`'s new `'Walmart 3P'` channel). Walmart sell-through now 1P+3P, consistent with 1P+3P sell-in. Walmart **promo** still 1P only (3P promo deferred). |
| 2026-06-09 | **Finance Actuals (Monthly) wired** (see §3.9). New `load_finance_actuals()` + `apply_finance_actuals()` write monthly GL spend (Promos/Advertising/Other Mktg; **Discounts ignored**) into the inline Finance Actuals (Monthly) rows at the LAST Mon-Sun week of each month ($→$k); pre-filled % of budget formulas now auto-compute. HTML gets matching `finance`-class rows. Initial source had an aggregate `Retail` channel → only TikTok Shop mappable. |
| 2026-06-10 | **Number-format bug fix.** `apply_number_formats` looped `range(3, 179)`, stopping mid-TikTok-Shop-section — rows 179-194 (TTS Finance Actuals / Budgeted / % rows) stayed unformatted `General`. Changed to `range(3, ws.max_row + 1)`. Only TTS was affected (other sections end before row 179). |
| 2026-06-10 | **Finance source swapped to `Wyze_Marketing_Spend_2026YTD.xlsx`** (wide `Channel × Category` tab; `load_finance_actuals()` rewritten for wide format). Retail now split into `Amazon 1P` / `Home Depot` / `Walmart` / `Other Retail`, so `FINANCE_ROW_MAP` + `FINANCE_HTML_MAP` expanded — **5 sections populate** (Amazon, Walmart, HD, Other Omni, TikTok Shop). `Other Retail` → Other Omni; HD `Advertising` tiny (ads booked under Other Mktg in GL); HD Canada → International (dashboard HD finance US-only). DTC/International/Service/Shared/Unallocated not mapped. |
| 2026-06-30 | **Full-year sell-in & sell-through TARGET line.** `weekly_retail_actuals.sql`: `sell_in_raw.week_start = COALESCE(DATE_TRUNC('week',END_DATE), DATEADD((WEEK_NUMBER-1)*7,'2025-12-29'))` so future weeks (END_DATE NULL — the view borrows dates from actuals; H10) still get a week_start; WHERE relaxed to `END_DATE>='2025-12-29' OR (YEAR=2026 AND END_DATE IS NULL)`; the final `per_channel` filter changed from `< CURRENT_DATE` to `BETWEEN '2025-12-29' AND '2026-12-28'` (keep full year). Validated: END_DATE is a clean +7/week series (wk1 = 2026-01-04), and the derived anchor (wk1 → Mon 2025-12-29) matches. Sell-through already had full-year forecast `TARGET_UNITS`. `build_outputs.py`: targets (sell-in/sell-through) now written for the full year; **actuals / single-week % / cumulative % (Target-YTD, Total-YTD) / NC / % of budget only through the last complete week** (`snapshot_date` passed in → `last_actual_col`; week complete iff its Sunday < snapshot). Future weeks show ONLY target — actuals are blank, never 0 (the view's COALESCE-to-0 future actual is simply not written). HTML still shows complete weeks only (unchanged). Cross-year first 3 days (2025-12-29..31) still ignored per user (2025 has no BUDGET_GROSS_REV; they'll annotate manually). |
| 2026-06-30 | **Misc fixes:** (a) `FINANCE_ROW_MAP` had been missed in the 212-remap (it sat outside the spliced block) — Walmart/HD/Other-Omni/TTS finance were writing to 252-row positions (TTS landed on phantom rows 227/234); fixed to 93/100, 129/136, 165/172, 195/202. (b) Sell-through first week (2025-12-29) had no target → now backfilled with that week's actual (per channel; Total Retail follows via the sum formula), so `% of Target - YTD` reverted to inclusive `SUM/SUM` (no SUMIF exclusion). (c) `apply_row_label_styles` now also copies the label cell's FILL onto data cells, so a greyed-out row (Amazon NC block) greys across. (d) Cell comment added on each sell-through "Target" row label explaining the first-week backfill. |
| 2026-06-29 | **Template → 212 rows; `% of budget` back to finance numerator; NC → flat 52% / pp.** (a) User trimmed Ads / Other-Marketing **detail blocks to 2 rows** (actual + Budgeted) — Finance + the three `% of budget` rows now live **only on Promo + Total Marketing**. (b) **`As % of budget` (month / QTD / YTD) now uses the block's monthly Finance Actuals row as numerator** for all three (not the bottom-up weekly actual): month = finance ÷ month budget, QTD = ΣQTD finance ÷ full-quarter budget, YTD = ΣYTD finance ÷ full-year budget. Supersedes the 2026-06-18/06-22 weekly-numerator approach. (c) **New Customers target = flat 52% constant** (no longer 52% × units). `% of Target` rows are **percentage-points** = actual share − 52%, `0%`-formatted (shows e.g. `5%`). Total Retail gets pp rows (R22 = `R21−R24`, R23 = `ΣNC/Σunits − 52%`); Amazon NC fills actual (SB+SD NTB) + 52% target only, ratio rows R45/46/47 blank. (d) Remapped ALL row constants to 212 layout; `MARKETING_BLOCKS` is now Promo + Total Marketing only (6-tuple, finance numerator); removed stale `A211` / `D4` overrides (template self-corrects, and `A211` would now corrupt the Other-Marketing label). `pp` unit → `0%` format. HTML synced (NC pp; Amazon actual+target). Verified on 6/23 retail. |
| 2026-06-22 | **Amazon NTB += SD.** `amazon_ntb_weekly` now `SB_NTB_1P_US UNION ALL SD_NTB_1P_US` (same `ntb_purchases` column/algo, 1P). SB earliest 2026-03-27, SD earliest 2026-04-17 (small); nothing before 3/27. SP NTB still unavailable → add a third branch when it lands. Only `weekly_retail_actuals.sql` changed (per user). |
| 2026-06-22 | **Walmart 3P promo added to dashboard.** `wmt_3p_weekly` (3P units from `SRC_CLEAN.WALMART_3P.*` order-lines, Mon-Sun) + `wmt_3p_promo_26` (Ecomm `recb89...` campaigns re-labelled `'Walmart 3P'`) unioned into `wmt_promo_spend`. Dashboard Walmart promo now 1P+3P, matching CAC `Walmart_promo.sql` and the 1P+3P sell-in/sell-through. 3P rides the .com/Ecomm campaigns (no 3P-specific Airtable rec ID). Toggle off by removing the two CTEs from the unions. (Amazon promo stays 1P-only — by design, 3P intentionally excluded.) |
| 2026-06-22 | **Walmart ads wired.** `DATA_MART.FINANCE.WALMART_ADS_DAILY` (Walmart Connect, `SUM(ad_spend)` by `report_date`, 1P+3P combined, channel column summed away). CAC: added `walmart_ads` daily source → `daily_ads` in `ads_cost.sql` + `v_ads_cost_weekly.sql` (now **7 sources**; Sun-Sat rollup). Retail: `walmart_ads_weekly` (Mon-Sun) → `ads_metrics` as `'Walmart'`. `build_outputs.py`: added `CH_ROWS['Walmart']['ads_tot']=122`, a Walmart write branch, and flipped the Walmart Ads `MARKETING_BLOCKS` row from `'none'`→`'weekly'` so its `% of budget` computes. Re-deploy `V_ADS_COST_WEEKLY`. |
| 2026-06-22 | **Walmart 1P pre-cutover units → manual backfill.** Scintilla `OMNI_SALES` reliable only from WM Sat 2026-04-25 (dashboard `week_start >= 2026-04-27`); earlier weeks `[2025-12-29, 2026-04-27)` now read `SRC_RAW.WMT_SCINTILLA.WALMART1P_UNITS_20250615_20260613` (daily manual pull). Daily → WM Sat-Fri → +2 = Mon; SKU via `walmart_item_number → ITEM_DIM.GTIN → UPC_SKU_MAPPING` (UPC can't map directly). No Store/Ecomm split → bucketed `'Walmart 1P Ecomm'` (synced pre-cutover campaigns; existing Ecomm promo join matches as-is). Changed `wyze_sell_through.sql` (sell-through, the view feeding `TEST_FORECAST` → dashboard Walmart sell-through) + `weekly_retail_actuals.sql` `wmt_weekly` (promo units). CAC side (`Walmart_promo.sql`, `promo_spend_all_channels.sql`) changed too, with Sun-Sat cutover `week_start >= 2026-04-26`. ⚠ Requires re-deploying `WYZE_SELL_THROUGH_DATA` + refreshing `TEST_FORECAST_SELL_THROUGH_PRODUCT_LINE` before the dashboard sell-through reflects it. |
| 2026-06-22 | **Template → 252 rows: Total Marketing section, QTD %, annual-budget YTD %, 52% NC target, integer + unified row colors (Excel).** (a) **New Customers target = 52% × sell-through actual** (`NC_TARGET_PCT`/`NC_TARGET_ROWS`) — only Total Retail (base R13) and Amazon NTB (base R36); drives their `% of Target` + `% of Target - YTD`. (b) Marketing blocks now 6 rows: added **`As % of quarterly budget (QTD)`** (QTD actual ÷ **full-quarter** budget) and changed **`As % of annual budget (YTD)`** to ÷ **full-year** budget (was cumulative-to-date). Both still one value per complete month at month-end col. (c) **New `Total Marketing` section** after Promo with Ads/Other as indented detail; **finance lands only in Promo + Total Marketing** (`FINANCE_ROW_MAP`: Advertising + Other Mktg both → Total Marketing row, `apply_finance_actuals` accumulates) — resolves the Amazon-ads-in-6102 classification issue; detail Ads/Other finance rows left blank. (d) `% of budget` numerator by row (`MARKETING_BLOCKS` num_src): Promo/Ads = weekly actual, Total Marketing = finance, Other Mktg + Walmart Ads = blank (no source). (e) **All numbers integer** + **each data row's cells now inherit its row-name color/bold/italic** (`apply_row_label_styles`). Remapped all row constants for the 252-row layout. **HTML not yet updated to this layout** (per user: Excel-first). |
| 2026-06-18 | **Amazon Advertising finance wired + Other-Marketing % of budget uses finance.** (a) Added `('Amazon 1P','Advertising') → R69` (Ads On-site+DSP finance row) to `FINANCE_ROW_MAP` / `FINANCE_HTML_MAP`. ⚠ GL 6601 Advertising is tiny / sometimes negative (Mar −$46K) — most Amazon ad spend is booked under 6102 Channel Marketing → Other Mktg row. **Known classification issue; using as-is for now** per user. (b) `apply_pct_of_budget_formulas`: Other / Total Marketing blocks (`OTHER_MARKETING_ACTUAL_ROWS`) have no weekly-actual feed, so their `As % of budget` numerator now uses the monthly **Finance Actuals** row instead of the (empty) weekly actual; Promo / Ads still use weekly actual. |
| 2026-06-18 | **"As % of budget" reworked → monthly, weekly-actual based, complete-months-only.** `apply_pct_of_budget_formulas` rewritten: numerator is now the user's **weekly actual** (block row 1) summed over the month — NOT Finance Actuals (finance is amortized GL, sometimes negative, not comparable to budget; finance row stays for display only). One value per month at the month-end column (same col as Budgeted): `this month` = SUM(month's weeks actual) ÷ month budget; `YTD` = SUM(Jan..month-end actual) ÷ SUM(Jan..month-end budget). **Incomplete (in-progress) months left blank** (gate: `last_mon_of_month_col > max_col`). Old formula (finance ÷ budget, every column) removed. Label stays "...- this month". |
| 2026-06-18 | **Template grew to 208 rows + integer formats.** User added a `% of Target - YTD` row (cumulative actual ÷ cumulative target) to every sell-in / sell-through / new-customers block (between `% of Target` and `Target`). Remapped ALL row constants (`CH_ROWS`, `TR_ROWS`, `YTD_PCT_ROWS`, `METRIC_BLOCKS`, `PLAN_TARGETS`, `FINANCE_ROW_MAP`) from the old 194-row layout; added `TARGET_YTD_PCT_ROWS` (14 rows) + a formula pass in `build_excel` and a `cum_pct_target` helper + rows in `build_channel_section` (HTML). **All numbers now integer** — `apply_number_formats` drops `ONE_DECIMAL_ROWS`; $k/#/# ppl → `#,##0`, % → `0%`; HTML `fmt_pct` default 0 decimals. TTS header override moved A165→A177. Verified: 2676 formulas, 60 finance cells, all 14 YTD formulas reference correct actual/target rows. |
| 2026-06-18 | **Finance source migrated from Excel → NetSuite CSV.** Now `inputs/finance/Marketing_Spend_2026YTD.csv` (long `PERIOD, CATEGORY, CHANNEL, SPEND` from `DATA_MART.FINANCE.NETSUITE_CPAM_DETAILS`, pre-aggregated by new `sql/finance_actuals.sql`). `load_finance_actuals()` rewritten to read the long CSV via `csv.DictReader` (replaces the openpyxl wide-tab reader). The query owns category/channel/sign logic: Discounts (4092/4093) excluded, Promo (4094/4095/4098) sign-flipped positive, and — corrected from the 2026-06-10 Excel — **HD Canada folds into Home Depot** + **TikTok Shop Mexico folds into TTS** (own sections, not International). `FINANCE_ROW_MAP` / `FINANCE_HTML_MAP` unchanged (query emits the same 5 channel keys). Verified: 60 finance cells written, values match ($k-scaled). Old `Wyze_Marketing_Spend_2026YTD.xlsx` now unused. |
| 2026-08-13 | Standalone handoff layout introduced: root launcher/config only; working inputs under `inputs/`; SQL under `sql/`; builder under `scripts/`; dated deliverables under `outputs/`; technical documentation consolidated here. All runtime paths updated and regression-tested. |
| 2026-08-13 | Local Snowflake queries switched from direct underlying-object access to two dedicated views: `WEEKLY_RETAIL_ACTUALS` and `MARKETING_SPEND_YTD` in `DATA_MART.RETAIL_DASHBOARD`. Consumer role reduced to warehouse/database/schema `USAGE` plus `SELECT` on those views. |
