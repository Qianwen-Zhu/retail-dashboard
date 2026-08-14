# Retail Business Dashboard

Start here. This folder is a standalone weekly dashboard workflow and is not
part of the CAC Tracker.

## Weekly refresh

From this folder, run one command:

```bash
.venv/bin/python run_retail_dashboard.py
```

No virtual-environment activation is required. The command:

1. Runs `sql/finance_actuals.sql`.
2. Updates `inputs/finance/Marketing_Spend_2026YTD.csv` only when Snowflake
   contains a newer month. If the latest month is unchanged, the file is not
   touched.
3. Runs `sql/weekly_retail_actuals.sql`.
4. Uses a temporary weekly CSV while building, then removes it automatically.
5. Saves the dated Excel and HTML files in `outputs/YYYY-MM-DD/`.

To run only the Finance month check:

```bash
.venv/bin/python run_retail_dashboard.py --finance-check-only
```

To assign a specific dashboard date:

```bash
.venv/bin/python run_retail_dashboard.py --snapshot-date YYYY-MM-DD
```

## Folder map

```text
Retail Business Dashboard/
├── README.md                     Start here
├── run_retail_dashboard.py       Weekly one-command entry point
├── .env                          Snowflake credentials; never share or commit
├── .env.example                  Safe configuration example
├── requirements.txt              Python packages
│
├── inputs/
│   ├── Dashboard_Template.xlsx   Fixed dashboard workbook template
│   ├── FY26_Marketing_Plan.xlsx  Monthly marketing plan input
│   ├── FY26_Promo_Plan.xlsx      Monthly promotion plan input
│   ├── finance/
│   │   └── Marketing_Spend_2026YTD.csv
│
├── sql/
│   ├── weekly_retail_actuals.sql
│   ├── finance_actuals.sql
│   └── grant_retail_dashboard_read_role.sql
│
├── scripts/
│   └── build_outputs.py          CSV → Excel + HTML builder
│
├── outputs/                      Final files grouped by dashboard date
├── docs/
│   └── TECHNICAL_REFERENCE.md    Metrics, mappings, caveats, and history
└── tests/                        Offline regression tests
```

## One-time setup

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp .env.example .env
```

Complete `.env` with the read-only Snowflake account. Store the PKCS#8 private
key directly in `SNOWFLAKE_PRIVATE_KEY` on one line, replacing every original
line break with the two characters `\n`:

```dotenv
SNOWFLAKE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----"
```

The script restores the line breaks in memory and does not use a separate key
file. Set `SNOWFLAKE_PRIVATE_KEY_PASSPHRASE` only for an encrypted key.

The read-only role queries only these two views:

- `DATA_MART.RETAIL_DASHBOARD.WEEKLY_RETAIL_ACTUALS`
- `DATA_MART.RETAIL_DASHBOARD.MARKETING_SPEND_YTD`

It needs `USAGE` on `SALES_WH`, `DATA_MART`, and
`DATA_MART.RETAIL_DASHBOARD`, plus `SELECT` on those two views. It does not
need permissions on the underlying source tables. The exact grant script is
`sql/grant_retail_dashboard_read_role.sql`.

## What may be replaced

- `inputs/Dashboard_Template.xlsx`: replace only when the approved dashboard
  format changes. Row mappings in `scripts/build_outputs.py` may also need an
  update.
- `inputs/FY26_Marketing_Plan.xlsx`: replace when the operating plan changes;
  keep the expected worksheet layout.
- `inputs/FY26_Promo_Plan.xlsx`: replace when the promo plan changes; keep the
  expected worksheet layout.
- `inputs/finance/Marketing_Spend_2026YTD.csv`: normally maintained
  automatically. Manual replacement is allowed only when the Snowflake result
  contains a newer month.

Do not edit files inside `outputs/` and then treat them as templates. Each run
starts from `inputs/Dashboard_Template.xlsx`.

## Manual fallback

If the automated Snowflake connection is unavailable:

1. Run `sql/finance_actuals.sql` manually and update the Finance CSV only if a
   newer month exists.
2. Run `sql/weekly_retail_actuals.sql` and save the CSV anywhere, with a date
   in its filename.
3. Run:

```bash
.venv/bin/python scripts/build_outputs.py \
  "/path/to/Retail Business Dashboard_YYYY-MM-DD-HHMM.csv"
```

The output is written to `outputs/YYYY-MM-DD/`. Use `--snapshot-date` or
`--output-dir` only when a non-default date or location is required.

## Handoff checks

Before sharing the weekly output:

- Confirm the highlighted week is the latest complete Mon–Sun week.
- Confirm all retail sections are present.
- Compare the latest Excel and HTML with the prior week for obvious anomalies.
- Do not rename input files unless their paths are also changed in
  `scripts/build_outputs.py`.

For metric definitions, source tables, known limitations, and the full change
history, see `docs/TECHNICAL_REFERENCE.md`.
