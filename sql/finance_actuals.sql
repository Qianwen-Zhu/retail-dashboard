/*
  Monthly Finance actuals for the Retail Business Dashboard.

  The view owner maintains access to the underlying NetSuite objects. The
  dashboard read role needs SELECT only on this view, plus USAGE on its parent
  database/schema and the reporting warehouse.
*/

SELECT
    PERIOD,
    CATEGORY,
    CHANNEL,
    SPEND
FROM DATA_MART.RETAIL_DASHBOARD.MARKETING_SPEND_YTD;
