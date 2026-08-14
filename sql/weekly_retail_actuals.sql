/*
  Weekly actuals for the Retail Business Dashboard.

  The view owner maintains access to all underlying retail, ads, promo, and
  customer objects. The dashboard read role needs SELECT only on this view,
  plus USAGE on its parent database/schema and the reporting warehouse.
*/

SELECT *
FROM DATA_MART.RETAIL_DASHBOARD.WEEKLY_RETAIL_ACTUALS;
