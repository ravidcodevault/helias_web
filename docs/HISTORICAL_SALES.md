# Historical sales and inactive items

Dashboard and Reports summary cards call `dbo.USP_GetOverallDashboardSales_V2`.
Excel exports call `dbo.USP_GetSalesDataForReports`. Their result shapes differ;
changing the Excel procedure alone cannot update the summary cards.

Apply `database/migrations/20261031_include_inactive_items_in_sales.sql` to the
database used by the hosted application's `ConnectionStrings:DefaultConnection`.
Use SSMS or your normal SQL deployment tooling. Publishing/restarting the web app
does not deploy stored procedure changes.

The migration removes the item `IsActive = 1` predicate from the invoice, sales
receipt, and credit memo branches of both procedures. It preserves distributor
activation filters, date ranges, pending/deleted transaction exclusions, sales
tax exclusions, credit signs, category calculations, and result columns. Existing
procedure customizations are retained. An already updated procedure is a no-op;
unrecognized item-filter formatting/counts abort the transaction for review.
Save the current definitions before production deployment if rollback is needed.

Confirm the hosted database first. An IIS process environment variable such as
`ConnectionStrings__DefaultConnection` can override `appsettings.json`. Inspect
the deployed app's effective configuration without sharing credentials, then run
`SELECT @@SERVERNAME AS SqlServer, DB_NAME() AS DatabaseName` on that connection.
The source checkout does not establish the hosted application's database target.

Validation should include an active item's sales and an inactive item's historical
sales, plus an inactive-item credit memo. Pending/deleted transactions and sales
tax must stay excluded. Re-run for the same year/month and refresh the UI/download
the Excel file. A month's totals will remain unchanged if there are no qualifying
inactive-item transactions in that period.
