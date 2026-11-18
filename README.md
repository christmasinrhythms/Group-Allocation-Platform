# Group Allocation Platform

[Open the workbench](https://christmasinrhythms.github.io/Group-Allocation-Platform/)

Manage table and game-group assignments in a browser. Table 1 is marked as the head table. The demo uses fictional names.

## Connect a private Excel workbook

1. Keep the workbook in a restricted Microsoft 365 SharePoint library or work OneDrive. Use a link for people with existing access and keep Excel AutoSave on. Use this cloud workbook as the shared working copy.
2. Register a single-tenant app in Microsoft Entra and add `https://christmasinrhythms.github.io/Group-Allocation-Platform/auth.html` under **Authentication → Single-page application**. Add Microsoft Graph's delegated `Files.ReadWrite` permission. Complete any approval your organization requires. A client secret is unnecessary.
3. Open the workbench. Paste the private workbook link and enter the Entra tenant and application IDs under **Microsoft 365 settings**. Sign in with an account that already has workbook editing rights.
4. To use the desktop app alongside the board, open the same cloud workbook from Excel Online and turn on AutoSave. The board checks for updates every 12 seconds while visible. Moves and swaps update the same workbook through Microsoft Graph. Wait for the save confirmation before leaving the board.

## Workbook layout

- `Assignments!A7:G134` holds the roster ID, name, department, business group, unit, game group and table.
- `Overview!A7:B18` holds game group numbers and names.
- Optional `Previous Seating!A2:B129` holds matching roster IDs and last year's table numbers. This sheet enables the historical seating and separation details.

The public repository contains the workbench, Microsoft sign-in library and fictional demo. Keep actual workbooks and historical seating in the restricted Microsoft 365 library. Workbook links and Microsoft tokens are held in browser memory during a connection.
