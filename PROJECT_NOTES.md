# Account Book React Project Notes

## Project

- Path: `D:\JAECHUN\react\account-book-react`
- Stack: React + Vite
- Storage: `localStorage`
- Main source files:
  - `src/main.jsx`
  - `src/styles.css`

## Commands

```powershell
npm install
npm run dev -- --host 127.0.0.1
npm run build
```

Local routes:

- `/` dashboard
- `/entry` income/expense input
- `/fixed` fixed income/expense management

## Current Features

- Monthly dashboard with income, expense, balance summary
- Category spending comparison chart
- Recent transaction list
- Manual income/expense spreadsheet-style input
- Thousand separator formatting for amount fields
- Excel upload from `/entry`
- Excel template download
- Excel upload preview before applying to input table
- Duplicate detection during Excel upload
  - Duplicate key: date + type + category + amount + memo
  - Compares against saved transactions and duplicates inside the same Excel file
- Fixed monthly income/expense items
- Lightweight URL routing using `history.pushState` and `popstate`
- Transaction edit modal for normal transactions
- Delete normal transactions
- Fixed transactions are managed through `/fixed`

## localStorage Keys

- Normal transactions: `basic-budget-transactions`
- Fixed transactions: `basic-budget-fixed-transactions`

## Important Implementation Notes

- The app does not use `react-router-dom`; routing is handled manually in `src/main.jsx`.
- Excel support uses `xlsx` as a dynamic import to avoid loading it in the main bundle.
- Amount inputs are stored as formatted strings while editing, then converted with `parseAmount`.
- Fixed monthly items are converted to virtual transactions for the selected month.
- Entry rows reset after save. Saved rows are visible in dashboard recent history and entry-side recent save/history panel.

## Known Notes

- `npm install xlsx` reported one high severity vulnerability. Build works, but audit before production use.
- There are no automated tests yet.
- No backend or authentication yet.

## Good Next Tasks

1. Category management
   - Let users add/edit/delete categories.
   - Persist categories in localStorage.

2. Budget settings
   - Monthly category budgets.
   - Show used percentage on dashboard.

3. Monthly trend chart
   - Recent 6 months income/expense/balance.

4. Excel export
   - Download saved transactions as `.xlsx`.

5. Search and advanced filters
   - Search by memo/category.
   - Filter by date range, amount range, type.

6. Better Excel upload mapping
   - Let users choose which Excel column maps to date/type/category/amount/memo.

## Recent User Preferences

- Korean UI.
- Dashboard should be statistics-first.
- Input page should use spreadsheet-style rows.
- Buttons should be refined and not look generic.
- Font changed to `IBM Plex Sans KR` with fallbacks.
- Fixed item button should be visually distinct.

