# FleetOps Callout Tracker

Internal operations dashboard for tracking contractor staffing, attendance, callouts, vehicles, reports, and leadership readiness metrics across multiple sites and shifts.

Live app: https://callout-tracker.vercel.app

## What It Solves

FleetOps Callout Tracker helps operations teams replace scattered spreadsheets and messages with one shared command center.

- Contractors can submit callouts from a simple landing page.
- Admins can track staffing availability by site and shift.
- Managers can record attendance, tardy status, no-shows, early departures, and time off.
- Leadership can review daily readiness, attendance trends, callout volume, and vehicle utilization.
- Teams can export reports and backup operational data.

## Core Features

- Operator callout form with name, date, shift, reason, and notes
- Admin dashboard with daily staffing KPIs
- Site support for San Francisco, Miami, and Chicago
- 1st Shift and 2nd Shift operations separation
- Staffing/contractor management with active, inactive, and offboarding status
- Weekly Shift Attendance tracker
- Calendar view for attendance and operational adjustments
- Callout Center with manual admin callout entry
- Vehicle inventory, CSV/XLSX import, sorting, filtering, and operator pairing
- Contractor performance scorecards
- Reports generator with CSV export
- Backup and restore for operational data
- Notifications feed for operations events
- Admin login, password reset, profile display name, and password change

## Tech Stack

- React
- Vite
- Supabase
- Vercel
- XLSX import support

## Local Setup

Install dependencies:

```bash
npm install
```

Start the local app:

```bash
npm run dev
```

Build for production:

```bash
npm run build
```

Preview the production build:

```bash
npm run preview
```

## Environment Variables

Create a local `.env` file with:

```bash
VITE_SUPABASE_URL=your_supabase_project_url
VITE_SUPABASE_ANON_KEY=your_supabase_anon_key
```

The app can run with local browser storage if Supabase variables are missing, but shared team data requires Supabase.

## Deployment

The app is deployed on Vercel.

Deploy production:

```bash
npx vercel --prod
```

Production URL:

https://callout-tracker.vercel.app

## Main Workflows

1. Operators open the callout landing page and submit a callout.
2. Admins review Daily Ops for scheduled, available, callout, and readiness status.
3. Managers update Shift Attendance and Daily Operator Status.
4. Admins maintain contractors and mark offboarding contractors when they should leave operations views.
5. Vehicle coordinators add/import vehicles and pair operators to vehicles.
6. Leadership exports reports or full spreadsheet data for review.

## Notes

- Offboarding contractors stay visible in Staffing but are hidden from operational dashboards.
- Daily statuses like Late, No Show, Leaving Early, and Day Off Requested reset back to Scheduled after the logged date passes.
- Delete actions require password confirmation.
- Backup/restore is available in Reports.
