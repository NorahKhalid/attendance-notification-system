# Attendance Notification System V12

## New in V12
- One Gmail draft per employee containing all monthly attendance cases.
- Monthly archive with employee-grouped cases.
- Mark each case as sent/not sent.
- Add reply notes for each case.
- Monthly filters: employee, case type, project, send status.
- Export archived month to Excel or CSV.
- Re-saving a month refreshes the analysis while preserving existing sent/reply tracking.
- Dashboard and analysis remain available while navigating pages.
- Saturday schedule rules: Saturday workers = 8h; non-Saturday workers = 9h Sunday-Thursday; Friday is flagged as overtime/special day.
- Gmail is draft/review only; nothing is sent automatically.

## Streamlit Cloud note
SQLite is local to the app runtime. Use the built-in backup/download regularly. For true cross-restart persistence on Streamlit Cloud, connect an external database (e.g. Supabase/Postgres) or a controlled GitHub persistence layer later.
