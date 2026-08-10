## V9.1 database migration fix

This version fixes the `KeyError: ['project', 'project_priority']` that occurs when
a Streamlit Cloud deployment already has an older SQLite database.

The app now:
- automatically adds missing `project` and `project_priority` columns;
- preserves existing employee records;
- keeps compatibility with older databases;
- continues to support the new project fields.

No Excel files or confidential employee data are included.
