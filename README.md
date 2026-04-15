# Project Performance

A Flask web app for uploading and processing project performance data — including labor hour estimates, actuals, costs, and scope information — for reporting and analysis.

## What It Does

- Accepts project data uploads via a web form (estimates, CADO actuals, team mapping files)
- Parses and transforms Excel/CSV files into clean, structured tables
- Maps employee cost centers and expense categories to function groups (e.g. RF, HW, SW)
- Computes labor hour variances between estimated and actual hours
- Outputs structured data across five tables:
  - `Project_Master` — project dates and classification
  - `Fact_LaborHours` — total estimated vs actual hours per project
  - `Fact_TeamHours` — hours broken down by cost center and function group
  - `Fact_Costs` — material, dev, and NPI costs at TG1/TG3/TG4
  - `Fact_ScopeClusters` — scope cluster, difficulty, and key functions involved
- Designed to write to a SQL Server database (connection currently commented out for local development)

## Project Structure

```
Project_Performance/
├── run (4).py          # Main Flask app
├── templates/          # HTML form templates
├── uploads/            # Uploaded files (not committed)
├── requirements.txt    # Python dependencies
├── Table_Design.sql    # Database schema
└── .gitignore
```

## Input Files

| File | Description |
|---|---|
| Estimate file | Excel file with `Expense Category` and `Original hourly estimates (hours)` columns, or a raw ToolReady Summary sheet |
| CADO file | Excel/CSV with actual hours by employee and cost center |
| Team mapping file | Excel/CSV mapping cost centers and expense categories to function groups |

## Database

The app is built to write to a Microsoft SQL Server database (`ProjectPerformance`). The connection is currently commented out in the code. To enable it, configure these environment variables and uncomment the engine setup in `run (4).py`:

```bash
export MSSQL_SERVER=your_server
export MSSQL_DB=ProjectPerformance
export MSSQL_USER=your_user
export MSSQL_PASSWORD=your_password
```

The database schema is defined in `Table_Design.sql`.

## Dependencies

- Flask
- pandas
- openpyxl
- SQLAlchemy
- pyodbc (for SQL Server connection)