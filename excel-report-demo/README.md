# Demo: raw CSV → formatted Excel report

Takes a plain sales export (`sales.csv`, 600 fake orders) and produces `report.xlsx` with:
- **Summary** sheet (total revenue, orders, average order value, best product/region)
- **By month** sheet with a line chart
- **By product** sheet with a bar chart
- **Region × month** pivot with a red-to-green heat map
- **Raw data** sheet with filters and frozen header

All sheets have styled headers, auto-sized columns and currency formatting.

## Run it
```bash
pip install pandas openpyxl
python make_sample_data.py          # creates sales.csv (fake data)
python make_report.py sales.csv report.xlsx
```
Sample output: [`report.xlsx`](report.xlsx).

Built by Clearcut Automation. Have a CSV export you keep reformatting by hand? This is a one-click job.
