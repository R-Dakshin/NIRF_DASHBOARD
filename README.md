# NIRF Engineering Analytics Dashboard

Interactive, production-ready analytics dashboard for NIRF Engineering Rankings (2016-2025), powered by Excel-driven computed analytics.

## Data Source

Data was collected from the official **NIRF (National Institutional Ranking Framework) website** of the Ministry of Education, Government of India.

- Official portal: [https://www.nirfindia.org](https://www.nirfindia.org)
- Category used: Engineering
- Time span: 2016 to 2025

## Key Capabilities

- Multi-page analytics dashboard with:
  - Overview
  - Descriptive analytics
  - Diagnostic analytics
  - Regression and feature importance
  - Predictive forecasting
  - Advanced analytics:
    - Rank Mobility
    - K-Means Clustering
    - PCA
    - Hypothesis Testing
    - Outlier Detection
    - Rank Stability
    - State Dominance
    - Score Validation
  - Explorer with search/filter/sort
- Runtime data computation from Excel (`data.xlsx`)
- Safe normalization of rank anomalies (e.g., `21A`, missing ranks)
- Local server support + Vercel deployment support

## Project Structure

- `dash.html` - Main dashboard frontend
- `sync_dashboard.py` - Core analytics/data-engine pipeline
- `dashboard_server.py` - Local development server (`/api/data`)
- `api/data.py` - Vercel serverless API endpoint
- `data.xlsx` - Primary dataset used by pipeline
- `requirements.txt` - Python dependencies
- `vercel.json` - Vercel routing configuration

## Local Setup

### 1) Install dependencies

```bash
pip install -r requirements.txt
```

### 2) Start local server

```bash
python dashboard_server.py
```

### 3) Open dashboard

- [http://127.0.0.1:8000/dash.html](http://127.0.0.1:8000/dash.html)


## Troubleshooting

- **Charts not loading**
  - Verify API response: open `/api/data` directly in browser.
- **Dashboard says server is not running**
  - Open through server URL, not `file://`.
  - Use `http://127.0.0.1:8000/dash.html`.
- **Python dependency errors**
  - Reinstall with `pip install -r requirements.txt`.
- **Port already in use**
  - Stop process using port `8000` or change `PORT` in `dashboard_server.py`.

## Disclaimer

This dashboard is for academic/analytical purposes only.  
All rights to original ranking data belong to NIRF / Ministry of Education, Government of India.
