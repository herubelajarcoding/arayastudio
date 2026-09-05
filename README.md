# Studio Control Board — Streamlit V1

V1 scope:
- Setup / Master Data
  - Projects
  - Staff
- Daily Activities
  - Work
  - Meeting
  - Other
- Weekly Schedule Dashboard
  - Month filter
  - Project filter
  - Week filter
  - Work / Meeting / Other lanes
  - Project grouping
  - Submission sign
  - Other remains independent of Project filter

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app seeds its initial SQLite database from:

`data/Studio_Control_Board_FINAL.xlsm`

After the first run, daily input is stored in `data/studio_control.db`.

## GitHub / Streamlit

Upload:
- `app.py`
- `requirements.txt`
- `data/Studio_Control_Board_FINAL.xlsm`

Then deploy `app.py` as the Streamlit entry point.

This V1 intentionally does not implement SDM, workload, Gantt, financial or other future dashboards yet. The master/activity data structures are retained so those modules can be added later.
