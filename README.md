# NHL Team Reports

This repository contains everything needed to generate the high-resolution NHL team reports and convert them to 12k x 16k PNG images. Only the files directly involved in building the reports are included.

## Contents
- `team_report_generator.py` – orchestrates data aggregation, layout, and export to PDF/PNG
- `pdf_report_generator.py` – shared ReportLab helpers used by the generator
- `advanced_metrics_analyzer.py` – utility methods for calculating possession, scoring, and momentum stats
- `nhl_api_client.py` – minimal NHL API wrapper used to pull live season data
- `win_probability_predictions_v2.json` – cached historical predictions that seed the report data set
- `Paper.png` – background texture used on every page
- `fire_1f525.png`, `upwards-black-arrow_2b06.png`, `direct-hit_1f3af.png` – clutch metric icon assets
- `RussoOne-Regular.ttf` – custom display font for headings and charts
- `requirements.txt` – Python dependencies

## Setup
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Generating a report
```python
from team_report_generator import TeamReportGenerator

generator = TeamReportGenerator()
generator.generate_team_report_image('PIT')  # Saves PNG to ~/Desktop by default
```

The PNG export performs an oversampled pdftocairo render to guarantee maximum sharpness when zooming. Temporary header files are created in `/tmp` and cleaned up automatically after each run.
