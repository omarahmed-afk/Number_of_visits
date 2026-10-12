# Chart scraping

Python script that collects WebPT schedule counts and writes an Excel report.

## Setup

Install Python and Google Chrome, then install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Copy `local_settings.example.json` to `local_settings.json` and enter your WebPT credentials. Alternatively, set the `WEBPT_USERNAME` and `WEBPT_PASSWORD` environment variables; these take precedence over the local file.

Review the output paths and clinic settings in `collect_visit.py`, then run:

```powershell
python collect_visit.py
```

Local credentials, generated spreadsheets, and logs are excluded from Git.
