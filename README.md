# Trade Email Builder

A Streamlit app that converts single-trade or multi-trade CSV/XLSX files into an editable Agile-style email table and downloadable Outlook `.eml` file.

## Features

- Reads CSV, XLSX, and XLSM files.
- Supports one trade, multiple trades across columns, or one trade per row.
- Maps source fields into the target email schema.
- Shows validation and review warnings.
- Lets users edit recognised fields before generating the email.
- Downloads an Outlook-compatible `.eml` file, HTML preview, and mapping JSON.
- Processes uploaded files in the active app session without intentionally saving them permanently.

## Repository structure

```text
trade-email-automation/
├── .streamlit/
│   └── config.toml
├── sample_data/
│   └── synthetic_multi_trade.csv
├── .gitattributes
├── .gitignore
├── app.py
├── field_mapping.json
├── GITHUB_AND_STREAMLIT_GUIDE.md
├── README.md
├── requirements.txt
└── trade_mapper.py
```

## Run locally

Use Python 3.12 for consistency with Streamlit Community Cloud.

```powershell
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Open `http://localhost:8501` if the browser does not open automatically.

## Deploy

Follow `GITHUB_AND_STREAMLIT_GUIDE.md` for the complete first-time GitHub and Streamlit Community Cloud setup.

## Important

- The cloud version downloads an `.eml` file; it does not open desktop Outlook directly.
- Review every draft before sending.
- Do not commit live trade files, client data, passwords, or Streamlit secrets to GitHub.
- Confirm internal approval before processing confidential data on an external cloud service.
