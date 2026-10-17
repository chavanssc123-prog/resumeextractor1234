# Resume Extractor Job Portal — Flask conversion

This project converts the uploaded single-file Streamlit app into a Flask/Jinja web app. It preserves the uploaded app's skill keyword list, role clusters, job dataset, matching functions, resume parsing/report functions, SQLite account/report storage, and adds HTML pages for login, registration, upload, results, profile, and reports.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000.

## Deploy to Vercel

1. Push this folder's contents to a GitHub repository.
2. Import the repository in Vercel.
3. Add environment variable `SECRET_KEY` with a long random value.
4. Deploy. Vercel's Python runtime can run Flask, but the SQLite database on serverless storage is ephemeral and not suitable for durable multi-user production data. For persistent users/reports, use a managed PostgreSQL database and update the database layer before production.
5. Do not commit real API keys or secrets.

## Notes / limitations

- Original password hashing (SHA-256) was retained for compatibility with the source's implementation; for a production system, migrate to Werkzeug's `generate_password_hash` / `check_password_hash`.
- Resume parsing is keyword/rule-based, as in the source. It is not an external AI model call.
- This conversion implements the core workflow, but it does not reproduce every pixel or every Streamlit widget. Review and test with sample resumes before public use.
- Max upload size is 10 MB. PDF extraction requires text-based PDFs; scanned PDFs need OCR.
