import pandas as pd
import io
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

FILE_IDS = {
    2024: "1niQOJRj4cNzaeZ8RA7Q_CT8qKlLDOmdHyTVRLXWW1gs",
    2025: "1XwzGK9bHGkW-8-3ifYu0t_qtKFa3Dqyl",
    2026: "1HXxWSZ4gouasd-STZpi0w0GX9C5u4K7V",
}

SHEET_NAMES = {
    2024: "2024 PTOT Tracking",
    2025: "2025 PTOT Tracking",
    2026: "2026 PTOT Tracking",
}

MONTH_COL = {2024: 22, 2025: 20, 2026: 20}

# Worker columns after the month name (offset from month col)
# Format: [total, tracy, worker2, worker3, worker4]
WORKER_COLS = {
    2024: ["Total", "Tracy", "Chandler", "Tricia", "Macy"],
    2025: ["Total", "Tracy", "Tricia", "Kristi"],
    2026: ["Total", "Tracy", "Tricia", "Kristi"],
}

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]


def get_drive_service():
    import os, json
    if "GOOGLE_SERVICE_ACCOUNT" in os.environ:
        info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT"])
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/drive.readonly"])
    elif hasattr(st, 'secrets') and "gcp_service_account" in st.secrets:
        creds = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive.readonly"])
    else:
        creds = service_account.Credentials.from_service_account_file(
            "service_account.json",
            scopes=["https://www.googleapis.com/auth/drive.readonly"])
    return build("drive", "v3", credentials=creds)

def download_excel(service, file_id):
    meta = service.files().get(fileId=file_id, fields="mimeType").execute()
    mime = meta.get("mimeType", "")
    if mime == "application/vnd.google-apps.spreadsheet":
        request = service.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf


def parse_monthly_totals(buf, year):
    df = pd.read_excel(buf, sheet_name=SHEET_NAMES[year], header=None)
    col = MONTH_COL[year]
    workers = WORKER_COLS[year]
    results = {}

    for _, row in df.iterrows():
        val = str(row.iloc[col]) if len(row) > col else ""
        if val in MONTHS:
            entry = {"total": 0.0}
            # Parse total + each worker column
            for i, w in enumerate(workers):
                try:
                    v = float(row.iloc[col + 3 + i]) if pd.notna(row.iloc[col + 3 + i]) else 0.0
                except (ValueError, TypeError):
                    v = 0.0
                entry[w.lower()] = v
            entry["total"] = entry.get("total", 0.0)
            results[val] = entry

    # Count jobs per month from job rows (rows with a valid date in col 1)
    job_counts = {m: 0 for m in MONTHS}
    for _, row in df.iterrows():
        date_val = row.iloc[1]
        try:
            date = pd.to_datetime(date_val, errors="coerce")
            if pd.notna(date) and 2000 <= date.year <= 2100:
                month_name = MONTHS[date.month - 1]
                job_counts[month_name] = job_counts.get(month_name, 0) + 1
        except Exception:
            pass

    for m in MONTHS:
        if m in results:
            results[m]["jobs"] = job_counts.get(m, 0)

    return results


def load_all_data():
    service = get_drive_service()
    all_data = {}
    for year in [2024, 2025, 2026]:
        buf = download_excel(service, FILE_IDS[year])
        all_data[year] = parse_monthly_totals(buf, year)
    return all_data


def build_monthly_df(all_data):
    rows = []
    for year, monthly in all_data.items():
        workers = WORKER_COLS[year]
        for month in MONTHS:
            entry = monthly.get(month, {})
            row = {
                "year": year,
                "month": month,
                "month_num": MONTHS.index(month) + 1,
                "total_revenue": entry.get("total", 0.0),
                "jobs": entry.get("jobs", 0),
            }
            for w in workers:
                row[f"worker_{w.lower()}"] = entry.get(w.lower(), 0.0)
            rows.append(row)
    return pd.DataFrame(rows).fillna(0)
