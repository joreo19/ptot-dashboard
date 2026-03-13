import pandas as pd
import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# --- File IDs from Google Drive ---
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

# Column index where month name starts (0-based)
MONTH_COL = {
    2024: 22,
    2025: 20,
    2026: 20,
}

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]


def get_drive_service():
    creds = service_account.Credentials.from_service_account_file(
        "service_account.json",
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    return build("drive", "v3", credentials=creds)


def download_excel(service, file_id):
    # Check the file's MIME type first
    meta = service.files().get(fileId=file_id, fields="mimeType").execute()
    mime = meta.get("mimeType", "")

    if mime == "application/vnd.google-apps.spreadsheet":
        # Native Google Sheet — export as Excel
        request = service.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        # Regular .xlsx file stored in Drive — download directly
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
    results = {}
    for _, row in df.iterrows():
        val = str(row.iloc[col]) if len(row) > col else ""
        if val in MONTHS:
            total_rev = row.iloc[col + 3]
            tracy_rev = row.iloc[col + 4]
            try:
                results[val] = {
                    "total": float(total_rev) if pd.notna(total_rev) else 0.0,
                    "tracy": float(tracy_rev) if pd.notna(tracy_rev) else 0.0,
                }
            except (ValueError, TypeError):
                results[val] = {"total": 0.0, "tracy": 0.0}
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
        for month in MONTHS:
            entry = monthly.get(month, {"total": 0.0, "tracy": 0.0})
            rows.append({
                "year": year,
                "month": month,
                "month_num": MONTHS.index(month) + 1,
                "total_revenue": entry["total"],
                "tracy_revenue": entry["tracy"],
            })
    return pd.DataFrame(rows)
