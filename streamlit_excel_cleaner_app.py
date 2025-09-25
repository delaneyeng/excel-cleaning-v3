import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st

# ------------------------------
# Page / layout
# ------------------------------
st.set_page_config(
    page_title="Excel Sensor Cleaner",
    page_icon="🧹",
    layout="wide",
)

st.title("🧹 Excel Sensor Cleaner (Dynamic Columns)")
st.caption(
    "Robust to changing numbers of sensor columns (e.g., ALARM_* or *_C). "
    "Handles text or true Excel datetimes, trims headers, and coerces sensor values to numeric."
)

# ------------------------------
# Helpers
# ------------------------------

def normalize_headers(columns):
    """Strip/condense spaces in headers and return list of strings."""
    norm = []
    for c in columns:
        if c is None:
            norm.append("")
        else:
            s = str(c).replace("\n", " ").strip()
            s = re.sub(r"\s+", " ", s)
            norm.append(s)
    return norm


SENSOR_DEFAULT_REGEX = r"^(ALARM_|.*_C$)"  # starts with ALARM_ OR ends with _C
META_GUESS = {"S/N", "S N", "SN", "Date/Time", "Date Time", "Datetime", "Date", "Time", "Days"}


def detect_sensor_columns(df: pd.DataFrame, pattern: str = SENSOR_DEFAULT_REGEX):
    """Return list of sensor column names by regex match on header.
    - pattern defaults to ALARM_* or *_C
    - excludes known meta columns
    """
    headers = normalize_headers(df.columns)
    sensor_cols = []
    rx = re.compile(pattern, re.IGNORECASE)
    for h in headers:
        if h.lower() in {m.lower() for m in META_GUESS}:
            continue
        if rx.search(h):
            sensor_cols.append(h)
    return sensor_cols


def find_datetime_col(df: pd.DataFrame):
    """Pick a likely datetime column (case-insensitive)."""
    headers = normalize_headers(df.columns)
    candidates = [
        "Date/Time",
        "Date Time",
        "Datetime",
        "Timestamp",
        "Date",
    ]
    for c in candidates:
        for h in headers:
            if h.lower() == c.lower():
                return h
    # fallback: first column that parses to many datetimes
    for h in headers:
        s = pd.to_datetime(df[h], errors="coerce", dayfirst=True)
        if s.notna().mean() > 0.7:
            return h
    return None


def coerce_datetime(series: pd.Series) -> pd.Series:
    """Coerce a series to datetime (accept text or Excel serial)."""
    try:
        # Try pandas automatic conversion with dayfirst True (fits DD/MM/YYYY HH:MM AM/PM)
        s = pd.to_datetime(series, errors="coerce", dayfirst=True)
    except Exception:
        s = pd.Series(pd.NaT, index=series.index)
    # If still many NaT, try Excel serial numbers
    if s.isna().mean() > 0.5:
        try:
            # Excel serial date handling via to_datetime with origin
            s2 = pd.to_datetime(pd.to_numeric(series, errors="coerce"), unit="D", origin="1899-12-30")
            s = s.fillna(s2)
        except Exception:
            pass
    return s


def coerce_numeric(df: pd.DataFrame, cols):
    for c in cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df


def add_days_column(df: pd.DataFrame, dt_col: str, insert_after: str | None = None):
    """Add/overwrite a 'Days' column measuring *fractional days* from the first timestamp's date.
    - If 'Days' already exists, it is replaced.
    - Insert next to dt_col if possible.
    """
    if dt_col is None or dt_col not in df.columns:
        return df
    dt = df[dt_col]
    if not pd.api.types.is_datetime64_any_dtype(dt):
        dt = coerce_datetime(dt)
    if dt.isna().all():
        # nothing to do
        return df
    # Start-of-day baseline from first non-null timestamp
    baseline = dt.dropna().iloc[0].normalize()
    days = (dt - baseline) / pd.Timedelta(days=1)
    if "Days" in df.columns:
        df = df.drop(columns=["Days"])  # ensure clean insert
    # Decide insert position
    insert_idx = len(df.columns)
    if insert_after and insert_after in df.columns:
        insert_idx = list(df.columns).index(insert_after) + 1
    elif dt_col in df.columns:
        insert_idx = list(df.columns).index(dt_col) + 1
    left = df.iloc[:, :insert_idx]
    right = df.iloc[:, insert_idx:]
    df = pd.concat([left, pd.DataFrame({"Days": days}), right], axis=1)
    return df


# ------------------------------
# Sidebar controls
# ------------------------------
with st.sidebar:
    st.header("1) Upload Excel")
    file = st.file_uploader("Select an .xlsx file", type=["xlsx"])  # openpyxl backend

    st.header("2) Options")
    pattern = st.text_input(
        "Sensor header pattern (regex)",
        value=SENSOR_DEFAULT_REGEX,
        help="Headers matching this regex will be treated as sensor columns. Default: starts with 'ALARM_' or ends with '_C'",
    )
    keep_blank_rows = st.checkbox("Keep completely blank rows", value=False)
    preview_rows = st.slider("Preview rows", 5, 100, 25)

    st.caption("Tip: If your vendor adds more sensors later, you don't need to touch the code.")

# ------------------------------
# Main logic
# ------------------------------
if file is None:
    st.info("Upload an Excel file to begin.")
    st.stop()

# Load workbook with pandas (list sheets)
try:
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names
except Exception as e:
    st.error(f"Failed to read the Excel file: {e}")
    st.stop()

sheet = st.selectbox("Choose a sheet to clean", options=sheet_names)

try:
    raw_df = pd.read_excel(xls, sheet_name=sheet, dtype=object, engine="openpyxl")
except Exception as e:
    st.error(f"Failed to read sheet '{sheet}': {e}")
    st.stop()

# Normalize headers early
raw_df.columns = normalize_headers(raw_df.columns)

# Guess columns
dt_col = find_datetime_col(raw_df)
sensor_cols = detect_sensor_columns(raw_df, pattern=pattern)

col1, col2, col3 = st.columns([1.2, 1, 1])
with col1:
    st.subheader("Detected columns")
    st.write(
        {
            "Datetime": dt_col,
            "# Sensors": len(sensor_cols),
            "Sensors": sensor_cols,
        }
    )
with col2:
    st.metric("Total rows", len(raw_df))
with col3:
    st.metric("Total columns", len(raw_df.columns))

st.divider()

st.subheader("Preview — Original")
st.dataframe(raw_df.head(preview_rows), use_container_width=True)

# Cleaning pipeline
clean_df = raw_df.copy()

# Coerce datetime
if dt_col is not None:
    clean_df[dt_col] = coerce_datetime(clean_df[dt_col])
else:
    st.warning("No obvious datetime column detected. You can continue, but 'Days' won't be added.")

# Drop fully empty rows unless asked to keep
if not keep_blank_rows:
    clean_df = clean_df.dropna(how="all").reset_index(drop=True)

# Coerce sensor columns to numeric
clean_df = coerce_numeric(clean_df, sensor_cols)

# Add/refresh Days column right after the datetime column
clean_df = add_days_column(clean_df, dt_col, insert_after=dt_col)

# Sort by datetime if available (keeps original order otherwise)
if dt_col and pd.api.types.is_datetime64_any_dtype(clean_df[dt_col]):
    clean_df = clean_df.sort_values(dt_col).reset_index(drop=True)

st.subheader("Preview — Cleaned")
st.dataframe(clean_df.head(preview_rows), use_container_width=True)

# Download cleaned Excel
out = io.BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    clean_df.to_excel(writer, sheet_name="Cleaned", index=False)
    # Also include a tiny README sheet with metadata
    meta = pd.DataFrame(
        {
            "Key": [
                "Generated",
                "Sheet",
                "Datetime column",
                "Sensor pattern",
                "# Sensors",
            ],
            "Value": [
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                sheet,
                dt_col or "<not detected>",
                pattern,
                len(sensor_cols),
            ],
        }
    )
    meta.to_excel(writer, sheet_name="README", index=False)

st.download_button(
    label="⬇️ Download Cleaned Excel",
    data=out.getvalue(),
    file_name=f"cleaned_{sheet}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.success("Done. If the vendor adds or removes sensor columns, this app will adapt automatically.")
