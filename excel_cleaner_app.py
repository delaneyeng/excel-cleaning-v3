import io
import re
import sys
import platform
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------------------------------
# Environment health check (openpyxl)
# ---------------------------------
missing_openpyxl = False
try:
    import openpyxl  # noqa: F401
except Exception:
    missing_openpyxl = True

st.set_page_config(page_title="Excel Sensor Cleaner", page_icon="🧹", layout="wide")
st.title("🧹 Excel Sensor Cleaner (Dynamic Columns)")

if missing_openpyxl:
    st.error(
        "**openpyxl is not installed in this Python environment**, so .xlsx files cannot be read yet.\n\n"
        "Install the dependencies *in the same interpreter* you use to launch Streamlit.")
    py = sys.executable or "python"
    st.code(
        f"""
# create/activate a local venv (recommended)
{py} -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate

# upgrade pip and install requirements
python -m pip install --upgrade pip
python -m pip install streamlit pandas openpyxl

# always launch Streamlit from the same interpreter
python -m streamlit run app.py
""",
        language="bash",
    )
    with st.expander("Why am I seeing this? (diagnostics)"):
        st.write(
            {
                "python": sys.version.split(" (", 1)[0],
                "executable": sys.executable,
                "platform": platform.platform(),
            }
        )
        st.info(
            "If you previously ran `pip install openpyxl` under a different Python, Streamlit may still run with another interpreter.\n"
            "Use `python -m pip ...` and `python -m streamlit ...` to guarantee the same interpreter.")
    st.stop()

st.caption(
    "Robust to changing numbers of sensor columns (e.g., ALARM_* or *_C). "
    "Parses text/Excel datetimes, adds a fractional 'Days' column, and coerces sensor readings to numeric."
)

# ------------------------------
# Helpers
# ------------------------------
import re as _re

def normalize_headers(columns):
    norm = []
    for c in columns:
        if c is None:
            norm.append("")
        else:
            s = str(c).replace("\n", " ").strip()
            s = _re.sub(r"\s+", " ", s)
            norm.append(s)
    return norm

SENSOR_DEFAULT_REGEX = r"^(ALARM_|.*_C$)"  # starts with ALARM_ OR ends with _C
META_GUESS = {"S/N", "S N", "SN", "Date/Time", "Date Time", "Datetime", "Date", "Time", "Days"}


def detect_sensor_columns(df: pd.DataFrame, pattern: str = SENSOR_DEFAULT_REGEX):
    headers = normalize_headers(df.columns)
    sensor_cols = []
    rx = _re.compile(pattern, _re.IGNORECASE)
    meta_lower = {m.lower() for m in META_GUESS}
    for h in headers:
        if h.lower() in meta_lower:
            continue
        if rx.search(h):
            sensor_cols.append(h)
    return sensor_cols


def find_datetime_col(df: pd.DataFrame):
    headers = normalize_headers(df.columns)
    candidates = ["Date/Time", "Date Time", "Datetime", "Timestamp", "Date"]
    for c in candidates:
        for h in headers:
            if h.lower() == c.lower():
                return h
    for h in headers:
        s = pd.to_datetime(df[h], errors="coerce", dayfirst=True)
        if s.notna().mean() > 0.7:
            return h
    return None


def coerce_datetime(series: pd.Series) -> pd.Series:
    try:
        s = pd.to_datetime(series, errors="coerce", dayfirst=True)
    except Exception:
        s = pd.Series(pd.NaT, index=series.index)
    if s.isna().mean() > 0.5:
        try:
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
    if dt_col is None or dt_col not in df.columns:
        return df
    dt = df[dt_col]
    if not pd.api.types.is_datetime64_any_dtype(dt):
        dt = coerce_datetime(dt)
        df[dt_col] = dt
    if dt.isna().all():
        return df
    baseline = dt.dropna().iloc[0].normalize()
    days = (dt - baseline) / pd.Timedelta(days=1)
    if "Days" in df.columns:
        df = df.drop(columns=["Days"])  # clean insert
    insert_idx = list(df.columns).index(dt_col) + 1 if dt_col in df.columns else len(df.columns)
    left = df.iloc[:, :insert_idx]
    right = df.iloc[:, insert_idx:]
    df = pd.concat([left, pd.DataFrame({"Days": days}) , right], axis=1)
    return df

# ------------------------------
# Sidebar controls
# ------------------------------
with st.sidebar:
    st.header("1) Upload Excel")
    file = st.file_uploader("Select an .xlsx file", type=["xlsx"])  # requires openpyxl

    st.header("2) Options")
    pattern = st.text_input(
        "Sensor header pattern (regex)",
        value=SENSOR_DEFAULT_REGEX,
        help="Headers matching this regex are treated as sensor columns. Default: starts with 'ALARM_' or ends with '_C'",
    )
    keep_blank_rows = st.checkbox("Keep completely blank rows", value=False)
    preview_rows = st.slider("Preview rows", 5, 100, 25)

# ------------------------------
# Main logic
# ------------------------------
if file is None:
    st.info("Upload an Excel .xlsx file to begin.")
    st.stop()

# List sheets safely
try:
    xls = pd.ExcelFile(file)  # engine auto-detected (needs openpyxl for .xlsx)
    sheet_names = xls.sheet_names
except Exception as e:
    msg = str(e)
    hint = ""
    if "openpyxl" in msg.lower():
        hint = (
            "\n\nIt looks like **openpyxl** is still missing in the interpreter running Streamlit. "
            "Install with: `python -m pip install openpyxl` and restart using `python -m streamlit run app.py`."
        )
    st.error(f"Failed to read the Excel file: {e}{hint}")
    st.stop()

sheet = st.selectbox("Choose a sheet to clean", options=sheet_names)

try:
    raw_df = pd.read_excel(xls, sheet_name=sheet, dtype=object)
except Exception as e:
    msg = str(e)
    hint = ""
    if "openpyxl" in msg.lower():
        hint = (
            "\n\nInstall **openpyxl** in this environment and relaunch the app.\n"
            "Commands: `python -m pip install openpyxl` then `python -m streamlit run app.py`.")
    st.error(f"Failed to read sheet '{sheet}': {e}{hint}")
    st.stop()

# Normalize headers
raw_df.columns = normalize_headers(raw_df.columns)

# Detect columns
dt_col = find_datetime_col(raw_df)
sensor_cols = detect_sensor_columns(raw_df, pattern=pattern)

col1, col2, col3 = st.columns([1.4, 1, 1])
with col1:
    st.subheader("Detected columns")
    st.write({"Datetime": dt_col, "# Sensors": len(sensor_cols), "Sensors": sensor_cols})
with col2:
    st.metric("Total rows", len(raw_df))
with col3:
    st.metric("Total columns", len(raw_df.columns))

st.divider()

st.subheader("Preview — Original")
st.dataframe(raw_df.head(preview_rows), use_container_width=True)

# Cleaning pipeline
clean_df = raw_df.copy()

# Datetime
if dt_col is not None:
    clean_df[dt_col] = coerce_datetime(clean_df[dt_col])
else:
    st.warning("No obvious datetime column detected. Proceeding without 'Days'.")

# Drop fully empty rows
if not keep_blank_rows:
    clean_df = clean_df.dropna(how="all").reset_index(drop=True)

# Coerce sensors
clean_df = coerce_numeric(clean_df, sensor_cols)

# Add Days next to datetime
clean_df = add_days_column(clean_df, dt_col, insert_after=dt_col)

# Sort by datetime if available
if dt_col and pd.api.types.is_datetime64_any_dtype(clean_df[dt_col]):
    clean_df = clean_df.sort_values(dt_col).reset_index(drop=True)

st.subheader("Preview — Cleaned")
st.dataframe(clean_df.head(preview_rows), use_container_width=True)

# Download cleaned Excel
out = io.BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    clean_df.to_excel(writer, sheet_name="Cleaned", index=False)
    meta = pd.DataFrame(
        {
            "Key": ["Generated", "Sheet", "Datetime column", "Sensor pattern", "# Sensors"],
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

st.success("Ready. If your vendor adds/removes sensor columns later, this app adapts automatically.")
