#streamlit run pptgen.py --server.port 8502

# app.py

import io
import os
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

from exportPPT import MonthlyPerformancePPT


# -----------------------------
# Page setup
# -----------------------------
st.set_page_config(page_title="Excel → PPT Generator (Quarter, In‑Memory)", page_icon="📑", layout="wide")
st.title("▶️ Generate PPTs from Excel/CSV (In‑Memory Downloads)")

st.caption(
    "Required columns (exact names): "
    "`Country`, `Bike_Model`, `Quarter` (e.g., Q1-2025), `Sales Units`, `Revenue_INR` (optional). "
    "Nothing is written to disk; download PPTs directly below."
)

# -----------------------------
# Session state
# -----------------------------
if "sheet_choices" not in st.session_state:
    st.session_state.sheet_choices = {}

# Persist generated results across reruns:
# Each item: (filename: str, blob: bytes)
if "ppt_results" not in st.session_state:
    st.session_state.ppt_results = []
if "ppt_generated_at" not in st.session_state:
    st.session_state.ppt_generated_at = None

# If uploads change, you may want to clear previous results (optional)
def clear_results():
    st.session_state.ppt_results = []
    st.session_state.ppt_generated_at = None

# -----------------------------
# Sidebar: uploads
# -----------------------------
with st.sidebar:
    st.header("Upload Excel/CSV")
    uploads = st.file_uploader(
        "Upload one or more files",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        on_change=clear_results,  # clear old results when new files are chosen
        key="uploads_widget",
    )

# -----------------------------
# Helpers
# -----------------------------
def parse_uploaded_file(file, chosen_sheet=None):
    """
    Returns (df, sheets, is_excel, error_msg)
    - Tries Excel first (openpyxl). If Excel, returns sheet names and uses chosen_sheet (or first).
    - Falls back to CSV if Excel parsing fails.
    """
    try:
        data = file.getvalue()  # bytes
        bio = io.BytesIO(data)

        # Try Excel
        try:
            xl = pd.ExcelFile(bio, engine="openpyxl")
            sheets = xl.sheet_names
            if not sheets:
                return None, [], False, "No worksheets found."
            sheet_to_use = chosen_sheet if (chosen_sheet in sheets) else sheets[0]

            bio2 = io.BytesIO(data)
            df = pd.read_excel(bio2, sheet_name=sheet_to_use, engine="openpyxl")
            return df, sheets, True, None
        except Exception as e:
            print(f"Excel parse failed for {getattr(file, 'name', '')}: {type(e).__name__} - {e}")

            # CSV fallback
            bio.seek(0)
            try:
                df_csv = pd.read_csv(bio)
                return df_csv, [], False, None
            except Exception as e_csv:
                return None, [], False, f"Could not parse as Excel or CSV: {e_csv}"

    except Exception as e:
        return None, [], False, f"Read error: {e}"

# -----------------------------
# 1) Upload previews (right pane)
# -----------------------------
st.subheader("1) Uploaded Files Preview")

if not uploads:
    st.info("Upload one or more Excel/CSV files from the sidebar.")
else:
    for i, up in enumerate(uploads):
        df0, sheets0, is_excel0, err0 = parse_uploaded_file(up)
        if err0:
            st.error(f"**{up.name}** — {err0}")
            st.divider()
            continue

        st.markdown(f"**File:** {up.name}")
        if is_excel0 and sheets0:
            default_sheet = st.session_state.sheet_choices.get(up.name, sheets0[0])
            chosen = st.selectbox(
                f"Worksheet for {up.name}",
                sheets0,
                index=(sheets0.index(default_sheet) if default_sheet in sheets0 else 0),
                key=f"sheet_{up.name}_{i}",
            )
            st.session_state.sheet_choices[up.name] = chosen
            df, sheets, is_excel, err = parse_uploaded_file(up, chosen_sheet=chosen)
        else:
            df, sheets, is_excel, err = df0, sheets0, is_excel0, err0

        if err:
            st.error(f"{err}")
        elif df is None or df.empty:
            st.info("No rows to display.")
        else:
            st.dataframe(df.head(50), use_container_width=True, height=320)

        st.divider()

# -----------------------------
# 2) Generate (in-memory) & persist results
# -----------------------------
st.subheader("2) Generate PPTs")
st.caption("We’ll combine all uploaded files and generate PPTs in memory. Downloads stay visible after you click them.")

if st.button("🚀 Generate PPTs", key="generate_btn"):
    if not uploads:
        st.error("Please upload at least one file.")
    else:
        with st.spinner("Generating Country-wise PowerPoints…"):
            # Combine data from all files / selected sheets
            dfs = []
            for up in uploads:
                sheet_choice = st.session_state.sheet_choices.get(up.name)
                df, _, _, err = parse_uploaded_file(up, chosen_sheet=sheet_choice)
                if err:
                    st.error(f"{up.name} — {err}")
                    continue
                if df is not None and not df.empty:
                    dfs.append(df)

            if not dfs:
                st.error("No valid data found in the uploaded files.")
            else:
                combined = pd.concat(dfs, ignore_index=True)

                gen = MonthlyPerformancePPT(
                    template_ppt=None,
                    logo_bytes=None,
                    last_n_quarters=6
                )
                try:
                    results = gen.generate_from_dataframe(combined)  # List[(filename, bytes)]
                    # ✅ Persist results so they survive reruns
                    st.session_state.ppt_results = results
                    st.session_state.ppt_generated_at = datetime.now()
                    st.success(f"✅ Generated {len(results)} PPTX file(s).")
                except Exception as e:
                    st.error(f"Generation failed: {e}")

st.markdown("---")

# -----------------------------
# 3) Downloads (read from session_state every run)
# -----------------------------
st.subheader("3) Downloads")
if not st.session_state.ppt_results:
    st.info("No PPTs generated yet.")
else:
    ts = st.session_state.ppt_generated_at.strftime("%Y-%m-%d %H:%M") if st.session_state.ppt_generated_at else "—"
    st.caption(f"Generated at: {ts}")

    # Individual downloads (stable keys so they don't clash)
    for idx, (fname, blob) in enumerate(st.session_state.ppt_results):
        st.download_button(
            label=f"⬇️ {fname}",
            data=blob,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key=f"dl_{idx}_{fname}",  # stable unique key
            use_container_width=True,
        )

    # Zip all
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, blob in st.session_state.ppt_results:
            zf.writestr(fname, blob)
    zip_buf.seek(0)

    st.download_button(
        label="📦 Download All (ZIP)",
        data=zip_buf,
        file_name=f"PPTs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
        mime="application/zip",
        key="dl_zip_all",
        use_container_width=True,
    )
