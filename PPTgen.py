#streamlit run pptgen.py --server.port 8502

# app.py

import os, io, glob, platform, subprocess
from datetime import datetime

import pandas as pd
import streamlit as st

from exportPPT import MonthlyPerformancePPT  # quarter-based generator


# -----------------------------
# Page setup
# -----------------------------
st.set_page_config(page_title="Excel → PPT Viewer", page_icon="📑", layout="wide")
st.title("▶️ Generate PPTs from Excel/CSV files (Quarter-based)")

# -----------------------------
# Sidebar: uploads + output folder
# -----------------------------
with st.sidebar:
    st.header("1) Upload Excel/CSV")
    uploads = st.file_uploader(
        "Upload one or more files",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True
    )

    st.markdown("---")
    st.header("2) Output PPT Folder")
    output_dir = st.text_input(
        "Where should I save PPTX files?",
        value=os.path.abspath("ppt_output"),
    )
    if st.button("🗂️ Ensure folder exists"):
        os.makedirs(output_dir, exist_ok=True)
        st.success(f"Ensured folder: {output_dir}")

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

            # Rebuild buffer and parse specific sheet
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

def list_pptx(folder: str) -> list[str]:
    if not os.path.isdir(folder):
        return []
    files = glob.glob(os.path.join(folder, "*.pptx"))
    files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return files

def open_locally(path: str):
    try:
        sysname = platform.system()
        if sysname == "Darwin":
            subprocess.Popen(["open", path])
        elif sysname == "Windows":
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", path])
        st.toast(f"Attempted to open {os.path.basename(path)} locally.")
    except Exception as e:
        st.error(f"Could not open locally: {e}")

# -----------------------------
# 1) Upload previews (right pane)
# -----------------------------
st.subheader("1) Uploaded Files Preview")

if "sheet_choices" not in st.session_state:
    st.session_state.sheet_choices = {}

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
# 2) Map Columns (Quarter-based)
# -----------------------------
st.subheader("2) Map Columns")
st.caption("Pick which columns correspond to Country, Model, Quarter (e.g., 'Q1-2025'), and Units (Revenue/Price optional).")

# Build candidate columns from all uploads
all_cols = set()
if uploads:
    for up in uploads:
        df0, _, _, err0 = parse_uploaded_file(up)
        if not err0 and df0 is not None:
            all_cols.update(df0.columns.tolist())
all_cols = sorted(all_cols)

def pick(label, required=True, key=None, default=None):
    opts = ["—"] + all_cols if not required else all_cols
    index = (opts.index(default) if (default and default in opts) else 0)
    return st.selectbox(label, opts, index=index, key=key)

# Defaults if present
default_country = next((c for c in all_cols if c.lower() == "country"), None)
default_model   = next((c for c in all_cols if "model" in c.lower()), None)
default_quarter = next((c for c in all_cols if "quarter" in c.lower() or c.lower() in ("qtr","qrtr","period")), None)
default_units   = next((c for c in all_cols if c.lower() in ("sales units","sales_units","units","sales","qty","quantity")), None)
default_revenue = next((c for c in all_cols if "revenue" in c.lower() or c.lower() in ("amount","salesvalue")), None)
default_price   = next((c for c in all_cols if c.lower() in ("unit price","unit_price","price","asp","avg price","average price")), None)

col_country = pick("Country *", required=True,  key="map_country", default=default_country)
col_model   = pick("Model *",   required=True,  key="map_model",   default=default_model)
col_quarter = pick("Quarter * (format like Q1-2025)", required=True, key="map_quarter", default=default_quarter)
col_units   = pick("Units *",   required=True,  key="map_units",   default=default_units)
col_revenue = pick("Revenue (optional)", required=False, key="map_revenue", default=default_revenue)
col_price   = pick("Unit Price (optional)", required=False, key="map_price", default=default_price)

column_map = {
    "country": col_country,
    "model":   col_model,
    "quarter": col_quarter,
    "units":   col_units,
    "revenue": (None if col_revenue in (None, "—") else col_revenue),
    "price":   (None if col_price   in (None, "—") else col_price),
}
missing_required = [k for k in ("country","model","quarter","units") if not column_map[k] or column_map[k] == "—"]

# -----------------------------
# 3) Generate button
# -----------------------------
st.subheader("3) Generate PPTs")
st.caption("Using the uploaded files and your column mapping, we’ll generate per‑country PowerPoints.")

if st.button("🚀 Generate PPTs"):
    if not uploads:
        st.error("Please upload at least one file.")
    elif missing_required:
        st.error(f"Please map required columns: {', '.join(missing_required)}")
    else:
        with st.spinner("Generating Country-wise PowerPoints…"):
            # Combine data
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

                os.makedirs(output_dir, exist_ok=True)
                gen = MonthlyPerformancePPT(
                    output_dir=output_dir,
                    template_ppt="BAL.pptx",   # or None
                    logo_path=None,            # optionally set a logo
                    last_n_quarters=6
                )
                try:
                    out_paths = gen.generate_from_dataframe(combined, column_map=column_map)
                    st.success(f"✅ Generated {len(out_paths)} PPTX file(s).")
                except Exception as e:
                    st.error(f"Generation failed: {e}")
                    out_paths = []

        st.markdown("---")

        # -----------------------------
        # 4) PPT folder listing
        # -----------------------------
        st.subheader("4) PPT Files in Output Folder")
        ppt_files = list_pptx(output_dir)
        if not ppt_files:
            st.info("No PPTX files found. Check your output folder path in the sidebar.")
        else:
            for p in ppt_files:
                fname = os.path.basename(p)
                mtime = datetime.fromtimestamp(os.path.getmtime(p)).strftime("%Y-%m-%d %H:%M")
                size_kb = os.path.getsize(p) // 1024
                with st.container(border=True):
                    st.markdown(f"**{fname}**  \n*Modified:* {mtime} • *Size:* {size_kb} KB")
                    c1, c2 = st.columns([1, 1], vertical_alignment="center")
                    with c1:
                        with open(p, "rb") as f:
                            st.download_button(
                                label="⬇️ Download PPTX",
                                data=f.read(),
                                file_name=fname,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True,
                            )
                    with c2:
                        if st.button(f"🖥️ Open locally", key=f"open_{fname}"):
                            open_locally(p)

st.caption("Tip: Quarter must look like Q1-2025 (case-insensitive). Other variants like '2025Q1' also work.")
