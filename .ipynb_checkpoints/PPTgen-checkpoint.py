#streamlit run pptgen.py --server.port 8502

import os
import io
import time
import glob
import platform
import subprocess
from datetime import datetime

import pandas as pd
import streamlit as st

# -----------------------------
# Page setup
# -----------------------------
st.set_page_config(page_title="Excel → PPT Viewer", page_icon="📑", layout="wide")
st.title("▶️ Generate PPTs from excel files")

# -----------------------------
# Sidebar: uploads + output folder
# -----------------------------
with st.sidebar:
    st.header("1) Upload Excel/CSV")
    uploads = st.file_uploader(
        "Upload one or more files",
        type=["xlsx", "csv"],
        accept_multiple_files=True
    )

    st.markdown("---")
    st.header("2) Output PPT Folder")
    output_dir = st.text_input(
        "Where should I save PPTX files?",
        value=os.path.abspath("ppt_output"),  # default: ./ppt_output
    )
    # if st.button("🗂️ Ensure folder exists"):
    #     os.makedirs(output_dir, exist_ok=True)
    #     st.success(f"Ensured folder: {output_dir}")

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

            # Rebuild buffer for parsing (ExcelFile may have advanced the pointer)
            bio2 = io.BytesIO(data)
            xl2 = pd.ExcelFile(bio2, engine="openpyxl")
            df = xl2.parse(sheet_to_use)
            return df, sheets, True, None
        except Exception:
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
        sys = platform.system()
        if sys == "Darwin":
            subprocess.Popen(["open", path])
        elif sys == "Windows":
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

if not uploads:
    st.info("Upload one or more Excel/CSV files from the sidebar.")
else:
    for i, up in enumerate(uploads):
        # First parse to discover whether it's Excel and its sheet list
        df0, sheets0, is_excel0, err0 = parse_uploaded_file(up)
        if err0:
            st.error(f"**{up.name}** — {err0}")
            continue

        st.markdown(f"**File:** {up.name}")
        if is_excel0 and sheets0:
            # Per-file sheet picker (unique key)
            chosen = st.selectbox(
                f"Worksheet for {up.name}",
                sheets0,
                index=0,
                key=f"sheet_{up.name}_{i}",
            )
            df, sheets, is_excel, err = parse_uploaded_file(up, chosen_sheet=chosen)
        else:
            # CSV or non-Excel — no sheet picker
            df, sheets, is_excel, err = df0, sheets0, is_excel0, err0

        if err:
            st.error(f"{err}")
            continue

        if df is None or df.empty:
            st.info("No rows to display.")
        else:
            st.dataframe(df.head(50), use_container_width=True, height=320)

        st.divider()

# -----------------------------
# 2) Generate button (spinner only)
# -----------------------------
st.subheader("2) Generate PPTs")
st.caption("Using the excel files uploaded, I am generating the export performance PPTs by country..please wait!")

if st.button("🚀 Generate PPTs"):
    with st.spinner("Generating Country-wise PowerPoints…"):
        time.sleep(3)  # simulate work
    st.success("Done! Listing PPTs below from your output folder.")

    st.markdown("---")

    # -----------------------------
    # 3) PPT folder listing
    # -----------------------------
    st.subheader("3) PPT Files in Output Folder")
    ppt_files = list_pptx(output_dir)
    
    if not ppt_files:
        st.info("No PPTX files found. Point to your folder in the sidebar or ensure it exists.")
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

st.caption("Tip: If your script writes to a different folder, just change the path in the sidebar.")
