# pptgen_module.py

# pptgen_module.py

import io
import os
import re
from datetime import datetime
from typing import List, Optional, Iterable, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# ----- Fixed column names (expected EXACTLY as below) -----
COL_COUNTRY    = "Country"
COL_MODEL      = "Bike_Model"
COL_QUARTER    = "Quarter"       # values like "Q1-2025", "2025Q1", "Q2 2024"
COL_UNITS      = "Sales Units"
COL_REVENUE    = "Revenue_INR"   # numeric (optional)

# ----- Quarter parsing helpers -----
_QUARTER_PATTERNS = [
    re.compile(r"^\s*Q([1-4])[-\s]?(\d{4})\s*$", re.IGNORECASE),   # Q1-2025 or Q1 2025
    re.compile(r"^\s*(\d{4})[-\s]?Q([1-4])\s*$", re.IGNORECASE),   # 2025-Q1 or 2025 Q1
    re.compile(r"^\s*([1-4])\s*[Qq][-]?\s*(\d{4})\s*$"),           # 1Q-2025 / 1q 2025
]

def _quarter_to_period(s: str) -> Optional[pd.Period]:
    """Parse various quarter strings to a pandas Period (Q-DEC)."""
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return None
    text = str(s).strip()
    for pat in _QUARTER_PATTERNS:
        m = pat.match(text)
        if m:
            if pat.pattern.startswith("^\\s*Q"):
                q = int(m.group(1)); year = int(m.group(2))
            elif pat.pattern.startswith("^\\s*(\\d{4})"):
                year = int(m.group(1)); q = int(m.group(2))
            else:
                q = int(m.group(1)); year = int(m.group(2))
            try:
                return pd.Period(f"{year}Q{q}", freq="Q")
            except Exception:
                return None
    # Accept already-like '2025Q1'
    try:
        return pd.Period(str(text).replace(" ", ""), freq="Q")
    except Exception:
        return None

def _period_to_label(p: pd.Period) -> str:
    """Return canonical display like '2025Q1'."""
    return f"{p.year}Q{p.quarter}"


class MonthlyPerformancePPT:
    """
    Generate per-country PowerPoints from *quarterly* performance data using fixed column names:
      Country | Bike_Model | Quarter | Sales Units | Revenue_INR

    All outputs are returned IN-MEMORY:
      - generate_from_dataframe(...) -> List[Tuple[filename:str, bytes]]
      - generate_from_file(...)      -> List[Tuple[filename:str, bytes]]
    """

    def __init__(
        self,
        template_ppt: Optional[str] = None,
        logo_bytes: Optional[bytes] = None,   # optional logo image bytes
        last_n_quarters: int = 6,
    ):
        self.template_ppt = template_ppt
        self.logo_bytes = logo_bytes
        self.last_n_quarters = max(1, int(last_n_quarters))

    # ---------- Public API (in-memory outputs) ----------

    def generate_from_file(
        self,
        input_path: str,
        sheet_name: Optional[str] = None,
    ) -> List[Tuple[str, bytes]]:
        df = self._read_table(input_path, sheet_name)
        work = self._prepare_dataframe(df)
        return self._generate_all(work)

    def generate_from_dataframe(self, df: pd.DataFrame) -> List[Tuple[str, bytes]]:
        work = self._prepare_dataframe(df)
        return self._generate_all(work)

    # ---------- Core flow ----------

    def _generate_all(self, work: pd.DataFrame) -> List[Tuple[str, bytes]]:
        if work.empty:
            raise ValueError("No rows after preprocessing. Check your input file and required columns.")

        latest_q = work["QuarterPeriod"].max()  # pandas Period
        start_q = latest_q - (self.last_n_quarters - 1)

        outputs: List[Tuple[str, bytes]] = []
        for country, df_c in work.groupby("Country"):
            fname, blob = self._generate_for_country(
                country=country,
                df_country=df_c.copy(),
                latest_q=latest_q,
                start_q=start_q,
            )
            outputs.append((fname, blob))
        return outputs

    def _generate_for_country(
        self,
        country: str,
        df_country: pd.DataFrame,
        latest_q: pd.Period,
        start_q: pd.Period,
    ) -> Tuple[str, bytes]:
        prs = Presentation(self.template_ppt) if (self.template_ppt and os.path.exists(self.template_ppt)) else Presentation()

        # ---- Title slide ----
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        if slide.shapes.title:
            slide.shapes.title.text = f"{country} — Quarterly Performance"
        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = (
                f"Reporting through: {_period_to_label(latest_q)}\n"
                f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            )
        if self.logo_bytes:
            self._add_picture_from_bytes(slide, self.logo_bytes, left=8.5, top=0.2, width=1.5)

        # ---- KPI slide ----
        self._add_kpi_slide(prs, df_country)

        # ---- Country trend (last N quarters) ----
        self._add_trend_slide(prs, df_country, country, start_q, latest_q)

        # ---- Model mix ----
        top_models = self._add_model_mix_slide(prs, df_country)

        # ---- Top-3 model trends ----
        self._add_top_models_slides(prs, df_country, country, top_models, start_q, latest_q)

        # ---- Save deck to bytes ----
        out_name = f"{country}_Quarterly_Performance_{_period_to_label(latest_q)}.pptx"
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        return out_name, buf.getvalue()

    # ---------- Slides ----------

    def _add_kpi_slide(self, prs: Presentation, df_c: pd.DataFrame) -> None:
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title + content
        slide.shapes.title.text = "Overview KPIs"

        content = slide.shapes.placeholders[1].text_frame
        content.clear()
        content.word_wrap = True

        latest_q = df_c["QuarterPeriod"].max()
        last_df = df_c[df_c["QuarterPeriod"] == latest_q]

        last_units = float(last_df["Sales_Units"].sum())
        if last_df["Revenue"].notna().any():
            last_rev = float(last_df["Revenue"].sum())
        else:
            last_rev = np.nan

        avg_price = self._safe_div(last_rev, last_units) if np.isfinite(last_rev) else float(last_df["Unit_Price"].mean())

        prev_q = latest_q - 1
        prev_units = float(df_c.loc[df_c["QuarterPeriod"] == prev_q, "Sales_Units"].sum())
        qoq_growth = (self._safe_div((last_units - prev_units), prev_units) * 100.0) if prev_units else 0.0

        def add_kpi_bullet(label: str, value: str, unit: str = ""):
            p = content.add_paragraph()
            run = p.add_run()
            run.text = f"{label}: "
            run.font.size = Pt(18)

            run_val = p.add_run()
            run_val.text = f"{value}{unit}"
            run_val.font.bold = True
            run_val.font.size = Pt(18)
            run_val.font.color.rgb = RGBColor(0, 0, 0)

        add_kpi_bullet("Latest Quarter", _period_to_label(latest_q))
        add_kpi_bullet("Total Units", f"{last_units:,.0f}")
        add_kpi_bullet("Revenue (INR)", f"{last_rev:,.0f}" if np.isfinite(last_rev) else "—")
        add_kpi_bullet("Avg Unit Price (INR)", f"{avg_price:,.0f}" if np.isfinite(avg_price) else "—")
        add_kpi_bullet("QoQ Growth (Units)", f"{qoq_growth:.1f}", "%")

    def _add_trend_slide(
        self,
        prs: Presentation,
        df_c: pd.DataFrame,
        country: str,
        start_q: pd.Period,
        latest_q: pd.Period,
    ) -> None:
        trend = (
            df_c[(df_c["QuarterPeriod"] >= start_q) & (df_c["QuarterPeriod"] <= latest_q)]
            .groupby("QuarterPeriod", as_index=False)["Sales_Units"].sum()
            .sort_values("QuarterPeriod")
        )
        if trend.empty:
            return

        x = trend["QuarterPeriod"].apply(_period_to_label)
        y = trend["Sales_Units"]

        fig = plt.figure(figsize=(8, 4))
        plt.plot(x, y)
        plt.title(f"Quarterly Sales Trend — {country} (Last {self.last_n_quarters} quarters)")
        plt.xlabel("Quarter")
        plt.ylabel("Sales Units")
        img_bytes = self._fig_bytes(fig)
        plt.close(fig)

        slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title-only
        slide.shapes.title.text = "Sales Trend"
        self._add_picture_from_bytes(slide, img_bytes, left=0.7, top=1.5, width=8.5)

    def _add_model_mix_slide(self, prs: Presentation, df_c: pd.DataFrame) -> List[str]:
        model_mix_df = (
            df_c.groupby("Model", as_index=False)["Sales_Units"]
            .sum()
            .sort_values("Sales_Units", ascending=False)
        )
        if model_mix_df.empty:
            return []

        fig, ax = plt.subplots(figsize=(6, 4))
        bars = ax.bar(model_mix_df["Model"], model_mix_df["Sales_Units"])
        for bar in bars:
            height = bar.get_height()
            ax.annotate(
                f"{height:,.0f}",
                xy=(bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 3),
                textcoords="offset points",
                ha="center",
                va="bottom",
                fontsize=10,
                fontweight="bold",
            )
        ax.set_title("Model Mix – Units", fontsize=14, fontweight="bold")
        ax.set_ylabel("Units")
        plt.xticks(rotation=30, ha="right")

        img_bytes = self._fig_bytes(plt.gcf())
        plt.close()

        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "Model Mix"
        self._add_picture_from_bytes(slide, img_bytes, left=1.0, top=1.5, width=8.0)

        top_models = model_mix_df.head(3)["Model"].tolist()
        return top_models

    def _add_top_models_slides(
        self,
        prs: Presentation,
        df_c: pd.DataFrame,
        country: str,
        top_models: Iterable[str],
        start_q: pd.Period,
        latest_q: pd.Period,
    ) -> None:
        for m in top_models:
            d = (
                df_c[
                    (df_c["Model"] == m)
                    & (df_c["QuarterPeriod"] >= start_q)
                    & (df_c["QuarterPeriod"] <= latest_q)
                ]
                .groupby("QuarterPeriod", as_index=False)["Sales_Units"].sum()
                .sort_values("QuarterPeriod")
            )
            if d.empty:
                continue

            fig = plt.figure(figsize=(8, 4))
            plt.plot(d["QuarterPeriod"].apply(_period_to_label), d["Sales_Units"])
            plt.title(f"Model Trend — {m} ({country})")
            plt.xlabel("Quarter")
            plt.ylabel("Sales Units")
            img_bytes = self._fig_bytes(fig)
            plt.close(fig)

            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = f"Model Trend — {m}"
            self._add_picture_from_bytes(slide, img_bytes, left=0.7, top=1.5, width=8.5)

    # ---------- Data prep ----------

    def _read_table(self, path: str, sheet: Optional[str]) -> pd.DataFrame:
        ext = os.path.splitext(path)[1].lower()
        if ext in (".xlsx", ".xls"):
            df = pd.read_excel(path, sheet_name=sheet)
        else:
            df = pd.read_csv(path)
        df.columns = [str(c) for c in df.columns]  # ensure strings
        return df

    def _prepare_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        # Validate required columns exist exactly
        required = [COL_COUNTRY, COL_MODEL, COL_QUARTER, COL_UNITS]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise KeyError(f"Missing required columns: {missing}. Expected {required} (and optional '{COL_REVENUE}').")

        work = df.copy()

        # Parse Quarter -> Period + canonical text label
        qp = work[COL_QUARTER].apply(_quarter_to_period)
        work = work.loc[qp.notna()].copy()
        work["QuarterPeriod"] = qp[qp.notna()].astype("period[Q]")
        work["Quarter"] = work["QuarterPeriod"].apply(_period_to_label)

        # Canonical columns used downstream
        work = work.rename(columns={
            COL_COUNTRY: "Country",
            COL_MODEL:   "Model",
        })

        work["Sales_Units"] = pd.to_numeric(work[COL_UNITS], errors="coerce").fillna(0)

        if COL_REVENUE in work.columns:
            work["Revenue"] = pd.to_numeric(work[COL_REVENUE], errors="coerce")
        else:
            work["Revenue"] = np.nan

        # Derive Unit_Price if Revenue present
        work["Unit_Price"] = np.nan
        mask = work["Revenue"].notna() & (work["Sales_Units"] > 0)
        work.loc[mask, "Unit_Price"] = work.loc[mask, "Revenue"] / work.loc[mask, "Sales_Units"]

        return work

    # ---------- Utils (in-memory) ----------

    @staticmethod
    def _safe_div(a, b) -> float:
        try:
            if b is None or (isinstance(b, float) and np.isnan(b)) or b == 0:
                return np.nan
            return float(a) / float(b)
        except Exception:
            return np.nan

    @staticmethod
    def _fig_bytes(fig) -> bytes:
        """Return a PNG image as bytes from a matplotlib figure."""
        buf = io.BytesIO()
        fig.tight_layout()
        fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
        buf.seek(0)
        return buf.getvalue()

    @staticmethod
    def _add_picture_from_bytes(slide, image_bytes: bytes, left: float, top: float, width: float) -> None:
        """Add an image from bytes to a slide at the given position."""
        stream = io.BytesIO(image_bytes)
        slide.shapes.add_picture(stream, Inches(left), Inches(top), width=Inches(width))
