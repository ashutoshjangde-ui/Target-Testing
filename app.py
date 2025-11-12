# app_masterfile.py

import io
import json
import re
from difflib import SequenceMatcher
from textwrap import dedent

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Masterfile Automation", page_icon="📦", layout="wide")

# =========================
# Helpers
# =========================
def norm(s: str) -> str:
    """Normalize header strings for robust matching."""
    if s is None:
        return ""
    x = str(s).strip().lower()
    # strip "- en-us" & variants
    x = re.sub(r"\s*-\s*en\s*[-_ ]\s*us\s*$", "", x)
    # normalize dashes
    x = x.replace("–", "-").replace("—", "-").replace("−", "-")
    # replace common separators with a space (keep '-' last in class)
    x = re.sub(r"[._/\\-]+", " ", x)
    # drop anything not alnum or space
    x = re.sub(r"[^0-9a-z\s]+", " ", x)
    # collapse spaces
    return re.sub(r"\s+", " ", x).strip()


def top_matches(query, candidates, k=3):
    q = norm(query)
    scored = [(SequenceMatcher(None, q, norm(c)).ratio(), c) for c in candidates]
    scored.sort(key=lambda t: t[0], reverse=True)
    return scored[:k]


def worksheet_used_cols(ws, header_rows=(1,), hard_cap=512, empty_streak_stop=8):
    """Heuristically detect meaningful column span by scanning header rows."""
    max_try = min(ws.max_column, hard_cap)
    last_nonempty, streak = 0, 0
    for c in range(1, max_try + 1):
        any_val = False
        for r in header_rows:
            v = ws.cell(row=r, column=c).value
            if v not in (None, ""):
                any_val = True
                break
        if any_val:
            last_nonempty = c
            streak = 0
        else:
            streak += 1
            if streak >= empty_streak_stop:
                break
    return max(last_nonempty, 1)


def nonempty_rows(df: pd.DataFrame) -> int:
    """Count rows that have at least one non-empty cell."""
    if df.empty:
        return 0
    tmp = df.replace("", pd.NA)
    return tmp.dropna(how="all").shape[0]


def pick_best_onboarding_sheet(uploaded_file, mapping_aliases_by_master):
    """
    Inspect all sheets and pick the best one:
    - Row 1 is treated as headers
    - Row 2+ as data
    - Score = number of mapping keys that find at least one alias in the sheet headers
              + small tie-breaker for non-empty data rows
    Returns (best_df, best_sheet_name, debug_info)
    """
    uploaded_file.seek(0)
    xl = pd.ExcelFile(uploaded_file)

    best = None
    best_score = -1
    best_info = ""
    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet_name=sheet, header=0, dtype=str)
            df = df.fillna("")
            df.columns = [str(c).strip() for c in df.columns]
        except Exception:
            continue

        header_set = {norm(c) for c in df.columns}
        matches = 0
        for _, aliases in mapping_aliases_by_master.items():
            if any(norm(a) in header_set for a in aliases):
                matches += 1

        rows = nonempty_rows(df)
        score = matches + min(rows, 1) * 0.01

        if score > best_score:
            best = (df, sheet)
            best_score = score
            best_info = f"matched headers: {matches}, non-empty rows: {rows}"

    if best is None:
        raise ValueError("No readable sheet found in onboarding workbook.")

    df, sheet_name = best
    return df, sheet_name, best_info


# Unique sentinel for the special "Listing Action" fill
SENTINEL_LISTING_ACTION = object()

# =========================
# UI
# =========================
st.title("📦 Masterfile Automation")
st.caption("Map onboarding columns to master template headers and generate a ready-to-upload masterfile.")

with st.expander("ℹ️ Instructions", expanded=True):
    instructions = dedent("""
    - **Masterfile template (.xlsx)**  
      - Row **1** = display labels  
      - Row **2** = internal keys/helper labels  
      - Data is written starting at **Row 3** (this tool preserves template styles).

    - **Onboarding sheet (.xlsx)**  
      - Row **1** = **headers**  
      - Row **2+** = data

    - **Mapping JSON**: keys are the **master display headers** (Row 1 of master).  
      Values are **lists of onboarding header aliases** in priority order (the tool will use **all found** aliases and merge them row-wise).

    **Example**
    ```json
    {
      "Partner SKU": ["Target SKU","Seller SKU","SKU","item_sku"],
      "Barcode": ["UPC/EAN","UPC","Product ID","barcode","barcode.value"],
      "Brand": ["Brand Name","brand_name","Walmart Brand Name - en-US"],
      "Product Title": ["Item Name","Product Name","Title"],
      "Description": ["Long Description","Product Description","Description"]
    }
    ```
    """)
    st.markdown(instructions)

st.divider()

colA, colB = st.columns([1, 1])
with colA:
    masterfile_file = st.file_uploader("📄 Upload Masterfile Template (.xlsx)", type=["xlsx"])
with colB:
    onboarding_file = st.file_uploader("🧾 Upload Onboarding Sheet (.xlsx)", type=["xlsx"])

st.markdown("#### 🔗 Mapping JSON")
mapping_tab1, mapping_tab2 = st.tabs(["Paste JSON", "Upload JSON file"])
mapping_json_text = ""
mapping_json_file = None
with mapping_tab1:
    mapping_json_text = st.text_area(
        "Paste mapping JSON here",
        height=220,
        placeholder='{\n  "Partner SKU": ["Seller SKU", "item_sku"]\n}',
    )
with mapping_tab2:
    mapping_json_file = st.file_uploader("Or upload mapping.json", type=["json"], key="mapping_file")

st.divider()
go = st.button("🚀 Generate Final Masterfile", type="primary")

log_area = st.container()
download_area = st.container()

# =========================
# Main Action
# =========================
if go:
    with log_area:
        st.markdown("### 📝 Log")
        log = st.empty()

        def slog(msg):
            log.markdown(msg)

        # Validate inputs
        if not masterfile_file or not onboarding_file:
            st.error("Please upload both **Masterfile Template** and **Onboarding Sheet**.")
            st.stop()

        # Parse mapping JSON
        mapping_raw = None
        if mapping_json_text.strip():
            try:
                mapping_raw = json.loads(mapping_json_text)
            except Exception as e:
                st.error(f"Mapping JSON could not be parsed. Error: {e}")
                st.stop()
        elif mapping_json_file is not None:
            try:
                mapping_raw = json.load(mapping_json_file)
            except Exception as e:
                st.error(f"Mapping JSON file could not be parsed. Error: {e}")
                st.stop()
        else:
            st.error("Please provide mapping JSON (paste or upload).")
            st.stop()

        # Normalize mapping keys and keep ordered aliases
        mapping_aliases_by_master = {}
        for k, v in mapping_raw.items():
            aliases = v[:] if isinstance(v, list) else [v]
            if k not in aliases:
                aliases = aliases + [k]
            mapping_aliases_by_master[norm(k)] = aliases

        slog("⏳ Reading workbooks…")
        try:
            master_wb = load_workbook(masterfile_file, keep_links=False)
            master_ws = master_wb.active
        except Exception as e:
            st.error(f"Could not read **Masterfile**: {e}")
            st.stop()

        # Pick best onboarding sheet
        try:
            best_df, best_sheet, info = pick_best_onboarding_sheet(onboarding_file, mapping_aliases_by_master)
            st.success(f"Using onboarding sheet: **{best_sheet}** ({info})")
        except Exception as e:
            st.error(f"Could not find a suitable onboarding sheet: {e}")
            st.stop()

        on_df = best_df
        on_headers = list(on_df.columns)
        series_by_alias = {norm(h): on_df[h] for h in on_headers}

        # Master headers (Row 1 display, Row 2 keys)
        used_cols = worksheet_used_cols(master_ws, header_rows=(1, 2))
        master_displays = [master_ws.cell(row=1, column=c).value or "" for c in range(1, used_cols + 1)]

        # Build master -> onboarding merged series map
        master_to_source = {}   # col -> (Series) or SENTINEL_LISTING_ACTION
        chosen_alias = {}       # col -> comma-joined aliases actually used
        unmatched = []
        report_lines = []
        report_lines.append("#### 🔎 Mapping Summary")

        for c, m_disp in enumerate(master_displays, start=1):
            disp_norm = norm(m_disp)
            if not disp_norm:
                continue

            aliases = mapping_aliases_by_master.get(disp_norm, [m_disp])

            # Collect all present synonym columns
            resolved_series_list = []
            resolved_aliases = []
            for a in aliases:
                a_norm = norm(a)
                if a_norm in series_by_alias:
                    resolved_series_list.append(series_by_alias[a_norm].astype(str))
                    resolved_aliases.append(a)

            if resolved_series_list:
                # Merge: first non-empty wins, left-to-right by alias priority
                combined = pd.Series(dtype=object)
                for s in resolved_series_list:
                    combined = combined.combine_first(s.replace({"nan": "", "None": ""}))
                master_to_source[c] = combined
                chosen_alias[c] = ", ".join(resolved_aliases)
                report_lines.append(f"- ✅ **{m_disp}** ← `{', '.join(resolved_aliases)}`")
            else:
                if disp_norm == norm("Listing Action (List or Unlist)"):
                    master_to_source[c] = SENTINEL_LISTING_ACTION
                    report_lines.append(f"- 🟨 **{m_disp}** ← (will fill `'List'`)")
                else:
                    unmatched.append(m_disp)
                    suggestions = top_matches(m_disp, on_headers, 3)
                    sug_txt = ", ".join(
                        f"`{name}` ({round(sc*100,1)}%)" for sc, name in suggestions
                    ) if suggestions else "*none*"
                    report_lines.append(f"- ❌ **{m_disp}** ← *no match*. Suggestions: {sug_txt}")

        st.markdown("\n".join(report_lines))

        # Write values to master starting row 3
        slog("🛠️ Writing data…")
        out_row = 3
        num_rows = len(on_df)

        for i in range(num_rows):
            for c in range(1, used_cols + 1):
                src = master_to_source.get(c, None)
                if src is None:
                    continue
                if src is SENTINEL_LISTING_ACTION:
                    master_ws.cell(row=out_row + i, column=c, value="List")
                elif isinstance(src, pd.Series):
                    if i < len(src):
                        val = src.iloc[i]
                        # Clean typical empties
                        if pd.isna(val):
                            val = ""
                        s = str(val).strip()
                        if s.lower() in ("nan", "none"):
                            s = ""
                        # Write only non-empty to avoid literal "nan"
                        if s != "":
                            master_ws.cell(row=out_row + i, column=c, value=s)

        # Save to buffer
        slog("💾 Saving…")
        bio = io.BytesIO()
        master_wb.save(bio)
        bio.seek(0)

        with download_area:
            st.success("✅ Final masterfile is ready!")
            st.download_button(
                "⬇️ Download Final Masterfile",
                data=bio.getvalue(),
                file_name="final_masterfile_real.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            if unmatched:
                st.info(
                    "Some master columns had no match and were left blank:\n\n- " +
                    "\n- ".join(unmatched)
                )
