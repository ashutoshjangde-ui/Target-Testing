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
# Mapping Library
# =========================
DEFAULT_MAPPING_LIBRARY = {
    "Walmart": {
        "Partner SKU": ["item_sku", "Seller SKU", "SKU", "Item SKU"],
        "Barcode": ["UPC/EAN", "barcode.value", "Product ID", "UPC", "EAN", "GTIN"],
        "Brand": ["Walmart Brand Name - en-US", "brand_name", "Brand Name", "Brand"],
        "Product Title": ["Item Name", "Product Name", "Title", "product_name"],
        "Description": ["Long Description", "Product Description", "Description", "short_description"],
        "Category": ["Product Category", "Category", "product_category"],
        "Price": ["Price", "Unit Price", "price", "unit_price"],
        "Quantity": ["Quantity", "Stock", "Available Quantity", "qty"],
        "Weight": ["Weight", "Item Weight", "Shipping Weight", "weight"],
        "Length": ["Length", "Product Length", "length"],
        "Width": ["Width", "Product Width", "width"],
        "Height": ["Height", "Product Height", "height"],
    },
    "Target": {
        "Partner SKU": ["Target SKU", "TCIN", "SKU", "Item Number"],
        "Barcode": ["UPC", "UPC/EAN", "Product ID", "GTIN"],
        "Brand": ["Brand Name", "Brand", "Manufacturer"],
        "Product Title": ["Item Name", "Product Name", "Title", "Description"],
        "Description": ["Long Description", "Product Description", "Item Description"],
        "Category": ["Category", "Product Category", "Department"],
        "Price": ["Price", "Retail Price", "Unit Price"],
        "Quantity": ["Quantity", "Stock Quantity", "Available"],
        "Weight": ["Weight", "Shipping Weight", "Item Weight"],
    },
    "Amazon": {
        "Partner SKU": ["seller-sku", "ASIN", "SKU", "Seller SKU"],
        "Barcode": ["UPC", "EAN", "product-id", "GTIN"],
        "Brand": ["brand_name", "Brand", "brand"],
        "Product Title": ["item_name", "Product Name", "Title", "product-name"],
        "Description": ["product_description", "Description", "item-description"],
        "Category": ["product_category", "Category", "item_type"],
        "Price": ["standard_price", "Price", "price"],
        "Quantity": ["quantity", "Stock", "inventory"],
        "Weight": ["item_weight", "Weight", "package-weight"],
    },
    "Shopify": {
        "Partner SKU": ["Variant SKU", "SKU", "sku"],
        "Barcode": ["Variant Barcode", "Barcode", "barcode"],
        "Brand": ["Vendor", "Brand", "vendor"],
        "Product Title": ["Title", "Product Title", "title"],
        "Description": ["Body (HTML)", "Description", "body_html"],
        "Category": ["Product Category", "Type", "product_type"],
        "Price": ["Variant Price", "Price", "price"],
        "Quantity": ["Variant Inventory Qty", "Quantity", "inventory_quantity"],
        "Weight": ["Variant Weight", "Weight", "weight"],
    },
    "Generic Retail": {
        "Partner SKU": ["SKU", "Item SKU", "Product SKU", "Seller SKU"],
        "Barcode": ["UPC", "EAN", "Barcode", "GTIN", "Product ID"],
        "Brand": ["Brand", "Brand Name", "Manufacturer"],
        "Product Title": ["Title", "Product Name", "Item Name", "Name"],
        "Description": ["Description", "Product Description", "Long Description"],
        "Category": ["Category", "Product Category", "Type"],
        "Price": ["Price", "Unit Price", "Retail Price"],
        "Quantity": ["Quantity", "Stock", "Inventory", "QTY"],
        "Weight": ["Weight", "Shipping Weight", "Item Weight"],
        "Length": ["Length", "Product Length"],
        "Width": ["Width", "Product Width"],
        "Height": ["Height", "Product Height"],
    }
}

# Initialize session state for mapping library
if 'mapping_library' not in st.session_state:
    st.session_state.mapping_library = DEFAULT_MAPPING_LIBRARY.copy()

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

def similarity_score(s1: str, s2: str) -> float:
    """Calculate similarity between two strings."""
    return SequenceMatcher(None, norm(s1), norm(s2)).ratio()

def top_matches(query, candidates, k=3):
    """Get top k matches for a query string from candidates."""
    q = norm(query)
    scored = [(SequenceMatcher(None, q, norm(c)).ratio(), c) for c in candidates]
    scored.sort(key=lambda t: t[0], reverse=True)
    return scored[:k]

def auto_map_columns(master_headers, onboarding_headers, threshold=0.6):
    """
    Automatically map master columns to onboarding columns using fuzzy matching.
    Returns: dict of {master_col: onboarding_col}
    """
    mapping = {}
    for master_col in master_headers:
        if not master_col or not str(master_col).strip():
            continue
        
        best_match = None
        best_score = 0
        
        for on_col in onboarding_headers:
            if not on_col or not str(on_col).strip():
                continue
            
            score = similarity_score(master_col, on_col)
            if score > best_score and score >= threshold:
                best_score = score
                best_match = on_col
        
        if best_match:
            mapping[master_col] = [best_match]
    
    return mapping

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
        for master_norm, aliases in mapping_aliases_by_master.items():
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

def show_interactive_mapper(master_headers, onboarding_headers, auto_suggestions=None):
    """Display interactive dropdowns for each master column."""
    st.markdown("#### 🔗 Interactive Column Mapping")
    st.caption("Map each master template column to an onboarding sheet column. Pre-filled suggestions are based on name similarity.")
    
    mapping = {}
    
    # Create a clean list of onboarding headers
    on_headers_clean = [h for h in onboarding_headers if h and str(h).strip()]
    
    # Table header
    col1, col2, col3 = st.columns([3, 1, 3])
    with col1:
        st.markdown("**📋 Master Template Column**")
    with col2:
        st.markdown("**→**")
    with col3:
        st.markdown("**📊 Onboarding Sheet Column**")
    
    st.markdown("---")
    
    for idx, master_col in enumerate(master_headers):
        if not master_col or not str(master_col).strip():
            continue
        
        col1, col2, col3 = st.columns([3, 1, 3])
        
        with col1:
            st.markdown(f"**{master_col}**")
        
        with col2:
            st.markdown("→")
        
        with col3:
            # Determine default selection
            default_idx = 0
            
            if auto_suggestions and master_col in auto_suggestions:
                suggested = auto_suggestions[master_col]
                if isinstance(suggested, list) and len(suggested) > 0:
                    suggested = suggested[0]
                try:
                    default_idx = on_headers_clean.index(suggested) + 1
                except (ValueError, AttributeError):
                    default_idx = 0
            
            selected = st.selectbox(
                f"map_{master_col}",
                ["(skip this column)"] + on_headers_clean,
                index=default_idx,
                label_visibility="collapsed",
                key=f"select_{idx}_{master_col}"
            )
            
            if selected != "(skip this column)":
                mapping[master_col] = [selected]
    
    return mapping

# Unique sentinel for the special "Listing Action" fill
SENTINEL_LISTING_ACTION = object()

# =========================
# UI
# =========================
st.title("📦 Masterfile Automation")
st.caption("Map onboarding columns to master template headers and generate a ready-to-upload masterfile.")

with st.expander("ℹ️ Instructions", expanded=False):
    instructions = dedent("""
    ### How to Use This Tool
    
    1. **Upload Files**
       - **Masterfile Template** (.xlsx): Row 1 = headers, Row 2 = keys (optional), Data starts Row 3
       - **Onboarding Sheet** (.xlsx): Row 1 = headers, Row 2+ = data
    
    2. **Choose Mapping Method**
       - 🤖 **Auto-detect**: Automatically matches columns by name similarity
       - 📚 **Use Template**: Select from pre-configured retailer mappings
       - 📝 **Manual Mapping**: Interactive dropdown selection for each column
       - 📄 **Upload JSON**: Use your own custom mapping file
    
    3. **Generate**: Click the button to create your final masterfile
    
    ### Mapping JSON Format Example
    ```json
    {
      "Partner SKU": ["Target SKU", "Seller SKU", "SKU"],
      "Barcode": ["UPC/EAN", "UPC", "Product ID"],
      "Brand": ["Brand Name", "brand_name"]
    }
    ```
    
    The tool will use the **first matching column** from the list for each master column.
    """)
    st.markdown(instructions)

st.divider()

# File uploads
colA, colB = st.columns([1, 1])
with colA:
    masterfile_file = st.file_uploader("📄 Upload Masterfile Template (.xlsx)", type=["xlsx"])
with colB:
    onboarding_file = st.file_uploader("🧾 Upload Onboarding Sheet (.xlsx)", type=["xlsx"])

# Initialize variables
master_headers = []
onboarding_headers = []
mapping_raw = None

# Preview files if uploaded
if masterfile_file and onboarding_file:
    with st.spinner("Reading files..."):
        try:
            # Read master headers
            masterfile_file.seek(0)
            master_wb_preview = load_workbook(masterfile_file, read_only=True)
            master_ws_preview = master_wb_preview.active
            used_cols = worksheet_used_cols(master_ws_preview, header_rows=(1, 2))
            master_headers = [
                master_ws_preview.cell(row=1, column=c).value or "" 
                for c in range(1, used_cols + 1)
            ]
            master_headers = [str(h).strip() for h in master_headers if h]
            master_wb_preview.close()
            
            # Read onboarding headers
            onboarding_file.seek(0)
            on_df_preview = pd.read_excel(onboarding_file, nrows=0)
            onboarding_headers = [str(c).strip() for c in on_df_preview.columns]
            
            st.success(f"✅ Found {len(master_headers)} master columns and {len(onboarding_headers)} onboarding columns")
            
        except Exception as e:
            st.error(f"Error reading files: {e}")
            st.stop()

st.divider()

# Mapping method selection
st.markdown("### 🔧 Choose Mapping Method")

mapping_method = st.radio(
    "How would you like to map columns?",
    ["🤖 Auto-detect (Recommended)", "📚 Use Retailer Template", "📝 Manual Mapping", "📄 Upload JSON File"],
    horizontal=True
)

# Handle different mapping methods
if mapping_method == "🤖 Auto-detect (Recommended)":
    st.info("The system will automatically match columns based on name similarity.")
    
    if master_headers and onboarding_headers:
        sensitivity = st.slider(
            "Matching Sensitivity",
            min_value=40,
            max_value=100,
            value=60,
            step=5,
            help="Higher values require closer name matches. Lower values are more lenient."
        )
        
        threshold = sensitivity / 100.0
        mapping_raw = auto_map_columns(master_headers, onboarding_headers, threshold)
        
        st.success(f"✅ Auto-detected {len(mapping_raw)} column mappings")
        
        with st.expander("🔍 Preview Auto-detected Mappings"):
            for master, onboarding_list in mapping_raw.items():
                st.write(f"**{master}** ← `{onboarding_list[0]}`")

elif mapping_method == "📚 Use Retailer Template":
    st.info("Select a pre-configured mapping template for common retailers.")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        template_name = st.selectbox(
            "Choose Retailer Template",
            list(st.session_state.mapping_library.keys())
        )
    
    with col2:
        st.markdown("")
        st.markdown("")
        if st.button("🔄 Refresh Templates"):
            st.rerun()
    
    mapping_raw = st.session_state.mapping_library[template_name].copy()
    
    with st.expander("📝 View/Edit Template Mapping"):
        edited_json = st.text_area(
            "Edit mapping if needed:",
            value=json.dumps(mapping_raw, indent=2),
            height=300
        )
        
        col_a, col_b = st.columns([1, 5])
        with col_a:
            if st.button("💾 Update Mapping"):
                try:
                    mapping_raw = json.loads(edited_json)
                    st.success("✅ Mapping updated!")
                except json.JSONDecodeError as e:
                    st.error(f"Invalid JSON: {e}")

elif mapping_method == "📝 Manual Mapping":
    if not master_headers or not onboarding_headers:
        st.warning("⚠️ Please upload both files first to use manual mapping.")
    else:
        # Get auto-suggestions to pre-fill the dropdowns
        auto_suggestions = auto_map_columns(master_headers, onboarding_headers, threshold=0.6)
        mapping_raw = show_interactive_mapper(master_headers, onboarding_headers, auto_suggestions)

elif mapping_method == "📄 Upload JSON File":
    st.info("Upload a custom JSON mapping file or paste JSON directly.")
    
    tab1, tab2 = st.tabs(["📎 Upload File", "📝 Paste JSON"])
    
    with tab1:
        mapping_json_file = st.file_uploader("Upload mapping.json", type=["json"])
        if mapping_json_file:
            try:
                mapping_raw = json.load(mapping_json_file)
                st.success("✅ JSON file loaded successfully!")
                with st.expander("Preview Mapping"):
                    st.json(mapping_raw)
            except Exception as e:
                st.error(f"Error parsing JSON file: {e}")
    
    with tab2:
        mapping_json_text = st.text_area(
            "Paste mapping JSON here:",
            height=300,
            placeholder='{\n  "Partner SKU": ["Seller SKU", "item_sku"],\n  "Barcode": ["UPC", "barcode"]\n}'
        )
        
        if mapping_json_text.strip():
            try:
                mapping_raw = json.loads(mapping_json_text)
                st.success("✅ JSON parsed successfully!")
            except json.JSONDecodeError as e:
                st.error(f"Invalid JSON: {e}")

st.divider()

# Generate button
col_left, col_center, col_right = st.columns([2, 2, 2])
with col_center:
    go = st.button("🚀 Generate Final Masterfile", type="primary", use_container_width=True)

log_area = st.container()
download_area = st.container()

# =========================
# Main Processing
# =========================
if go:
    with log_area:
        st.markdown("### 📝 Processing Log")
        
        # Validate inputs
        if not masterfile_file or not onboarding_file:
            st.error("❌ Please upload both **Masterfile Template** and **Onboarding Sheet**.")
            st.stop()
        
        if not mapping_raw:
            st.error("❌ Please configure column mapping using one of the methods above.")
            st.stop()
        
        # Normalize mapping keys and keep ordered aliases
        mapping_aliases_by_master = {}
        for k, v in mapping_raw.items():
            aliases = v[:] if isinstance(v, list) else [v]
            if k not in aliases:
                aliases = aliases + [k]
            mapping_aliases_by_master[norm(k)] = aliases
        
        st.info("⏳ Reading workbooks...")
        
        try:
            # Read masterfile with openpyxl to preserve styles
            masterfile_file.seek(0)
            master_wb = load_workbook(masterfile_file, keep_links=False)
            master_ws = master_wb.active
        except Exception as e:
            st.error(f"❌ Could not read **Masterfile**: {e}")
            st.stop()
        
        # Pick best onboarding sheet
        try:
            onboarding_file.seek(0)
            best_df, best_sheet, info = pick_best_onboarding_sheet(
                onboarding_file, 
                mapping_aliases_by_master
            )
            st.success(f"✅ Using onboarding sheet: **{best_sheet}** ({info})")
        except Exception as e:
            st.error(f"❌ Could not find a suitable onboarding sheet: {e}")
            st.stop()
        
        on_df = best_df
        on_headers = list(on_df.columns)
        
        # Build normalized lookup
        series_by_alias = {norm(h): on_df[h] for h in on_headers}
        
        # Master headers
        used_cols = worksheet_used_cols(master_ws, header_rows=(1, 2))
        master_displays = [
            master_ws.cell(row=1, column=c).value or "" 
            for c in range(1, used_cols + 1)
        ]
        
        # Build master -> onboarding series map
        master_to_source = {}
        chosen_alias = {}
        unmatched = []
        report_lines = []
        
        report_lines.append("#### 🔎 Mapping Summary")
        
        for c, m_disp in enumerate(master_displays, start=1):
            disp_norm = norm(m_disp)
            if not disp_norm:
                continue
            
            aliases = mapping_aliases_by_master.get(disp_norm, [m_disp])
            resolved_series = None
            resolved_alias = None
            
            for a in aliases:
                a_norm = norm(a)
                if a_norm in series_by_alias:
                    resolved_series = series_by_alias[a_norm]
                    resolved_alias = a
                    break
            
            if resolved_series is not None:
                master_to_source[c] = resolved_series
                chosen_alias[c] = resolved_alias
                report_lines.append(f"- ✅ **{m_disp}** ← `{resolved_alias}`")
            else:
                if disp_norm == norm("Listing Action (List or Unlist)"):
                    master_to_source[c] = SENTINEL_LISTING_ACTION
                    report_lines.append(f"- 🟨 **{m_disp}** ← (auto-filled with `'List'`)")
                else:
                    unmatched.append(m_disp)
                    suggestions = top_matches(m_disp, on_headers, 3)
                    sug_txt = ", ".join(
                        f"`{name}` ({round(sc*100,1)}%)" 
                        for sc, name in suggestions
                    ) if suggestions else "*none*"
                    report_lines.append(f"- ❌ **{m_disp}** ← _no match_. Suggestions: {sug_txt}")
        
        st.markdown("\n".join(report_lines))
        
        # Write values to master
        st.info("🛠️ Writing data to masterfile...")
        
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
                        # Keep as string to preserve formatting
                        master_ws.cell(row=out_row + i, column=c, value=str(val) if val else "")
        
        # Save to buffer
        st.info("💾 Saving final masterfile...")
        bio = io.BytesIO()
        master_wb.save(bio)
        bio.seek(0)
        
        with download_area:
            st.success(f"✅ **Final masterfile is ready!** ({num_rows} rows processed)")
            
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.download_button(
                    "⬇️ Download Final Masterfile",
                    data=bio.getvalue(),
                    file_name="final_masterfile.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                # Option to save mapping
                if st.button("💾 Save This Mapping", use_container_width=True):
                    st.session_state.show_save_dialog = True
            
            if st.session_state.get('show_save_dialog', False):
                with st.form("save_mapping_form"):
                    st.markdown("##### Save Mapping Template")
                    save_name = st.text_input(
                        "Template Name:",
                        placeholder="e.g., 'My Custom Walmart Mapping'"
                    )
                    
                    col_a, col_b = st.columns(2)
                    with col_a:
                        if st.form_submit_button("💾 Save", use_container_width=True):
                            if save_name.strip():
                                st.session_state.mapping_library[save_name.strip()] = mapping_raw
                                st.success(f"✅ Saved as '{save_name.strip()}'")
                                st.session_state.show_save_dialog = False
                                st.rerun()
                            else:
                                st.error("Please enter a template name")
                    
                    with col_b:
                        if st.form_submit_button("Cancel", use_container_width=True):
                            st.session_state.show_save_dialog = False
                            st.rerun()
            
            if unmatched:
                with st.expander("⚠️ Unmatched Columns (left blank)"):
                    st.warning(
                        "The following master columns had no match and were left blank:\n\n- " +
                        "\n- ".join(unmatched)
                    )
            
            # Show statistics
            st.markdown("---")
            stat_col1, stat_col2, stat_col3 = st.columns(3)
            with stat_col1:
                st.metric("✅ Mapped Columns", len(master_to_source))
            with stat_col2:
                st.metric("❌ Unmatched Columns", len(unmatched))
            with stat_col3:
                st.metric("📊 Rows Processed", num_rows)

# Footer
st.divider()
st.caption("💡 Tip: Save frequently used mappings as templates for faster processing in the future!")
