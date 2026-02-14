import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide", page_title="School ↔ SKU Mapper 2025")
st.title("School ↔ SKU Mapping + Entity ID Enrichment")

# ────────────────────────────────────────────────
# Helpers
# ────────────────────────────────────────────────

def norm(s):
    if pd.isna(s):
        return ""
    return str(s).strip().upper().replace("  ", " ")


def to_str(df):
    for col in df.columns:
        df[col] = df[col].astype(str).replace(['nan','NaN','None','<NA>'], '')
    return df


def safe_name(n):
    n = str(n)
    n = re.sub(r'[\\/*?:[\]]', '_', n).strip()[:31].strip('_ ')
    return n or "Sheet"


# ────────────────────────────────────────────────
# 1. School master CSV
# ────────────────────────────────────────────────

school_uploader = st.file_uploader("1. Upload schools_export.csv", type="csv")

school_df = pd.DataFrame()

if school_uploader:
    try:
        school_df = pd.read_csv(school_uploader, dtype=str, keep_default_na=False)
        school_df.columns = school_df.columns.str.strip()

        # Try to find important columns (case insensitive)
        col_lower = school_df.columns.str.lower()

        entity_col   = next((c for c in school_df.columns if "entity" in c.lower() or "id" in c.lower()), "entity_id")
        name_col     = next((c for c in school_df.columns if "name" in c.lower()), "school_name")
        code_col     = next((c for c in school_df.columns if "code" in c.lower() or "party" in c.lower()), "school_code")
        brand_col    = next((c for c in school_df.columns if "brand" in c.lower()), None)
        zone_col     = next((c for c in school_df.columns if "zone" in c.lower()), None)

        # Rename to standard names
        rename = {}
        if entity_col != "entity_id": rename[entity_col] = "entity_id"
        if name_col   != "school_name": rename[name_col]   = "school_name"
        if code_col   != "school_code": rename[code_col]   = "school_code"
        if brand_col and brand_col != "brand": rename[brand_col] = "brand"
        if zone_col  and zone_col  != "zone":  rename[zone_col]  = "zone"

        school_df.rename(columns=rename, inplace=True)

        # Create normalized versions
        if "school_code" in school_df.columns:
            school_df["code_norm"] = school_df["school_code"].apply(norm)
        if "school_name" in school_df.columns:
            school_df["name_norm"] = school_df["school_name"].apply(norm)

        school_df = to_str(school_df)

        st.success(f"School master loaded — {len(school_df)} rows")

        with st.expander("School master columns & sample"):
            st.write("Columns:", list(school_df.columns))
            st.dataframe(school_df.head(6))

    except Exception as e:
        st.error(f"Failed to read schools CSV\n{e}")


# ────────────────────────────────────────────────
# 2. Item master Excel
# ────────────────────────────────────────────────

excel_uploader = st.file_uploader("2. Upload Master Final For Ecom.xlsx", type="xlsx")

if excel_uploader:
    try:
        xl = pd.ExcelFile(excel_uploader)

        sheet = st.selectbox(
            "Select sheet",
            xl.sheet_names,
            index=next((i for i,s in enumerate(xl.sheet_names) if "item" in s.lower() or "master" in s.lower() or "report" in s.lower()), 0)
        )

        # Read raw → find header
        raw = pd.read_excel(excel_uploader, sheet_name=sheet, header=None)
        header_idx = None
        for r in range(15):
            if any("ITEM CODE" in str(x).upper() for x in raw.iloc[r]):
                header_idx = r
                break

        if header_idx is None:
            st.error("Cannot find row with 'ITEM CODE' in first 15 rows")
            st.stop()

        st.success(f"Header found at row {header_idx}")

        # Read real data
        df = pd.read_excel(excel_uploader, sheet_name=sheet, header=header_idx)
        df.columns = df.columns.astype(str).str.strip()
        df = to_str(df)

        # Basic item master
        items = pd.DataFrame({
            "item_code": df.get("ITEM CODE", "").apply(norm),
            "sku":       df.get("ADDL ITEM CODE", "").apply(norm),
            "name":      df.get("ITEM NAME", "").str.strip()
        }).query("sku != ''").drop_duplicates("sku").reset_index(drop=True)

        st.subheader("Item master")
        c1, c2 = st.columns(2)
        c1.metric("Rows", f"{len(df):,}")
        c2.metric("Unique SKUs", len(items))

        st.dataframe(items.head(10), use_container_width=True)

        # ──────────────────────────────
        # Find school columns
        # ──────────────────────────────

        end_col = 15
        cols_lower = [c.lower() for c in df.columns]
        if "multiple item codes" in cols_lower:
            end_col = cols_lower.index("multiple item codes") + 1

        school_cols = df.columns[end_col:]

        if len(school_cols) == 0:
            st.warning("No school columns detected after master data")
        else:
            st.write(f"Found {len(school_cols)} potential school columns")

            # Build mapping
            mapping = []
            for col in school_cols:
                col_str = str(col).strip()
                if not col_str:
                    continue

                skus = df[col].dropna().astype(str).str.strip().str.upper()
                valid = skus[(skus != "") & (skus != "NAN")]

                for sku in valid:
                    mapping.append({
                        "school_raw": col_str,
                        "school_norm": norm(col_str),
                        "sku": sku
                    })

            map_df = pd.DataFrame(mapping).drop_duplicates()

            if map_df.empty:
                st.warning("No valid SKUs found in school columns")
            else:
                st.subheader("Raw school → SKU mapping")
                st.dataframe(map_df.head(20), use_container_width=True)

                # ──────────────────────────────
                # ENRICHMENT
                # ──────────────────────────────

                enriched = map_df.copy()

                if not school_df.empty:
                    st.write("Enriching...")

                    # 1. Try exact code match
                    if "code_norm" in school_df.columns:
                        enriched = enriched.merge(
                            school_df[["entity_id", "school_code", "brand", "zone", "code_norm"]],
                            left_on="school_raw",
                            right_on="code_norm",
                            how="left"
                        ).drop(columns=["code_norm"], errors="ignore")

                    # 2. If still many missing → try name match (simple contains)
                    missing = enriched["entity_id"].isna()
                    if missing.any() and "name_norm" in school_df.columns:
                        for idx in enriched[missing].index:
                            s_norm = enriched.at[idx, "school_norm"]
                            candidates = school_df[school_df["name_norm"].str.contains(s_norm, na=False)]
                            if len(candidates) == 1:
                                row = candidates.iloc[0]
                                enriched.at[idx, "entity_id"]   = row.get("entity_id", "")
                                enriched.at[idx, "school_code"] = row.get("school_code", "")
                                enriched.at[idx, "brand"]       = row.get("brand", "")
                                enriched.at[idx, "zone"]        = row.get("zone", "")

                    enriched = to_str(enriched)

                st.subheader("Enriched mapping")
                st.dataframe(enriched.head(25), use_container_width=True)

                # Counts
                groupby_cols = [c for c in ["entity_id", "school_code", "school_norm", "brand", "zone"] if c in enriched.columns]
                if groupby_cols:
                    counts = enriched.groupby(groupby_cols, as_index=False, dropna=False)\
                                     .agg(sku_count=("sku", "nunique"))\
                                     .sort_values("sku_count", ascending=False)
                    st.subheader("SKUs per school / entity")
                    st.dataframe(counts, use_container_width=True)

    except Exception as e:
        st.error(f"Error during processing:\n{str(e)}")
        import traceback
        st.code(traceback.format_exc(), language="python")

st.caption("Clean & safe version • All string • Simple name fallback • Feb 2025")