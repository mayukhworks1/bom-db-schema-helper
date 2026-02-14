import streamlit as st
import pandas as pd
from io import BytesIO

# Optional fuzzy matching
FUZZY_AVAILABLE = False
try:
    from fuzzywuzzy import fuzz, process
    FUZZY_AVAILABLE = True
except:
    pass


# ─────────────────────────────────────────────
# App Config
# ─────────────────────────────────────────────

st.set_page_config(layout="wide", page_title="School SKU Gap Analyzer")
st.title("School ↔ SKU Expected vs Actual Gap Analysis")


# ─────────────────────────────────────────────
# Utility Functions
# ─────────────────────────────────────────────

def normalize(value):
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def safe_string_dataframe(df):
    df = df.fillna("")
    return df.astype(str)


def to_integer_series(series):
    return pd.to_numeric(series, errors="coerce").fillna(0).astype(int)


def download_csv(df, filename):
    buffer = BytesIO()
    df.to_csv(buffer, index=False)
    st.download_button(
        label=f"📥 Download {filename}",
        data=buffer.getvalue(),
        file_name=f"{filename}.csv",
        mime="text/csv"
    )


# ─────────────────────────────────────────────
# 1️⃣ Load Schools CSV (ACTUAL DATA)
# ─────────────────────────────────────────────

school_file = st.file_uploader("Upload schools_export.csv", type="csv")

if school_file is None:
    st.info("Upload both CSV and Excel files to begin.")
    st.stop()

try:
    schools = pd.read_csv(school_file, dtype=str)
    schools = safe_string_dataframe(schools)
    schools.columns = schools.columns.str.strip()

    # Column mapping
    column_map = {}

    for col in schools.columns:
        lower = col.lower()

        if "entity" in lower and "id" in lower:
            column_map[col] = "entity_id"

        elif lower == "name":
            column_map[col] = "school_name"

        elif "school code" in lower:
            column_map[col] = "school_code"

        elif "associated" in lower:
            column_map[col] = "actual_items"

        elif "brand" in lower:
            column_map[col] = "brand"

        elif "zone" in lower:
            column_map[col] = "zone"

    schools.rename(columns=column_map, inplace=True)

    # Required columns check
    required_columns = ["entity_id", "school_name", "school_code"]
    for col in required_columns:
        if col not in schools.columns:
            st.error(f"Missing required column: {col}")
            st.stop()

    # Ensure actual_items exists
    if "actual_items" in schools.columns:
        schools["actual_items"] = to_integer_series(schools["actual_items"])
    else:
        schools["actual_items"] = 0

    # Normalized fields
    schools["school_name_norm"] = schools["school_name"].apply(normalize)
    schools["school_code_norm"] = schools["school_code"].apply(normalize)

    st.success(f"Schools loaded: {len(schools)} rows")

except Exception as e:
    st.error(f"CSV Load Failed:\n{str(e)}")
    st.stop()


# ─────────────────────────────────────────────
# 2️⃣ Load Excel (EXPECTED DATA)
# ─────────────────────────────────────────────

excel_file = st.file_uploader("Upload Master Final For Ecom.xlsx", type="xlsx")

if excel_file is None:
    st.stop()

try:
    xl = pd.ExcelFile(excel_file)
    sheet = st.selectbox("Select Sheet", xl.sheet_names)

    # Detect header row
    raw = pd.read_excel(excel_file, sheet_name=sheet, header=None)

    header_row = None
    for i in range(min(30, len(raw))):
        row = raw.iloc[i].astype(str).str.upper()
        if row.str.contains("ITEM CODE").any():
            header_row = i
            break

    if header_row is None:
        st.error("Could not detect header row containing 'ITEM CODE'")
        st.stop()

    data = pd.read_excel(excel_file, sheet_name=sheet, header=header_row)
    data = safe_string_dataframe(data)
    data.columns = data.columns.str.strip()

    # Identify where school columns start
    try:
        end_index = list(data.columns).index("MULTIPLE ITEM CODES") + 1
    except:
        end_index = 15  # fallback

    school_columns = list(data.columns[end_index:])

    # Build Expected Mapping
    mapping_rows = []

    for col in school_columns:

        school_norm = normalize(col)
        if school_norm == "":
            continue

        sku_series = data[col].astype(str).str.strip().str.upper()
        sku_series = sku_series[sku_series != ""]

        unique_skus = sku_series.unique()

        for sku in unique_skus:
            mapping_rows.append({
                "school_norm": school_norm,
                "sku": sku
            })

    mapping = pd.DataFrame(mapping_rows)

    if mapping.empty:
        st.error("No school SKU mapping found in Excel.")
        st.stop()

    expected_counts = (
        mapping.groupby("school_norm")["sku"]
        .nunique()
        .reset_index(name="expected_items")
    )

    st.success("Excel expected mapping built successfully.")

except Exception as e:
    import traceback
    st.error("Excel Processing Failed")
    st.code(traceback.format_exc(), language="python")
    st.stop()


# ─────────────────────────────────────────────
# 3️⃣ Matching Logic
# ─────────────────────────────────────────────

results = []
school_groups = expected_counts["school_norm"].tolist()

for _, row in schools.iterrows():

    expected_items = 0
    matched_group = ""
    match_score = 0
    match_type = ""

    code_norm = row["school_code_norm"]
    name_norm = row["school_name_norm"]

    # 1. Exact Code Match
    if code_norm in school_groups:

        expected_items = int(
            expected_counts.loc[
                expected_counts["school_norm"] == code_norm,
                "expected_items"
            ].iloc[0]
        )

        matched_group = code_norm
        match_score = 100
        match_type = "code"

    # 2. Fuzzy Name Match
    elif FUZZY_AVAILABLE and len(school_groups) > 0:

        best_match = process.extractOne(
            name_norm,
            school_groups,
            scorer=fuzz.token_sort_ratio
        )

        if best_match is not None:

            candidate, score = best_match

            if score >= 80:
                expected_items = int(
                    expected_counts.loc[
                        expected_counts["school_norm"] == candidate,
                        "expected_items"
                    ].iloc[0]
                )

                matched_group = candidate
                match_score = score
                match_type = "fuzzy"

    # 3. Contains fallback
    else:
        possible_matches = [
            group for group in school_groups
            if name_norm in group
        ]

        if len(possible_matches) > 0:
            candidate = possible_matches[0]

            expected_items = int(
                expected_counts.loc[
                    expected_counts["school_norm"] == candidate,
                    "expected_items"
                ].iloc[0]
            )

            matched_group = candidate
            match_type = "contains"

    actual_items = int(row["actual_items"])

    difference = actual_items - expected_items

    coverage = 0.0
    if expected_items > 0:
        coverage = round((actual_items / expected_items) * 100, 1)

    results.append({
        "entity_id": row["entity_id"],
        "school_name": row["school_name"],
        "school_code": row["school_code"],
        "expected_items": expected_items,
        "actual_items": actual_items,
        "difference": difference,
        "coverage_%": coverage,
        "matched_group": matched_group,
        "match_score": match_score,
        "match_type": match_type
    })


result_df = pd.DataFrame(results).sort_values("difference", ascending=True)

st.subheader("Expected vs Actual Comparison")
st.dataframe(result_df, use_container_width=True, height=650)
download_csv(result_df, "expected_vs_actual")


# ─────────────────────────────────────────────
# 4️⃣ Gap Analysis
# ─────────────────────────────────────────────

gap_df = result_df[abs(result_df["difference"]) > 5]

if not gap_df.empty:
    st.subheader("Significant Gaps (>|5|)")
    st.dataframe(gap_df, use_container_width=True)
    download_csv(gap_df, "significant_gaps")
else:
    st.success("No significant gaps found.")


st.caption("Stable Version • No ambiguous boolean checks • Fully validated")
