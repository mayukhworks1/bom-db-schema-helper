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

st.set_page_config(layout="wide", page_title="School SKU Mapping + Gap Analyzer")
st.title("School ↔ SKU Mapping + Expected vs Actual Verification")


# ─────────────────────────────────────────────
# Utility Functions
# ─────────────────────────────────────────────

def normalize(value):
    if value is None or pd.isna(value):
        return ""
    return str(value).strip().upper()

def safe_df(df):
    return df.fillna("").astype(str)

def to_int(series):
    return pd.to_numeric(series, errors="coerce").fillna(0).astype(int)

def download_csv(df, name):
    buffer = BytesIO()
    df.to_csv(buffer, index=False)
    st.download_button(
        f"📥 Download {name}",
        buffer.getvalue(),
        f"{name}.csv",
        "text/csv"
    )


# ─────────────────────────────────────────────
# 1️⃣ LOAD SCHOOLS (ACTUAL DATA)
# ─────────────────────────────────────────────

schools_file = st.file_uploader("1️⃣ Upload schools_export.csv", type="csv")
if schools_file is None:
    st.stop()

schools = pd.read_csv(schools_file, dtype=str)
schools = safe_df(schools)
schools.columns = schools.columns.str.strip()

column_map = {}
for col in schools.columns:
    l = col.lower()
    if "entity" in l and "id" in l:
        column_map[col] = "entity_id"
    elif l == "name":
        column_map[col] = "school_name"
    elif "school code" in l:
        column_map[col] = "school_code"
    elif "associated" in l:
        column_map[col] = "actual_items"

schools.rename(columns=column_map, inplace=True)

required_cols = ["entity_id", "school_name", "school_code"]
for col in required_cols:
    if col not in schools.columns:
        st.error(f"Missing column: {col}")
        st.stop()

schools["actual_items"] = to_int(schools.get("actual_items", 0))
schools["school_name_norm"] = schools["school_name"].apply(normalize)
schools["school_code_norm"] = schools["school_code"].apply(normalize)

st.success(f"Schools Loaded: {len(schools)} rows")


# ─────────────────────────────────────────────
# 2️⃣ LOAD MASTER EXCEL (EXPECTED DATA)
# ─────────────────────────────────────────────

excel_file = st.file_uploader("2️⃣ Upload Master Final For Ecom.xlsx", type="xlsx")
if excel_file is None:
    st.stop()

xl = pd.ExcelFile(excel_file)
sheet = st.selectbox("Select Sheet", xl.sheet_names)

raw = pd.read_excel(excel_file, sheet_name=sheet, header=None)

header_row = None
for i in range(min(30, len(raw))):
    if raw.iloc[i].astype(str).str.upper().str.contains("ITEM CODE").any():
        header_row = i
        break

if header_row is None:
    st.error("Header row with 'ITEM CODE' not found.")
    st.stop()

data = pd.read_excel(excel_file, sheet_name=sheet, header=header_row)
data = safe_df(data)
data.columns = data.columns.str.strip()

try:
    start_index = list(data.columns).index("MULTIPLE ITEM CODES") + 1
except:
    start_index = 15

school_columns = list(data.columns[start_index:])

mapping_rows = []

for col in school_columns:
    school_norm = normalize(col)
    if school_norm == "":
        continue

    sku_series = data[col].str.strip().str.upper()
    sku_series = sku_series[sku_series != ""].unique()

    for sku in sku_series:
        mapping_rows.append({
            "school_norm": school_norm,
            "sku": sku
        })

mapping = pd.DataFrame(mapping_rows)

if mapping.empty:
    st.error("No school SKU mapping found.")
    st.stop()

expected_counts = (
    mapping.groupby("school_norm")["sku"]
    .nunique()
    .reset_index(name="expected_items")
)

st.success("Master Excel Processed Successfully")


# ─────────────────────────────────────────────
# 3️⃣ LOAD ITEMS (FOR MAPPING OUTPUT)
# ─────────────────────────────────────────────

items_file = st.file_uploader("3️⃣ Upload Items CSV", type="csv")
if items_file is None:
    st.stop()

items = pd.read_csv(items_file, dtype=str)
items = safe_df(items)
items.columns = items.columns.str.strip()

item_map = {}
for col in items.columns:
    l = col.lower()
    if "id" in l and "zoho" not in l:
        item_map[col] = "item_id"
    elif "item code" in l:
        item_map[col] = "item_code"
    elif "sku" in l:
        item_map[col] = "sku"

items.rename(columns=item_map, inplace=True)


# ─────────────────────────────────────────────
# 4️⃣ MATCHING LOGIC (IDENTICAL TO ORIGINAL)
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

    elif FUZZY_AVAILABLE:
        best_match = process.extractOne(
            name_norm,
            school_groups,
            scorer=fuzz.token_sort_ratio
        )
        if best_match and best_match[1] >= 80:
            candidate, score = best_match
            expected_items = int(
                expected_counts.loc[
                    expected_counts["school_norm"] == candidate,
                    "expected_items"
                ].iloc[0]
            )
            matched_group = candidate
            match_score = score
            match_type = "fuzzy"

    else:
        possible = [g for g in school_groups if name_norm in g]
        if possible:
            candidate = possible[0]
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

result_df = pd.DataFrame(results).sort_values("difference")


# ─────────────────────────────────────────────
# 5️⃣ GENERATED MAPPING (SEPARATE SECTION)
# ─────────────────────────────────────────────

school_match_df = result_df[["entity_id","school_name","matched_group"]]

school_sku_join = school_match_df.merge(
    mapping,
    left_on="matched_group",
    right_on="school_norm",
    how="inner"
)

final_mapping = school_sku_join.merge(
    items,
    on="sku",
    how="left"
)[["entity_id","school_name","sku","item_code","item_id"]].drop_duplicates()

st.subheader("Generated Mapping (School ID → Item)")
st.dataframe(final_mapping, use_container_width=True, height=600)
download_csv(final_mapping, "school_item_mapping")

# ─────────────────────────────────────────────
# 🔍 VERIFICATION: GENERATED MAPPING vs MASTER
# ─────────────────────────────────────────────

st.subheader("Mapping Verification Summary")

# Normalize items SKU once for accurate checks
items["sku"] = items["sku"].str.strip().str.upper()

# Detect duplicate SKUs in items file
duplicate_item_skus = set(
    items[items["sku"].duplicated()]["sku"]
)

verification_rows = []

for _, row in result_df.iterrows():

    entity_id = row["entity_id"]
    school_name = row["school_name"]
    matched_group = row["matched_group"]

    if matched_group == "":
        verification_rows.append({
            "entity_id": entity_id,
            "school_name": school_name,
            "expected_skus": 0,
            "generated_skus": 0,
            "missing_skus": 0,
            "extra_skus": 0,
            "unresolved_items": 0,
            "unresolved_skus": "",
            "unresolved_reason": "School not matched to Master",
            "status": "NO MATCH"
        })
        continue

    # Expected SKUs from Master
    expected_set = set(
        mapping.loc[
            mapping["school_norm"] == matched_group,
            "sku"
        ]
    )

    # Generated mapping for school
    generated_df = final_mapping.loc[
        final_mapping["entity_id"] == entity_id
    ]

    generated_set = set(generated_df["sku"])

    missing = expected_set - generated_set
    extra = generated_set - expected_set

    unresolved_skus_list = []
    reason_set = set()

    for _, sku_row in generated_df.iterrows():

        sku = sku_row["sku"]
        item_id = sku_row["item_id"]
        item_code = sku_row["item_code"]

        if item_id == "" or item_code == "":

            unresolved_skus_list.append(sku)

            # Diagnose reason
            if sku not in set(items["sku"]):
                reason_set.add("SKU missing in Items file")

            elif sku in duplicate_item_skus:
                reason_set.add("Duplicate SKU in Items file")

            elif item_id == "":
                reason_set.add("Blank item_id in Items")

            elif item_code == "":
                reason_set.add("Blank item_code in Items")

            else:
                reason_set.add("Unknown resolution issue")

    unresolved_skus_list = sorted(set(unresolved_skus_list))
    unresolved_skus_str = ", ".join(unresolved_skus_list)

    if len(unresolved_skus_list) == 0:
        unresolved_reason = ""
    else:
        unresolved_reason = "; ".join(sorted(reason_set))

    # Determine status
    if len(missing) == 0 and len(extra) == 0 and len(unresolved_skus_list) == 0:
        status = "VERIFIED"
    elif len(unresolved_skus_list) > 0:
        status = "ITEM RESOLUTION ISSUE"
    elif len(missing) > 0:
        status = "MISSING SKUs"
    elif len(extra) > 0:
        status = "EXTRA SKUs"
    else:
        status = "CHECK"

    verification_rows.append({
        "entity_id": entity_id,
        "school_name": school_name,
        "expected_skus": len(expected_set),
        "generated_skus": len(generated_set),
        "missing_skus": len(missing),
        "extra_skus": len(extra),
        "unresolved_items": len(unresolved_skus_list),
        "unresolved_skus": unresolved_skus_str,
        "unresolved_reason": unresolved_reason,
        "status": status
    })

verification_df = pd.DataFrame(verification_rows)

st.dataframe(verification_df, use_container_width=True)
download_csv(verification_df, "mapping_verification_summary")

# ─────────────────────────────────────────────
# 6️⃣ EXPECTED VS ACTUAL (ORIGINAL LOGIC)
# ─────────────────────────────────────────────

st.subheader("Expected vs Actual Comparison")
st.dataframe(result_df, use_container_width=True, height=650)
download_csv(result_df, "expected_vs_actual")


# ─────────────────────────────────────────────
# 7️⃣ SIGNIFICANT GAPS
# ─────────────────────────────────────────────

gap_df = result_df[result_df["difference"].abs() > 5]

if not gap_df.empty:
    st.subheader("Significant Gaps (>|5|)")
    st.dataframe(gap_df, use_container_width=True)
    download_csv(gap_df, "significant_gaps")
else:
    st.success("No significant gaps found.")

st.caption("Stable • Original Gap Logic Preserved • Mapping Verified Against Master")
