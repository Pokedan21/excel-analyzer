import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import io

st.set_page_config(page_title="Excel Analyzer", layout="wide")
st.title("📊 Excel Analyzer: Filter, Explore, and Visualize")

uploaded_file = st.file_uploader("📥 Upload an Excel file", type=["xlsx"])

@st.cache_data
def get_valid_sheet_names(upload):
    try:
        excel_file = pd.ExcelFile(upload, engine='openpyxl')
        return [s for s in excel_file.sheet_names if s.strip().lower() != "zinformation"]
    except Exception as e:
        st.error(f"❌ Unable to read Excel file: {e}")
        return []

@st.cache_data
def load_sample_sheet(upload, sheet_name, sample_size=500):
    df = pd.read_excel(upload, sheet_name=sheet_name, engine='openpyxl', nrows=sample_size)
    df.columns = df.columns.map(str)
    return df

@st.cache_data
def load_full_sheet(upload, sheet_name):
    df = pd.read_excel(upload, sheet_name=sheet_name, engine='openpyxl')
    df.columns = df.columns.map(str)
    return df

def try_convert_dates(df):
    for col in df.columns:
        if 'date' in col.lower():
            try:
                df[col] = pd.to_datetime(df[col])
            except:
                pass
    return df

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=True)
    return output.getvalue()

if uploaded_file:
    sheet_names = get_valid_sheet_names(uploaded_file)
    selected_sheet = st.selectbox("📄 Select sheet to load:", sheet_names)

    with st.spinner("🔄 Loading sample data..."):
        sample_df = try_convert_dates(load_sample_sheet(uploaded_file, selected_sheet))
        sample_df["Start year"] = pd.to_numeric(sample_df.get("Start year", pd.Series(dtype="float")), errors="coerce")
    st.success("✅ Sample loaded.")

    st.markdown("### 🧾 Sample Columns in File")
    col_df = pd.DataFrame({
        "Column Name": sample_df.columns,
        "Data Type": [str(sample_df[col].dtype) for col in sample_df.columns]
    })
    st.dataframe(col_df)

    if st.checkbox("🔍 Load full data from this sheet (may take time)"):
        with st.spinner("📂 Loading full dataset... please wait."):
            df = try_convert_dates(load_full_sheet(uploaded_file, selected_sheet))
            df["Start year"] = pd.to_numeric(df.get("Start year", pd.Series(dtype="float")), errors="coerce")
        st.success("✅ Full data loaded successfully.")
    else:
        df = sample_df.copy()

    st.sidebar.header("🔎 Filter Options")
    logic_mode = st.sidebar.radio("Combine all filters using:", ["AND", "OR"], index=0)

    if st.sidebar.button("🔄 Reset All Filters"):
        for key in list(st.session_state.keys()):
            if key.endswith("_select_all") or key.endswith("_multiselect") or key.endswith("_logic"):
                del st.session_state[key]
        st.experimental_rerun()

    filter_conditions = []

    # Numeric filter: Capacity
    if "Capacity (MW)" in df.columns:
        min_cap, max_cap = float(df["Capacity (MW)"].min()), float(df["Capacity (MW)"].max())
        min_cap_input = st.sidebar.number_input("Min Capacity (MW)", value=min_cap)
        max_cap_input = st.sidebar.number_input("Max Capacity (MW)", value=max_cap)
        cap_range = (min_cap_input, max_cap_input)
        condition = (df["Capacity (MW)"] >= cap_range[0]) & (df["Capacity (MW)"] <= cap_range[1])
        filter_conditions.append(condition)

    # Numeric filter: Start year
    if "Start year" in df.columns:
        year_col = df["Start year"].dropna()
        if not year_col.empty:
            min_year, max_year = int(year_col.min()), int(year_col.max())
            min_year_input = st.sidebar.number_input("Min Start Year", value=min_year, step=1)
            max_year_input = st.sidebar.number_input("Max Start Year", value=max_year, step=1)
            selected_years = (min_year_input, max_year_input)
            condition = (df["Start year"] >= selected_years[0]) & (df["Start year"] <= selected_years[1])
            filter_conditions.append(condition)

    if "GEM unit/phase ID" in df.columns:
        include_gem_only = st.sidebar.checkbox("✅ Only include GEM units", value=False)
        if include_gem_only:
            condition = df["GEM unit/phase ID"].notna()
            filter_conditions.append(condition)

    st.sidebar.markdown("### 🔍 Column Search and Unique Value Filtering")
    show_unique_only = st.sidebar.checkbox("Show only unique values per column", value=False)

    for column in df.columns:
        if column in ["Capacity (MW)", "Start year", "GEM unit/phase ID"]:
            continue

        col_data = df[column].dropna().astype(str)
        if show_unique_only:
            col_data = col_data[col_data.map(col_data.value_counts()) == 1]

        if pd.api.types.is_string_dtype(col_data):
            with st.sidebar.expander(f"🔤 Filter: {column}", expanded=False):
                search_term = st.text_input(f"Search {column}", key=f"{column}_search")
                if search_term:
                    col_data = col_data[col_data.str.contains(search_term, case=False, na=False)]

                logic_key = f"{column}_logic"
                logic_mode_col = st.radio("Match Mode", ["Match ANY (OR)", "Match ALL (AND)"], key=logic_key)
                unique_values = sorted(col_data.unique().tolist())
                selected_options = st.multiselect(f"{column}", unique_values, key=f"{column}_multiselect")
                if selected_options:
                    if logic_mode_col.startswith("Match ANY"):
                        condition = df[column].astype(str).isin(selected_options)
                    else:
                        condition = df[column].astype(str).apply(lambda x: all(opt in str(x) for opt in selected_options))
                    filter_conditions.append(condition)