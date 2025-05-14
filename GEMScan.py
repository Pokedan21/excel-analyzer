import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import io

st.set_page_config(page_title="Excel Analyzer", layout="wide")
st.title("📊 Excel Analyzer: Filter, Explore, and Visualize")

uploaded_file = st.file_uploader("📥 Upload an Excel file", type=["xlsx"])

@st.cache_data
def dedup_columns(df):
    df.columns = pd.io.parsers.ParserBase({'names': df.columns})._maybe_dedup_names(df.columns)
    df.columns = df.columns.map(str)
    return df

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
    return dedup_columns(df)

@st.cache_data
def load_full_sheet(upload, sheet_name):
    df = pd.read_excel(upload, sheet_name=sheet_name, engine='openpyxl')
    return dedup_columns(df)

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
    else:
        df = sample_df.copy()

    # Dynamically find relevant columns
    cap_col = next((c for c in df.columns if c.startswith("Capacity (MW)")), None)
    year_col = next((c for c in df.columns if c.startswith("Start year")), None)

    if cap_col: df[cap_col] = pd.to_numeric(df[cap_col], errors="coerce")
    if year_col: df[year_col] = pd.to_numeric(df[year_col], errors="coerce")

    # Rename type-like column
    if "Plant Type" in df.columns:
        df.rename(columns={"Plant Type": "Type"}, inplace=True)
    elif "Technology" in df.columns:
        df.rename(columns={"Technology": "Type"}, inplace=True)

    st.sidebar.header("🔎 Filter Options")
    logic_mode = st.sidebar.radio("Combine all filters using:", ["AND", "OR"], index=0)

    if st.sidebar.button("🔄 Reset All Filters"):
        for key in list(st.session_state.keys()):
            if key.endswith("_select_all") or key.endswith("_multiselect") or key.endswith("_logic"):
                del st.session_state[key]
        st.experimental_rerun()

    filter_conditions = []

    if cap_col:
        min_cap, max_cap = float(df[cap_col].min()), float(df[cap_col].max())
        cap_range = st.sidebar.slider("Capacity (MW)", min_value=min_cap, max_value=max_cap, value=(min_cap, max_cap))
        filter_conditions.append((df[cap_col] >= cap_range[0]) & (df[cap_col] <= cap_range[1]))

    if year_col:
        year_data = df[year_col].dropna()
        if not year_data.empty:
            min_year, max_year = int(year_data.min()), int(year_data.max())
            selected_years = st.sidebar.slider("Start year", min_year, max_year, (min_year, max_year))
            filter_conditions.append((df[year_col] >= selected_years[0]) & (df[year_col] <= selected_years[1]))

    if "GEM unit/phase ID" in df.columns:
        if st.sidebar.checkbox("✅ Only include GEM units", value=False):
            filter_conditions.append(df["GEM unit/phase ID"].notna())

    for column in df.columns:
        if column in [cap_col, year_col, "GEM unit/phase ID"]:
            continue
        col_data = df[column].dropna()
        if pd.api.types.is_string_dtype(col_data):
            with st.sidebar.expander(f"🔤 Filter: {column}", expanded=False):
                logic_mode_col = st.radio("Match Mode", ["Match ANY (OR)", "Match ALL (AND)"], key=f"{column}_logic")
                options = sorted(col_data.unique().tolist())
                selected = st.multiselect(f"{column}", options, key=f"{column}_multiselect")
                if selected:
                    if logic_mode_col.startswith("Match ANY"):
                        filter_conditions.append(df[column].isin(selected))
                    else:
                        filter_conditions.append(df[column].apply(lambda x: all(opt in str(x) for opt in selected)))

    if filter_conditions:
        combined = filter_conditions[0]
        for cond in filter_conditions[1:]:
            combined = combined & cond if logic_mode == "AND" else combined | cond
        filtered_df = df[combined]
    else:
        filtered_df = df.copy()

    if filtered_df.empty:
        st.warning("⚠️ No data matches your filters.")
        st.stop()

    sort_column = st.selectbox("🗂️ Sort by column:", df.columns)
    sort_order = st.radio("Sort Order", ["Ascending", "Descending"])
    filtered_df = filtered_df.sort_values(by=sort_column, ascending=(sort_order == "Ascending"))

    st.write(f"✅ Showing {len(filtered_df)} matching record(s):")
    st.dataframe(filtered_df.head(100))

    st.download_button(
        label="📥 Download filtered table as Excel",
        data=to_excel_bytes(filtered_df),
        file_name="filtered_table.xlsx",
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    # === Permanent Cumulative Chart ===
    if cap_col and year_col and "Type" in df.columns:
        st.subheader("📈 Cumulative 10–300 MW Plants Built Since 2015 (by Type)")
        valid = df[
            (df[cap_col] >= 10) & (df[cap_col] <= 300) &
            (df[year_col] >= 2015) & df["Type"].notna()
        ].copy()
        if valid.empty:
            st.warning("⚠️ No matching plants found.")
        else:
            annual = valid.groupby([year_col, "Type"]).size().unstack(fill_value=0).sort_index()
            cumulative = annual.cumsum()
            st.dataframe(cumulative)

            fig, ax = plt.subplots(figsize=(10, 6))
            for col in cumulative.columns:
                ax.plot(cumulative.index, cumulative[col], label=col)
            ax.set_title("Cumulative 10–300 MW Plants Since 2015 by Type")
            ax.set_xlabel("Start Year")
            ax.set_ylabel("Cumulative Count")
            ax.legend()
            ax.grid(True)
            st.pyplot(fig)

            img_buf = io.BytesIO()
            fig.savefig(img_buf, format='png')
            st.download_button("📥 Download chart as PNG", data=img_buf.getvalue(), file_name="cumulative_chart.png", mime="image/png")
            st.download_button("📥 Download table as Excel", data=to_excel_bytes(cumulative), file_name="cumulative_table.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("⚠️ Missing 'Capacity (MW)', 'Start year' or 'Type' column.")

    else:
        st.info("👆 Please upload an Excel file to begin.")

        st.warning("⚠️ Cannot show cumulative chart — required columns missing: 'Start year', 'Capacity (MW)', or a type column like 'Plant Type'.")

else:
    st.info("👆 Please upload an Excel file to begin.")
