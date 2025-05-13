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

    if "Capacity (MW)" in df.columns:
        min_cap, max_cap = float(df["Capacity (MW)"].min()), float(df["Capacity (MW)"].max())
        cap_range = st.sidebar.slider("Capacity (MW)", min_value=min_cap, max_value=max_cap, value=(min_cap, max_cap))
        condition = (df["Capacity (MW)"] >= cap_range[0]) & (df["Capacity (MW)"] <= cap_range[1])
        filter_conditions.append(condition)

    if "Start year" in df.columns:
        year_col = df["Start year"].dropna()
        if not year_col.empty:
            min_year, max_year = int(year_col.min()), int(year_col.max())
            selected_years = st.sidebar.slider("Start year", min_year, max_year, (min_year, max_year))
            condition = (df["Start year"] >= selected_years[0]) & (df["Start year"] <= selected_years[1])
            filter_conditions.append(condition)

    if "GEM unit/phase ID" in df.columns:
        include_gem_only = st.sidebar.checkbox("✅ Only include GEM units", value=False)
        if include_gem_only:
            condition = df["GEM unit/phase ID"].notna()
            filter_conditions.append(condition)

    for column in df.columns:
        if column in ["Capacity (MW)", "Start year", "GEM unit/phase ID"]:
            continue

        col_data = df[column].dropna()
        if pd.api.types.is_string_dtype(col_data):
            with st.sidebar.expander(f"🔤 Filter: {column}", expanded=False):
                logic_key = f"{column}_logic"
                logic_mode_col = st.radio("Match Mode", ["Match ANY (OR)", "Match ALL (AND)"], key=logic_key)
                unique_values = sorted(col_data.unique().tolist())
                selected_options = st.multiselect(f"{column}", unique_values, key=f"{column}_multiselect")
                if selected_options:
                    if logic_mode_col.startswith("Match ANY"):
                        condition = df[column].isin(selected_options)
                    else:
                        condition = df[column].apply(lambda x: all(opt in str(x) for opt in selected_options))
                    filter_conditions.append(condition)

    if filter_conditions:
        combined_filter = filter_conditions[0]
        for cond in filter_conditions[1:]:
            combined_filter = combined_filter & cond if logic_mode == "AND" else combined_filter | cond
        filtered_df = df[combined_filter]
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

    # ===== PERMANENT CUMULATIVE CHART BY TYPE =====

    # Rename columns if necessary
    if "Plant Type" in df.columns:
        df.rename(columns={"Plant Type": "Type"}, inplace=True)
    elif "Technology" in df.columns:
        df.rename(columns={"Technology": "Type"}, inplace=True)

    if all(col in df.columns for col in ["Start year", "Capacity (MW)", "Type"]):
        st.subheader("📈 Cumulative 10–300 MW Plants Built Since 2015 (by Type)")

        valid_plants = df[
            (pd.to_numeric(df["Capacity (MW)"], errors="coerce") >= 10) &
            (pd.to_numeric(df["Capacity (MW)"], errors="coerce") <= 300) &
            (pd.to_numeric(df["Start year"], errors="coerce") >= 2015) &
            df["Type"].notna()
        ].copy()

        valid_plants["Start year"] = pd.to_numeric(valid_plants["Start year"], errors="coerce")
        valid_plants["Type"] = valid_plants["Type"].astype(str)

        if valid_plants.empty:
            st.warning("⚠️ No plants match the criteria for 10–300 MW and Start Year ≥ 2015.")
        else:
            annual = valid_plants.groupby(["Start year", "Type"]).size().unstack(fill_value=0).sort_index()
            cumulative = annual.cumsum()

            st.subheader("📊 Cumulative Table")
            st.dataframe(cumulative)

            fig, ax = plt.subplots(figsize=(10, 6))
            for col in cumulative.columns:
                ax.plot(cumulative.index, cumulative[col], label=col)
            ax.set_title("Cumulative number of 10–300 MW plants built since 2015,\nsplit by plant type")
            ax.set_xlabel("Year")
            ax.set_ylabel("Cumulative Plant Count")
            ax.legend()
            ax.grid(True)
            st.pyplot(fig)

            img_buf = io.BytesIO()
            fig.savefig(img_buf, format='png')
            st.download_button(
                label="📥 Download chart as PNG",
                data=img_buf.getvalue(),
                file_name="cumulative_plant_type_chart.png",
                mime="image/png"
            )

            st.download_button(
                label="📥 Download table as Excel",
                data=to_excel_bytes(cumulative),
                file_name="cumulative_plant_type_table.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ Cannot show cumulative chart — required columns missing: 'Start year', 'Capacity (MW)', or a type column like 'Plant Type'.")

else:
    st.info("👆 Please upload an Excel file to begin.")
