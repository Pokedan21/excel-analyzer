import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

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

if uploaded_file:
    sheet_names = get_valid_sheet_names(uploaded_file)
    if not sheet_names:
        st.warning("⚠️ No usable sheets found.")
        st.stop()

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

    # === Sidebar Filters ===
    st.sidebar.header("🔎 Filter Options")

    if st.sidebar.button("🔄 Reset All Filters"):
        for key in list(st.session_state.keys()):
            if key.endswith("_select_all") or key.endswith("_multiselect") or key.endswith("_logic"):
                del st.session_state[key]
        st.experimental_rerun()

    filter_conditions = []

    # Numeric filters
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

    # === String Filters with Per-Filter AND/OR ===
    for column in df.columns:
        if column in ["Capacity (MW)", "Start year", "GEM unit/phase ID"]:
            continue

        col_data = df[column].dropna()

        if pd.api.types.is_string_dtype(col_data):
            with st.sidebar.expander(f"🔤 Filter: {column}", expanded=False):
                logic_key = f"{column}_logic"
                logic_mode = st.radio("Match Mode", ["Match ANY (OR)", "Match ALL (AND)"], key=logic_key)
                unique_values = sorted(col_data.unique().tolist())
                selected_options = st.multiselect(f"{column}", unique_values, key=f"{column}_multiselect")
                if selected_options:
                    if logic_mode.startswith("Match ANY"):
                        condition = df[column].isin(selected_options)
                    else:
                        condition = df[column].apply(lambda x: all(opt in str(x) for opt in selected_options))
                    filter_conditions.append(condition)

    # Apply filters
    if filter_conditions:
        combined_filter = filter_conditions[0]
        for cond in filter_conditions[1:]:
            combined_filter &= cond
        filtered_df = df[combined_filter]
    else:
        filtered_df = df.copy()

    if filtered_df.empty:
        st.warning("⚠️ No data matches your filters.")
        st.stop()

    # === Table ===
    sort_column = st.selectbox("🗂️ Sort by column:", df.columns)
    sort_order = st.radio("Sort Order", ["Ascending", "Descending"])
    filtered_df = filtered_df.sort_values(by=sort_column, ascending=(sort_order == "Ascending"))
    st.write(f"✅ Showing {len(filtered_df)} matching record(s):")
    st.dataframe(filtered_df.head(100))

    # === Interactive Chart Section ===
    st.header("📈 Customize Your Chart")

    chart_type = st.selectbox("Choose chart type", ["Line", "Bar", "Area", "Pie"])
    x_col = st.selectbox("X-axis column", options=filtered_df.columns)
    y_col = st.selectbox("Y-axis column (for counting, use a column with categories)", options=filtered_df.columns)
    group_col = st.selectbox("Group by (optional)", ["None"] + list(filtered_df.columns))
    group_col = None if group_col == "None" else group_col

    show_grid = st.checkbox("Show grid", True)
    rotate_labels = st.checkbox("Rotate x labels", True)
    title = st.text_input("Chart title", value=f"{chart_type} Chart of {y_col} vs {x_col}")
    y_label = st.text_input("Y-axis label", value="Count" if chart_type != "Pie" else "")
    x_label = st.text_input("X-axis label", value=x_col)

    # Build chart data
    if group_col:
        plot_df = filtered_df.groupby([x_col, group_col])[y_col].count().unstack().fillna(0)
    else:
        plot_df = filtered_df.groupby(x_col)[y_col].count()

    # Plot
    fig, ax = plt.subplots()

    if chart_type == "Line":
        plot_df.plot(kind="line", ax=ax, marker="o")
    elif chart_type == "Bar":
        plot_df.plot(kind="bar", ax=ax)
    elif chart_type == "Area":
        plot_df.plot(kind="area", ax=ax, stacked=True)
    elif chart_type == "Pie":
        if isinstance(plot_df, pd.Series):
            plot_df = plot_df[plot_df > 0].nlargest(10)
            plot_df.plot(kind="pie", ax=ax, autopct='%1.1f%%', ylabel='')
        else:
            st.warning("Pie chart only supports a single y-axis series.")

    ax.set_title(title)
    if chart_type != "Pie":
        ax.set_xlabel(x_label)
        ax.set_ylabel(y_label)
        if rotate_labels:
            ax.tick_params(axis="x", rotation=45)
        if show_grid:
            ax.grid(True)

    st.pyplot(fig)

    # === Download
    st.download_button(
        label="📥 Download filtered data as Excel",
        data=filtered_df.to_excel(index=False, engine='openpyxl'),
        file_name=f"{selected_sheet}_filtered_output.xlsx",
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
else:
    st.info("👆 Please upload an Excel file to begin.")
