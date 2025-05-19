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

# Excel byte generator for download buttons
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
            df = sample_df.copy()
    unique_columns = st.sidebar.multiselect("Select columns to filter for unique values only", options=df.columns.tolist())

        st.sidebar.header("🔎 Filter Options")

        # Universal filter logic toggle
        logic_mode = st.sidebar.radio("Combine all filters using:", ["AND", "OR"], index=0)

        # Reset button
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

        # Boolean filter: GEM unit
        if "GEM unit/phase ID" in df.columns:
            include_gem_only = st.sidebar.checkbox("✅ Only include GEM units", value=False)
            if include_gem_only:
                condition = df["GEM unit/phase ID"].notna()
                filter_conditions.append(condition)


        st.sidebar.markdown("### 🔍 Column Search and Unique Value Filtering")
        show_unique_only = st.sidebar.checkbox("Show only unique values per column", value=False)
    # String filters with Match ANY / Match ALL logic



        for i, column in enumerate(df.columns):
                if column in ["Capacity (MW)", "Start year", "GEM unit/phase ID"]:
                    continue

                if pd.api.types.is_string_dtype(df[column]):
                    with st.sidebar.expander(f"🔤 Filter: {column}", expanded=False):
                        search_term = st.text_input(f"Search {column}", key=f"search_{i}_{column}")
                        logic_mode_col = st.radio("Match Mode", ["Match ANY (OR)", "Match ALL (AND)"], key=f"radio_{i}_{column}")

                        col_data = df[column].dropna().astype(str)

                        if column in unique_columns:
                            vc = df[column].value_counts()
                            unique_vals = vc[vc == 1].index.tolist()
                            col_data = col_data[col_data.isin(unique_vals)]

                        if search_term:
                            col_data = col_data[col_data.str.contains(search_term, case=False, na=False)]

                        unique_values = sorted(col_data.unique().tolist())
                        selected_options = st.multiselect(f"{column}", unique_values, key=f"multi_{i}_{column}")

                        if selected_options:
                            if logic_mode_col.startswith("Match ANY"):
                                condition = df[column].astype(str).isin(selected_options)
                                condition = df[column].astype(str).apply(lambda x: all(opt in str(x) for opt in selected_options))
                            filter_conditions.append(condition)
        # Apply combined filter
        if filter_conditions:
            combined_filter = filter_conditions[0]
            for cond in filter_conditions[1:]:
                combined_filter = combined_filter & cond if logic_mode == "AND" else combined_filter | cond
            filtered_df = df[combined_filter]
            filtered_df = df.copy()

        if filtered_df.empty:
            st.warning("⚠️ No data matches your filters.")
            st.stop()

            # Sort and show
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

        # Chart Builder
        st.header("📈 Customize Your Chart")

        chart_type = st.selectbox("Choose chart type", ["Line", "Bar", "Area", "Pie"])
        x_col = st.selectbox("X-axis column", options=filtered_df.columns)
        y_col = st.selectbox("Y-axis column", options=filtered_df.columns)
        group_col = st.selectbox("Group by (optional)", ["None"] + list(filtered_df.columns))
        group_col = None if group_col == "None" else group_col

        show_grid = st.checkbox("Show grid", True)
        rotate_labels = st.checkbox("Rotate x labels", True)
        title = st.text_input("Chart title", value=f"{chart_type} Chart of {y_col} vs {x_col}")
        y_label = st.text_input("Y-axis label", value="Count" if chart_type != "Pie" else "")
        x_label = st.text_input("X-axis label", value=x_col)

        if group_col:
            plot_df = filtered_df.groupby([x_col, group_col])[y_col].count().unstack().fillna(0)
            plot_df = filtered_df.groupby(x_col)[y_col].count()

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

        st.download_button(
            label="📊 Download chart data as Excel",
            data=to_excel_bytes(plot_df if isinstance(plot_df, pd.DataFrame) else plot_df.to_frame()),
            file_name="chart_data.xlsx",
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

        # Cumulative Chart
        st.header("📈 Cumulative Plant Chart (Filtered Data)")
        if "Capacity (MW)" in filtered_df.columns and "Start year" in filtered_df.columns:
            group_10_300 = filtered_df[(filtered_df["Capacity (MW)"] >= 10) & (filtered_df["Capacity (MW)"] <= 300)]
            group_300_plus = filtered_df[filtered_df["Capacity (MW)"] > 300]

            cum_10_300 = group_10_300["Start year"].value_counts().sort_index().cumsum()
            cum_300_plus = group_300_plus["Start year"].value_counts().sort_index().cumsum()

            cum_df = pd.DataFrame({
                "10–300 MW": cum_10_300,
                "300+ MW": cum_300_plus
            }).fillna(method='ffill').fillna(0)

            fig2, ax2 = plt.subplots()
            cum_df.plot(ax=ax2, marker='o')
            ax2.set_title("Cumulative Count of Filtered Plants by Size")
            ax2.set_ylabel("Cumulative Count")
            ax2.set_xlabel("Start Year")
            ax2.grid(True)
            st.pyplot(fig2)

            st.download_button(
                label="📈 Download cumulative data as Excel",
                data=to_excel_bytes(cum_df),
                file_name="cumulative_data.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        # ✅ New robust cumulative chart by plant type using full dataset
        st.header("📈 Cumulative 10–300 MW Plants by Type Since 2015")


        # Optional: Allow switching between full vs filtered dataset
        use_filtered_data = st.checkbox("Use filtered data for chart", value=False)
        source_df = filtered_df if use_filtered_data else df

        type_col = None
        for col in ["Type", "Plant Type", "Technology"]:
            if col in source_df.columns:
                type_col = col
                break

        if type_col is None:
            st.warning("⚠️ Could not find a column for plant type (e.g., 'Type', 'Plant Type', or 'Technology').")
            cum_df = source_df[
                (source_df["Capacity (MW)"] >= 10) &
                (source_df["Capacity (MW)"] <= 300) &
                (source_df["Start year"] >= 2015)
            ].copy()

            cum_df = cum_df[[type_col, "Start year"]].dropna()
            cum_df["Start year"] = pd.to_numeric(cum_df["Start year"], errors="coerce").astype("Int64")

            st.write("Using data rows:", len(cum_df))

            group_counts = (
                cum_df
                .groupby(["Start year", cum_df[type_col]])
                .size()
                .unstack()
                .fillna(0)
            )

            if group_counts.empty:
                st.warning("⚠️ No valid data found to generate the cumulative chart.")
                group_counts = group_counts.sort_index().cumsum()
                group_counts = group_counts.apply(pd.to_numeric, errors='coerce').fillna(0)

                fig3, ax3 = plt.subplots(figsize=(10, 6))
                group_counts.plot(ax=ax3, marker='o')
                ax3.set_title("Cumulative number of 10 - 300 MW plants built starting 2015,\nsplit into different plant types")
                ax3.set_xlabel("Year")
                ax3.set_ylabel("Cumulative Count")
                ax3.grid(True)
                ax3.legend(title=type_col)
                st.pyplot(fig3)

                st.download_button(
                    label="📥 Download cumulative chart data",
                    data=to_excel_bytes(group_counts),
                    file_name="cumulative_by_type.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )


else:
    st.info("👆 Please upload an Excel file to begin.")