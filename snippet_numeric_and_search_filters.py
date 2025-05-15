# Modified numeric inputs and enhanced filtering
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