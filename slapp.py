import streamlit as st
import pandas as pd
import plotly.express as px

# Function to load and concatenate data from multiple files and selected sheets
def load_and_combine_data(files_and_sheets):
    combined_df = pd.DataFrame()
    for file, sheet in files_and_sheets:
        df = pd.read_excel(file, sheet_name=sheet)
        combined_df = pd.concat([combined_df, df], ignore_index=True)
    return combined_df

# Title of the Streamlit app
st.title("Excel File Viewer and Graph Creator")

# File uploader
uploaded_files = st.file_uploader("Choose Excel files", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    # Display file names and sheet selectors
    st.write("### Uploaded Files and Sheet Selection")
    
    files_and_sheets = []
    for file in uploaded_files:
        st.write(f"**{file.name}**")
        xls = pd.ExcelFile(file)
        sheet_names = xls.sheet_names
        selected_sheet = st.selectbox(f"Select Sheet for {file.name}", sheet_names, key=file.name)
        files_and_sheets.append((file, selected_sheet))
    
    if files_and_sheets:
        # Load and combine data from the selected sheets across all uploaded files
        combined_data = load_and_combine_data(files_and_sheets)
        
        # Display the combined data
        st.write("### Data Preview: Combined Sheets from Multiple Files")
        st.dataframe(combined_data)

        # Sidebar for column selection
        st.sidebar.write("### Select Columns for Report")
        columns = combined_data.columns.tolist()
        
        selected_columns = st.sidebar.multiselect("Select Columns", columns, default=columns)
        
        if selected_columns:
            selected_data = combined_data[selected_columns]
            st.write("### Selected Data")
            st.dataframe(selected_data)
            
            # Sidebar for slicing and dicing options
            st.sidebar.write("### Filter Data")
            filters = {}
            for column in selected_columns:
                if pd.api.types.is_numeric_dtype(selected_data[column]):
                    min_value = float(selected_data[column].min())
                    max_value = float(selected_data[column].max())
                    step = (max_value - min_value) / 100
                    filters[column] = st.sidebar.slider(f"Filter {column}", min_value, max_value, (min_value, max_value), step=step)
                elif pd.api.types.is_categorical_dtype(selected_data[column]) or selected_data[column].nunique() < 10:
                    filters[column] = st.sidebar.multiselect(f"Filter {column}", selected_data[column].unique(), default=list(selected_data[column].unique()))
                else:
                    filters[column] = st.sidebar.text_input(f"Filter {column} (substring match)")
            
            # Apply filters
            filtered_data = selected_data.copy()
            for column, filter_value in filters.items():
                if isinstance(filter_value, tuple):
                    filtered_data = filtered_data[(filtered_data[column] >= filter_value[0]) & (filtered_data[column] <= filter_value[1])]
                elif isinstance(filter_value, list):
                    filtered_data = filtered_data[filtered_data[column].isin(filter_value)]
                else:
                    filtered_data = filtered_data[filtered_data[column].astype(str).str.contains(filter_value, case=False, na=False)]
            
            st.write("### Filtered Data")
            st.dataframe(filtered_data)
            
            # Select columns for the x and y axis for the bar chart
            st.sidebar.write("### Select Columns for Bar Chart")
            x_axis = st.sidebar.selectbox("Select X-axis", selected_columns)
            y_axis = st.sidebar.selectbox("Select Y-axis", selected_columns)
            
            if x_axis and y_axis:
                fig = px.bar(filtered_data, x=x_axis, y=y_axis)
                st.plotly_chart(fig)
            
            # Option to download the selected data as a CSV file
            csv = filtered_data.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="Download data as CSV",
                data=csv,
                file_name='filtered_data.csv',
                mime='text/csv',
            )
else:
    st.write("Please upload Excel files to view and combine their contents.")
