import streamlit as st
import pandas as pd
import os

# Exact file names
EXCEL_FILES = [
    "[ENGLISH] [INTERNAL] [2024] BeltMetrics Training Quiz (1-15).xlsx",
    "[ENGLISH] [INTERNAL] [2024] General Product Training(1-78).xlsx",
    "[ENGLISH] [INTERNAL] [2024] PortaMetrics Gen 2 Training Quiz(1-6).xlsx",
    "LoaderMetrics™ Gen 2 Features Monitoring on MetricsManager Pro [EN](1-2).xlsx",
    "[ENGLISH] [INTERNAL] [2024] TruckMetrics Training Quiz(1-39).xlsx",
    "[ENGLISH] [INTERNAL] [2024] LoaderMetrics Training Quiz - TVM (1-37).xlsx",
    "[ENGLISH] [INTERNAL] [2024] ShovelMetrics Training Quiz - TVFWPM (1-52).xlsx",
    "ShovelMetrics™ Gen 3 Training Overview - Onboarding - [EN](1-8).xlsx",
    "[EN] ShovelMetrics™ Gen 3 Features G.E.T Monitoring on MetricsManager Pro(1-8).xlsx",
    "ShovelMetrics™ Gen 3 Features Rock Monitoring on MetricsManager Pro [EN](1-9).xlsx",
]

# Path to the folder containing Excel files
EXCEL_FOLDER = r"https://github.com/ValentinaAkpan/Support-Training"

# Streamlit app title
st.title("Excel File Overview")

# Check if the folder exists
if not os.path.exists(EXCEL_FOLDER):
    st.error(f"The folder '{EXCEL_FOLDER}' does not exist. Please check the path.")
else:
    # Get only specified Excel files from the folder
    excel_files = [f for f in os.listdir(EXCEL_FOLDER) if f in EXCEL_FILES]

    if not excel_files:
        st.warning("No matching Excel files found in the specified folder.")
    else:
        data = []

        for filename in excel_files:
            file_path = os.path.join(EXCEL_FOLDER, filename)
            title = filename  # Use the filename as the title

            try:
                # Load the Excel file
                df = pd.ExcelFile(file_path)

                unique_names_per_title = set()  # To track unique names for this title

                for sheet_name in df.sheet_names:
                    sheet_data = df.parse(sheet_name)

                    # Check for columns where names might exist
                    name_column = None
                    if "Name" in sheet_data.columns:
                        name_column = "Name"
                    elif "Please add your First name and Surname" in sheet_data.columns:
                        name_column = "Please add your First name and Surname"

                    # Process rows if a valid name column exists
                    if name_column and all(col in sheet_data.columns for col in ["Total points", "Start time"]):
                        for _, row in sheet_data.iterrows():
                            taker_name = row[name_column]

                            # Fallback if name is empty or placeholder
                            if pd.isna(taker_name) or not taker_name.strip() or taker_name == "Please add your First name and Surname":
                                taker_name = "Name Not Provided"

                            # Add unique names only
                            if taker_name not in unique_names_per_title:
                                unique_names_per_title.add(taker_name)
                                data.append({
                                    "Date": row["Start time"],  # Use "Start time" as date
                                    "Name": taker_name,  # Extract the taker's name
                                    "Title": title,  # Use the filename as title
                                    "Total Points": row["Total points"],  # Extract total points
                                })
                    else:
                        st.warning(f"Sheet '{sheet_name}' in file '{filename}' is missing required columns.")
            except Exception as e:
                st.error(f"Error processing file '{filename}': {e}")

        # Display the collected data as a table
        if data:
            df_result = pd.DataFrame(data)
            st.write("File Overview:")
            st.dataframe(df_result)

            # Option to download the result
            csv = df_result.to_csv(index=False)
            st.download_button(
                label="Download Results as CSV",
                data=csv,
                file_name="file_overview.csv",
                mime="text/csv",
            )
