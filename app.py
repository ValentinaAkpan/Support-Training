import streamlit as st
import pandas as pd
import requests
from io import BytesIO

# List of raw GitHub URLs for your Excel files
EXCEL_FILE_URLS = {
    "[ENGLISH] [INTERNAL] [2024] BeltMetrics Training Quiz (1-15)": "https://raw.githubusercontent.com/ValentinaAkpan/Support-Training/main/BeltMetrics_Training_Quiz.xlsx",
    "[ENGLISH] [INTERNAL] [2024] General Product Training(1-78)": "https://raw.githubusercontent.com/ValentinaAkpan/Support-Training/main/General_Product_Training.xlsx",
    "[ENGLISH] [INTERNAL] [2024] PortaMetrics Gen 2 Training Quiz(1-6)": "https://raw.githubusercontent.com/ValentinaAkpan/Support-Training/main/PortaMetrics_Training_Quiz.xlsx",
    "LoaderMetrics™ Gen 2 Features Monitoring on MetricsManager Pro [EN](1-2)": "https://raw.githubusercontent.com/ValentinaAkpan/Support-Training/main/LoaderMetrics_Features_Monitoring.xlsx",
    "[ENGLISH] [INTERNAL] [2024] TruckMetrics Training Quiz(1-39)": "https://raw.githubusercontent.com/ValentinaAkpan/Support-Training/main/TruckMetrics_Training_Quiz.xlsx",
    "[ENGLISH] [INTERNAL] [2024] LoaderMetrics Training Quiz - TVM (1-37)": "https://raw.githubusercontent.com/ValentinaAkpan/Support-Training/main/LoaderMetrics_Training_Quiz.xlsx",
    "[ENGLISH] [INTERNAL] [2024] ShovelMetrics Training Quiz - TVFWPM (1-52)": "https://raw.githubusercontent.com/ValentinaAkpan/Support-Training/main/ShovelMetrics_Training_Quiz.xlsx",
    "ShovelMetrics™ Gen 3 Training Overview - Onboarding - [EN](1-8)": "https://raw.githubusercontent.com/ValentinaAkpan/Support-Training/main/ShovelMetrics_Training_Overview.xlsx",
    "[EN] ShovelMetrics™ Gen 3 Features G.E.T Monitoring on MetricsManager Pro(1-8)": "https://raw.githubusercontent.com/ValentinaAkpan/Support-Training/main/ShovelMetrics_GET_Monitoring.xlsx",
    "ShovelMetrics™ Gen 3 Features Rock Monitoring on MetricsManager Pro [EN](1-9)": "https://raw.githubusercontent.com/ValentinaAkpan/Support-Training/main/ShovelMetrics_Rock_Monitoring.xlsx",
}



# Streamlit app title
st.title("Excel File Overview")

if not EXCEL_FILE_URLS:
    st.warning("No Excel file URLs provided.")
else:
    data = []

    for filename, url in EXCEL_FILE_URLS.items():
        try:
            # Fetch the file from the URL
            response = requests.get(url)
            response.raise_for_status()  # Check for request errors

            # Load the Excel file
            file_content = BytesIO(response.content)
            df = pd.ExcelFile(file_content)

            product_name = file_mapping.get(filename, "Unknown Product")  # Extract product name using the mapping

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
                                "Title": filename,  # Use the exact filename as title
                                "Product": product_name,  # Include the product name
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
