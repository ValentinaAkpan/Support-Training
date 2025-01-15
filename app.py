import streamlit as st
import pandas as pd
import os
import requests

# GitHub raw file links (replace with correct raw URLs)
EXCEL_FILES = [
    "[ENGLISH] [INTERNAL] [2024] BeltMetrics Training Quiz (1-15).xlsx",
    "[ENGLISH] [INTERNAL] [2024] General Product Training(1-78).xlsx",
    "[ENGLISH] [INTERNAL] [2024] PortaMetrics Gen 2 Training Quiz(1-6).xlsx",
    # Add remaining file names
]

# Base URL for GitHub raw files
BASE_URL = "https://raw.githubusercontent.com/ValentinaAkpan/Support-Training/main/"

# Streamlit app title
st.title("Excel File Overview")

# Function to download files
def download_excel_file(filename, download_folder="temp"):
    os.makedirs(download_folder, exist_ok=True)
    file_path = os.path.join(download_folder, filename)
    url = BASE_URL + filename
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()
        with open(file_path, "wb") as f:
            f.write(response.content)
        return file_path
    except Exception as e:
        st.error(f"Failed to download {filename}: {e}")
        return None

# Process files
data = []
for file_name in EXCEL_FILES:
    file_path = download_excel_file(file_name)
    if file_path:
        try:
            # Load Excel file
            df = pd.ExcelFile(file_path)
            for sheet_name in df.sheet_names:
                sheet_data = df.parse(sheet_name)
                if "Name" in sheet_data.columns and "Total points" in sheet_data.columns:
                    for _, row in sheet_data.iterrows():
                        data.append({
                            "Date": row.get("Start time", "N/A"),
                            "Name": row.get("Name", "Name Not Provided"),
                            "Title": file_name,
                            "Total Points": row.get("Total points", "N/A"),
                        })
        except Exception as e:
            st.warning(f"Error processing {file_name}: {e}")

# Display results
if data:
    df_result = pd.DataFrame(data)
    st.dataframe(df_result)
    csv = df_result.to_csv(index=False)
    st.download_button("Download CSV", csv, "overview.csv", "text/csv")
else:
    st.warning("No data to display.")
