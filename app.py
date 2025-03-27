import streamlit as st
import pandas as pd

# Streamlit UI
st.title("Excel Processing App")
st.write("Upload two XLS files, and the app will generate an output file.")

# File uploaders with notes
file1 = st.file_uploader("Upload 1st XLS file (e.g., 'SID List Report Full Detail TNT')", type=['xls'])
file2 = st.file_uploader("Upload 2nd XLS file (e.g., 'F.I.R. Time Difference AnalysisTNT')", type=['xls'])

if file1 and file2:
    try:
        # Read Excel files
        df1 = pd.read_excel(file1, engine='xlrd', header=None)
        df2 = pd.read_excel(file2, engine='xlrd', header=None)

        # Extract Police Station Name (B5 cell in 2nd file)
        police_station_name = df2.iloc[4, 1]  # B5 cell

        # Prepare Output File Name
        output_filename = f"{police_station_name}.xlsx"

        # Extract Data for Case Number 1, Case Number 2, and FIR Number
        case_number_1 = df1.iloc[3:, 2].reset_index(drop=True)  # C4 and below
        case_number_2 = df1.iloc[3:, 10].reset_index(drop=True)  # K4 and below
        fir_number = df2.iloc[4:, 2].reset_index(drop=True)  # C5 and below

        # Create Output DataFrame
        output_df = pd.DataFrame({
            "Case Number 1": case_number_1,
            "Case Number 2": case_number_2,
            "FIR Number": fir_number
        })

        # Generate Final Output
        output_df["Final Output"] = output_df["FIR Number"].apply(
            lambda x: x if x in case_number_1.values or x in case_number_2.values else None
        )

        # Count Filled Rows
        fir_filled_count = output_df["FIR Number"].count()
        final_filled_count = output_df["Final Output"].count()

        # Create Summary Row
        count_row = pd.DataFrame({
            "Case Number 1": [output_df["Case Number 1"].iloc[0]],
            "Case Number 2": [output_df["Case Number 2"].iloc[0]],
            "FIR Number": [output_df["FIR Number"].iloc[0]],
            "Final Output": [output_df["Final Output"].iloc[0]],
            "FIR Number Count": [f"Filled: {fir_filled_count}"],
            "Final Output Count": [f"Filled: {final_filled_count}"],
            "Police Station": [police_station_name]
        })

        # Remove Blank Space in First 4 Columns
        output_df = output_df.iloc[1:].reset_index(drop=True)

        # Concatenate Summary Row with Data
        output_df = pd.concat([count_row, output_df], ignore_index=True)

        # Save to Excel
        output_df.to_excel(output_filename, index=False, engine='openpyxl')

        # Provide Download Link
        with open(output_filename, "rb") as file:
            st.download_button(label="Download Processed Excel File", data=file, file_name=output_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.success("File processed successfully! Download your output file above.")

    except Exception as e:
        st.error(f"An error occurred: {e}")

