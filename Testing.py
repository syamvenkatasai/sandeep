import pandas as pd
from zipfile import BadZipFile
from openpyxl.utils.exceptions import InvalidFileException
import logging
import os
 
# Configure logging with specific path
LOG_FILE = os.path.join(os.path.dirname(__file__), "logs", "error_log.txt")
 
# Create logs directory if it doesn't exist
os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
 
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
 
def analyze_pay_equity(df, threshold, compensation_column):
    """
    Pay equity analysis focusing on lowest paid employees within each group
    """
    # Ensure numeric for compensation
    df[compensation_column] = pd.to_numeric(df[compensation_column], errors="coerce")
 
    # Store results in Pay Equity column
    df["Pay Equity"] = ""
 
    # Group by Job Code + Department Code
    grouped = df.groupby(["Job Code", "Department Code"], dropna=False)
 
    # Process each group
    results = []
    for (job, dept), group_data in grouped:
        # Sort descending by compensation (highest to lowest)
        group_data = group_data.sort_values(by=compensation_column, ascending=False).reset_index(drop=True)
 
        # Skip groups with single employee
        if len(group_data) < 2:
            group_data["Pay Equity"] = ""
            results.append(group_data)
            continue
 
        # Process each employee in descending order
        for n in range(len(group_data) - 1, -1, -1):
            current_employee = group_data.iloc[n]
            comp_diff = group_data.iloc[0][compensation_column] - current_employee[compensation_column]
 
            # Skip if compensation difference is within the threshold
            if comp_diff <= threshold:
                continue
 
            # Create sets B and C
            set_B = group_data[group_data[compensation_column] > current_employee[compensation_column] + threshold]
            set_C = group_data[group_data[compensation_column] <= current_employee[compensation_column] + threshold]
 
            # Gender Bias Logic
            gender_bias_flag = False
            if set_C.empty and current_employee["Gender"] not in set_B["Gender"].unique():
                gender_bias_flag = True
                gender_bias_reason = f"Pay Disparity detected with respect to Gender. Employee Number: {set_B.iloc[0]['Employee Number']}"
            elif current_employee["Gender"] not in set_B["Gender"].unique() and all(set_C["Gender"] == current_employee["Gender"]):
                gender_bias_flag = True
                gender_bias_reason = f"Pay Disparity detected with respect to Gender. Employee Number: {set_B.iloc[0]['Employee Number']}"
 
            # Ethnicity Bias Logic
            ethnicity_bias_flag = False
            if set_C.empty and current_employee["Ethnicity"] not in set_B["Ethnicity"].unique():
                ethnicity_bias_flag = True
                ethnicity_bias_reason = f"Pay Disparity detected with respect to Ethnicity. Employee Number: {set_B.iloc[0]['Employee Number']}"
            elif current_employee["Ethnicity"] in set_B["Ethnicity"].unique():
                ethnicity_bias_flag = False
            elif not all(eth in set_C["Ethnicity"].unique() for eth in set_B["Ethnicity"].unique()):
                missing_ethnicity = [eth for eth in set_B["Ethnicity"].unique() if eth not in set_C["Ethnicity"].unique()]
                missing_ethnicity_employee = set_B[set_B["Ethnicity"] == missing_ethnicity[0]].iloc[0]["Employee Number"]
                ethnicity_bias_flag = True
                ethnicity_bias_reason = f"Pay Disparity detected with respect to Ethnicity {missing_ethnicity[0]}. Employee Number: {missing_ethnicity_employee}"
 
            # Update findings
            findings = []
            if gender_bias_flag:
                findings.append(gender_bias_reason)
            if ethnicity_bias_flag:
                findings.append(ethnicity_bias_reason)
 
            if findings:
                group_data.at[n, "Pay Equity"] = " | ".join(findings)
 
        results.append(group_data)
 
    # Combine results
    final_df = pd.concat(results).sort_index()
    return final_df
 
# Simplified output function
def save_report_to_existing_file(df_result, file_path, sheet_name):
    """
    Save the updated DataFrame back to the same Excel file.
    """
    with pd.ExcelWriter(file_path, mode="a", if_sheet_exists="replace") as writer:
        df_result.to_excel(writer, sheet_name=sheet_name, index=False)
 
# Update the main execution code
if __name__ == "__main__":
    file_path = r"filePath"
 
    try:
        # Check if file exists
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"The file '{file_path}' does not exist.")
 
        # Process Salaried sheet first
        salaried_sheet_name = "Salaried"
        df_salaried = pd.read_excel(file_path, sheet_name=salaried_sheet_name)
        threshold_salaried = 2200  # Threshold for salaried employees
        df_salaried_result = analyze_pay_equity(df_salaried, threshold_salaried, "Total Compensation")
        save_report_to_existing_file(df_salaried_result, file_path, salaried_sheet_name)
 
        # Process Hourly sheet
        hourly_sheet_name = "Hourly"
        df_hourly = pd.read_excel(file_path, sheet_name=hourly_sheet_name)
        threshold_hourly = 0.01  # Threshold for hourly employees
        df_hourly_result = analyze_pay_equity(df_hourly, threshold_hourly, "Calculated Compensation")
        save_report_to_existing_file(df_hourly_result, file_path, hourly_sheet_name)
 
        print("Successfully processed both Salaried and Hourly sheets.")
 
    except FileNotFoundError as e:
        logging.error(str(e))
        print(f"Error: {str(e)}")
    except (BadZipFile, InvalidFileException) as e:
        logging.error(f"Invalid Excel file: {str(e)}")
        print(f"Error: The file '{file_path}' is not a valid Excel file or is corrupted.")
    except KeyError as e:
        logging.error(f"Sheet not found: {str(e)}")
        print(f"Error: Sheet '{e}' not found in the Excel file.")
    except Exception as e:
        logging.error(f"Unexpected error: {str(e)}")
        print(f"An unexpected error occurred: {str(e)}")
