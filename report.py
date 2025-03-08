import pandas as pd
import requests
import base64
from openpyxl import load_workbook


# Load settings from the "Settings" sheet
def load_settings(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name="Settings", header=None)
        settings = {df.iloc[i, 0]: str(df.iloc[i, 1]).strip() for i in range(len(df))}
        return settings
    except Exception as e:
        print(f"Error loading settings: {e}")
        return None


# Fetch BambooHR custom report
def fetch_bamboohr_report(file_path):
    # settings = load_settings(file_path)
    # if not settings:
    #     print("Error: Could not load settings.")
    #    return
    
    # api_key = settings.get("bamboo_key")
    # subdomain = settings.get("bamboo_domain")
    # report_id = settings.get("bamboo_comp_report")

    api_key = "4d5052be2a809b964d915109d871332d92ba7aea"
    subdomain = "cognite"
    report_id = "2299"

    if not api_key or not subdomain or not report_id:
        print("Error: Missing API key, domain, or report ID.")
        return

    url = f"https://api.bamboohr.com/api/gateway.php/{subdomain}/v1/reports/{report_id}"
    print(f"Requesting Report URL: {url}")

    headers = {
        "Authorization": "Basic " + base64.b64encode(f"{api_key}:".encode()).decode(),
        "Accept": "application/json" #,
        # "Content-Type": "application/json; charset=UTF-8"
    }

    try:
        response = requests.get(url, headers=headers)
        print(f"Response Code: {response.status_code}")

        if response.status_code != 200:
            print("Failed to fetch report. Check API permissions and report ID.")
            return

        report_data = response.json()
        print(report_data)

        if "fields" not in report_data or "employees" not in report_data:
            print("No valid data found in report.")
            return

        # Extract headers and data
        field_headers = [field["name"] for field in report_data["fields"]]
        field_ids = [field["id"] for field in report_data["fields"]]

        employees_data = [[emp.get(field_id, "") for field_id in field_ids] for emp in report_data["employees"]]

        # Save to Excel
        with pd.ExcelWriter(file_path, mode="a", if_sheet_exists="replace") as writer:
            df = pd.DataFrame(employees_data, columns=field_headers)
            df.to_excel(writer, sheet_name=report_id, index=False)

        print(f"Report '{report_id}' successfully updated in '{file_path}'.")
    except Exception as e:
        print(f"Error fetching BambooHR report: {e}")


# Fetch all employees from BambooHR
def fetch_bamboohr_employees(file_path):
    settings = load_settings(file_path)
    if not settings:
        print("Error: Could not load settings.")
        return

    api_key = settings.get("bamboo_key")
    subdomain = settings.get("bamboo_domain")

    if not api_key or not subdomain:
        print("Error: Missing API key or domain.")
        return

    url = f"https://api.bamboohr.com/api/gateway.php/{subdomain}/v1/employees/directory"
    print(f"Requesting URL: {url}")

    headers = {
        "Authorization": "Basic " + base64.b64encode(f"{api_key}:".encode()).decode(),
        "Accept": "application/json"
    }

    try:
        response = requests.get(url, headers=headers)
        print(f"Response Code: {response.status_code}")

        if response.status_code == 401:
            print("Authentication Error: Check API key and subdomain.")
            return

        employees_data = response.json().get("employees", [])

        # Prepare DataFrame
        df = pd.DataFrame(employees_data)
        df = df[["id", "firstName", "lastName", "workEmail", "jobTitle", "department", "location"]]
        df.columns = ["Employee ID", "First Name", "Last Name", "Email", "Job Title", "Department", "Location"]

        # Save to Excel
        with pd.ExcelWriter(file_path, mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name="All Employees", index=False)

        print("Employee directory successfully updated.")
    except Exception as e:
        print(f"Error fetching BambooHR data: {e}")

# Example Usage
# if __name__ == "__main__":
#    excel_file = "bamboohr_data.xlsx"  # Change this to your Excel file path
#    fetch_bamboohr_report(excel_file)
#    fetch_bamboohr_employees(excel_file)


fetch_bamboohr_report("Report.xlsx")