import requests
import pandas as pd
from datetime import datetime, timedelta
import os

# Get yesterday's date in ISO 8601 format
yesterday = datetime.now() - timedelta(days=1)
yesterday_iso = yesterday.isoformat()
yesterday_formatted = yesterday.strftime("%Y-%m-%d")
yesterday_year = yesterday.strftime("%Y")
yesterday_month = yesterday.strftime("%B")

# Define the base URL and construct the complete URL with the date filter
base_url = "https://api.capacity.parade.ai/exports/quotes?created_at__gte="
url = f"{base_url}{yesterday_iso}"

headers = {
    'Authorization': 'Token 0bd30d9a200c94b17ea30fc0fdc01e9f22892f47'
}

response = requests.get(url, headers=headers)

# Check the response status code to ensure it's a successful request (e.g., 200)
if response.status_code == 200:
    data = response.json()  # This will parse the JSON response into a Python data structure

    # Turn Results into a df
    results_data = data.get("results", [])

    df = pd.DataFrame(results_data)

    #Check if custom folder exists if not create it
    custom_folder = f"C:\\Users\\PASVC\\Bridge Logistics Inc\\BL-Bi Team - Documents\\02 BL-Areas\\bridge_iq\\Data Sets\\Parade\\parade_quotes_api_files\\{yesterday_year}\\{yesterday_month}\\"
    if not os.path.exists(custom_folder):
        os.makedirs(custom_folder)

    excel_file_name = f"{yesterday_formatted}.xlsx"
    excel_file_path = f"{custom_folder}{excel_file_name}"

    # Save the DataFrame to the custom folder and file name
    df.to_excel(excel_file_path, index=False)  # Set index to False to exclude index column

    print("Data saved to", excel_file_path)
else:
    print("Request failed with status code:", response.status_code)