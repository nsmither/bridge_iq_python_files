import requests
import pandas as pd
from datetime import datetime, timedelta
import os

# Get yesterday's date in ISO 8601 format
yesterday = datetime.now() - timedelta(days=1)
yesterday_iso = yesterday.isoformat()
yesterday_formatted = yesterday.strftime("%Y-%m-%d")  # Format the date without colons
yesterday_year = yesterday.strftime("%Y")
yesterday_month = yesterday.strftime("%B")

# Define the base URL and construct the complete URL with the date filter
base_url = "https://api.capacity.parade.ai/exports/loads?archived_at__gte="
url = f"{base_url}{yesterday_iso}"

headers = {
    'Authorization': 'Token 0bd30d9a200c94b17ea30fc0fdc01e9f22892f47'
}

response = requests.get(url, headers=headers)

# Check the response status code to ensure it's a successful request (e.g., 200)
if response.status_code == 200:
    data = response.json()  # This will parse the JSON response into a Python data structure

    # Assuming "results" is the key containing the data you want in the DataFrame
    results_data = data.get("results", [])

    df = pd.DataFrame(results_data)

    # Specify the custom folder and file name with backslashes
    #custom_folder = "C:\\Users\\nsmither.bridgelogistics\\Bridge Logistics Inc\\Business Intelligence - Documents\\bridge_iq\\Data Sets\\parade_orders_api_files\\"
    custom_folder = f"C:\\Users\\PASVC\\Bridge Logistics Inc\\BL-Bi Team - Documents\\02 BL-Areas\\bridge_iq\\Data Sets\\Parade\\parade_orders_api_files\\{yesterday_year}\\{yesterday_month}\\"
    if not os.path.exists(custom_folder):
        os.makedirs(custom_folder)


    excel_file_name = f"{yesterday_formatted}.xlsx"
    excel_file_path = f"{custom_folder}{excel_file_name}"

    # Save the DataFrame to the custom folder and file name
    df.to_excel(excel_file_path, index=False)  # Set index to False to exclude index column

    print("Data saved to", excel_file_path)
else:
    print("Request failed with status code:", response.status_code)
