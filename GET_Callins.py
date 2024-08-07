import requests
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime, timedelta
import os

current_datetime = datetime.now()
one_hour_ago = current_datetime - timedelta(hours=1)

# Format the dates as string
current_date = datetime.now()
first_day_of_the_month = current_date.replace(day=1)
current_date_string = current_datetime.strftime("%m%d%Y %H:%M")
one_hour_ago_time_string = one_hour_ago.strftime("%m/%d/%Y %H:%M")  # Updated format
month_name_year = first_day_of_the_month.strftime('%B%Y')
current_year_string = current_date.strftime('%Y')

url = f"https://tms.bridgelogisticsinc.com/ws/callins/search?entered_date=>{one_hour_ago_time_string}"

payload = {}
headers = {
    'Authorization': 'Basic YXBpdXNlcjpicmxvYXBpdXNlcg=='
}

response = requests.get(url, headers=headers, data=payload)

# Check if the request was successful (status code 200)
if response.status_code == 200:
    # Parse the XML response
    root = ET.fromstring(response.text)

    # Extract data from XML and create a list of dictionaries
    data_list = []
    for callin_elem in root.findall('.//callin'):  # Adjust the path based on your XML structure
        data = {}
        # Extract data from callin element and add to the dictionary
        # Convert call_date_time to ISO 8601 format
        call_date_time = callin_elem.get('call_date_time')
        if call_date_time:
            call_date_time = datetime.strptime(call_date_time, "%Y%m%d%H%M%S%z").isoformat()
            data['call_date_time'] = call_date_time
        else:
            print("Warning: call_date_time not found for a callin element.")
            continue

        data['entered_by'] = callin_elem.get('entered_by')
        data['initiated_type'] = callin_elem.get('initiated_type')
        data['movement_id'] = callin_elem.get('movement_id')
        data['payee_id'] = callin_elem.get('payee_id')
        data['id'] = callin_elem.get('id')
        # Add more keys as needed

        data_list.append(data)

    # Create a DataFrame from the list of dictionaries
    new_df = pd.DataFrame(data_list)

    # Create the custom folder if it doesn't exist
    custom_folder = f"C:\\Users\\PASVC\\Bridge Logistics Inc\\BL-Bi Team - Documents\\bridge_iq\\Data Sets\\Callins\\"
    if not os.path.exists(custom_folder):
        os.makedirs(custom_folder)

    # Save the DataFrame to an Excel file
    excel_file_name = f"{month_name_year}.xlsx"
    excel_file_path = f"{custom_folder}{current_year_string}\\{excel_file_name}"

    # Check if the Excel file already exists
    if os.path.exists(excel_file_path):
        # Load the existing Excel file
        existing_df = pd.read_excel(excel_file_path)

        # Check for duplicate IDs in the new data
        duplicate_ids = existing_df[existing_df['id'].isin(new_df['id'])]['id'].tolist()

        # Filter out rows with duplicate IDs from the new data
        new_df = new_df[~new_df['id'].isin(duplicate_ids)]

        # Concatenate the new DataFrame with the existing data
        updated_df = pd.concat([existing_df, new_df], ignore_index=True)

        # Save the updated DataFrame to the Excel file
        updated_df.to_excel(excel_file_path, index=False)
    else:
        # Save the new DataFrame to the Excel file
        new_df.to_excel(excel_file_path, index=False)

    print("Data saved to", excel_file_path)

else:
    print(f"Request failed with status code: {response.status_code}")
    print(response.text)