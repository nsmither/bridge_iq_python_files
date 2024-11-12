import xmlrpc.client
import pandas as pd
from datetime import datetime
import pytz
import os

url = "https://bridge-logistics-inc.odoo.com"
db = "bridge-logistics-master-297285"
username = "nsmither@bridgelogisticsinc.com"
password = "83b75e5b1ed1151f12bbfc05892a2963d4a24d73"

# Get the date
current_date = datetime.now()
first_day_of_the_month = current_date.replace(day=1)
formatted_date = first_day_of_the_month.strftime('%Y-%m-%d')
month_name_year = first_day_of_the_month.strftime('%B%Y')
common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
current_year_string = current_date.strftime('%Y')

# Authentication
uid = common.authenticate(db, username, password, {})

if not uid:
    print("Authentication failed")
else:
    models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))

    # Search method
    ids = models.execute_kw(db, uid, password, 'mail.message', 'search', [['&', ['date', '>=', formatted_date], ['mail_activity_type_id', '!=', False]]])
    fields = models.execute_kw(db, uid, password, 'mail.message', 'fields_get', [], {'attributes': ['string', 'help', 'type']})
    records = models.execute_kw(db, uid, password, 'mail.message', 'read', [ids], {'fields': ['date', 'author_id', 'mail_activity_type_id', 'model', 'res_id']})

    # Create a DataFrame from the records
    activities = pd.DataFrame(records)

    # Convert the 'date' column to Eastern Time
    eastern = pytz.timezone('US/Eastern')
    activities['date'] = pd.to_datetime(activities['date']).dt.tz_localize(pytz.utc).dt.tz_convert(eastern)

    # Format the 'date' column as a string
    activities['date'] = activities['date'].dt.strftime('%Y-%m-%d %H:%M:%S')

    custom_folder = f"C:\\Users\\PASVC\\Bridge Logistics Inc\\BL-Bi Team - Documents\\02 BL-Areas\\bridge_iq\\Data Sets\\behaviors_api_files\\{current_year_string}\\"
    if not os.path.exists(custom_folder):
        os.makedirs(custom_folder)

    # Save the DataFrame to Excel
    excel_file_name = f"{month_name_year}.xlsx"
    excel_file_path = f"{custom_folder}{excel_file_name}"
    activities.to_excel(excel_file_path, index=False, engine='xlsxwriter')

    print("Data saved to", excel_file_path)