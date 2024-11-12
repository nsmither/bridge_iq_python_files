import xmlrpc.client
import pandas as pd
from datetime import datetime
import pytz

url = "https://bridge-logistics-inc.odoo.com"
db = "bridge-logistics-master-297285"
username = "nsmither@bridgelogisticsinc.com"
password = "83b75e5b1ed1151f12bbfc05892a2963d4a24d73"

# Specify the start and end dates
start_date = "2024-11-01"  # Change this to your desired start date
end_date = "2024-11-30"  # Change this to your desired end date

start_date_obj = datetime.strptime(start_date, '%Y-%m-%d')
start_date_month_name_year = start_date_obj.strftime('%B%Y')

common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))

# Authentication
uid = common.authenticate(db, username, password, {})

if not uid:
    print("Authentication failed")
else:
    models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))

    # Search method with date range criteria
    ids = models.execute_kw(db, uid, password, 'mail.message', 'search', [['&', ['date', '>=', start_date], ['date', '<=', end_date], ['mail_activity_type_id', '!=', False]]])
    fields = models.execute_kw(db, uid, password, 'mail.message', 'fields_get', [], {'attributes': ['string', 'help', 'type']})
    records = models.execute_kw(db, uid, password, 'mail.message', 'read', [ids], {'fields': ['date', 'author_id', 'mail_activity_type_id', 'model', 'res_id']})

    # Create a DataFrame from the records
    activities = pd.DataFrame(records)

    #convert the 'date' column to Eastern Time
    eastern = pytz.timezone('US/Eastern')
    activities['date'] = pd.to_datetime(activities['date']).dt.tz_localize(pytz.utc).dt.tz_convert(eastern)

     # Format the 'date' column as a string
    activities['date'] = activities['date'].dt.strftime('%Y-%m-%d %H:%M:%S')

    # Save the DataFrame
    custom_folder = "C:\\Users\\nsmither.bridgelogistics\\Bridge Logistics Inc\\Business Intelligence - Documents\\02 BL-Areas\\bridge_iq\\Data Sets\\behaviors_api_files\\"
    excel_file_name = f"{start_date_month_name_year}.xlsx"
    excel_file_path = f"{custom_folder}{excel_file_name}"
    activities.to_excel(excel_file_path, index=False)

    print("Data saved to", excel_file_path)
