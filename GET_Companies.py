import xmlrpc.client
import pandas as pd

url = "https://bridge-logistics-inc.odoo.com"
db = "bridge-logistics-master-297285"
username = "nsmither@bridgelogisticsinc.com"
password = "83b75e5b1ed1151f12bbfc05892a2963d4a24d73"

common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))

# Authentication
uid = common.authenticate(db, username, password, {})

if not uid:
    print("Authentication failed")
else:
    models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))

    # Search method
    ids = models.execute_kw(db, uid, password, 'res.partner', 'search', [[['is_company', '=', True]]])
    fields = models.execute_kw(db, uid, password, 'res.partner', 'fields_get', [], {'attributes': ['string', 'help', 'type']})
    records = models.execute_kw(db, uid, password, 'res.partner', 'read', [ids], {'fields': ['name', 'x_studio_zoominfo_company_id', 'opportunity_count', 'parent_id', 'city', 'phone', 'state_id']})

    # Create a DataFrame from the records
    crm= pd.DataFrame(records)

    #print(crm)
    
    #Save exel file
    custom_folder = "C:\\Users\\PASVC\\Bridge Logistics Inc\\BL-Bi Team - Documents\\02 BL-Areas\\bridge_iq\\Data Sets\\CRM\\"
    excel_file_name = "companies.xlsx"
    excel_file_path = f"{custom_folder}{excel_file_name}"

    crm.to_excel(excel_file_path, index=False)

    print("Data saved to", excel_file_path)