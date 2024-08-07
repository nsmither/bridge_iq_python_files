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
    ids = models.execute_kw(db, uid, password, 'crm.lead', 'search', [[]])
    fields = models.execute_kw(db, uid, password, 'crm.lead', 'fields_get', [], {'attributes': ['string', 'help', 'type']})
    records = models.execute_kw(db, uid, password, 'crm.lead', 'read', [ids], {'fields': ['partner_id','name','x_studio_mcleod_code','user_id','x_studio_sdr','stage_id','date_last_stage_update','activity_date_deadline','date_open','partner_id', 'x_studio_do_not_move','x_studio_lead_type','x_studio_originally_passed_to','x_studio_lead_passed_date','x_studio_enrichment_date','city','x_studio_revenue_range','x_studio_primary_industry','state_id','phone','type','x_date_last_prospect_stage_change','x_stage_id']})

    # Create a DataFrame from the records
    crm= pd.DataFrame(records)

    #print(crm)
    
    #Save exel file
    custom_folder = "C:\\Users\\PASVC\\Bridge Logistics Inc\\BL-Bi Team - Documents\\bridge_iq\\Data Sets\\CRM\\"
    excel_file_name = "CRM_api.xlsx"
    excel_file_path = f"{custom_folder}{excel_file_name}"

    crm.to_excel(excel_file_path, index=False)

    print("Data saved to", excel_file_path)