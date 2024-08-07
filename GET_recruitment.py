import xmlrpc.client
import pandas as pd

url = "https://bridge-logistics-inc.odoo.com"
db = "bridge-logistics-master-297285"
username = "nsmither@bridgelogisticsinc.com"
password = "83b75e5b1ed1151f12bbfc05892a2963d4a24d73"

common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))

#Authentication
uid = common.authenticate(db, username, password, {})

if not uid:
    print("Authentication Failed")
else:
    models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    
    #Search method
    ids = models.execute_kw(db, uid, password, 'hr.applicant', 'search', [[['active', 'in', [True, False]]]])
    fields = models.execute_kw(db, uid, password, 'hr.applicant', 'fields_get', [], {'attributes': ['string', 'help', 'type']})
    records = models.execute_kw(db, uid, password, 'hr.applicant', 'read', [ids], {'fields' :['name', 'job_id', 'stage_id', 'user_id', 'create_date', 'source_id', 'active', 'x_studio_start_date', 'x_studio_team', 'x_studio_sdr_start_date', 'x_studio_d_team_start_date', 'x_studio_henry_start_date', 'x_studio_mcleod_user_code','x_studio_mcleod_salesperson_code', 'x_studio_termination_date', 'x_studio_dat_license', 'x_studio_bi_license', 'x_studio_bamboo_employee_number']})

    #Create a DataFrame from the records
    recruitment = pd.DataFrame(records)

    #Save exel file
    custom_folder = "C:\\Users\\PASVC\\Bridge Logistics Inc\\BL-Bi Team - Documents\\bridge_iq\\Data Sets\\Recruitment\\"
    excel_file_name = "recruitment_api.xlsx"
    excel_file_path = f"{custom_folder}{excel_file_name}"

    recruitment.to_excel(excel_file_path, index=False)

    print("Data saved to", excel_file_path)