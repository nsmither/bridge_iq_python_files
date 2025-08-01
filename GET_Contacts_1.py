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
    models = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object')

    # Search for all active partner IDs
    ids = models.execute_kw(db, uid, password, 'res.partner', 'search', [[['active', '=', True]]])

    # Define fields to fetch
    fields_to_fetch = [
        'is_company','x_studio_mcleod_code','parent_id','name','city','phone','parent_id','email',
        'x_studio_zoominfo_contact_id','user_id','x_studio_sdr','x_studio_last_stage_update','x_studio_zoominfo_company_id',
        'activity_date_deadline','x_studio_lead_type','x_studio_originally_passed_to','x_studio_do_not_move',
        'x_studio_lead_passed_date','x_studio_enrichment_date','x_studio_revenue_range','create_date',
        'x_studio_primary_industry','state_id','type','x_res_partner_prospecting_stage','x_studio_last_activity_date',
        'x_studio_working_start_date','x_studio_working_end_date','x_studio_no_for_now_start_date','x_studio_no_for_now_end_date',
        'x_studio_hope_island_start_date','x_studio_hope_island_end_date','x_studio_qualified_start_date','x_studio_qualified_end_date',
        'x_studio_quoted_start_date','x_studio_quoted_end_date','x_studio_credit_app_start_date','x_studio_credit_app_end_date',
        'x_studio_1_3_loads_start_date','x_studio_1_3_loads_end_date','x_studio_customer_start_date','x_studio_customer_end_date','category_id'
    ]

    # Helper function to split list into batches
    def chunked(iterable, size):
        for i in range(0, len(iterable), size):
            yield iterable[i:i + size]

    # Fetch records in batches
    all_records = []
    for chunk in chunked(ids, 50000):
        try:
            batch = models.execute_kw(
                db, uid, password,
                'res.partner', 'read',
                [chunk],
                {'fields': fields_to_fetch}
            )
            all_records.extend(batch)
        except Exception as e:
            print(f"Failed to process chunk with IDs {chunk[:5]}...: {e}")

    # Convert to DataFrame
    crm = pd.DataFrame(all_records)

    # === Convert category_id (list of tag IDs) to display names ===
    all_tag_ids = set()
    for row in crm['category_id']:
        if isinstance(row, list):
            all_tag_ids.update(row)

    tag_map = {}
    if all_tag_ids:
        tag_records = models.execute_kw(
            db, uid, password,
            'res.partner.category', 'read',
            [list(all_tag_ids)],
            {'fields': ['id', 'display_name']}
        )
        tag_map = {tag['id']: tag['display_name'] for tag in tag_records}

    def convert_category_ids_to_names(row):
        if isinstance(row, list):
            return ', '.join([tag_map.get(tag_id, f"Unknown({tag_id})") for tag_id in row])
        return ''

    crm['category_names'] = crm['category_id'].apply(convert_category_ids_to_names)

    crm.drop(columns=['category_id'], inplace=True)

    # === Add user_login from res.users based on user_id ===
    user_ids = set()
    for user in crm['user_id']:
        if isinstance(user, list) and len(user) >= 1:
            user_ids.add(user[0])

    user_id_to_login = {}
    if user_ids:
        user_records = models.execute_kw(
            db, uid, password,
            'res.users', 'read',
            [list(user_ids)],
            {'fields': ['id', 'login']}
        )
        user_id_to_login = {user['id']: user['login'] for user in user_records}

    def extract_user_login(user_id_val):
        if isinstance(user_id_val, list) and len(user_id_val) >= 1:
            return user_id_to_login.get(user_id_val[0], f"Unknown({user_id_val[0]})")
        return None

    crm['user_login'] = crm['user_id'].apply(extract_user_login)

    # Save to Excel
    custom_folder = "C:\\Users\\PASVC\\Bridge Logistics Inc\\BL-Bi Team - Documents\\02 BL-Areas\\bridge_iq\\Data Sets\\CRM\\"
    excel_file_name = "new_contacts.xlsx"
    excel_file_path = f"{custom_folder}{excel_file_name}"
    crm.to_excel(excel_file_path, index=False)

    print("Data saved to", excel_file_path)
