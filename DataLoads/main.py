import openpyxl
from pathlib import Path
import requests
def getLocation(obj):

def readAssetSheet(sheet):
    for i in range(4,sheet.max_row):
        params="&serial_no={0}&user_email=&asset_no={1}&usage_type=Live&gef_id_number=1&machine_make=1&machine_serial_no=1&hdd_make=1&hdd_serial_no=1&processor=Dual_Core&warranty_start_date_month=11&warranty_start_date_day=11&warranty_start_date_year=1940&amc_start_date_month=11&amc_start_date_day=11&amc_start_date_year=1940&user_acceptance_date_month=11&user_acceptance_date_day=11&user_acceptance_date_year=1940&OS=Windows&ms_office_version=Office95&OEM_Volume=on&AutoCAD=on&Visio=on&SAP=on&Status=1&user_name=1&location=Hyderabad&emp_id=yaseen1596&machine_type=Laptop&domain_workgroup=Workgroup&machine_model_no=1&hdd=500MB&hdd_model=1&ram=2GB&processor_purchase_date_month=11&processor_purchase_date_day=11&processor_purchase_date_year=1940&warranty_end_date_month=11&warranty_end_date_day=11&warranty_end_date_year=1940&amc_end_date_month=11&amc_end_date_day=11&amc_end_date_year=1940&user_handed_over_date_month=11&user_handed_over_date_day=11&user_handed_over_date_year=1940&Operating_System_Version=Windows+XP&ms_office=on&Antivirus=on&Adobe_acrobate=on&Access=on&Remarks=Under+Warranty".format(sheet['B'+str(i)].value,sheet['D'+str(i)].value,getLocation())
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        #r = requests.post('http://localhost:8000/assetapp/assets_edit', data=params, headers=headers)
        print(params)



def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    file_path="/home/abed/Downloads"
    xlsx_file = Path(file_path, 'GEF_IT_Asset_Report-KPT.xlsx')
    wb_obj = openpyxl.load_workbook(xlsx_file)
    asset_sheet=wb_obj["IT Asset's"]
    readAssetSheet(asset_sheet)



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
