import openpyxl
from pathlib import Path
import datetime
import requests

#@csrf_exempt
def getValue(sheet):
    x=""
    if sheet!=None:
        if sheet != "Blank":
            x = sheet
    if isinstance(x, str)  :
        return x.strip()
    return x


def getdate_month_year(sheet):
    #print(type(sheet))
    a = 12
    b = 12
    c = 1980
    try:
        if sheet!=None :
            x = sheet.split(".")
            a = x[0]
            b = x[1]
            c = x[2]
            x = datetime.datetime(a, b, c)
    except:
        a = 12
        b = 12
        c = 1980
    return b,a,c
def updateMSOFFICE(obj):
    mso=(
    ('MS Office Standard 2010', 'MS Office Standard 2010'),
    ('MS Office Standard 2013', 'MS Office Standard 2013'),
    ('MS Office Standard 2016', 'MS Office Standard 2016'),
    ('MS office standard 2013','MS Office Standard 2013'),
    ('MS Office Standard 2019', 'MS Office Standard 2019'),
    )

    for x in mso:
        if x[0] == obj:
            return x[1]
    return obj;
def updateOS(obj):
    OS = (
        ('', ''),
        ('Win.7', 'Win.7'),
        ('Win.10', 'Win.10'),
        ('Ser.2012', 'Ser.2012'),
        ('Blank', ''),
        ('Win.8', 'Win.8'),
        ('Win 8', 'Win.8'),
        ('Win.XP', 'Win.XP'),
        ('Win-7', 'Win.7'),
        ('Ser.2016', 'Ser.2016'),
    )
    for x in OS:
        if x[0]==obj:
            return x[1]
    return obj;
def readAssetSheet(sheet):
    count=0
    for i in range(4,sheet.max_row):
        a = getValue(sheet['C' + str(i)].value) # location
        b = updateOS(getValue(sheet['Z' + str(i)].value))
        c = updateMSOFFICE(getValue(sheet['AA' + str(i)].value))
        d = isOEM(getValue(sheet['AC' + str(i)].value))
        e = getValue(sheet['AE' + str(i)].value)
        f = getValue(sheet['AG' + str(i)].value)
        g = getValue(sheet['AI' + str(i)].value)
        h = getValue(sheet['AJ' + str(i)].value)
        j = getValue(sheet['F' + str(i)].value)
        q = getValue(sheet['E' + str(i)].value)
        r = getValue(sheet['H' + str(i)].value)
        p = getValue(sheet['J' + str(i)].value)
        s = getValue(sheet['L' + str(i)].value)
        t = getValue(sheet['N' + str(i)].value)
        u = getValue(sheet['P' + str(i)].value)
        v = getValue(sheet['R' + str(i)].value)
        aj = getValue(sheet['Y' + str(i)].value)
        ak = getValue(sheet['AB' + str(i)].value)
        al = getValue(sheet['AD' + str(i)].value)
        am = getValue(sheet['AF' + str(i)].value)
        an = getValue(sheet['AH' + str(i)].value)
        ao = getValue(sheet['AK' + str(i)].value)
        ba = getValue(sheet['B' + str(i)].value)
        bb = getValue(sheet['D' + str(i)].value)
        bc = getValue(sheet['G' + str(i)].value)
        bd = getValue(sheet['I' + str(i)].value)
        be = getValue(sheet['K' + str(i)].value)
        bf = getValue(sheet['M' + str(i)].value)
        bg = getValue(sheet['O' + str(i)].value)
        value = getValue(sheet['Q' + str(i)].value)
        bh = value
        bi = getValue(sheet['S' + str(i)].value)
        k = getdate_month_year(sheet['U' + str(i)].value)  # warranty_start_date
        l = getdate_month_year(sheet['V' + str(i)].value)  # warranty_end_date
        m = getdate_month_year(sheet['T' + str(i)].value)  # Purchase Date
        n = getdate_month_year(sheet['W' + str(i)].value)  # AMC Start Date
        o = getdate_month_year(sheet['X' + str(i)].value)  # AMC end Date
        bj = getdate_month_year(sheet['W' + str(i)].value)  # user_acceptance_date
        bk = getdate_month_year(sheet['X' + str(i)].value)  # user_handed_over_date
        params = "&serial_no={0}&user_email=&asset_no={1}&usage_type={2}&gef_id_number={3}&machine_make={4}&machine_serial_no={5}&hdd_make={6}&hdd_serial_no={7}&processor={8}&warranty_start_date_month={9}&warranty_start_date_day={10}&warranty_start_date_year={11}&amc_start_date_month={12}&amc_start_date_day={13}&amc_start_date_year={14}&user_acceptance_date_month=12&user_acceptance_date_day=12&user_acceptance_date_year=1980&OS={15}&ms_office_version={16}{17}&AutoCAD={18}&Visio={19}&SAP={20}&Status={21}&user_name={22}&location={23}&emp_id={24}&machine_type={25}&domain_workgroup={26}&machine_model_no={27}&hdd={28}&hdd_model={29}&ram={30}&processor_purchase_date_month={31}&processor_purchase_date_day={32}&processor_purchase_date_year={33}&warranty_end_date_month={34}&warranty_end_date_day={35}&warranty_end_date_year={36}&amc_end_date_month={37}&amc_end_date_day={38}&amc_end_date_year={39}&user_handed_over_date_month=12&user_handed_over_date_day=12&user_handed_over_date_year=1980&Operating_System_Version={40}&ms_office={41}&Antivirus={42}&Adobe_acrobate={43}&Access={44}&Remarks={45}&Domain_User_Name=NA&SAP_User_ID=NA".format(ba,bb,bc,bd,be,bf,bg,bh,bi,k[0],k[1],k[2],n[0],n[1],n[2],b,c,d,e,f,g,h,j,a,q,r,p,s,t,u,v,m[0],m[1],m[2],l[1],l[0],l[2],o[1],o[0],o[2],aj,ak,al,am,an,ao)
        #params="&serial_no={0}&user_email=&asset_no={1}&usage_type={2}&gef_id_number={3}&machine_make={4}&machine_serial_no={5}&hdd_make={6}&hdd_serial_no={7}&processor={8}&warranty_start_date_month={9}&warranty_start_date_day={10}&warranty_start_date_year={11}&amc_start_date_month={12}&amc_start_date_day={13}&amc_start_date_year={14}&user_acceptance_date_month=&user_acceptance_date_day=&user_acceptance_date_year=&OS={15}&ms_office_version={16}&OEM_Volume={17}&AutoCAD={18}&Visio={19}&SAP={20}&Status={21}&user_name={22}&location={23}&emp_id={24}&machine_type={25}&domain_workgroup={26}&machine_model_no={27}&hdd={28}&hdd_model={29}&ram={30}&processor_purchase_date_month={31}&processor_purchase_date_day={32}&processor_purchase_date_year={33}&warranty_end_date_month={34}&warranty_end_date_day={35}&warranty_end_date_year={36}&amc_end_date_month={37}&amc_end_date_day={38}&amc_end_date_year={39}&user_handed_over_date_month=&user_handed_over_date_day=&user_handed_over_date_year=&Operating_System_Version={40}&ms_office={41}&Antivirus={42}&Adobe_acrobate={43}&Access={44}&Remarks={45}&Domain_User_Name=&SAP_User_ID=".format(ba,bb,bc,bd,be,bf,bg,bh,bi,k[0],k[1],k[2],n[0],n[1],n[2],b,c,d,e,f,g,h,j,a,q,r,p,s,t,u,v,m[0],m[1],m[2],l[1],l[0],l[2],o[1],o[0],o[2],aj,ak,al,am,an,ao)
        #params="&user_name={0}&user_email=&location={1}&asset_no={2}&serial_no={3}&emp_id={4}&usage_type={5}&machine_type={6}&gef_id_number={}&domain_workgroup={}&Domain_User_Name={}&machine_make={}&machine_model_no={}&machine_serial_no={}&hdd={}&hdd_make={}&hdd_model={}&hdd_serial_no={}&ram={}&processor={}&processor_purchase_date={}&warranty_start_date={}&warranty_end_date={}&amc_start_date={}&amc_end_date={}&user_acceptance_date={}&user_handed_over_date={}&Operating_System_Version={}&OS={}&OEM_Volume={}&ms_office={}&ms_office_version={}&Antivirus={}&AutoCAD={}&Adobe_acrobate={}&Visio={}&Access={}&SAP={}&SAP_User_ID={}&Status={}&Remarks={}"
        headers = {'Content-Type': 'application/x-www-form-urlencoded'}
        r = requests.post( 'http://localhost:8000/assets_entry', data=params, headers=headers)
        print(".")
        count=count+1
    print(count)
gc=0
def isOEM(val):
    val=val.strip()

    if val=='OEM':
        return "&OEM_Volume=on"
    return ""
def checkit():
    '''k="serial_no=1&user_email=ibasha%40gmail.comb&asset_no=2&usage_type=Live&gef_id_number=1&machine_make=1&machine_serial_no=1&hdd_make=1&hdd_serial_no=1&processor=Dual_Core&warranty_start_date_month=11&warranty_start_date_day=11&warranty_start_date_year=1940&amc_start_date_month=11&amc_start_date_day=11&amc_start_date_year=1940&user_acceptance_date_month=11&user_acceptance_date_day=11&user_acceptance_date_year=1940&OS=Windows&ms_office_version=Office95&OEM_Volume=on&AutoCAD=on&Visio=on&SAP=on&Status=1&user_name=1&location=Hyderabad&emp_id=yaseen1596&machine_type=Laptop&domain_workgroup=Workgroup&machine_model_no=1&hdd=500MB&hdd_model=1&ram=2GB&processor_purchase_date_month=11&processor_purchase_date_day=11&processor_purchase_date_year=1940&warranty_end_date_month=11&warranty_end_date_day=11&warranty_end_date_year=1940&amc_end_date_month=11&amc_end_date_day=11&amc_end_date_year=1940&user_handed_over_date_month=11&user_handed_over_date_day=11&user_handed_over_date_year=1940&Operating_System_Version=Windows+XP&ms_office=on&Antivirus=on&Adobe_acrobate=on&Access=on&Remarks=Under+Warranty&Domain_User_Name=irshad&SAP_User_ID=sap1234"
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    r = requests.post( 'http://localhost:8000/assetapp/assets_entry', data=k, headers=headers)'''

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    file_path="/home/abed/Downloads"
    xlsx_file = Path(file_path, 'GEF_IT_Asset_Report-KPT.xlsx')
    wb_obj = openpyxl.load_workbook(xlsx_file)
    asset_sheet=wb_obj["IT Asset's"]
    readAssetSheet(asset_sheet)
    #checkit()



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

