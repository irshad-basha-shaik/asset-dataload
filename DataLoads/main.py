import openpyxl
from pathlib import Path
import datetime
import requests
def getValue(sheet):
    x = sheet
    return x


def getdate_month_year(sheet):
    #print(type(sheet))
    a = 0
    b = 0
    c = 0
    try:
        if sheet!=None :
            x = sheet.split(".")
            a = x[0]
            b = x[1]
            c = x[2]
    except:
        a = 0
        b = 0
        c = 0
    return b,a,c

#def getoFFICEvERSION(obj):
def readAssetSheet(sheet):
    for i in range(4,sheet.max_row):
        a = getValue(sheet['C' + str(i)].value) # location
        b = getValue(sheet['Z' + str(i)].value)
        c = getValue(sheet['AA' + str(i)].value)
        d = getValue(sheet['AC' + str(i)].value)
        e = getValue(sheet['AE' + str(i)].value)
        f = getValue(sheet['AG' + str(i)].value)
        g = getValue(sheet['AI' + str(i)].value)
        h = getValue(sheet['AJ' + str(i)].value)
        j = getValue(sheet['F' + str(i)].value)
        q = getValue(sheet['E' + str(i)].value)
        r = getValue(sheet['H' + str(i)].value)
        k = getdate_month_year(sheet['U' + str(i)].value)#warranty_start_date
        l = getdate_month_year(sheet['V' + str(i)].value)#warranty_end_date
        m = getdate_month_year(sheet['T' + str(i)].value)#Purchase Date
        n = getdate_month_year(sheet['W' + str(i)].value)#AMC Start Date
        o = getdate_month_year(sheet['X' + str(i)].value)#AMC end Date
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
        bh = getValue(sheet['Q' + str(i)].value)
        bi = getValue(sheet['S' + str(i)].value)


        #params = "&Operating_System_Version={0}&ms_office={1}&Antivirus={2}&Adobe_acrobate={3}&Access={4}&Remarks={5}".format(aj,ak,al,am,an,ao)

        params="&serial_no={0}&user_email=&asset_no={1}&usage_type={2}&gef_id_number={3}&machine_make={4}&machine_serial_no={5}&hdd_make={6}&hdd_serial_no={7}&processor={8}&warranty_start_date_month={9}&warranty_start_date_day={10}&warranty_start_date_year={11}&amc_start_date_month={12}&amc_start_date_day={13}&amc_start_date_year={14}&user_acceptance_date_month=&user_acceptance_date_day=&user_acceptance_date_year=&OS={15}&ms_office_version={16}&OEM_Volume={17}&AutoCAD={18}&Visio={19}&SAP={20}&Status={21}&user_name={22}&location={23}&emp_id={24}&machine_type={25}&domain_workgroup={26}&machine_model_no={27}&hdd={28}&hdd_model={29}&ram={30}&processor_purchase_date_month={31}&processor_purchase_date_day={32}&processor_purchase_date_year={33}&warranty_end_date_month={34}&warranty_end_date_day={35}&warranty_end_date_year={36}&amc_end_date_month={37}&amc_end_date_day={38}&amc_end_date_year={39}&user_handed_over_date_month=&user_handed_over_date_day=&user_handed_over_date_year=&Operating_System_Version={40}&ms_office={41}&Antivirus={42}&Adobe_acrobate={43}&Access={44}&Remarks={45}".format(ba,bb,bc,bd,be,bf,bg,bh,bi,k[0],k[1],k[2],n[0],n[1],n[2],b,c,d,e,f,g,h,j,a,q,r,p,s,t,u,v,m[0],m[1],m[2],l[1],l[0],l[2],o[1],o[0],o[2],aj,ak,al,am,an,ao)
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
