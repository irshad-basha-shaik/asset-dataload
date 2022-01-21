import openpyxl
from pathlib import Path
import datetime
import requests
serial ={
    "=B4" :1,
    "=B4+1" :2
}
def getSerialValue(obj):
    try:

        if isinstance(obj, str):
            temp =obj.split("+")
            temp1 = temp[0].replace("=B","",1)
            temp2 =int(temp1)
            if isinstance(temp2, int):
                obj = (temp2-3)+int(temp[1])
    except :
        a=1
    return obj
#@csrf_exempt
def getValue(sheet):
    x=""

    if sheet!=None:
        if sheet != "Blank":
            if isinstance(sheet, str):
                sheet = sheet.replace("+", "%2B", 1)
            x = sheet
    if isinstance(x, str)  :
        return x.strip()
    return x

def getHDD(obj):
    HDD = [
        ('', ''),
        ('320 GB', '320 GB'),
        ('1 TB', '1 TB'),
        ('160 GB', '160 GB'),
        ('500 GB', '500 GB'),
        ('512 GB SSD', '512 GB SSD'),
        ('512 GB', '512 GB'),
        ('1 TB+256 SSD', '1 TB+256 SSD'),
        ('256 SSD+1 TB', '256 SSD+1 TB'),
        ('1 TB SSD', '1 TB SSD'),
        ('1 TB  SSD', '1 TB SSD'),
        ('240 GB', '240 GB'),
        ('1 TB 256 SSD', '1 TB 256 SSD'),
        ('250 GB', '250 GB'),
        ('1 TB /256 GB SSD', '1 TB /256 GB SSD'),
        ('350 GB', '350 GB'),
        ('2TB+4TB', '2TB+4TB'),
        ('1TB+256GB', '1TB+256GB'),
        ('1 TB  SSD', '1 TB  SSD'),
        ('256 SSD 1 TB', '256 SSD 1 TB'),

    ]
    for x in HDD:
        if x[0]==obj:
            return x[1]
    return obj
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
def getMSOFficeVersion(obj):
    VER=[
        ('Libre Office', ''),
        ('MS Office Standard 2010', 'MS Office Standard 2010'),
        ('Ms Office Standard 2010', 'MS Office Standard 2010'),
        ('MS Office Standard 2013', 'MS Office Standard 2013'),
        ('MS Office Standard 2016', 'MS Office Standard 2016'),
        ('ms Office Standard 2016', 'MS Office Standard 2016'),
        ('Ms Office Standard 2016', 'MS Office Standard 2016'),
        ('MS Office Standard 2019', 'MS Office Standard 2019'),
        ('Open Office', ''),
        ('MS Office 365', 'MS Office 365'),
        ('', '')
    ]
def updateMSOFFICE(obj):
    mso=(
    ('MS Office Standard 2010', 'MS Office Standard 2010'),
    ('Ms Office Standard 2010', 'MS Office Standard 2010'),

    ('MS Office Standard 2013', 'MS Office Standard 2013'),
    ('MS Office Standard 2016', 'MS Office Standard 2016'),
    ('ms Office Standard 2016', 'MS Office Standard 2016'),
    ('Ms Office Standard 2016', 'MS Office Standard 2016'),

    ('MS office standard 2013','MS Office Standard 2013'),
    ('MS Office Standard 2019', 'MS Office Standard 2019'),
    ('Libre Office', ''),
    ('MS Office Standard 2010', 'MS Office Standard 2010'),
    ('Ms Office Standard 2010', 'MS Office Standard 2010'),
    ('MS Office Standard 2013', 'MS Office Standard 2013'),
    ('MS Office Standard 2016', 'MS Office Standard 2016'),
    ('ms Office Standard 2016', 'MS Office Standard 2016'),
    ('Ms Office Standard 2016', 'MS Office Standard 2016'),
    ('MS Office Standard 2019', 'MS Office Standard 2019'),
    ('Open Office', ''),
    ('MS Office 365', 'MS Office 365'),
    ('', ''),
    (' ', ' ')

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
def getProcessor(obj):
    PROCESSOR = [
        ('Core i-3', 'Core i-3'),
        ('Core i-3 2.00 GHz', 'Core i-3 2.00 GHz'),
        ('Core i-3 2.40 GHZ', 'Core i-3 2.40 GHZ'),
        ('Core i-3 3.70 GHZ', 'Core i-3 3.70 GHZ'),
        ('Core i-3, 2.10 GHZ', 'Core i-3, 2.10 GHZ'),
        ('Core i-5', 'Core i-5'),
        ('Core i-5 1.60 Ghz', 'Core i-5 1.60 Ghz'),
        ('Core i-5 1.80 Ghz', 'Core i-5 1.80 Ghz'),
        ('Core i-5 2.19 GHZ', 'Core i-5 2.19 GHZ'),
        ('Core i-5 2.20 GHZ', 'Core i-5 2.20 GHZ'),
        ('Core i-5 2.20GHZ' , 'Core i-5 2.20 GHZ'),
        ('Core i-5 2.30 GHZ', 'Core i-5 2.30 GHZ'),
        ('Core i-5 2.40 GHZ', 'Core i-5 2.40 GHZ'),
        ('Core I-5 2.40 Ghz', 'Core i-5 2.40 GHZ'),
        ('Core i-5 2.40 Ghz', 'Core i-5 2.40 GHZ'),
        ('Core i-5 2.50 GHZ', 'Core i-5 2.50 GHZ'),
        ('Core i-5, 1.60 GHz', 'Core i-5, 1.60 GHz'),
        ('Core i-5, 2.11 GHz', 'Core i-5, 2.11 GHz'),
        ('Core i-5, 2.18 GHz', 'Core i-5, 2.18 GHz'),
        ('Core i-5, 2.20 GHz', 'Core i-5, 2.20 GHz'),
        ('Core I-5, 2.40 GHz', 'Core I-5, 2.40 GHz'),
        ('Core i-5, 2.50 GHZ', 'Core i-5, 2.50 GHZ'),
        ('Core i-5, 2.50 GHz', 'Core i-5, 2.50 GHz'),
        ('Core i-5, 2.60 GHZ', 'Core i-5, 2.60 GHZ'),
        ('Core i-5, 2.60 GHz', 'Core i-5, 2.60 GHz'),
        ('Core i-5, 2.70 GHZ', 'Core i-5, 2.70 GHZ'),
        ('Core i-5, 3.20 GHz', 'Core i-5, 3.20 GHz'),
        ('core i-5, 3.20 Ghz', 'core i-5, 3.20 Ghz'),
        ('core i-5, 3.20 GHz', 'core i-5, 3.20 GHz'),
        ('Core i-5,2.20 GHz', 'Core i-5,2.20 GHz'),
        ('core i-5,2.60 GHz', 'core i-5,2.60 GHz'),
        ('Core i-7', 'Core i-7'),
        ('Core I-7 1.80 GHz', 'Core I-7 1.80 GHz'),
        ('Core i-7 1.90 GHz', 'Core i-7 1.90 GHz'),
        ('Core I-7 1.90 GHz', 'Core I-7 1.90 GHz'),
        ('Core i-7, 2.30 GHz', 'Core i-7, 2.30 GHz'),
        ('Core i-7, 2.90 GHz', 'Core i-7, 2.90 GHz'),
        ('Core i-7,1.19 GHz', 'Core i-7,1.19 GHz'),
        ('Core i3  3.60 GHz', 'Core i3  3.60 GHz'),
        ('Core i3  3.70 GHz', 'Core i3  3.70 GHz'),
        ('Core i5  1.60GHz', 'Core i5  1.60GHz'),
        ('Core i5  2.20GHz', 'Core i5  2.20GHz'),
        ('Core i5 2.20 GHZ', 'Core i5 2.20 GHZ'),
        ('Core i5 2.20 GHz', 'Core i5 2.20 GHz'),
        ('Core i5 2.50 GHZ', 'Core i5 2.50 GHZ'),
        ('Core i5 2.60 GHz', 'Core i5 2.60 GHz'),
        ('Core i5 2.90 GHZ', 'Core i5 2.90 GHZ'),
        ('Core i5 3.2 GHZ', 'Core i5 3.2 GHZ'),
        ('Core i5 3.20 GHz', 'Core i5 3.20 GHz'),
        ('core i5 3.30 GHZ', 'core i5 3.30 GHZ'),
        ('Core i5 3.7GHz', 'Core i5 3.7GHz'),
        ('Core i5 4.1GHz', 'Core i5 4.1GHz'),
        ('Core i5, 1.80 GHz', 'Core i5, 1.80 GHz'),
        ('Core i5-2.20 GHZ', 'Core i5-2.20 GHZ'),
        ('Core i5-2.40 GHZ', 'Core i5-2.40 GHZ'),
        ('Core i5-2.50 GHZ', 'Core i5-2.50 GHZ'),
        ('Core i5-2.60 GHZ', 'Core i5-2.60 GHZ'),
        ('Core I5-6th 2.3 Ghz', 'Core I5-6th 2.3 Ghz'),
        ('Core-i-7,  2.30 GHz', 'Core-i-7,  2.30 GHz'),
        ('Core-i3 8thgen', 'Core-i3 8thgen'),
        ('Core-i5', 'Core-i5'),
        ('Core-i5  1.60GHz', 'Core-i5  1.60GHz'),
        ('Core-i5  2.20GHz', 'Core-i5  2.20GHz'),
        ('Core-i5  2.50GHz', 'Core-i5  2.50GHz'),
        ('Core-i5  2.60GHz', 'Core-i5  2.60GHz'),
        ('Core-i5  3.20GHz', 'Core-i5  3.20GHz'),
        ('Core-i5  3.90GHz', 'Core-i5  3.90GHz'),
        ('Core-i5 10th gen', 'Core-i5 10th gen'),
        ('Core-i5 11th gen', 'Core-i5 11th gen'),
        ('Core-i5 2.20GHz', 'Core-i5 2.20GHz'),
        ('Core-i5 2.30GHz', 'Core-i5 2.30GHz'),
        ('Core-i5 2.50 GHZ', 'Core-i5 2.50 GHZ'),
        ('CORE-i5 3.90 GHZ', 'CORE-i5 3.90 GHZ'),
        ('Core-i5 7thgen', 'Core-i5 7thgen'),
        ('Core-i5 8thgen', 'Core-i5 8thgen'),
        ('Core-i7  1.80GHz', 'Core-i7  1.80GHz'),
        ('Core-i7 10th gen', 'Core-i7 10th gen'),
        ('Core2dual 2.40 GHz', 'Core2dual 2.40 GHz'),
        ('Core2Dual 2.93 GHz', 'Core2Dual 2.93 GHz'),
        ('Core2Duo 2.40 Ghz', 'Core2Duo 2.40 Ghz'),
        ('Core2Duo 2.93 Ghz', 'Core2Duo 2.93 Ghz'),
        ('Core2Duo 2.93 GHz', 'Core2Duo 2.93 GHz'),
        ('Core2Duo 3.30 Ghz', 'Core2Duo 3.30 Ghz'),
        ('Core2Duo-2.40 GHZ', 'Core2Duo-2.40 GHZ'),
        ('Core2Duo-2.93 GHZ', 'Core2Duo-2.93 GHZ'),
        ('Corei-5 1.60 GHz', 'Corei-5 1.60 GHz'),
        ('Corei-5 2.50 GHz', 'Corei-5 2.50 GHz'),
        ('Corei3 3.60 GHz', 'Corei3 3.60 GHz'),
        ('Corei3 3.70GHz', 'Corei3 3.70GHz'),
        ('Corei3-3.10 GHZ', 'Corei3-3.10 GHZ'),
        ('Corei3-3.30 GHZ', 'Corei3-3.30 GHZ'),
        ('Corei3-3.7 GHZ', 'Corei3-3.7 GHZ'),
        ('Corei3-6100-3.70GHZ', 'Corei3-6100-3.70GHZ'),
        ('Corei5 3.30 GHz', 'Corei5 3.30 GHz'),
        ('Corei5-4590-3.30GHZ', 'Corei5-4590-3.30GHZ'),
        ('Corei5-6500-3.20GHZ', 'Corei5-6500-3.20GHZ'),
        ('Corei5-9500, 3.2GHZ', 'Corei5-9500, 3.2GHZ'),
        ('i-3, 2.30Ghz', 'i-3, 2.30Ghz'),
        ('I-5 3.20 Ghz', 'I-5 3.20 Ghz'),
        ('I-7, 1.90 Ghz', 'I-7, 1.90 Ghz'),
        ('i3 3.6 GHZ', 'i3 3.6 GHZ'),
        ('i3 3.9 GHZ', 'i3 3.9 GHZ'),
        ('I3-7th Gen 3.7 Ghz', 'I3-7th Gen 3.7 Ghz'),
        ('I3-7th Gen 3.9 Ghz', 'I3-7th Gen 3.9 Ghz'),
        ('I3-8th Gen 3.6 Ghz', 'I3-8th Gen 3.6 Ghz'),
        ('I5 processor', 'I5 processor'),
        ('I5, Gen 3.0 Ghz', 'I5, Gen 3.0 Ghz'),
        ('I5-7th Gen', 'I5-7th Gen'),
        ('I5-8th Gen', 'I5-8th Gen'),
        ('I5-8th Gen 3.0 Ghz', 'I5-8th Gen 3.0 Ghz'),
        ('I5-8th Gen 3.0Ghz', 'I5-8th Gen 3.0Ghz'),
        ('I7-9th Gen 3.0Ghz to 4.7Ghz', 'I7-9th Gen 3.0Ghz to 4.7Ghz'),
        ('I7-9th Gen 4.7 Ghz', 'I7-9th Gen 4.7 Ghz'),
        ('Intel Xeon Silver 4110 2.10GHZ', 'Intel Xeon Silver 4110 2.10GHZ'),
        ('InteL- 3.10 Ghz', 'InteL- 3.10 Ghz'),
        ('InteL- 3.50 Ghz', 'InteL- 3.50 Ghz'),
        ('Intel®Xeon 3.50GHZ', 'Intel®Xeon 3.50GHZ'),
        ('P Dualcore-3 GHZ', 'P Dualcore-3 GHZ')
    ]
    for x in PROCESSOR:
        if x[0]==obj:
            return x[1]
    return obj

def getOSVersion(obj):
    PROCESSOR_VERSION= [
        ('Win- 8.1 Pro 64 Bit', 'Win- 8.1 Pro 64 Bit'),
    ('Win-10 Home Single Lan.', 'Win-10 Home Single Lan.'),
    ('Win-10 Pro 64 Bit', 'Win-10 Pro 64 Bit'),
    ('Win-10 Pro 64 bit', 'Win-10 Pro 64 bit'),
    ('Win-7 Pro.32 Bit', 'Win-7 Pro.32 Bit'),
    ('win-7 Pro.32 Bit', 'Win-7 Pro.32 Bit'),
    ('Win-7 Pro.64 Bit', 'Win-7 Pro.64 Bit'),
    ('Win-8.1 Pro 32 Bit', 'Win-8.1 Pro 32 Bit'),
    ('Win-8.1 Pro 64 Bit', 'Win-8.1 Pro 64 Bit'),
    ('Win-8.1 Pro.32 Bit', 'Win-8.1 Pro.32 Bit'),
    ('Win-8.1pro 64 Bit', 'Win-8.1pro 64 Bit'),
    ('Win-Server-2012', 'Win-Server-2012'),
    ('Win-Server-2016 Std', 'Win-Server-2016 Std')
            ]
    for x in PROCESSOR_VERSION:
        if x[0]==obj:
            return x[1]
    return obj

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
        t = getHDD(getValue(sheet['N' + str(i)].value))
        u = getValue(sheet['P' + str(i)].value)
        v = getValue(sheet['R' + str(i)].value)
        aj = getOSVersion(getValue(sheet['Y' + str(i)].value))
        ak = getValue(sheet['AB' + str(i)].value)
        al = getValue(sheet['AD' + str(i)].value)
        am = getValue(sheet['AF' + str(i)].value)
        an = getValue(sheet['AH' + str(i)].value)
        ao = getValue(sheet['AK' + str(i)].value)
        ba = getValue(getSerialValue(sheet['B' + str(i)].value))
        bb = getValue(sheet['D' + str(i)].value)
        bc = getValue(sheet['G' + str(i)].value)
        bd = getValue(sheet['I' + str(i)].value)
        be = getValue(sheet['K' + str(i)].value)
        bf = getValue(sheet['M' + str(i)].value)
        bg = getValue(sheet['O' + str(i)].value)
        value = getValue(sheet['Q' + str(i)].value)
        bh = value
        bi = getProcessor(getValue(sheet['S' + str(i)].value))
        k = getdate_month_year(sheet['U' + str(i)].value)  # warranty_start_date
        l = getdate_month_year(sheet['V' + str(i)].value)  # warranty_end_date
        m = getdate_month_year(sheet['T' + str(i)].value)  # Purchase Date
        n = getdate_month_year(sheet['W' + str(i)].value)  # AMC Start Date
        o = getdate_month_year(sheet['X' + str(i)].value)  # AMC end Date
        bj = getdate_month_year(sheet['W' + str(i)].value)  # user_acceptance_date
        bk = getdate_month_year(sheet['X' + str(i)].value)  # user_handed_over_date
        #slno = [500077,500260,500976,500905,500079,500075,501646,501645,501647,502056,502057,502225,500284,500910,501196,501048,500998,500281,501046,500283,500902,500006,500878,500303,501015,500876,500007,501022,500355,500966,500325,500994,500326,500338,500329,500356,500901,500965,500963,500999,500964,500975,500993,500900,501057,501058,501059,501061,501069,501142,500063,500066,501047,501673,501674,501181,500331,501537,501538,501578,501049,501624,501627,501625,501625,501708,501707,501790,501791,501926,501927,501925,501942,900861,501943,501945]
        params = "&serial_no={0}&user_email=&asset_no={1}&usage_type={2}&gef_id_number={3}&machine_make={4}&machine_serial_no={5}&hdd_make={6}&hdd_serial_no={7}&processor={8}&warranty_start_date_month={9}&warranty_start_date_day={10}&warranty_start_date_year={11}&amc_start_date_month={12}&amc_start_date_day={13}&amc_start_date_year={14}&user_acceptance_date_month=12&user_acceptance_date_day=12&user_acceptance_date_year=1980&OS={15}&ms_office_version={16}{17}&AutoCAD={18}&Visio={19}&SAP={20}&Status={21}&user_name={22}&location={23}&emp_id={24}&machine_type={25}&domain_workgroup={26}&machine_model_no={27}&hdd={28}&hdd_model={29}&ram={30}&processor_purchase_date_month={31}&processor_purchase_date_day={32}&processor_purchase_date_year={33}&warranty_end_date_month={34}&warranty_end_date_day={35}&warranty_end_date_year={36}&amc_end_date_month={37}&amc_end_date_day={38}&amc_end_date_year={39}&user_handed_over_date_month=12&user_handed_over_date_day=12&user_handed_over_date_year=1980&Operating_System_Version={40}&ms_office={41}&Antivirus={42}&Adobe_acrobate={43}&Access={44}&Remarks={45}&Domain_User_Name=NA&SAP_User_ID=NA".format(ba,bb,bc,bd,be,bf,bg,bh,bi,k[0],k[1],k[2],n[0],n[1],n[2],b,c,d,e,f,g,h,j,a,q,r,p,s,t,u,v,m[0],m[1],m[2],l[1],l[0],l[2],o[1],o[0],o[2],aj,ak,al,am,an,ao)
        #params="&serial_no={0}&user_email=&asset_no={1}&usage_type={2}&gef_id_number={3}&machine_make={4}&machine_serial_no={5}&hdd_make={6}&hdd_serial_no={7}&processor={8}&warranty_start_date_month={9}&warranty_start_date_day={10}&warranty_start_date_year={11}&amc_start_date_month={12}&amc_start_date_day={13}&amc_start_date_year={14}&user_acceptance_date_month=&user_acceptance_date_day=&user_acceptance_date_year=&OS={15}&ms_office_version={16}&OEM_Volume={17}&AutoCAD={18}&Visio={19}&SAP={20}&Status={21}&user_name={22}&location={23}&emp_id={24}&machine_type={25}&domain_workgroup={26}&machine_model_no={27}&hdd={28}&hdd_model={29}&ram={30}&processor_purchase_date_month={31}&processor_purchase_date_day={32}&processor_purchase_date_year={33}&warranty_end_date_month={34}&warranty_end_date_day={35}&warranty_end_date_year={36}&amc_end_date_month={37}&amc_end_date_day={38}&amc_end_date_year={39}&user_handed_over_date_month=&user_handed_over_date_day=&user_handed_over_date_year=&Operating_System_Version={40}&ms_office={41}&Antivirus={42}&Adobe_acrobate={43}&Access={44}&Remarks={45}&Domain_User_Name=&SAP_User_ID=".format(ba,bb,bc,bd,be,bf,bg,bh,bi,k[0],k[1],k[2],n[0],n[1],n[2],b,c,d,e,f,g,h,j,a,q,r,p,s,t,u,v,m[0],m[1],m[2],l[1],l[0],l[2],o[1],o[0],o[2],aj,ak,al,am,an,ao)
        #params="&user_name={0}&user_email=&location={1}&asset_no={2}&serial_no={3}&emp_id={4}&usage_type={5}&machine_type={6}&gef_id_number={}&domain_workgroup={}&Domain_User_Name={}&machine_make={}&machine_model_no={}&machine_serial_no={}&hdd={}&hdd_make={}&hdd_model={}&hdd_serial_no={}&ram={}&processor={}&processor_purchase_date={}&warranty_start_date={}&warranty_end_date={}&amc_start_date={}&amc_end_date={}&user_acceptance_date={}&user_handed_over_date={}&Operating_System_Version={}&OS={}&OEM_Volume={}&ms_office={}&ms_office_version={}&Antivirus={}&AutoCAD={}&Adobe_acrobate={}&Visio={}&Access={}&SAP={}&SAP_User_ID={}&Status={}&Remarks={}"
        #params="serial_no={0}".format(ba)
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
    xlsx_file = Path(file_path, 'GEF_IT_Asset_Report_of_all_location15-06-21.xlsx')
    wb_obj = openpyxl.load_workbook(xlsx_file)
    asset_sheet=wb_obj["IT Asset's"]
    readAssetSheet(asset_sheet)
    #checkit()



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')

