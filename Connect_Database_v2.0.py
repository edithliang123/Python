import pymssql
import xlrd
import datetime
from dateutil.relativedelta import relativedelta
import calendar
import sys

file_dict = dict.fromkeys(['database_name', 'file_location'])

def find_date():
    # 获取当月的第一天和最后一天
    x, y = calendar.monthrange(datetime.date.today().year, datetime.date.today().month)
    #HK&TW
    if datetime.date.today().day == y:
        sheet_month = str((datetime.date.today() + relativedelta(months=+1)).month).zfill(2)
        sheet_year = str(datetime.date.today().year)
        change_id = 1
        #如果HK&TW的月份为01月那么当前年份+1为实际修改年份
        if sheet_month == "01":
            sheet_year = str((datetime.date.today() + relativedelta(years=+1)).year)
    #CN
    else:
        sheet_month = str((datetime.date.today().month)).zfill(2)
        sheet_year = str(datetime.date.today().year)
        change_id = 0
    return sheet_year, sheet_month, change_id
    
def read_file():
    f = open(r"D:\setting_config.txt",'r')
    for line in f:
        for i in file_dict.keys():
            if i in line:
                value = (line.replace("\n", "")).split("=") #eliminate \n then split by =
                file_dict[i] = value[1]
    f.close()

def read_database():
    database_name = file_dict['database_name']
    conn = pymssql.connect('IP address','username','password',database_name)
    cur = conn.cursor()
    cur.execute('SELECT id,currencyName,currencyMask,tradeUnit,rate,companyid FROM workflow_currency1')
    row = cur.fetchone()
    while row is not None:
        print(row)
        row = cur.fetchone()
    cur.close()
    conn.close()
    
def update_database(number, change_id):
    database_name = file_dict['database_name']
    conn2 = pymssql.connect('IP address','username','password', database_name)
    cur2 = conn2.cursor()
    if (change_id == 1):#HK
        cur2.execute('''update workflow_currency1 set rate = '''+ str(number[0]) +''' where companyid=13 and id=13;
                     update workflow_currency1 set rate = '''+ str(number[1]) +'''where companyid=13 and id=14;
                     update workflow_currency1 set rate = '''+ str(number[2]) +'''where companyid=13 and id=16;
                     update workflow_currency1 set rate = '''+ str(number[3]) +'''where companyid=13 and id=17;
                     update workflow_currency1 set rate = '''+ str(number[4]) +'''where companyid=13 and id=18;
                     update workflow_currency1 set rate = '''+ str(number[5]) +'''where companyid=13 and id=20;
                     update workflow_currency1 set rate = '''+ str(number[6]) +''' where companyid=14 and id=25;
                     update workflow_currency1 set rate = '''+ str(number[7]) +'''where companyid=14 and id=26;
                     update workflow_currency1 set rate = '''+ str(number[8]) +'''where companyid=14 and id=27;
                     update workflow_currency1 set rate = '''+ str(number[9]) +'''where companyid=14 and id=29;
                     update workflow_currency1 set rate = '''+ str(number[10]) +'''where companyid=14 and id=30;
                     update workflow_currency1 set rate = '''+ str(number[11]) +'''where companyid=14 and id=32''')
    if (change_id == 0):#CN
        cur2.execute('''update workflow_currency1 set rate = '''+ str(number[0]) +''' where companyid=12 and id=2;
                     update workflow_currency1 set rate = '''+ str(number[1]) +'''where companyid=12 and id=3;
                     update workflow_currency1 set rate = '''+ str(number[2]) +'''where companyid=12 and id=4;
                     update workflow_currency1 set rate = '''+ str(number[3]) +'''where companyid=12 and id=5;
                     update workflow_currency1 set rate = '''+ str(number[4]) +'''where companyid=12 and id=6;
                     update workflow_currency1 set rate = '''+ str(number[5]) +'''where companyid=12 and id=7;
                     update workflow_currency1 set rate = '''+ str(number[6]) +'''where companyid=12 and id=8;
                     update workflow_currency1 set rate = '''+ str(number[7]) +'''where companyid=12 and id=9;
                     update workflow_currency1 set rate = '''+ str(number[8]) +'''where companyid=12 and id=10;
                     update workflow_currency1 set rate = '''+ str(number[9]) +'''where companyid=12 and id=11;
                     update workflow_currency1 set rate = '''+ str(number[10]) +'''where companyid=12 and id=12''')
    conn2.commit()
    cur2.close()
    conn2.close()
        
def read_excel(change_id, sheet_year, sheet_month):
    rate_list = [];
    file_location = file_dict['file_location']
    # 打开文件
    print("sheet year= ", sheet_month)
    if(change_id == 1):
        try:
            workBook = xlrd.open_workbook(file_location+'Y'+sheet_year+'M'+sheet_month+'.xlsx');
        except(FileNotFoundError):
            workBook = xlrd.open_workbook(file_location+'Y'+sheet_year+'M'+sheet_month+'.xls');
        except(FileNotFoundError):
            input('HK&TW Currency File Not Found')
            sys.exit()
    else:
        try:
            workBook = xlrd.open_workbook(file_location+sheet_year+sheet_month+'Exchange Rate.xlsx');
        except(FileNotFoundError):
            input('CN Currency File Not Found')
            sys.exit()
    # 获取sheet的名字
    # 获取所有sheet的名字(list类型)
    allSheetNames = workBook.sheet_names();
    #print(allSheetNames);
    if not allSheetNames:
        input("No sheet exists")
        sys.exit()

    # 按索引号获取sheet的名字（string类型）
    if(change_id == 0):
        sheetName = workBook.sheet_names()[-1];
    #print(sheet1Name);

    # 获取sheet内容
    if(change_id == 1):
        sheet1_content1 = workBook.sheet_by_name('Cross rate table')
        sheet1_content2 = workBook.sheet_by_name('Update form (OCB-TW) ')
    else:
        sheet1_content3 = workBook.sheet_by_name(sheetName)

    # 获取单元格内容(三种方式)
    if (change_id==1):
        #修改HK汇率
        rate_list.append(sheet1_content1.cell(11, 1).value);
        rate_list.append(sheet1_content1.cell(6, 1).value);
        rate_list.append(sheet1_content1.cell(10, 1).value);
        rate_list.append(sheet1_content1.cell(7, 1).value);
        rate_list.append(sheet1_content1.cell(8, 1).value);
        rate_list.append(sheet1_content1.cell(12, 1).value);
        #修改TW汇率
        rate_list.append(sheet1_content2.cell(10, 2).value);
        rate_list.append(sheet1_content2.cell(12, 2).value);
        rate_list.append(sheet1_content2.cell(5, 2).value);
        rate_list.append(sheet1_content2.cell(6, 2).value);
        rate_list.append(sheet1_content2.cell(7, 2).value);
        rate_list.append(sheet1_content2.cell(11, 2).value);
    else:
        rate_list.append(sheet1_content3.cell(8, 8).value);
        rate_list.append(sheet1_content3.cell(11, 8).value);
        rate_list.append(sheet1_content3.cell(88, 5).value);
        rate_list.append(sheet1_content3.cell(10, 8).value);
        rate_list.append(sheet1_content3.cell(80, 5).value);
        rate_list.append(sheet1_content3.cell(44, 5).value);
        rate_list.append(sheet1_content3.cell(9, 8).value);
        rate_list.append(sheet1_content3.cell(35, 5).value);
        rate_list.append(sheet1_content3.cell(85, 5).value);
        rate_list.append(sheet1_content3.cell(61, 5).value);
        rate_list.append(sheet1_content3.cell(56, 5).value);
    return rate_list;
    #print(sheet1_content1.cell_value(2, 2));
    #print(sheet1_content1.row(2)[2].value);

    # 获取单元格内容的数据类型
    #print(sheet1_content1.cell(11, 1).ctype);


if __name__ == '__main__':
    sheet_year, sheet_month, change_id = find_date()
    print("original database:")
    read_file()
    read_database()
    update_value = [];
    #读取汇率表里将要更新到数据库里的汇率
    update_value = read_excel(change_id, sheet_year, sheet_month);
    update_database(update_value,change_id);
    print("updated database:")
    read_database()
    input('Press Enter to exit...')