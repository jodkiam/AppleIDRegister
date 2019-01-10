import xlrd
import xlwt
from datetime import datetime
from xlrd import xldate_as_tuple
import json

data = xlrd.open_workbook('/Users/kenchoi/PycharmProjects/untitled4/qmqb.xlsx')
table = data.sheet_by_name(u'Sheet4')
nrows = table.nrows
ncols = table.ncols
print(nrows,ncols)
wbk = xlwt.Workbook()
sheet = wbk.add_sheet('Sheet_handle')


with open("/Users/kenchoi/PycharmProjects/untitled4/json.txt",'r') as load_f:
    load_dict = json.load(load_f)
    # print(load_dict)
# print(load_dict)

for i in range(nrows):
    cell_cols1 = table.cell(i,0).value
    cell_cols2 = table.cell(i,1).value
    sheet.write(i, 0, cell_cols1)

    if i == 0:
        sheet.write(i,1,cell_cols2)
    else:
        flag = 0
        for ele in load_dict:
            # print(ele["distinct_id"])
            if '`' + ele["distinct_id"] == cell_cols1:
                date_handle = datetime.strptime(ele["properties"]["$signup_time"], "%Y-%m-%d %H:%M:%S")
                    # datetime(ele["properties"]["$signup_time"],0).strftime("%Y-%m-%d %H:%M:%S")
                sheet.write(i, 1, date_handle)
                print(ele["properties"]["$signup_time"])
                flag = 1
            else:
                print("")
                # continue
        if flag == 0:
            d=datetime(*xldate_as_tuple(cell_cols2,0))
            e=d.strftime("%Y-%m-%d %H:%M:%S")
            sheet.write(i,1,e)
wbk.save('/Users/kenchoi/PycharmProjects/untitled4/qmqb_handle.xlsx')