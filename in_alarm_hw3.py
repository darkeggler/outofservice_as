# _*_ coding : utf-8 _*_

import xlrd,sqlite3,time,os,json

def xldate_as_unix(xdt):
    return int((xdt-25569)*86400-28800)

def pick_id(s):
    l = ["基站编号=","基站名称=","扇区标识="]
    r = []
    for flag in l:
        p = s.find(flag)
        r.append(s[p+5 : s.find(',',p) if s.find(',',p)+1 else None] if p+1 else None)
    return r

#g_ : Generator
def g_alarmxls2(f):
    for i in range(1,4):
        b = xlrd.open_workbook(f % i)
        s = b.sheet_by_index(0) #by_name("sheet1")

        for i in range(3,s.nrows):
            info = s.cell_value(i,7)
            out = pick_id(info)
            l = [s.cell_value(i,0),
                s.cell_value(i,3),
                xldate_as_unix(s.cell_value(i,4)) if s.cell_type(i,4) == xlrd.XL_CELL_DATE else None,
                xldate_as_unix(s.cell_value(i,5)) if s.cell_type(i,5) == xlrd.XL_CELL_DATE else None,
                s.cell_value(i,6),
                info
                ]
            out.extend(l)
            out.extend(s.row_values(i,8,16))
            yield out

def g_alarmxls(f):
    b = xlrd.open_workbook(f)
    s = b.sheet_by_index(0) #by_name("sheet1")

    for i in range(3,s.nrows):
        info = s.cell_value(i,7)
        out = pick_id(info)
        l = [s.cell_value(i,0),
            s.cell_value(i,3),
            xldate_as_unix(s.cell_value(i,4)) if s.cell_type(i,4) == xlrd.XL_CELL_DATE else None,
            xldate_as_unix(s.cell_value(i,5)) if s.cell_type(i,5) == xlrd.XL_CELL_DATE else None,
            s.cell_value(i,6),
            info
            ]
        out.extend(l)
        out.extend(s.row_values(i,8,16))
        yield out

#f = r"D:\新工作空间\07.告警分析\20160118_20160124_原始告警\bsc_0%d.xls"

#for p in g_alarmxls(f):
#    print(p)

cmd = r'INSERT INTO t_raw_alarm2 VALUES(%s)' % ','.join('?'*17)
conn = sqlite3.connect(r'D:\Code\py_falcon\RedAlarm.sqlite')

dn = r"D:\新工作空间\07.告警分析\raw_data\alarm_3G"
fn = r"D:\新工作空间\07.告警分析\raw_data\readme.json"

with open(fn, "r") as fp :
    d  = json.load(fp)

for f in os.listdir(dn):
    if f not in d['alarm_3G']:
        with conn:
            t = time.time()
            print("...",f,t)
            conn.executemany(cmd,g_alarmxls(os.path.join(dn,f)))
            conn.commit()
            d['alarm_3G'][f] = t

with open(fn, "w") as fp :
    json.dump(d,fp)















#打开xls, wb > sheet > cell 
#存储为静态list {重构时考虑迭代器,yield}
#写入sql
#{重构时考虑例外}

#20160201
#注册一个函数，并使用触发器，逻辑简单
#但是 python create_function 注册的函数，在sqlite中无法使用，或许我哪里弄错了。
#妥协的办法

#20160216
'''
新增功能，原始文件自动添加，json啥的，readme.json
'''
