#coding=utf8

import xlrd


class Area():
    def __init__(self, id, name):
        self.id = id
        self.name = name
        self.citys = []


class City():
    def __init__(self, id, name):
        self.id = id
        self.name = name
        self.countys = []

        
class County():
    def __init__(self, id, name):
        self.id = id
        self.name = name


workbook = xlrd.open_workbook('quhua.xls')

sheet0 = workbook.sheet_by_index(0)
sheet1 = workbook.sheet_by_index(1)
sheet2 = workbook.sheet_by_index(2)

areas = []
unkown_countys = []

for i in range(sheet0.nrows):
    id = str(int(sheet0.cell(i, 0).value))
    name = sheet0.cell(i, 2).value.encode("utf8").strip()
    area = Area(id, name)
    areas.append(area)

for i in range(sheet1.nrows):
    id = str(int(sheet1.cell(i, 0).value))
    name = sheet1.cell(i, 1).value.encode("utf8").strip()
    area_id = id[:2]
    city = City(id, name)
    for area in areas:
        if area.id == area_id:
            area.citys.append(city)

for i in range(sheet2.nrows):
    id = str(int(sheet2.cell(i, 0).value))
    name = sheet2.cell(i, 1).value.encode("utf8").strip()
    city_id = str(int(sheet2.cell(i, 2).value))
    area_id = id[:2]
    county = County(id, name)
    flag = False
    for area in areas:
        if area.id == area_id:
            for city in area.citys:
                if city.id == city_id:
                    flag = True
                    city.countys.append(county)
    if not flag:
        unkown_countys.append(county)

lines = []

lines.append("var area_array=[];")
lines.append("var sub_array=[];")
lines.append('area_array[0] = "请选择";')

for area in areas:
    lines.append('area_array[%s]="%s";' % (area.id, area.name))
    lines.append('sub_array[%s]=[];' % area.id)
    lines.append('sub_array[%s][0]="请选择";' % area.id)
    for city in area.citys:
        lines.append('sub_array[%s][%s]="%s";' % (area.id, city.id, city.name))

lines.append("var l_arr=[];")
lines.append("var sub_arr=[];")

for area in areas:
    for city in area.citys:
        lines.append('l_arr[%s]="%s";' % (city.id, city.name))
        lines.append('sub_arr[%s]=[];' % city.id)
        lines.append('sub_arr[%s][0]="请选择";' % city.id)
        for county  in city.countys:
            lines.append('sub_arr[%s][%s]="%s";' % (city.id, county.id, county.name))

for county in unkown_countys:
    lines.append('sub_arr[%s][%s]="%s";' % (county.id[:4], county.id, county.name))

output = "\r\n".join(lines)
open("quhua_output.txt", "w").write(output)

print "OK!"
