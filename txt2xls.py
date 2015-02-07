# coding=utf8


import xlwt


s = open("quhua.txt").read()

s = s.replace("\xef\xbb\xbf", "")
s = s.replace("\r", "")
s = s.decode("utf8").encode("gbk")


area_array = [None for i in range(9999)]
sub_array = [None for i in range(9999)]
l_arr = [None for i in range(9999)]
sub_arr = [None for i in range(9999)]

for line in s.split("\n"):
    if "var" in line:
        continue
    if u"请选择".encode("gbk") in line:
        continue
    if "=[]" in line:
        if "sub_arr[" in line:
            line = line.replace("=[]", "=[None for i in range(999999)]")
        elif "sub_array[" in line:
            line = line.replace("=[]", "=[None for i in range(9999)]")      
    print line
    exec(line)




workBook = xlwt.Workbook()
workBook.add_sheet(u"省")
workBook.add_sheet(u"市")
workBook.add_sheet(u"县")
sheet0 = workBook.get_sheet(0)
sheet1 = workBook.get_sheet(1)
sheet2 = workBook.get_sheet(2)


print u"生成省表数据".encode("gbk")

n = 0
for i in range(len(area_array)):
    area = area_array[i]
    if not area:
        continue
    print area
    sheet0.write(n, 0, str(i))
    sheet0.write(n, 1, "000")
    sheet0.write(n, 2, area.decode("gbk"))
    n += 1


print u"生成市表数据".encode("gbk")

n = 0
for i in range(len(sub_array)):
    if not sub_array[i]:
        continue
    for j in range(len(sub_array[i])):
        city = sub_array[i][j]
        if not city:
            continue
        print city
        sheet1.write(n, 0, str(j))
        sheet1.write(n, 1, city.decode("gbk"))
        n += 1


print u"生成县表数据".encode("gbk")

n = 0
for i in range(len(l_arr)):
    city = l_arr[i]
    if not city:
        continue
    print "===============", city, "==============="
    for j in range(len(sub_arr[i])):
        county = sub_arr[i][j]
        if not county:
            continue
        print county
        sheet2.write(n, 0, str(j))
        sheet2.write(n, 1, county.decode("gbk"))
        sheet2.write(n, 2, str(i))
        sheet2.write(n, 3, city.decode("gbk"))
        n += 1

workBook.save("quhua_output.xls")


raw_input("OK!")