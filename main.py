# -*— coding: utf-8 -*-
__author__ = 'Andliage Pox'
__date__ = '2019/01/22 几点我忘了'


import requests
import pytesseract
import xlwt
import json
from PIL import Image
from bs4 import BeautifulSoup


class Item:
    def __init__(self, tr):
        tds = tr.select('td')
        self.name = tds[3].get_text().replace('（', '(').replace('）', ')')
        self.xf = float(tds[6].get_text())
        self.jd = float(tds[7].get_text())
        self.gs = tds[12].get_text()
        self.bz = ""
        if tds[4].get_text() == "必修课" or tds[4].get_text() == "院级选修课":
            self.xz = tds[4].get_text()
        else:
            self.xz = "校级选修%s类" % tds[5].get_text()[5]
        if "体育" in tds[3].get_text():
            self.xz = "必修课"
        cj_text = tds[8].get_text()
        cj_text = cj_text.replace(".0", "")
        if cj_text.isdigit():
            self.cj = float(tds[8].get_text())
        else:
            if cj_text == "优秀":
                self.cj = 90
            elif cj_text == "良好":
                self.cj = 80
            elif cj_text == "中等":
                self.cj = 70
            elif cj_text == "及格":
                self.cj = 60
            else:
                self.cj = 0
            self.bz = "五级制：" + cj_text


def parse_secret_code():
    im = Image.open("ccimg.png")
    im = im.convert("RGBA")
    pixs = im.load()
    for y in range(im.size[1]):
        for x in range(im.size[0]):
            if ((pixs[x, y][0] > 20 or pixs[x, y][1] > 20) and pixs[x, y][2] < 100) or \
                    (pixs[x, y][0] > 100 or pixs[x, y][1] > 100):
                pixs[x, y] = (255, 255, 255, 255)
    rst = pytesseract.image_to_string(im).replace(" ", "")
    return rst[0: 4]


def build_xls(items):
    xls = xlwt.Workbook(encoding='utf-8')
    sheet = xls.add_sheet('Sheet1')
    sheet.col(0).width = 8000
    sheet.col(1).width = 4000
    sheet.col(5).width = 4000
    sheet.col(6).width = 8000
    sheet.col(8).width = 6000
    al_center = xlwt.Alignment()
    al_center.horz = xlwt.Alignment.HORZ_CENTER
    al_center.vert = xlwt.Alignment.VERT_CENTER
    al_left = xlwt.Alignment()
    al_left.horz = xlwt.Alignment.HORZ_LEFT
    al_left.vert = xlwt.Alignment.VERT_CENTER
    style_head = xlwt.easyxf('font: bold on, height 220, name 宋体')
    style_head.alignment = al_center
    style_name = xlwt.easyxf()
    style_name.alignment = al_left
    style_num_float = xlwt.easyxf('font: height 220, name 宋体', num_format_str='0.00')
    style_num_float.alignment = al_center
    style_num_int = xlwt.easyxf('font: height 220, name 宋体', num_format_str='0')
    style_num_int.alignment = al_center
    head_text = ("课程名称", "课程性质", "学分", "绩点", "成绩", "成绩备注", "课程归属")
    sum_list_text = ("校级选修A类", "校级选修B类", "校级选修C类", "校级选修D类",
                     "校级选修总学分", "院级选修总学分", "必修总学分", "总学分",
                     " ", "加权均分", "加权均绩点")
    need_list = (4, '/', 2, 2, 12, 18, 164, 194, ' ', '/', '/')
    for index in range(len(head_text)):
        sheet.write(0, index, head_text[index], style_head)
    sheet.write(0, 8, "合计列表", style_head)
    sheet.write(0, 9, "分数", style_head)
    sheet.write(0, 10, "需求", style_head)
    for index in range(len(sum_list_text)):
        sheet.write(index + 1, 8, sum_list_text[index], style_head)
    sheet.write(1, 9, xlwt.Formula('SUMIF(B2:B150,"校级选修A类",C2:C150)'), style_num_float)
    sheet.write(2, 9, xlwt.Formula('SUMIF(B2:B150,"校级选修B类",C2:C150)'), style_num_float)
    sheet.write(3, 9, xlwt.Formula('SUMIF(B2:B150,"校级选修C类",C2:C150)'), style_num_float)
    sheet.write(4, 9, xlwt.Formula('SUMIF(B2:B150,"校级选修D类",C2:C150)'), style_num_float)
    sheet.write(5, 9, xlwt.Formula('SUM(J2:J5)'), style_num_float)
    sheet.write(6, 9, xlwt.Formula('SUMIF(B2:B150,"院级选修",C2:C150)'), style_num_float)
    sheet.write(7, 9, xlwt.Formula('SUMIF(B2:B150,"必修课",C2:C150)'), style_num_float)
    sheet.write(8, 9, xlwt.Formula('SUM(C2:C150)'), style_num_float)
    sheet.write(10, 9, xlwt.Formula('SUMPRODUCT(E2:E150,C2:C150)/SUM(C2:C150)'), style_num_float)
    sheet.write(11, 9, xlwt.Formula('SUMPRODUCT(D2:D150,C2:C150)/SUM(C2:C150)'), style_num_float)
    for index in range(len(need_list)):
        sheet.write(index + 1, 10, need_list[index], style_num_float)
    for index in range(len(items)):
        sheet.write(index + 1, 0, items[index].name, style_name)
        sheet.write(index + 1, 1, items[index].xz, style_num_float)
        sheet.write(index + 1, 2, items[index].xf, style_num_float)
        sheet.write(index + 1, 3, items[index].jd, style_num_float)
        sheet.write(index + 1, 4, items[index].cj, style_num_int)
        sheet.write(index + 1, 5, items[index].bz, style_num_int)
        sheet.write(index + 1, 6, items[index].gs, style_name)
    xls.save(config["filepath"])


config = json.load(open("config.json", encoding="utf-8"))
xh = config["username"]
pswd = config["password"]
header = {'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) \
Chrome/48.0.2564.116 Safari/537.36'}
sc_try_count = 0
se = requests.session()
se.headers = header
se.get("http://202.200.112.210")
print("尝试识别验证码登录...")
while True:
    sc_try_count += 1
    ccimg = se.get("http://202.200.112.210/CheckCode.aspx")
    ccfile = open("ccimg.png", "wb")
    ccfile.write(ccimg.content)
    ccfile.close()
    cc = parse_secret_code()
    print("try: %s times:%d" % (cc, sc_try_count))
    postdata = {
        "__VIEWSTATE": "dDwtNTE2MjI4MTQ7Oz74/gDxTawfZAV831VtlWiI90NFVg==",
        "__VIEWSTATEGENERATOR": "92719903",
        "txtUserName": xh,
        "TextBox2": pswd,
        "txtSecretCode": cc,
        "Button1": ""
    }
    rs = se.post("http://202.200.112.210/default2.aspx?xh=%s" % xh, data=postdata)
    if "成绩查询" in rs.text:
        break
print("登录成功，生成xls...")
header = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) \
    Chrome/48.0.2564.116 Safari/537.36",
    "Referer": "http://202.200.112.210/xscj_gc.aspx?xh=%s" % xh
}
postdata = {
    "__VIEWSTATE": "dDwxODI2NTc3MzMwO3Q8cDxsPHhoOz47bDwzMTcwOTIxMDQ3Oz4+O\
    2w8aTwxPjs+O2w8dDw7bDxpPDE+O2k8Mz47aTw1PjtpPDc+O2k8OT47aTwxMT47aTwxMz\
    47aTwxNj47aTwyNj47aTwyNz47aTwyOD47aTwzNT47aTwzNz47aTwzOT47aTw0MT47aTw\
    0NT47PjtsPHQ8cDxwPGw8VGV4dDs+O2w85a2m5Y+377yaMzE3MDkyMTA0Nzs+Pjs+Ozs+\
    O3Q8cDxwPGw8VGV4dDs+O2w85aeT5ZCN77ya5byg5a6H5by6Oz4+Oz47Oz47dDxwPHA8b\
    DxUZXh0Oz47bDzlrabpmaLvvJrorqHnrpfmnLrnp5HlrabkuI7lt6XnqIvlrabpmaI7Pj\
    47Pjs7Pjt0PHA8cDxsPFRleHQ7PjtsPOS4k+S4mu+8mjs+Pjs+Ozs+O3Q8cDxwPGw8VGV\
    4dDs+O2w86L2v5Lu25bel56iLOz4+Oz47Oz47dDxwPHA8bDxUZXh0Oz47bDzooYzmlL/n\
    j63vvJrova/ku7YxNzI7Pj47Pjs7Pjt0PHA8cDxsPFRleHQ7PjtsPDIwMTcwOTIxOz4+O\
    z47Oz47dDx0PHA8cDxsPERhdGFUZXh0RmllbGQ7RGF0YVZhbHVlRmllbGQ7PjtsPFhOO1\
    hOOz4+Oz47dDxpPDM+O0A8XGU7MjAxOC0yMDE5OzIwMTctMjAxODs+O0A8XGU7MjAxOC0\
    yMDE5OzIwMTctMjAxODs+Pjs+Ozs+O3Q8cDw7cDxsPG9uY2xpY2s7PjtsPHdpbmRvdy5w\
    cmludCgpXDs7Pj4+Ozs+O3Q8cDw7cDxsPG9uY2xpY2s7PjtsPHdpbmRvdy5jbG9zZSgpX\
    Ds7Pj4+Ozs+O3Q8cDxwPGw8VmlzaWJsZTs+O2w8bzx0Pjs+Pjs+Ozs+O3Q8QDA8Ozs7Oz\
    s7Ozs7Oz47Oz47dDxAMDw7Ozs7Ozs7Ozs7Pjs7Pjt0PEAwPDs7Ozs7Ozs7Ozs+Ozs+O3Q\
    8O2w8aTwwPjtpPDE+O2k8Mj47aTw0Pjs+O2w8dDw7bDxpPDA+O2k8MT47PjtsPHQ8O2w8\
    aTwwPjtpPDE+Oz47bDx0PEAwPDs7Ozs7Ozs7Ozs+Ozs+O3Q8QDA8Ozs7Ozs7Ozs7Oz47O\
    z47Pj47dDw7bDxpPDA+O2k8MT47PjtsPHQ8QDA8Ozs7Ozs7Ozs7Oz47Oz47dDxAMDw7Oz\
    s7Ozs7Ozs7Pjs7Pjs+Pjs+Pjt0PDtsPGk8MD47PjtsPHQ8O2w8aTwwPjs+O2w8dDxAMDw\
    7Ozs7Ozs7Ozs7Pjs7Pjs+Pjs+Pjt0PDtsPGk8MD47aTwxPjs+O2w8dDw7bDxpPDA+Oz47\
    bDx0PEAwPHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj47Pjs7Ozs7Ozs7Ozs+Ozs+Oz4+O\
    3Q8O2w8aTwwPjs+O2w8dDxAMDxwPHA8bDxWaXNpYmxlOz47bDxvPGY+Oz4+Oz47Ozs7Oz\
    s7Ozs7Pjs7Pjs+Pjs+Pjt0PDtsPGk8MD47PjtsPHQ8O2w8aTwwPjs+O2w8dDxwPHA8bDx\
    UZXh0Oz47bDxYQVVUOz4+Oz47Oz47Pj47Pj47Pj47dDxAMDw7Ozs7Ozs7Ozs7Pjs7Pjs+\
    Pjs+Pjs+/59vDHvz5idNBQ5w4XC0o4eD2XA=",
    "__VIEWSTATEGENERATOR": "DB0F94E3",
    "ddlXN": "",
    "ddlXQ": "",
    "Button1": ""
}
rs = se.post("http://202.200.112.210/xscj_gc.aspx?xh=%s" % xh, data=postdata, headers=header)
soup = BeautifulSoup(rs.text, "html.parser")
table = soup.select("table")
items_soup = table[0].select("tr")
item_list = []
for i in range(1, len(items_soup)):
    item_list.append(Item(items_soup[i]))
build_xls(item_list)
print("excle表已成功生成")
