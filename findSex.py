# coding:utf-8
import sys,re,xlrd
from xlwt import *
reload(sys)
sys.setdefaultencoding('utf8')
from uiautomator import device as d
import unittest
import time

#打开excel
def openExcel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print (str(e))

def floatToInt(list):
    realList = []
    for value in list:
        if (type(value) == float):
            value = str(value)
            value = re.sub('\.0*$', "", value)
        value = str(value).rstrip()
        realList.append(value)
    return realList





class Mytest(unittest.TestCase):
    #初始化工作
    def setUp(self):
        print ("--------------初始化工作")
    #退出清理工作
    # def tearDown(self):
    #     print ("--------------退出清理工作")

    #测试
    def test_33(self):

        # d.screen.on()

        file='1.xlsx'
        print("check[+] %s" % file)


        data = openExcel(file)
        sheets = data.sheets()
        table=sheets[0]

        writeFile = Workbook(encoding='utf-8')
        # 指定file以utf-8的格式打开
        writeTable = writeFile.add_sheet('微信号-性别')

        for row in range(table.nrows):
            tel=floatToInt(table.row_values(row))[0]
            #tel='18283037249'
            if not d(resourceId="com.tencent.mm:id/h2").exists:
                if d(resourceId='com.tencent.mm:id/h7').exists:
                    d(resourceId='com.tencent.mm:id/h7').click()
            if  d(resourceId="com.tencent.mm:id/h2").exists:
                d(resourceId="com.tencent.mm:id/h2").clear_text()
                d(resourceId="com.tencent.mm:id/h2").set_text(tel)
                d(resourceId='com.tencent.mm:id/b20').click()

                d.watcher(tel).when(resourceId="com.tencent.mm:id/aes").when(text="确定") \
                    .click(text="确定")
                d.watchers.run()
                isTriggered=d.watcher(tel).triggered
                print(isTriggered)
                if  isTriggered:
                    d.watcher(tel).remove()
                if not isTriggered:
                    if d(resourceId='com.tencent.mm:id/agf').exists:
                        contentDescription=d(resourceId='com.tencent.mm:id/agf').info['contentDescription']
                        print("tel=%s,sex=%s"%(tel,contentDescription))
                        writeTable.write(row, 0, tel)
                        writeTable.write(row, 1, contentDescription)

                    else:
                        print("tel=%s,sex=%s" % (tel,"用户没有设置性别"))
                        writeTable.write(row, 0, tel)
                        writeTable.write(row, 1, "用户没有设置性别")
                    if d(resourceId='com.tencent.mm:id/h7').exists:
                        d(resourceId='com.tencent.mm:id/h7').click()
                else:
                    print("tel=%s,sex=%s" % (tel, "没有该用户"))
                    writeTable.write(row, 0, tel)
                    writeTable.write(row, 1, "没有该用户")
            time.sleep(3)

        writeFile.save('wx_tel_sex.xlsx')
        print ("--------------测试1")



if __name__ == '__main__':
    # from uiautomator import device as d
    #
    #
    # print(d.info)
    #
    # d.press.power()
    unittest.main()