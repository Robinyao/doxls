# coding: utf-8
import sys
reload(sys)
sys.setdefaultencoding("utf-8")
import xlrd

__author__ = 'Yoyo'

class Excels(object):

    def __init__(self, data):
        self.data = data

    def create_sql(self):
        lines = self.data.nrows
        last = lines - 1
        table_name = self.data.row_values(0)[0].strip()

        print 'CREATE TABLE %s (' % table_name
        # 循环读取除最后一行
        for i in range(2, lines - 2):
            if (self.data.row_values(i)[1].strip()[:4] == 'CHAR'):
                lenth = self.data.row_values(i)[1].strip()[4:]
                temp_type = 'VARCHAR2'
                print '  %s %s,' % (self.data.row_values(i)[0].strip(), temp_type + lenth)
            else:
                print '  %s %s,' % (self.data.row_values(i)[0].strip(), self.data.row_values(i)[1].strip())
        # 读取最后一行
        if (self.data.row_values(last)[1].strip()[:4] == 'CHAR'):
            lenth = self.data.row_values(last)[1].strip()[4:]
            temp_type = 'VARCHAR2'
            print '  %s %s' % (self.data.row_values(last)[0].strip(), temp_type + lenth)
        else:
            print '  %s %s' % (self.data.row_values(last)[0].strip(), self.data.row_values(last)[1].strip())
        print ')\n'
        # 注释语句
        for i in range(2, last):
            print "COMMENT ON COLUMN %s.%s IS '%s';" %(table_name, self.data.row_values(i)[0].strip(), self.data.row_values(i)[2].strip())

    def select_sql(self):
        lines = self.data.nrows
        last = lines - 1
        table_name = self.data.row_values(0)[0].strip()

        print 'SELECT '
        for j in range(2, lines - 2):
            print '\t%s,' % self.data.row_values(j)[0].strip()
        print '\t%s' % self.data.row_values(last)[0].strip()
        print 'FROM %s' % table_name

if __name__ == '__main__':
    infile = raw_input('Input excel file name : ')
    # 打开excel文件
    excel = xlrd.open_workbook(infile)
    # 读取第一张表
    # data1 = excel.sheets()[0]
    # excel1 = Excels(data1)
    # excel1.create_sql()
    for s in range(len(excel.sheets())):
        data = excel.sheets()[s]
        table = Excels(data)
        table.create_sql()
        print '\n'
        print '-+- ' * 25
        print '\n'
