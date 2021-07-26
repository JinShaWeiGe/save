from selenium import webdriver

import time
import sys
from xlrd import open_workbook  # xlrd用于读取xld
import xlwt
import multiprocessing
import shutil
import os
import math
from multiprocessing import Process
import multiprocessing


class XMLData(object):
    def __init__(self, xls_name):
        self.path_name = xls_name
        self.xls_name = xls_name.split('/')[-1]
        self.workbook = ''
        self.sheet = ''
        self.data_list_code = ''
        self.data_list_ex = ''
        self.data_list_name = ''

        self.browser = webdriver.Edge()
        self.site_dic = {'000300closeweight.xls': 'http://www.csindex.com.cn/uploads/file/autofile/closeweight/000300closeweight.xls?t=1614850899',
                         '000300cons.xls': 'http://www.csindex.com.cn/uploads/file/autofile/cons/000300cons.xls?t=1609930653'}

    def get_xls(self):
        if os.path.exists("C:\\Users\\ZHUWEI\\Downloads\\" + self.xls_name):
            os.remove("C:\\Users\\ZHUWEI\\Downloads\\" + self.xls_name)
        self.xls_site = self.site_dic[self.xls_name]
        self.browser.get(self.xls_site)
        for i in range(20):
            if os.path.exists("C:\\Users\\ZHUWEI\\Downloads\\" + self.xls_name):
                shutil.move("C:\\Users\\ZHUWEI\\Downloads\\" + self.xls_name, 'index/' + self.xls_name)
                break
            else:
                time.sleep(1)
                print('wait xls')
        #self.browser.quit()

    def get_xls_data(self):
        if 'index' in self.path_name:
            self.get_xls()
        self.workbook = open_workbook(self.path_name)
        self.sheet = self.workbook.sheet_by_index(0)
        return self.sheet.col_values


if __name__ == '__main__':
    xd = XMLData('index' + '/' + '000300cons.xls')
    data = xd.get_xls_data()
