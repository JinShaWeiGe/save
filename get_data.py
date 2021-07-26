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


class Browser(object):
    def __init__(self, xls_name):
        self.xls_name = xls_name
        self.workbook = ''
        self.sheet = ''
        self.data_list_code = ''
        self.data_list_ex = ''
        self.data_list_name = ''

    def get_xls(self):
        if os.path.exists("C:\\Users\\ZHUWEI\\Downloads\\" + self.xls_name):
            os.remove("C:\\Users\\ZHUWEI\\Downloads\\" + self.xls_name)
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
        self.workbook = open_workbook(self.xls_name)
        self.sheet = self.workbook.sheet_by_index(0)
        self.data_list_code = self.sheet.col_values(4)[1:]
        self.data_list_ex = self.sheet.col_values(7)[1:]
        self.data_list_name = self.sheet.col_values(5)[1:]

    def check_text(self, button):
        count = 0
        ans = ''
        while count < 10:
            try:
                ans = self.browser.find_element_by_xpath(button).text
            except:
                time.sleep(1)
            count += 1
            if ans != '':
                return ans
        return ''

    def get_data(self):
        if '000300' in self.xls_name:
            self.get_a_data()
        else:
            self.get_hk_data()

    def get_a_data(self):
        PEG = '//*[@id="Table2"]/tbody/tr[3]/td[4]'
        PE = '//*[@id="Table2"]/tbody/tr[3]/td[6]'
        ROE = '//*[@id="Table3"]/tbody/tr[3]/td[4]'
        for count in range(len(self.data_list_code)):
            if self.data_list_ex[count] == 'SHH':
                exchange = 'SH'
            else:
                exchange = 'SZ'
            self.browser.get(self.site + exchange + self.data_list_code[count])
            print(self.data_list_code[count],
                  self.check_text(PEG),
                  self.check_text(PE),
                  self.check_text(ROE),
                  self.data_list_name[count],
                  sep='\t')

    def get_hk_data(self):
        # http://quote.eastmoney.com/hk/00001.html
        devided = '//*[@id="base_info"]/ul[3]/li[5]/i'
        pe = '//*[@id="base_info"]/ul[4]/li[3]/i'
        roe = '//*[@id="financial_analysis"]/tbody/tr[1]/td[10]'
        A_site = '//*[@id="quote_brief"]/div[3]/div/p[1]/a'
        AH_price = '//*[@id="quote_brief"]/div[3]/div/div[1]/span'

        for count in range(len(self.data_list_code)):
            self.browser.get(self.site + self.data_list_code[count].zfill(5) + '.html')
            if 'A股行情' not in self.check_text(A_site):
                print(self.data_list_code[count],
                      self.check_text(devided),
                      self.check_text(pe),
                      self.check_text(roe),
                      '\t',
                      self.data_list_name[count],
                      sep='\t')
            else:
                print(self.data_list_code[count],
                      self.check_text(devided),
                      self.check_text(pe),
                      self.check_text(roe),
                      self.check_text(AH_price),
                      self.data_list_name[count],
                      sep='\t')

            count += 1

        print('end')

    def run(self):
        self.browser = webdriver.Edge()
        if '000300' in self.xls_name:
            self.site = "http://f10.eastmoney.com/IndustryAnalysis/Index?type=web&code="
            self.xls_site = "http://www.csindex.com.cn/uploads/file/autofile/cons/000300cons.xls?t=1609930653"
            self.get_xls()
        elif 'H50069' in self.xls_name:
            self.site = 'http://quote.eastmoney.com/hk/'
            self.xls_site = "http://www.csindex.com.cn/uploads/file/autofile/cons/H50069cons.xls?t=1611478182"
            self.get_xls()
        self.xls_name = 'index/' + self.xls_name
        self.get_xls_data()
        self.get_data()


def main():
    xml_name = '000300cons.xls'
    b = Browser(xml_name)
    b.run()

if __name__ == '__main__':
    main()
