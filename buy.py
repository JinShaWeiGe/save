from selenium import webdriver

import time
import sys
import xlwt
import multiprocessing
import shutil
import os
import math
from multiprocessing import Process
import multiprocessing

from get_xls import XMLData


class Browser(object):
    def __init__(self, xls_name):
        self.xls_name = xls_name
        self.data_list_code = ''
        self.data_list_ex = ''
        self.data_list_name = ''
        self.out_code = []  # 不在表格内的股票

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
        if '000300.xls' in self.xls_name:
            self.get_a_data()
        elif '000300closeweight.xls' in self.xls_name:
            self.get_buy_data()
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

    def write_xls(self, data, list1, list2):
        workbook = xlwt.Workbook(encoding='ascii')
        worksheet = workbook.add_sheet('My Worksheet')
        for j in range(11):
            for i in range(len(data(0))):
                if j <= 8:
                    worksheet.write(i, j, data(j)[i])
                elif j == 9:
                    worksheet.write(i, j, list1[i])
                elif j == 10:
                    worksheet.write(i, j, list2[i])

        if not os.path.exists('result'):
            os.mkdir('result')
        workbook.save('result/test.xls')  # 保存文件

    def get_buy_data(self):
        money = '//*[@id="hq_2"]/span'
        money2 = '//*[@id="hq_9"]/span'

        xd = XMLData(self.xls_name)
        data_origin = xd.get_xls_data()
        name_origin = data_origin(5)[1:]
        code_origin = data_origin(4)[1:]
        ratio_orgin = data_origin(8)[1:]
        ex_origin = data_origin(7)[1:]

        '''汇总股数'''
        code = []
        ex = []
        number = []
        for x in os.listdir('self'):
            xd = XMLData('self/' + x)
            data = xd.get_xls_data()
            for i in range(1, len(data(4))):
                if data(4)[i] in code:
                    position = code.index(data(4)[i])
                    number[position] = str(int(number[position]) + int(data(9)[i]))
                else:
                    if data(4)[i] in code_origin:
                        code.append(data(4)[i])
                        ex.append(data(7)[i])
                        number.append(data(9)[i])
                    else:
                        if data(5)[i] not in self.out_code:
                            self.out_code.append(data(5)[i])

        all = 0
        price_list = []
        for count in range(len(code)):
            if ex[count] == 'SHH':
                exchange = 'SH'
            else:
                exchange = 'SZ'
            print(self.site + exchange + code[count])
            self.browser.get(self.site + exchange + code[count])
            money_str = self.check_text(money)
            if money_str == '-':
                money_str = self.check_text(money2)
            price = float(money_str) * int(number[count])
            price_list.append(price)
            all += price

        print('price:', all)

        out_list1 = ['']
        out_list2 = ['']

        for count in range(len(code_origin)):
            if ex_origin[count] == 'SHH':
                exchange = 'SH'
            else:
                exchange = 'SZ'

            radio = 0
            if code_origin[count] in code:
                self.browser.get(self.site + exchange + code_origin[count])
                index = code.index(code_origin[count])
                radio = price_list[index] / all * 100
            print(name_origin[count], radio, ratio_orgin[count] - radio, sep='\t')
            out_list1.append(radio)
            out_list2.append(ratio_orgin[count] - radio)

        self.write_xls(data_origin, out_list1, out_list2)
        print('未纳入股票: ')
        for i in self.out_code:
            print(i)


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
        if '000300.xls' in self.xls_name:
            self.site = "http://f10.eastmoney.com/IndustryAnalysis/Index?type=web&code="
        elif 'H50069' in self.xls_name:
            self.site = 'http://quote.eastmoney.com/hk/'
        elif '000300closeweight.xls' in self.xls_name:
            self.site = "http://f10.eastmoney.com/IndustryAnalysis/Index?type=web&code="
        self.get_data()


def main():
    xml_name = 'index/000300closeweight.xls'
    b = Browser(xml_name)
    b.run()


if __name__ == '__main__':
    main()
