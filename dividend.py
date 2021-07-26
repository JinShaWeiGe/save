# 研究分红

from selenium import webdriver

import time
import csv


class CSV():
    def __init__(self, name):
        self.values = csv.writer(open(name, 'w+', newline=''))

    def write(self, row):
        multi_line = row.split('\n')
        for line in multi_line:
            ans = []
            space_count = 0
            for c in line.split(' '):
                if c == '':
                    if space_count == 0:
                        ans.append(c)
                        space_count = 1
                    else:
                        space_count = 0
                else:
                    ans.append(c)
            self.values.writerow(ans)

class Browser():
    def __init__(self):
        self.browser = webdriver.Edge()
        self.path = "http://data.eastmoney.com/xg/xg/dxsyl.html"
        self.csv = CSV('test.csv')

    def click_button(self, button):
        try:
            self.browser.find_element_by_xpath(button).click()
        except:
            print('')
            time.sleep(1)

    def check_data(self, data1, data2):
        for i in range(1, 100):
            if data1 == data2:
                return True
            time.sleep(0.1)
        return False

    def open_page(self):
        self.browser.get(self.path)

    def get_data(self):
        for i in range(1, 51):
        #for i in range(1, 3):
            data = self.browser.find_element_by_xpath('//*[@id="dt_1"]/tbody').text
            self.check_data(i * 50 - 49, int(data.split(' ')[0]))
            data = self.browser.find_element_by_xpath('//*[@id="dt_1"]/tbody').text
            print(data)
            self.csv.write(data)
            self.click_button("//*[contains(text(),'下一页')]")


def main():
    b = Browser()
    b.open_page()
    b.get_data()


if __name__ == '__main__':
    main()


    '''
    f = csv.reader(open('test.csv', 'r'))
    for i in f:
        print(i)
    '''
