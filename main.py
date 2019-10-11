#!/usr/bin/env python
# -*- coding: utf-8 -*-


import os
import json
import time
from pprint import pprint
import shutil
from datetime import datetime
import numpy as np
import pandas as pd
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import requests
import bs4


DATA_FOLDER = './data'

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36'\
                    + ' (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36 QIHU 360SE'
}


CODE_DICT = {
    '煤炭开采业': '332',
    '橡胶塑料制造业': '355',
    '黑色金属冶炼和压延加工业': '357',
    '汽车制造业': '362',
    '土木工程建筑业': '374',
    '石油和天然气开采业': '333',
    '黑色金属矿采选业': '334',
    '有色金属矿采选业': '335',
    '纺织业': '343',
    '皮革、毛皮、羽毛及其制品和制鞋业': '345',
    '造纸和纸制品业': '348',
    '石油加工、炼焦和核燃料加工业': '351',
    '化学原料和化学制品制造业': '352',
    '医药制造业': '353',
    '有色金属冶炼和压延加工业': '358'
}

COMPANY_LIST_URL_PREFIX = 'http://webstock.quote.hermes.hexun.com/a/sortlist?block='
COMPANY_LIST_URL_SUFFIX = '&callback=stocklistrequest.sortlistback&commodityid=0&title=15'\
                            + '&direction=0&start=0&number=10000&input=undefined&time=224500'\
                            + '&column=code,name,price,updownrate,LastClose,open,high,low,volume,'\
                            + 'priceweight,amount,exchangeratio,VibrationRatio,VolumeRatio'


DATES = ('2019', '2018', '2017', '2016', '2015', '2014', '2013', '2012', '2011', '2010')

COLUMNS = ['公司名称', '负债和所有者（或股东权益）合计', 
        '经营活动产生的现金流量净额', '投资活动产生的现金流量净额',
        '筹资活动产生的现金流量净', '权益负债比率',
        '总资产收益率']


def build_company_list_url(company_code):
    return COMPANY_LIST_URL_PREFIX + company_code + COMPANY_LIST_URL_SUFFIX


def request(url):
    return requests.get(url, headers=HEADERS)


def get_companies_data(url):
    """
    Returns
    -------
    company_list : list
        The element is tuple type which is (company_code, company_name)
    """
    res = request(url)
    res.encoding = 'utf-8'
    start = res.text.index('(') + 1
    end = res.text.index(')')
    companies_data = json.loads(res.text[start:end])
    return [(company[0], company[1]) for company in companies_data['Data'][0]]


def get_dateurl_dates(company_code, company_name, url_prefix):
    res = request(url_prefix + company_code + '.shtml')
    res.encoding = 'gbk'
    soup = bs4.BeautifulSoup(res.text, 'html.parser')
    script = soup.find_all(id='zaiyaocontent')[0].find_all('script')[0]
    parts = script.text.split(';')
    dateurl_parts = parts[0].split('=')[1:]
    dateurl_parts.pop()
    dateurl = eval('='.join(dateurl_parts) + '"')
    dates = eval(parts[1].split(' = ')[1])
    return dateurl, dates

def get_table(dateurl, date):
    query_url = dateurl + '=' + date
    res = request(query_url)
    res.encoding = 'gbk'
    soup = bs4.BeautifulSoup(res.text, 'html.parser')
    div = soup.find_all(id='zaiyaocontent')[0]
    return div.table.find_all('tr')

def get_zcfz(company_code, company_name):
    dateurl, dates = get_dateurl_dates(company_code, company_name, 
                                       'http://stockdata.stock.hexun.com/2009_zcfz_')
    zcfz_dict = {}
    for i in range(len(dates) - 1, -1, -1):
        table = get_table(dateurl, dates[i][0])
        zcfz = table[-2].find_all('td')[1].text
        zcfz_dict[dates[i][0]] = zcfz
    return zcfz_dict

def get_xjll(company_code, company_name, index):
    dateurl, dates = get_dateurl_dates(company_code, company_name, 
                                      'http://stockdata.stock.hexun.com/2009_xjll_')
    
    xjll_dict = {}
    for i in range(len(dates) - 1, -1, -1):
        table = get_table(dateurl, dates[i][0])
        xjll = table[index].find_all('td')[1].text
        xjll_dict[dates[i][0]] = xjll
    return xjll_dict

def get_cwbl(company_code, company_name, index):
    dateurl, dates = get_dateurl_dates(company_code, company_name, 
                                      'http://stockdata.stock.hexun.com/2009_cwbl_')
    cwbl_dict = {}
    for i in range(len(dates) - 1, -1, -1):
        table = get_table(dateurl, dates[i][0])
        cwbl = table[index].find_all('td')[1].text
        cwbl_dict[dates[i][0]] = cwbl
    return cwbl_dict


def calc_avg(date, data_dict):
    total = 0.0
    times = 1
    for dt in data_dict:
        if date in dt:
            try:
                value = float(data_dict[dt].replace(',', ''))
            except Exception as e:
                print_with_datetime(data_dict[dt].replace(',', '')\
                                    + " can't be convert to number, program will set it to 0 by default")
                continue
            total += value
            times += 1
    return str(total / times)


def print_with_datetime(msg):
    print("[%s]: %s" % (datetime.now(), msg))


def main():
    num = 0
    shutil.copyfile(os.path.join(DATA_FOLDER, 'template.xlsx'), os.path.join(DATA_FOLDER, 'data.xlsx'))
    for k, key in enumerate(CODE_DICT):
        url = build_company_list_url(CODE_DICT[key])
        company_list = get_companies_data(url)

        for i, company in enumerate(company_list):
            start_time = time.time()
            company_code = company[0]
            company_name = company[1]
            print_with_datetime("start to fetch data of company %s" % company_name)
            
            try:
                zcfz_dict = get_zcfz(company_code, company_name)
                print_with_datetime("finish fetching 负债和所有者（或股东权益）合计 of %s" % company_name)

                xjll_dict_01 = get_xjll(company_code, company_name, 13)
                print_with_datetime("finish fetching 经营活动产生的现金流量净额 of %s" % company_name)

                xjll_dict_02 = get_xjll(company_code, company_name, 28)
                print_with_datetime("finish fetching 投资活动产生的现金流量净额 of %s" % company_name)

                xjll_dict_03 = get_xjll(company_code, company_name, 40)
                print_with_datetime("finish fetching 筹资活动产生的现金流量净额 of %s" % company_name)

                cwbl_dict_01 = get_cwbl(company_code, company_name, 13)
                print_with_datetime("finish fetching 权益负债比率 of %s" % company_name)

                cwbl_dict_02 = get_cwbl(company_code, company_name, 26)
                print_with_datetime("finish fetching 总资产收益率 of %s" % company_name)
            except Exception as e:
                print_with_datetime("failed to fetch data of %s, program will skip the company" % company_name)
                continue

            for date in DATES:
                zcfz_avg = calc_avg(date, zcfz_dict)
                xjll_avg_01 = calc_avg(date, xjll_dict_01)
                xjll_avg_02 = calc_avg(date, xjll_dict_02)
                xjll_avg_03 = calc_avg(date, xjll_dict_03)
                cwbl_avg_01 = calc_avg(date, cwbl_dict_01)
                cwbl_avg_02 = calc_avg(date, cwbl_dict_02)
                df = pd.DataFrame(columns=COLUMNS)
                df.loc[0] = [
                    company_name, zcfz_avg, xjll_avg_01, 
                    xjll_avg_02, xjll_avg_03, 
                    cwbl_avg_01, cwbl_avg_02
                ]
                path = os.path.join(DATA_FOLDER, 'data.xlsx')
                with pd.ExcelWriter(path, \
                                    engine='openpyxl', mode='a') as writer:
                    book = load_workbook(path)
                    writer.book = book
                    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                    df.to_excel(writer, sheet_name=date, index=False, startrow=num, header=False)
            num += 1
            end_time = time.time()
            print_with_datetime("finish saving data of %s, cost: %.2f, line: %s, index: %s, category: %s"\
                    % (company_name, end_time - start_time, num, i, key))


if __name__ == "__main__":
    main()
