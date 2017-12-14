import os
import re
import shutil
from time import sleep

import sys

import pickle

from models import SourceFile, Product, Purchase, Sales, show_progress
from settings import *

def init():

    print("清理所有错误日志...")
    shutil.rmtree(os.path.join(ERROR_DIR))
    os.mkdir(ERROR_DIR)
    print("清理完成.\n")


def save_data(data, file_name):
    with open(file_name, 'wb') as fout:
        pickle.dump(data, fout, pickle.HIGHEST_PROTOCOL)


def load_data(file_name):
    try:
        with open(file_name, 'rb') as fin:
            data = pickle.load(fin)
    except FileNotFoundError:
        return None
    print('从文件{}中读取数据成功!\n'.format(file_name))
    return data


if __name__ == '__main__':

    # ==============商品目录处理===============
    index_sheet = SourceFile('index_new.xls')
    index_sheet.set_header(0)
    index_sheet.extract_data(model_class=Product)
    products = index_sheet.items
    # index_sheet.save_error()

    sales_list = load_data('sales_list.pickle')

    if sales_list is None:

        init()

        # ==============提取数据表===============
        print("正在预处理销售数据, 请稍候...")
        purchase_sheets = []
        sales_sheets = []

        for dirpath, dirnames, filenames in os.walk(DATA_ROOT):
            dir_match = re.match(r'.*(?P<new>201\d\-\d)', dirpath)
            if not dir_match:
                continue
            for file_name in filenames:
                file_path = os.path.join(dirpath, file_name)
                if '销售' in file_path:
                    year = os.path.basename(os.path.dirname(dirpath))
                    matches = re.match(r'^\d{4}年(?P<month>\d+)', file_name)
                    month = matches.groupdict()['month']
                    sales_sheet = SourceFile(file_path)
                    sales_sheets.append((year, month, sales_sheet))
        print('预处理完成!\n')

        # ==============生成销售数据================
        print('开始生成销售数据模型...')
        for year, month, sheet in sales_sheets:
            sheet.set_header(0)
            sheet.set_time(year, month)
            sheet.extract_data(model_class=Sales)
            sheet.save_error()
        print("模型生成完成!\n")

        print("储存模型...")
        sales_list = []
        for i, (year, month, sheet) in enumerate(sales_sheets):
            sales_list += sheet.items
        save_data(sales_list, "sales_list.pickle")

    # ==============生成系统编码================
    print('开始对应销售商品系统编码...\n')
    num_sales = len(sales_list)
    no_matches = 0
    print("\t共有{}条销售数据待处理".format(num_sales))
    show_progress(0, num_sales)

    count = 0
    for i, item in enumerate(sales_list):
        if (i + 1) % 5 == 0:
            show_progress(i + 1, num_sales)
        if i + 1 > 0 and (i + 1) % 50 == 0:
            save_data(sales_list, "sales_list.pickle")

        if item.processed():
            count += 1
            continue

        for product in products:
            item.match_product(product)
            if item.has_serial():
                break
        if not item.has_serial():
            no_matches += 1

        item.scanned = True

    print()
    print(count)

    print("对应完成, 共有{}条数据 ({}%) 未找到系统编码".format(
        no_matches, no_matches / num_purchases * 100))
