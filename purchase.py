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


def load_products(reload=False):
    if reload:
        index_sheet = SourceFile('index_new.xls')
        index_sheet.set_header(0)
        index_sheet.extract_data(model_class=Product)
        products = index_sheet.items
        save_data(products, "product_list.pickle")
    else:
        product = load_data('product_list.pickle')


    products = load_data('product_list.pickle')
    if not products:
        index_sheet = SourceFile('index_new.xls')
        index_sheet.set_header(0)
        index_sheet.extract_data(model_class=Product)
        products = index_sheet.items
    return products


if __name__ == '__main__':

    # ==============商品目录处理===============
    index_sheet = SourceFile('index_new.xls')
    index_sheet.set_header(0)
    index_sheet.extract_data(model_class=Product)
    products = index_sheet.items
    # index_sheet.save_error()

    purchase_list = load_data('purchase_list.pickle')

    if purchase_list is None:

        init()

        # ==============提取数据表===============
        print("正在预处理购进/销售数据, 请稍候...")
        purchase_sheets = []
        sales_sheets = []

        for dirpath, dirnames, filenames in os.walk(DATA_ROOT):
            dir_match = re.match(r'.*(?P<new>201\d\-\d)', dirpath)
            if not dir_match:
                continue
            for file_name in filenames:
                file_path = os.path.join(dirpath, file_name)
                if '购进' in file_path:
                    year = os.path.basename(os.path.dirname(dirpath))
                    matches = re.match(r'^\d{4}\.?(?P<month>\d+)', file_name)
                    month = matches.groupdict()['month']
                    purchase_sheet = SourceFile(file_path)
                    purchase_sheets.append((year, month, purchase_sheet))
                elif '销售' in file_path:
                    year = os.path.basename(os.path.dirname(dirpath))
                    matches = re.match(r'^\d{4}年(?P<month>\d+)', file_name)
                    month = matches.groupdict()['month']
                    sales_sheet = SourceFile(file_path)
                    sales_sheets.append((year, month, sales_sheet))
        print('预处理完成!\n')

        # ==============生成购进数据================
        print('开始生成购进数据模型...')
        for year, month, sheet in purchase_sheets:
            sheet.set_header(0)
            sheet.set_time(year, month)
            sheet.extract_data(model_class=Purchase)
            sheet.save_error()
        print("模型生成完成!\n")

        # ==============生成销售数据================
        print('开始生成销售数据模型...')
        for year, month, sheet in sales_sheets:
            sheet.set_header(0)
            sheet.set_time(year, month)
            sheet.extract_data(model_class=Sales)
            sheet.save_error()
        print("模型生成完成!\n")

        print("储存模型...")
        purchase_list = []
        for i, (year, month, sheet) in enumerate(purchase_sheets):
            purchase_list += sheet.items
        save_data(purchase_list, "purchase_list.pickle")

    # ==============生成系统编码================
    print('开始对应采购商品系统编码...\n')
    num_purchases = len(purchase_list)
    no_matches = 0
    print("\t共有{}条采购数据待处理".format(num_purchases))
    show_progress(0, num_purchases)

    count = 0
    for i, item in enumerate(purchase_list):
        if (i + 1) % 5 == 0:
            show_progress(i + 1, num_purchases)
        if i + 1 > 0 and (i + 1) % 500 == 0:
            save_data(purchase_list, "purchase_list.pickle")

        if item.processed():
            count += 1
            continue

        item.scanned = True

        for product in products:
            item.match_product(product)
            if item.has_serial():
                break
        if not item.has_serial():
            no_matches += 1

    print()
    print(count)

    print("对应完成, 共有{}条数据 ({}%) 未找到系统编码".format(
        no_matches, no_matches / num_purchases * 100))
