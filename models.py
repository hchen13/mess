import re
import sys
from difflib import SequenceMatcher

import numpy as np
import xlrd

from settings import *


def show_progress(iteration, total,
                  prefix='Progress', suffix='Complete',
                  decimals=1, length=50, fill='█'):
    precision = "0:.{}f".format(decimals)
    pattern = "{{{}}}".format(precision)
    percent = pattern.format(100 * (iteration / float(total)))
    filled_length = int(length * iteration // total)
    bar = fill * filled_length + '-' * (length - filled_length)
    sys.stdout.write("\r{} |{}| {}% {}".format(prefix, bar, percent, suffix))
    sys.stdout.flush()
    if iteration == total:
        print()


class Item:
    KEY_HEADERS = {}
    REVERSE_HEADERS = {}
    VALIDATE_FIELDS = []

    def __init__(self, *args, **kwargs):
        self.potential_matches = []
        self.scanned = False
        self.unlikely = False
        for key in kwargs:
            value = str(kwargs[key]).lower()
            self.__setattr__(key, value)

    def validate(self):
        self.errors = {}
        valid = True
        for field in self.VALIDATE_FIELDS:
            if not getattr(self, field):
                valid = False
                self.errors[self.REVERSE_HEADERS[field]] = '取值为空'
            value = getattr(self, field)
            if field == 'vendor' and value == '42':
                valid = False
                self.errors[self.REVERSE_HEADERS[field]] = '错误值: error 42'
        return valid

    def is_valid(self):
        return self.validate()

    def has_serial(self):
        return self.serial is not None

    def processed(self):
        if self.scanned:
            return True
        return self.has_serial() or len(self.potential_matches) > 0

    def normalize_product_info(self):
        self.name = "" if self.name is None else self.name
        self.dose = "" if self.dose is None else self.dose
        self.manufacturer = "" if self.manufacturer is None else self.manufacturer
        self.name = self.name.lower()
        self.dose = str(self.dose).lower()
        self.manufacturer = str(self.manufacturer).lower()

        try:
            self.amount = float(self.amount)
        except ValueError as e:
            error_dict = {
                "0.4g": 300,
                "8粒": 100,
                "新乡市亚太": 10,
                "becton dickinson and": 2
            }
            self.amount = error_dict[self.amount]


    @property
    def get_product_info(self):
        return "{}{}{}".format(self.name, self.dose, self.manufacturer).lower()

    def match_product(self, product):

        def get_similarity(a, b):
            a, b = str(a), str(b)
            return SequenceMatcher(None, a, b).ratio()

        def check_similarity(similarities, threshold):
            for sim in similarities:
                if sim < threshold:
                    return False
            return True

        similarities = [0, 0, 0]
        similarities[0] = get_similarity(self.name, product.name)
        similarities[1] = get_similarity(self.dose, product.dose)
        similarities[2] = get_similarity(self.manufacturer, product.manufacturer)
        if self.dose is None:
            self.dose = ""
        if product.dose is None:
            product.dose = ""
        this = self.name + str(self.dose) + self.manufacturer
        pd = product.name + str(product.dose) + product.manufacturer
        overall = get_similarity(this, pd)

        if check_similarity(similarities, .7):
            self.serial = product.serial
            self.scanned = True
        elif overall >= .8:
            self.serial = product.serial
            self.scanned = True
        elif overall >= .5:
            self.potential_matches.append(product)
        else:
            self.unlikely = True
            return overall


class Product(Item):
    KEY_HEADERS = {
        '系统编码': 'serial',
        '品名': 'name',
        '规格': 'dose',
        '生产企业': 'manufacturer'
    }
    REVERSE_HEADERS = dict(zip(KEY_HEADERS.values(), KEY_HEADERS.keys()))
    VALIDATE_FIELDS = ['serial', 'name']

    name, manufacturer, dose, serial = None, None, None, None

    def __repr__(self):
        return "<商品 品名={} 规格={} 生产企业={}, 编号={}>".format(self.name, self.dose, self.manufacturer, self.serial)


class Purchase(Item):
    KEY_HEADERS = {
        '供应商': 'vendor',
        '品名': 'name',
        '规格': 'dose',
        '生产企业': 'manufacturer',
        '数量': 'amount',
        '供货商': 'vendor',
        '购进金额': '_price',
        '购进单价': 'unit_price',
        '购进含税金额': '_price'
    }
    REVERSE_HEADERS = dict(zip(KEY_HEADERS.values(), KEY_HEADERS.keys()))
    VALIDATE_FIELDS = ['vendor', 'name', 'amount']

    vendor, name, dose = None, None, None
    serial, manufacturer, amount = None, None, None
    year, month = None, None
    _price, unit_price = None, None

    @property
    def time(self):
        result = "{}.{:02}".format(self.year, int(self.month))
        return result

    @property
    def price(self):
        if self._price is not None:
            return self._price
        return 0

    def validate_data(self):
        self.amount = float(self.amount)
        self.unit_price = float(self.unit_price)

        if not self._price:
            self._price = self.amount * self.unit_price

    def __repr__(self):
        return "<采购 商品={} 供应商={} 单价={} 数量={}{}>".format(
            self.serial if self.serial else self.name,
            self.vendor,
            self.unit_price,
            self.amount,
            " 时间={}.{}".format(self.year, self.month) if self.year else ""
        )


class Sales(Item):
    KEY_HEADERS = {
        '商品去向': 'client',
        '品名': 'name',
        '规格': 'dose',
        '生产企业': 'manufacturer',
        '数量': 'amount',
        '含税金额': 'total_price',
        '金额': 'total_price',
    }
    REVERSE_HEADERS = dict(zip(KEY_HEADERS.values(), KEY_HEADERS.keys()))
    VALIDATE_FIELDS = ['client', 'name', 'amount', 'total_price']

    client, name, dose = None, None, None
    manufacturer, amount, total_price = None, None, None
    serial, year, month = None, None, None

    @property
    def unit_price(self):
        try:
            result = float(self.total_price) / float(self.amount)
        except ValueError:
            print(self)
            result = ""
        return result

    def validate_data(self):
        if not self.total_price:
            self.total_price = self.amount * self.unit_price

    @property
    def time(self):
        result = "{}.{:02}".format(self.year, int(self.month))
        return result

    def __repr__(self):
        return "<销售 商品={} 客户={} 数量={} 总金额={}{}>".format(
            self.serial if self.serial else self.name,
            self.client,
            self.amount,
            self.total_price,
            " 时间={}.{}".format(self.year, self.month) if self.year else ""
        )

    def show_product(self):
        print("{} {} {}".format(self.name, self.dose, self.manufacturer))


class SourceFile:
    year, month = None, None

    @property
    def time(self):
        result = "{}.{:02}".format(self.year, int(self.month))
        return result

    def __init__(self, file_path):
        if not file_path.startswith(DATA_ROOT):
            file_path = os.path.join(DATA_ROOT, file_path)
        self.file_dir, self.file_name = os.path.dirname(file_path), os.path.basename(file_path)
        # print("正在载入文件{}...".format(self.file_name))

        self.book = xlrd.open_workbook(file_path)
        self.sheet = self.book.sheet_by_index(0)
        self.error_items = []

        # print("文件载入成功!\n")

    def set_header(self, x):
        self.header_col = x
        header_line = self.sheet.row(x)
        self.header2idx = {}
        self.idx2header = {}
        for i, header in enumerate(header_line):
            header = header.value
            if not len(header):
                continue
            self.header2idx[header] = i
            self.idx2header[i] = header

    def set_time(self, year, month):
        self.year, self.month = year, month

    def extract_data(self, model_class):
        if "购进金额" in self.header2idx:
            self.items = []
            return
        rows, cols = self.sheet.nrows, self.sheet.ncols
        print("开始读取文件数据, 共有 {} 行 {} 列".format(rows, cols))
        items = []
        brokens = []
        for i in range(self.header_col + 1, rows):
            row = self.sheet.row(i)
            if self.year and self.month:
                fields = {
                    'year': self.year,
                    'month': self.month,
                    'row': i
                }
            else:
                fields = {}
            for key, val in model_class.KEY_HEADERS.items():
                if key not in self.header2idx:
                    continue
                value = row[self.header2idx[key]].value
                fields[val] = value
            item = model_class(**fields)
            if item.is_valid():
                items.append(item)
            else:
                brokens.append((i, item))
        self.items = items

        self.error_items = brokens

        print("\n读取完成, {}共生成{}组数据.\n".format(
            "检测到{}组无效数据, ".format(len(brokens)) if len(brokens) else "",
            len(items)
        ))

    def find_serial(self, products):
        total = len(products)
        no_matches = 0

        total_progress = len(self.items)
        print('\t文件中共{}条数据需处理'.format(total_progress))
        show_progress(0, total_progress)
        for i, item in enumerate(self.items):
            for product in products:
                item.match_product(product)
                if item.has_serial():
                    break
            if not item.has_serial():
                no_matches += 1
            if i % 5 == 0:
                show_progress(i + 1, total_progress)
        return total, no_matches


    def save_error(self):
        if not len(self.error_items):
            return
        print("正在记录文档{}的{}处错误信息...".format(self.file_name, len(self.error_items)))
        file_path = os.path.join(ERROR_DIR, self.file_name)
        file_path += '.txt'
        with open(file_path, 'w') as error:
            for i, item in self.error_items:
                error_message = "第{}行: {}\n".format(i+1, item.errors)
                error.write(error_message)
        print('记录完成, 结果保存于{}\n'.format(file_path))
