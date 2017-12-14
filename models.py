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
        for key in kwargs:
            self.__setattr__(key, kwargs[key])

    def validate(self):
        self.errors = {}
        valid = True
        for field in self.VALIDATE_FIELDS:
            if not getattr(self, field):
                valid = False
                self.errors[self.REVERSE_HEADERS[field]] = '取值为空'
        return valid

    def is_valid(self):
        return self.validate()

    def has_serial(self):
        return self.serial is not None

    def processed(self):
        if self.scanned == True:
            return True
        return self.has_serial() or len(self.potential_matches) > 0

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
        average = np.average(similarities)

        if check_similarity(similarities, .7):
            self.serial = product.serial
        elif average >= .5:
            self.potential_matches.append(product)

        # print(self.name, product.name, similarities[0])
        # print(self.dose, product.dose, similarities[1])
        # print(self.manufacturer, product.manufacturer, similarities[2])
        # print(average)

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
    }
    REVERSE_HEADERS = dict(zip(KEY_HEADERS.values(), KEY_HEADERS.keys()))
    VALIDATE_FIELDS = ['vendor', 'name', 'amount']

    vendor, name, dose = None, None, None
    serial, manufacturer, amount = None, None, None
    year, month = None, None

    def __repr__(self):
        return "<采购 商品={} 供应商={} 数量={}{}>".format(
            self.serial if self.serial else self.name,
            self.vendor,
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

    # client, name, dose = None, None, None
    # manufacturer, amount, total_price = None, None, None
    # serial, year, month = None, None, None

    def __repr__(self):
        return "<销售 商品={} 客户={} 数量={} 总金额={}{}>".format(
            self.serial if self.serial else self.name,
            self.client,
            self.amount,
            self.total_price,
            " 时间={}.{}".format(self.year, self.month) if self.year else ""
        )


class SourceFile:
    year, month = None, None

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
        rows, cols = self.sheet.nrows, self.sheet.ncols
        print("开始读取文件数据, 共有 {} 行 {} 列".format(rows, cols))
        items = []
        brokens = []
        for i in range(self.header_col + 1, rows):
            row = self.sheet.row(i)
            if self.year and self.month:
                fields = {
                    'year': self.year,
                    'month': self.month
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
