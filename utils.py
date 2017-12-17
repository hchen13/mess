import re
import pickle
from settings import *

from models import SourceFile, Purchase, Sales, Product


def file_exists(file_path):
    return os.path.exists(file_path)


def save_data(data, file_name, silent=True):
    if not silent:
        print("在文件{}中保存数据...".format(file_name))
    with open(file_name, 'wb') as fout:
        pickle.dump(data, fout, pickle.HIGHEST_PROTOCOL)
    if not silent:
        print("保存完成\n")


def load_data(file_path, silent=True):
    try:
        with open(file_path, "rb") as fin:
            data = pickle.load(fin)
    except FileNotFoundError:
        return None
    if not silent:
        print("从文件{}中读取数据成功!\n".format(file_path))
    return data


def _read_product_info():
    index_sheet = SourceFile(PRODUCT_FILE)
    index_sheet.set_header(0)
    index_sheet.extract_data(model_class=Product)
    return index_sheet.items


def load_products(reload=False):
    print("读取商品目录...")
    if not file_exists(PRODUCT_FILE) or reload:
        products = _read_product_info()
    else:
        products = load_data(PRODUCT_FILE, silent=False)
    print("读取成功!\n")
    return products


def _preprocess_sheets():

    def dir_match(path):
        match = re.match(r'.*(?P<new>201\d-\d)', path)
        if not match:
            return False
        return True

    def get_time_info(full_path):
        pattern = r'.*(?P<year>201\d)[\.|年]?(?P<month>\d+).*\.xls'
        matches = re.match(pattern, full_path)
        if matches is not None:
            r = matches.groupdict()
            return r['year'], r['month']
        print("未能成功匹配年月信息")
        return None, None

    print("正在预处理采购/销售数据...")
    purchase_sheets, sales_sheets = [], []
    for dirpath, dirnames, filenames in os.walk(DATA_ROOT):
        if not dir_match(dirpath):
            continue
        for file_name in filenames:
            file_path = os.path.join(dirpath, file_name)
            if '购进' in file_path:
                year, month = get_time_info(file_path)
                sheet = SourceFile(file_path)
                purchase_sheets.append((year, month, sheet))
            elif '销售' in file_path:
                year, month = get_time_info(file_path)
                sheet = SourceFile(file_path)
                sales_sheets.append((year, month, sheet))
    print("预处理完成!\n")
    return purchase_sheets, sales_sheets


# def load_purchases(reload=False):
#     print("正在读取采购数据...")
#     if not file_exists(PURCHASE_FILE) or reload:
#         purchase_list = _generate_ops_list(target='purchase')
#     else:
#         purchase_list = load_data(PURCHASE_FILE, silent=False)
#     print("采购数据读取成功!\n")
#     return purchase_list


# def load_sales(reload=False):
#     print("正在读取销售数据...")
#     if not file_exists(SALES_FILE) or reload:
#         sales_list = _generate_ops_list(target='sales')
#     else:
#         sales_list = load_data(SALES_FILE, silent=False)
#     print("销售数据读取成功!\n")
#     return sales_list


def load_ops_list(target='both', reload=False):
    target_texts = {
        "purchase": "采购",
        "sales": "销售",
        "both": "采购/销售"
    }
    print("正在读取{}数据...".format(target_texts[target]))
    if target == 'both':
        p, _ = load_ops_list(target='purchase', reload=reload)
        _, s = load_ops_list(target='sales', reload=reload)
        return p, s

    target_file = PURCHASE_FILE if target == 'purchase' else SALES_FILE
    if not file_exists(target_file) or reload:
        data_list = _generate_ops_list(target)
    else:
        from_file = load_data(target_file, silent=False)
        if target == 'purchase':
            data_list = from_file, []
        else:
            data_list = [], from_file
    print("{}数据读取成功!\n".format(target_texts[target]))
    return data_list


def _generate_ops_list(target='both'):
    p, s = _preprocess_sheets()
    purchase_list, sales_list = [], []

    if target == 'both' or target == 'purchase':
        print('开始生成购进数据模型...')
        for year, month, sheet in p:
            sheet.set_header(0)
            sheet.set_time(year, month)
            sheet.extract_data(model_class=Purchase)
            sheet.save_error()
            purchase_list += sheet.items
        print("模型生成完成!\n")

    if target == 'both' or target == 'sales':
        print('开始生成销售数据模型...')
        for year, month, sheet in s:
            sheet.set_header(0)
            sheet.set_time(year, month)
            sheet.extract_data(model_class=Sales)
            sheet.save_error()
            sales_list += sheet.items
        print("模型生成完成!\n")

    return purchase_list, sales_list


if __name__ == '__main__':
    pass
