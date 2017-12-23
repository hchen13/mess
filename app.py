from models import show_progress
from utils import *


def display_stats():
    purchase_list, sales_list = load_ops_list()
    print("=============数据统计=============\n")
    count_has_serial, count_has_potential = 0, 0
    for i in purchase_list:
        if i.has_serial():
            count_has_serial += 1
            continue
        if len(i.potential_matches) > 1:
            count_has_potential += 1
    print("采购数据共{}项".format(len(purchase_list)))
    print("成功匹配: {}\t\t有候选: {}\t\t几乎无匹配: {}\n".format(
        count_has_serial,
        count_has_potential,
        len(purchase_list) - count_has_serial - count_has_potential
    ))

    count_has_serial, count_has_potential = 0, 0
    for i in sales_list:
        if i.has_serial():
            count_has_serial += 1
            continue
        if len(i.potential_matches) > 1:
            count_has_potential += 1
    print("销售数据共{}项".format(len(sales_list)))
    print("成功匹配: {}\t\t有候选: {}\t\t几乎无匹配: {}\n".format(
        count_has_serial,
        count_has_potential,
        len(sales_list) - count_has_serial - count_has_potential
    ))


def handle_matches():

    def report_no_match(item):
        item_class = item.__class__

        title_pattern = "{year}年{month}月{type}数据 " \
                        "{src} [品名] {name} [规格] {dose} [生产商] {manu}\n"
        params = {
            "year": item.year,
            "month": item.month,
            "name": item.name,
            "dose": item.dose,
            "manu": item.manufacturer
        }

        if item_class is Purchase:
            item_type = "采购"
            src_pattern = "[供应商] {}".format(item.vendor)
        else:
            item_type = "销售"
            src_pattern = "[商品去向] {}".format(item.client)
        params['type'] = item_type
        params['src'] = src_pattern
        title = title_pattern.format(**params)

        if not len(item.potential_matches):
            info = "\t该数据没找到相似商品\n"
            return title + info
        info = ""
        for p in item.potential_matches:
            info += '\t' + str(p) + '\n'

        return title + info

    def write_errors(ops_list, file_prefix):
        count, buff, num = 0, [], 1
        for i in ops_list:
            if i.has_serial():
                continue
            report = report_no_match(i)
            buff.append(report)
            count += 1
            if count >= 500:
                file_name = "{}{}.txt".format(file_prefix, num)
                with open(file_name, 'w') as fout:
                    for line in buff:
                        fout.write(line)
                count = 0
                buff = []
                num += 1
        file_name = "{}{}.txt".format(file_prefix, num)
        with open(file_name, 'w') as fout:
            for line in buff:
                fout.write(line)

    purchase_list, sales_list = load_ops_list()

    init()

    purchase_error_file = os.path.join(ERROR_DIR, '采购数据整理建议')
    sales_error_file = os.path.join(ERROR_DIR, '销售数据整理建议')

    write_errors(purchase_list, purchase_error_file)
    write_errors(sales_list, sales_error_file)


def match_purchases(auto_save=False):
    products = load_products()
    purchase_list, _ = load_ops_list(target='purchase')

    print("开始对应采购商品系统编码...\n")
    num_purchases = len(purchase_list)
    show_progress(0, num_purchases)
    for i, item in enumerate(purchase_list):
        if (i + 1) % 5 == 0:
            show_progress(i + 1, num_purchases)
        if auto_save and i > 0 and i % 500 == 0:
            save_data(purchase_list, PURCHASE_FILE)

        if item.has_serial():  # or len(item.potential_matches):
            continue

        item.potential_matches = []

        best_match, best_similarity = None, 0
        for product in products:
            result = item.match_product(product)
            if item.has_serial():
                break
            if result is not None and result > best_similarity:
                best_similarity = result
                best_match = product
        if not len(item.potential_matches):
            item.potential_matches.append(best_match)
    save_data(purchase_list, PURCHASE_FILE)


def match_sales(auto_save=False):
    products = load_products()
    _, sales_list = load_ops_list(target='sales')

    print('开始对应销售商品系统编码...\n')
    num_sales = len(sales_list)
    show_progress(0, num_sales)
    for i, item in enumerate(sales_list):
        if (i + 1) % 5 == 0:
            show_progress(i + 1, num_sales)
        if auto_save and i > 0 and i % 500 == 0:
            save_data(sales_list, SALES_FILE)

        if item.has_serial():  # or len(item.potential_matches):
            continue

        item.potential_matches = []

        best_match, best_similarity = None, 0
        for product in products:
            result = item.match_product(product)
            if item.has_serial():
                break
            if result is not None and result > best_similarity:
                best_similarity = result
                best_match = product
        if not len(item.potential_matches):
            item.potential_matches.append(best_match)
    save_data(sales_list, SALES_FILE)


def preprocess():
    p, s = load_ops_list(reload=True)
    for i in p:
        i.normalize_product_info()
    for i in s:
        i.normalize_product_info()
    save_data(p, PURCHASE_FILE)
    save_data(s, SALES_FILE)


if __name__ == "__main__":
    init()
    preprocess()
    match_purchases()
    match_sales()
    handle_matches()
    display_stats()
