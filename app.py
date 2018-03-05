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

    def write_errors(ops_list, error_type, file_prefix):
        count, buff, num = 0, [], 1
        for i in ops_list:
            if i.has_serial():
                continue
            buff.append(i)
            count += 1
            if count >= 500:
                file_name = "{}{}.xlsx".format(file_prefix, num)
                write_match_errors(buff, file_name)
                count = 0
                buff = []
                num += 1

        file_name = "{}{}.xlsx".format(file_prefix, num)
        write_match_errors(buff, error_type, file_name)

    purchase_list, sales_list = load_ops_list()

    init()

    purchase_error_file = os.path.join(ERROR_DIR, '采购数据整理建议')
    sales_error_file = os.path.join(ERROR_DIR, '销售数据整理建议')

    write_errors(purchase_list, '采购', purchase_error_file)
    write_errors(sales_list, '销售', sales_error_file)


def match_purchases(auto_save=False):
    products = load_products()
    purchase_list, _ = load_ops_list(target='purchase')

    print("开始对应采购商品系统编码, 共{}项纪录...\n".format(len(purchase_list)))
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

    print('开始对应销售商品系统编码, 共{}项纪录...\n'.format(len(sales_list)))
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


def preprocess(save=False):
    p, s = load_ops_list(reload=False)
    for i in p:
        i.normalize_product_info()
        i.validate_data()
    for i in s:
        i.normalize_product_info()
        i.validate_data()
    if save:
        save_data(p, PURCHASE_FILE)
        save_data(s, SALES_FILE)


def match_serial_main():
    init()
    preprocess()
    match_purchases()
    match_sales()
    handle_matches()
    display_stats()


def dump_sheet_main():
    preprocess(save=True)
    p, s = load_ops_list(target='both', reload=False)
    p = list(filter(lambda x: x.has_serial(), p))
    s = list(filter(lambda x: x.has_serial(), s))
    s = list(filter(lambda x: float(x.total_price) >= 0, s))
    p = list(filter(lambda x: isinstance(x.vendor, str), p))
    s = list(filter(lambda x: isinstance(x.client, str), s))
    clear_output_dirs()
    batch_purchases(p)
    group_sales(s)


def test():
    products = load_products(reload=False)

    plist, slist = load_ops_list('both', reload=True)
    save_data(plist, PURCHASE_FILE)
    save_data(slist, SALES_FILE)



def divide_large_files(dir):
    files = os.listdir(dir)
    docs = list( filter( is_excel, files ) )


if __name__ == "__main__":
    match_serial_main()
    # dump_sheet_main()
    # display_stats()
    # divide_large_files(SALES_OUT_DIR)