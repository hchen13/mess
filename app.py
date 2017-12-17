from models import show_progress
from utils import *


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

    purchase_list, sales_list = load_ops_list()

    purchase_error_file = os.path.join(ERROR_DIR, '采购数据整理建议.txt')
    sales_error_file = os.path.join(ERROR_DIR, '销售数据整理建议.txt')

    with open(purchase_error_file, 'w') as fout:
        iter, tn = 0, len(purchase_list)
        show_progress(0, tn)
        for i in purchase_list:
            report = report_no_match(i)
            fout.write(report)
            iter += 1
            if not (iter % 50):
                show_progress(iter, tn)

    with open(sales_error_file, 'w') as fout:

        iter, tn = 0, len(sales_list)
        show_progress(0, tn)
        for i in sales_list:
            report = report_no_match(i)
            fout.write(report)
            iter += 1
            if not (iter % 50):
                show_progress(iter, tn)


if __name__ == "__main__":
    handle_matches()
