import os

BASE_DIR = os.path.dirname(__file__)
DATA_ROOT = os.path.join(BASE_DIR, 'data')
ERROR_DIR = os.path.join(BASE_DIR, 'errors')
PURCHASE_OUT_DIR = os.path.join(BASE_DIR, '采购导入')
SALES_OUT_DIR = os.path.join(BASE_DIR, '销售导入')

PURCHASE_FILE = "purchase_list.pickle"
SALES_FILE = "sales_list.pickle"
PRODUCT_FILE = "index_new.xls"