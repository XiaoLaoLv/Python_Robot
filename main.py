import pandas as pd
import json

from src import Bottom5Stock
from common import ExcelHandler


if __name__ == '__main__':
    # stock_list = Bottom5Stock.pick_Bottom5Stock()
    # print(stock_list)
    account_stock = Bottom5Stock.current_price()
    print(account_stock)


