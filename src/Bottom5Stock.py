from StockDataCollection import save_data2excel, filter_st_stock, read_datafromexcel, filter_limitup_limitdown_stock
import pandas as pd
from common import ExcelHandler
import json


def pick_Bottom5Stock():
    """
        选择市值最小的5个股票
    :return:
    """
    # 获取实时数据并保存到文件中
    save_data2excel()
    # 获取非ST的股票代码
    code_list = filter_st_stock()
    # 从文件中读取刚才爬取的数据
    stock_list = read_datafromexcel()
    # 过滤文件中的含有ST且停牌的股票
    code_list = [stock for stock in code_list if stock in stock_list.index]
    stock_list = stock_list.loc[code_list]
    # 过滤涨跌停的股票
    stock_list = filter_limitup_limitdown_stock(stock_list)
    # 选出市值最低的5只
    stock_list = stock_list.sort_values(by='market_capital', axis=0, ascending=True)
    stock_list = stock_list[['current', 'market_capital']].head(5)
    return stock_list


def current_price():
    """
    获取当前持有五只股票的最新价格
    """
    # 获取实时数据并保存到文件中
    # save_data2excel()
    # 从文件中读取刚才爬取的数据
    stock_list = read_datafromexcel()
    # 从文件中读取持股信息
    account_stock = read_account()

    # print(stock_list.loc[account_stock.index])
    # print(stock_list.loc[account_stock.index]['current'])
    account_stock['current'] = stock_list.loc[account_stock.index]['current']
    print(account_stock)



def read_account():
    """
    读取持仓信息
    """
    testRead = ExcelHandler.readExcel(readExcelPath=r'D:\test\Account.xlsx', readSheets=['all'], contentType='json')
    json_testRead = json.loads(testRead)
    # print(json_testRead["Sheet1"])

    col_code = 1
    col_in_price = 1

    for item in json_testRead["Sheet1"]:
        if item['content'] == 'code':
            col_code = item['col']

        if item['content'] == 'in_price':
            col_in_price = item['col']

    item_code_list = [item for item in json_testRead["Sheet1"] if item['col'] == col_code and item['content'] != '']
    item_in_price_list = [item for item in json_testRead["Sheet1"] if
                          item['col'] == col_in_price and item['content'] != '']
    # print(item_code_list)
    # print(item_in_price_list)

    code_list = []
    in_price_list = []
    for index in range(1, len(item_code_list)):
        if item_code_list[index]['row'] == index + 1:
            code_list.append(item_code_list[index]['content'])
        else:
            break

    for index in range(1, len(item_in_price_list)):
        if item_in_price_list[index]['row'] == index + 1:
            in_price_list.append(item_in_price_list[index]['content'])
        else:
            break

    df_account = pd.DataFrame(in_price_list, index=code_list, columns=['in_price'])
    # print(df_account)
    return df_account


