"""
    从雪球爬取实时数据并进行数据整理
"""
import json
import requests
import pandas as pd
import tushare as ts
import config.settings as config

pro = ts.pro_api(config.token)


def get_data_from_xueqiu(page_id):
    """
    从雪球获取实时数据
    :return:
    """
    headers = {
        'host': 'xueqiu.com',
        'Referer': 'https://xueqiu.com/hq',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest'
    }

    # 雪球Ajax获取翻页数据
    url = 'https://xueqiu.com/service/v5/stock/screener/quote/list?page=%s&size=90&order=asc&orderby=code&order_by=symbol&market=CN&type=sh_sz&_=1632365438642' % page_id
    res = requests.get(url, headers=headers)

    # 将数据转换成DataFrame
    obj_json = json.loads(res.text)
    stock_list = obj_json['data']['list']
    stock_list = pd.DataFrame(stock_list)
    return stock_list


def save_data2excel():
    """
    读取数据并将数据保存到excel
    :return:
    """
    stock_list = pd.DataFrame()
    for page_i in range(1, 52):
        result = get_data_from_xueqiu(page_i)
        stock_list = stock_list.append(result)

    stock_list = stock_list.set_index('symbol', drop=True)
    stock_list.to_excel(config.filePath)
    # print(stock_list)


def filter_st_stock():
    """
        筛选非ST，非*ST
    """
    columns = ['ts_code', 'symbol', 'name']
    stock_list = pro.query('stock_basic', exchange='', list_status='L', fields=columns, market='主板')
    stock_list = stock_list[~stock_list['name'].str.contains('ST')]
    code_list = stock_list['symbol']
    return code_list


def read_datafromexcel():
    """
        从本地excel文件中读取数据
    :return:
    """
    # file = r'D:/test/Result.xlsx'
    df = pd.DataFrame(pd.read_excel(config.filePath))
    df['symbol'] = df['symbol'].apply(lambda x: x[2:])
    stock_list = df.set_index('symbol', drop=True)
    return stock_list


def filter_limitup_limitdown_stock(stock_list):
    """
        过滤涨停和跌停的股票
    :return:
    """
    stock_list = stock_list[stock_list['percent'].notna()]
    stock_list = stock_list[stock_list['percent'] < 9]
    stock_list = stock_list[stock_list['percent'] > -9]

    return stock_list


# def pick_Bottom5Stock():
#     """
#         选择市值最小的5个股票
#     :return:
#     """
#     save_data2excel()
#     code_list = filter_st_stock()
#     stock_list = read_datafromexcel()
#     code_list = [stock for stock in code_list if stock in stock_list.index]
#     stock_list = stock_list.loc[code_list]
#     stock_list = filter_limitup_limitdown_stock(stock_list)
#     stock_list = stock_list.sort_values(by='float_market_capital', axis=0, ascending=True)
#     print(stock_list[['current', 'float_market_capital']].head(5))

