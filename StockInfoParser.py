import io
import os
import re
import sys
import time
import pandas
import datetime
import requests
import mplfinance
from matplotlib import dates

# Basic Data
file_name = __file__[:-3]
absolute_path = os.path.dirname(os.path.abspath(__file__))


# <editor-fold desc='common'>
def load_json_config():
    global file_directory
    config_file = os.path.join(os.sep, absolute_path, 'Config.cfg')
    with open(config_file, 'r') as file_handler:
        config_data = file_handler.read()
    regex = 'FILE_DIRECTORY=.*'
    match = re.findall(regex, config_data)
    file_directory = match[0].split('=')[1].strip()


# </editor-fold>


# <editor-fold desc='daily update'>


def save_dict_to_file(dic, txt):
    f = open(txt, 'w', encoding='utf-8')
    f.write(dic)
    f.close()


def load_dict_from_file(txt):
    f = open(txt, 'r', encoding='utf-8')
    data = f.read()
    f.close()
    return eval(data)


def crawl_price(date=datetime.datetime.now()):
    date_str = str(date).split(' ')[0].replace('-', '')
    r = requests.post('http://www.twse.com.tw/exchangeReport/MI_INDEX?response=csv&date=' + date_str + '&type=ALL')
    ret = pandas.read_csv(io.StringIO('\n'.join([i.translate({ord(c): None for c in ' '}) for i in r.text.split('\n') if
                                                 len(i.split(',')) == 17 and i[0] != '='])), header=0,
                          index_col='證券代號')
    ret['成交金額'] = ret['成交金額'].str.replace(',', '')
    ret['成交股數'] = ret['成交股數'].str.replace(',', '')
    return ret


def original_crawl_price(date='2011-01-01 00:00:00'):
    print('Begin: original_crawl_price!')
    data = {}
    success = False
    dateFormatter = '%Y-%m-%d %H:%M:%S'
    date = datetime.datetime.strptime(date, dateFormatter)
    while not success:
        print('parsing', date)
        try:
            data[date.date()] = crawl_price(date)
            print('success!')
            success = True
        except pandas.errors.EmptyDataError:
            # 假日爬不到
            print('fail! check the date is holiday')
        # 減一天
        date += datetime.timedelta(days=1)
        time.sleep(10)
    writer = pandas.ExcelWriter(stock_file_path, engine='xlsxwriter')
    stock_volume = pandas.DataFrame({k: d['成交股數'] for k, d in data.items()}).transpose()
    stock_volume.index = pandas.to_datetime(stock_volume.index)
    stock_volume.to_excel(writer, sheet_name='stock_volume', index=True)
    stock_open = pandas.DataFrame({k: d['開盤價'] for k, d in data.items()}).transpose()
    stock_open.index = pandas.to_datetime(stock_open.index)
    stock_open.to_excel(writer, sheet_name='stock_open', index=True)
    stock_close = pandas.DataFrame({k: d['收盤價'] for k, d in data.items()}).transpose()
    stock_close.index = pandas.to_datetime(stock_close.index)
    stock_close.to_excel(writer, sheet_name='stock_close', index=True)
    stock_high = pandas.DataFrame({k: d['最高價'] for k, d in data.items()}).transpose()
    stock_high.index = pandas.to_datetime(stock_high.index)
    stock_high.to_excel(writer, sheet_name='stock_high', index=True)
    stock_low = pandas.DataFrame({k: d['最低價'] for k, d in data.items()}).transpose()
    stock_low.index = pandas.to_datetime(stock_low.index)
    stock_low.to_excel(writer, sheet_name='stock_low', index=True)
    writer.save()
    print('End: original_crawl_price!')


def update_stock_info():
    print('Begin: update_stock_info!')
    data = {}
    count = 1
    fail_count = 0
    allow_continuous_fail_count = 20
    try:
        pandas.read_excel(stock_file_path, sheet_name='stock_volume', index_col=0)
        print(r'{} Exist.'.format(stock_file_path))
    except FileNotFoundError:
        print(r'{} Not Exist.'.format(stock_file_path))
        original_crawl_price()
    stock_volume_old = pandas.read_excel(stock_file_path, sheet_name='stock_volume', index_col=0)
    stock_volume_old.index = pandas.to_datetime(stock_volume_old.index)
    stock_open_old = pandas.read_excel(stock_file_path, sheet_name='stock_open', index_col=0)
    stock_open_old.index = pandas.to_datetime(stock_open_old.index)
    stock_close_old = pandas.read_excel(stock_file_path, sheet_name='stock_close', index_col=0)
    stock_close_old.index = pandas.to_datetime(stock_close_old.index)
    stock_high_old = pandas.read_excel(stock_file_path, sheet_name='stock_high', index_col=0)
    stock_high_old.index = pandas.to_datetime(stock_high_old.index)
    stock_low_old = pandas.read_excel(stock_file_path, sheet_name='stock_low', index_col=0)
    stock_low_old.index = pandas.to_datetime(stock_low_old.index)
    last_date = stock_volume_old.index[-1]
    dateFormatter = '%Y-%m-%d %H:%M:%S'
    date = datetime.datetime.strptime(str(last_date), dateFormatter)
    date += datetime.timedelta(days=1)
    if date > datetime.datetime.now():
        print('Finish update_stock_info!')
        sys.exit(0)
    while date < datetime.datetime.now() and count <= 100:
        print('parsing', date)
        try:
            data[date.date()] = crawl_price(date)
            print('success {} times!'.format(count))
            fail_count = 0
            count += 1
        except pandas.errors.EmptyDataError:
            # 假日爬不到
            print('fail! check the date is holiday')
            fail_count += 1
            if fail_count == allow_continuous_fail_count:
                raise
        date += datetime.timedelta(days=1)
        time.sleep(10)
    writer = pandas.ExcelWriter(stock_file_path, engine='xlsxwriter')

    stock_volume_new = pandas.DataFrame({k: d['成交股數'] for k, d in data.items()}).transpose()
    stock_volume_new.index = pandas.to_datetime(stock_volume_new.index)
    stock_volume = pandas.concat([stock_volume_old, stock_volume_new], join='outer')
    stock_volume.to_excel(writer, sheet_name='stock_volume', index=True)

    stock_open_new = pandas.DataFrame({k: d['開盤價'] for k, d in data.items()}).transpose()
    stock_open_new.index = pandas.to_datetime(stock_open_new.index)
    stock_open = pandas.concat([stock_open_old, stock_open_new], join='outer')
    stock_open.to_excel(writer, sheet_name='stock_open', index=True)

    stock_close_new = pandas.DataFrame({k: d['收盤價'] for k, d in data.items()}).transpose()
    stock_close_new.index = pandas.to_datetime(stock_close_new.index)
    stock_close = pandas.concat([stock_close_old, stock_close_new], join='outer')
    stock_close.to_excel(writer, sheet_name='stock_close', index=True)

    stock_high_new = pandas.DataFrame({k: d['最高價'] for k, d in data.items()}).transpose()
    stock_high_new.index = pandas.to_datetime(stock_high_new.index)
    stock_high = pandas.concat([stock_high_old, stock_high_new], join='outer')
    stock_high.to_excel(writer, sheet_name='stock_high', index=True)

    stock_low_new = pandas.DataFrame({k: d['最低價'] for k, d in data.items()}).transpose()
    stock_low_new.index = pandas.to_datetime(stock_low_new.index)
    stock_low = pandas.concat([stock_low_old, stock_low_new], join='outer')
    stock_low.to_excel(writer, sheet_name='stock_low', index=True)

    writer.save()
    print('End: update_stock_info!')
    if date > datetime.datetime.now():
        print('Finish update_stock_info!')
        sys.exit(0)
    update_stock_info()


def xlsx_to_csv_pd():
    data_xls = pandas.read_excel(stock_file_path, sheet_name='stock_volume', index_col=0, engine='openpyxl')
    data_xls.to_csv('C:/Users/ch032/Desktop/project/stock/stock_volume.csv', encoding='utf-8')
    data_xls = pandas.read_excel(stock_file_path, sheet_name='stock_open', index_col=0, engine='openpyxl')
    data_xls.to_csv('C:/Users/ch032/Desktop/project/stock/stock_open.csv', encoding='utf-8')
    data_xls = pandas.read_excel(stock_file_path, sheet_name='stock_close', index_col=0, engine='openpyxl')
    data_xls.to_csv('C:/Users/ch032/Desktop/project/stock/stock_close.csv', encoding='utf-8')
    data_xls = pandas.read_excel(stock_file_path, sheet_name='stock_high', index_col=0, engine='openpyxl')
    data_xls.to_csv('C:/Users/ch032/Desktop/project/stock/stock_high.csv', encoding='utf-8')
    data_xls = pandas.read_excel(stock_file_path, sheet_name='stock_low', index_col=0, engine='openpyxl')
    data_xls.to_csv('C:/Users/ch032/Desktop/project/stock/stock_low.csv', encoding='utf-8')


# </editor-fold>


def show_stock_data(stock_no='2330', select_type='month'):
    stock_volume = pandas.read_csv(volume_file_path, index_col=0, low_memory=False)
    stock_volume.index = pandas.to_datetime(stock_volume.index)
    stock_open = pandas.read_csv(open_file_path, index_col=0, low_memory=False)
    stock_open.index = pandas.to_datetime(stock_open.index)
    stock_close = pandas.read_csv(close_file_path, index_col=0, low_memory=False)
    stock_close.index = pandas.to_datetime(stock_close.index)
    stock_high = pandas.read_csv(high_file_path, index_col=0, low_memory=False)
    stock_high.index = pandas.to_datetime(stock_high.index)
    stock_low = pandas.read_csv(low_file_path, index_col=0, low_memory=False)
    stock_low.index = pandas.to_datetime(stock_low.index)
    date_list = list()

    if select_type == 'month':
        numbers = -20 * 3
    elif select_type == 'year':
        numbers = -240 * 3
    else:
        numbers = 0

    for date_stamp in stock_volume.index.tolist()[numbers:-1]:
        date_list.append(dates.date2num(date_stamp))
    stock_data_list = {
        'Open': stock_open[stock_no].dropna().astype(float).tolist()[numbers:-1],
        'Close': stock_close[stock_no].dropna().astype(float).tolist()[numbers:-1],
        'High': stock_high[stock_no].dropna().astype(float).tolist()[numbers:-1],
        'Low': stock_low[stock_no].dropna().astype(float).tolist()[numbers:-1],
        'Volume': stock_volume[stock_no].dropna().astype(float).tolist()[numbers:-1],
    }
    df1 = pandas.DataFrame(stock_data_list)
    df1.index = stock_volume.index[numbers:-1]
    df1.index.name = 'Date'
    market_colors = mplfinance.make_marketcolors(up='red', down='green', edge='i', wick='i', volume='in', inherit=True)
    style = mplfinance.make_mpf_style(base_mpf_style='yahoo', y_on_right=False, marketcolors=market_colors)
    kwargs = dict(type='candle', mav=(5, 20, 60), volume=True, title='Stock {}'.format(stock_no), ylabel='Price',
                  ylabel_lower='Volume', warn_too_much_data=10000, style=style)
    mplfinance.plot(df1, **kwargs)


# get global variable
file_directory = str()
load_json_config()
if not file_directory:
    print('file_directory is empty.')
    sys.exit(0)

# daily update file/directory
stock_file_path = file_directory + 'stock.xlsx'

# .csv file [volume, open, close, high, low]
volume_file_path = file_directory + 'volume.csv'
open_file_path = file_directory + 'open.csv'
close_file_path = file_directory + 'close.csv'
high_file_path = file_directory + 'high.csv'
low_file_path = file_directory + 'low.csv'
