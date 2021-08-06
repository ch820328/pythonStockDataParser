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


# <editor-fold desc='monthly update'>


def monthly_report(year, month):
    # 假如是西元，轉成民國
    if year > 1990:
        year -= 1911
    date_string = str(year) + '_' + str(month)
    url = 'https://mops.twse.com.tw/nas/t21/sii/t21sc03_' + date_string + '_0.html'
    if year <= 98:
        url = 'https://mops.twse.com.tw/nas/t21/sii/t21sc03_' + date_string + '.html'

    # 下載該年月的網站，並用pandas轉換成 dataframe
    r = requests.get(url, headers=headers)
    r.encoding = 'big5'

    dataframes = pandas.read_html(io.StringIO(r.text), encoding='big-5')

    dataframe = pandas.concat([dataframe for dataframe in dataframes if 11 >= dataframe.shape[1] > 5])

    if 'levels' in dir(dataframe.columns):
        dataframe.columns = dataframe.columns.get_level_values(1)
    else:
        dataframe = dataframe[list(range(0, 10))]
        column_index = dataframe.index[(dataframe[0] == '公司代號')][0]
        dataframe.columns = dataframe.iloc[column_index]

    dataframe['當月營收'] = pandas.to_numeric(dataframe['當月營收'], 'coerce')
    dataframe = dataframe[~dataframe['當月營收'].isnull()]
    dataframe = dataframe[dataframe['公司代號'] != '合計']

    # 偽停頓
    time.sleep(5)
    writer = pandas.ExcelWriter(report_dir_path + date_string + '.xlsx', engine='openpyxl')
    dataframe.to_excel(excel_writer=writer, index=False)
    writer.save()
    print(r'monthly_report {} save.'.format(date_string + '.xlsx'))
    return dataframe


def update_monthly_report(year, month):
    print('Begin: update_monthly_report!')
    if year > 1990:
        year -= 1911
    report_list = os.listdir(report_dir_path)
    date = datetime.datetime.now()
    months_count = int(year) * 12 + int(month)
    now_months_count = (int(str(date).split('-')[0]) - 1911) * 12 + int(str(date).split('-')[1].replace('0', ''))
    while months_count < now_months_count and int(str(datetime.datetime.now()).split('-')[2].split(' ')[0]) > 15:
        date_string = str(year) + '_' + str(month) + '.xlsx'
        if date_string not in report_list:
            print(r'Download monthly_report {}.'.format(date_string))
            monthly_report(year, month)
        months_count += 1
        month += 1
        if month > 12:
            month -= 12
            year += 1
    print('End: update_monthly_report!')


# </editor-fold>


# get global variable
file_directory = str()
load_json_config()
if not file_directory:
    print('file_directory is empty.')
    sys.exit(0)

# daily update file/directory
report_dir_path = os.path.join(os.sep, file_directory, 'report')
financial_dir_path = os.path.join(os.sep, file_directory, 'financial')

headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/39.0.2171.95 Safari/537.36'}

