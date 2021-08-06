import StockInfoParser
import MonthlyReportParser

date = '2011-01-01 00:00:00'

MonthlyReportParser.update_monthly_report(int(date.split('-')[0].strip()), int(date.split('-')[1].strip()))
StockInfoParser.update_stock_info()
StockInfoParser.xlsx_to_csv_pd()

