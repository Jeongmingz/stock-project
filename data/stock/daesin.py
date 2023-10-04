import win32com.client
import pandas as pd
from datetime import datetime
from datetime import timedelta
instCpCybos = win32com.client.Dispatch('CpUtil.CpCybos')
instCpStockCode = win32com.client.Dispatch('CpUtil.CpStockCode')
instCpCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
instStockChart = win32com.client.Dispatch('CpSysDib.StockChart')


def check_connection():
    if instCpCybos.IsConnect == 1:
        print('정상적으로 연결')
    else:
        print('연결 실패')


def return_stock(data, dataInt):
    '''
    주식명을 통해 주식 종목코드를 가지고 오는 함수.

    :param dataInt: int
    :param data: 주식명
    :return: dictionary : stock_code, stock_name
    '''

    if instCpCybos.IsConnect == 1:
        stockNum = instCpStockCode.GetCount()

        result = {}
        for item in range(stockNum):
            if instCpStockCode.GetData(dataInt, item) == data:
                result['stock_code'] = instCpStockCode.GetData(0, item)
                result['stock_name'] = instCpStockCode.GetData(1, item)

        import pprint
        pprint.pprint(result)

        return result
    else:
        print('Cybos is not connected')


def get_stock_list_ETF():
    rows = list()
    CPE_MARKET_KIND = {'KOSPI': 1, 'KOSDAQ': 2}

    for key, value in CPE_MARKET_KIND.items():
        codeList = instCpCodeMgr.GetStockListByMarket(value)
        for code in codeList:
            name = instCpCodeMgr.CodeToName(code)
            sectionKind = instCpCodeMgr.GetStockSectionKind(code)
            if sectionKind == 10:
                row = [code, name]
                rows.append(row)

    print(rows)
    return rows


# 특정 범위 일자 종목 데이터 가져오기
def get_stock_info(stock_code, start_day, end_day, type):
    import time

    instStockChart.SetInputValue(0, stock_code)  # 종목명
    instStockChart.SetInputValue(1, ord('1'))  # 1 : 기간으로 요청, 2: 개수로 요청
    instStockChart.SetInputValue(3, start_day)  # 요청 시작일
    instStockChart.SetInputValue(2, end_day)  # 요청 종료일
    # instStockChart.SetInputValue(4, request_num) # 요청 개수
    '''
    # 요청할 데이터 종류(리스트 형태로 요청 가능)
    0 : 날짜, 1 : 시간 - hhmm, 2 : 시가, 3 : 고가, 4 : 저가, 5 : 종가
    6 : 전일대비, 8 : 거래량, 9 : 거래대금, 13 : 시가총액
    '''
    instStockChart.SetInputValue(5,
                                 [0, 1, 2, 3, 4, 5, 6, 8, 9, 10, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25,
                                  26])
    instStockChart.SetInputValue(6, ord(type))  # 'D' : 일봉, 'm' : 분봉
    instStockChart.SetInputValue(9, ord('1'))

    instStockChart.BlockRequest()  # 위 정보로 요청

    numrow, numcolumn = instStockChart.GetHeaderValue(3), instStockChart.GetHeaderValue(2)

    index = []
    for i in range(numrow):
        index_ = str(instStockChart.GetDataValue(0, i))
        index.append(index_)

    stock_info = pd.DataFrame(columns=numcolumn[1:], index=index)

    for num in range(numrow):
        for col in range(len(numcolumn)):
            # 1,2,3,4,5,6,7,8,9, 10
            stock_info.iloc[num, col - 1] = str(instStockChart.GetDataValue(col, num))

    return stock_info


check_connection()
get_stock_list_ETF()
get_stock_info

# return_stock('삼성전자', 1)
# return_stock('A005930', 0)



