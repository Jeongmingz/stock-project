import win32com.client

instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
total_cnt = instCpStockCode.GetCount()


def returnStock(stock_name: str):
    '''
    대신증권 API에서 유저가 원하는 stock_name을 검색하여,\n
    해당 stock_name값과 같튼 주식 종목 정보를 Dict 형태로 리턴한다.

    :param stock_name:
    :return: {"stockNo": int , "stockName": str}: dict
    '''
    for i in range(total_cnt):
        if instCpStockCode.GetData(1, i) == stock_name:
            return {"stockNo": instCpStockCode.GetData(0, i),
                    "stockName": instCpStockCode.GetData(1, i)}


