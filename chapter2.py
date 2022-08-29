import win32com.client
from datetime import datetime
import pandas as pd
import time
import sys

# 크레온 플러스 오브젝트
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpWatch = win32com.client.Dispatch('CpSysDib.CpMarketWatch')  # 특징주 포착

# 뉴스 알림을 완료한 리스트
news_list = []


def get_watch_list(code):
    """특징주 리스트 반환"""
    cpWatch.SetInputValue(0, code)  # 종목코드
    cpWatch.SetInputValue(1, '1')
    cpWatch.SetInputValue(2, 0)
    cpWatch.BlockRequest()
    count = cpWatch.GetHeaderValue(2)
    if count == 0:
        return None
    columns = ['code', 'name', 'type', 'contents']
    index = []
    rows = []
    for i in range(count):
        index.append(cpWatch.GetDataValue(0, i))  # 첫 번째 칼럼에서 날짜데이터를 구해 index리스트에 추가
        rows.append([cpWatch.GetDataValue(1, i), cpWatch.GetDataValue(2, i),
                     cpWatch.GetDataValue(3, i), cpWatch.GetDataValue(4, i)])
    df = pd.DataFrame(rows, columns=columns, index=index)  # DataFrame 저장

    return df


def print_watch_list():
    """특징주 알림"""
    watch_list = get_watch_list('*')

    if watch_list is not None:
        for i, row in watch_list.iterrows():
            if row.type == 1:  # 종목 뉴스
                out = f"{str(cpCodeMgr.CodeToName(row.code))} : {str(row.contents)}"
                if out not in news_list:
                    news_list.append(out)
                    print(out)


def start():
    try:
        print(f"시작 : {datetime.now()}")
        time.sleep(5)
        while True:
            now = datetime.now()
            start_time = now.replace(hour=8, minute=0, second=0, microsecond=0)
            end_time = now.replace(hour=22, minute=0, second=0, microsecond=0)
            if start_time < now < end_time:
                print_watch_list()
                time.sleep(10)
            else:
                print('프로그램 종료')
                break
        sys.exit(0)
    except Exception as e:
        print(f"run() exception : {str(e)}")


if __name__ == '__main__':
    start()
