import requests
import pandas as pd

def test_twse():
    url = 'https://openapi.twse.com.tw/v1/exchangeReport/MI_INDEX'
    res = requests.get(url)
    data = res.json()
    print("TWSE MI_INDEX sample:", data[:2])
    
def test_tpex():
    url = 'https://www.tpex.org.tw/openapi/v1/t187ap03_L'
    res = requests.get(url)
    data = res.json()
    print("TPEx t187ap03_L sample:", data[:2])
    
if __name__ == '__main__':
    test_twse()
    test_tpex()
