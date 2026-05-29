import requests
import json

url = 'https://finance.naver.com/api/sise/etfItemList.nhn'
res = requests.get(url)
data = json.loads(res.text)
items = data['result']['etfItemList']
if items:
    print(items[0].keys())
    print(items[0])
