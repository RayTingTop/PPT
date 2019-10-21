import requests
from conf import conf

# 获取指定数据
def getData(sitename, function):
    url = conf.urls[function]
    par = {'token': "chenksoft!@!", 'domain': conf.sites[sitename]}
    header = {'Content-Type': 'application/x-www-form-urlencoded'}

    response = requests.post(url, par, header)
    if response.status_code == 200:
        # 返回数据
        return response.json()['data']


if __name__ == "__main__":
    data = getData("文件档案管理系统", "首页_功能")
    print(len(data))
    for i in data:
        print(i)
