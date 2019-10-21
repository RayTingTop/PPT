import requests
from conf import conf


# 获取指定数据
def request(sitename, function):
    url = conf.urls[function]
    par = {'token': "chenksoft!@!", 'domain': conf.sites[sitename]}
    header = {'Content-Type': 'application/x-www-form-urlencoded'}

    response = requests.post(url, par, header)
    if response.status_code == 200:
        # 返回数据
        return response.json()['data']


if __name__ == "__main__":
    data = request("资产设备管理系统", "首页_解决方案")
    print("记录数量：", len(data))
    for i in data:
        print(i)
