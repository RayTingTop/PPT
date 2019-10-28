import requests
from conf import conf


# 获取指定数据
def request(sitename, function):
    print("-------getdata",sitename,function)
    url = conf.urls[function]
    par = {'token': "chenksoft!@!", 'domain': conf.sites[sitename]}
    header = {'Content-Type': 'application/x-www-form-urlencoded'}

    response = requests.post(url, par, header)
    if response.status_code == 200:
        # 返回数据
        return response.json()['data']


# 首页_功能标题
def index_func_title(sitename):
    data = request(sitename, "首页_功能标题")
    return data[0]["index_func_title"]


# 首页_功能
def index_func(sitename):
    data = request(sitename, "首页_功能")
    return data,


# 首页_解决方案标题
def index_solution_title(sitename):
    data = request(sitename, "首页_解决方案标题")
    return data[0]["index_solution_title"]


# 首页_解决方案标题
def index_solution(sitename):
    data = request(sitename, "首页_解决方案")
    return data[0]["index_solution_title"]


if __name__ == "__main__":
    data = request("资产设备管理系统", "首页_功能")
    data = index_func("资产设备管理系统")
    print("记录数量：", len(data))
    for i in data:
        print(i)
