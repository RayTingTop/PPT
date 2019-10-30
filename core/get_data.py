import requests
from conf import conf


# 获取指定数据，type：主页/产品
def request(type, name, function):
    # 判断域名
    domain = conf.products[name] if (type == "产品") else conf.sites[name]
    print("..............getdata", type, name, function)

    url = conf.urls[function]
    par = {'token': "chenksoft!@!", 'domain': domain}
    header = {'Content-Type': 'application/x-www-form-urlencoded'}

    response = requests.post(url, par, header)
    if response.status_code == 200:
        # 返回数据
        return response.json()['data']


# 下载文件到本地
def downfile(name, typename):
    # 图片地址
    url = conf.urls["下载"]
    par = {'method': "downfile", 'fid': name, 'filename': name, 'domainid': 1}
    header = {'Content-Type': 'application/x-www-form-urlencoded'}

    filename = str(name) + typename
    filepath = conf.path_image + filename
    response = requests.post(url, par, header)
    if response.status_code == 200:
        # 保存图片到图片路径
        with open(filepath, "wb")as f:
            f.write(response.content)
    response.close()
    return filepath


def testrequest():
    data = request("产品", "晨科图书管理系统", "产品页_顶部内容")
    # data = index_func("资产设备管理系统")
    print("记录数量：", len(data))
    for i in data:
        print(i)


if __name__ == "__main__":
    # print(downfile(330, ".png"))
    testrequest()
