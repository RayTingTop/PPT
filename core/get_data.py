import requests
from conf import conf


# 获取指定数据
def request(sitename, function):
    print("-------getdata", sitename, function)
    url = conf.urls[function]
    par = {'token': "chenksoft!@!", 'domain': conf.sites[sitename]}
    header = {'Content-Type': 'application/x-www-form-urlencoded'}

    response = requests.post(url, par, header)
    if response.status_code == 200:
        # 返回数据
        return response.json()['data']


# 下载文件
def downfile(name, type):
    # 图片地址
    url = conf.urls["下载"]
    par = {'method': "downfile", 'fid': name, 'filename': name, 'domainid': 1}
    header = {'Content-Type': 'application/x-www-form-urlencoded'}

    filename = str(name) + type
    filepath = conf.path_images + filename
    response = requests.post(url, par, header)
    if response.status_code == 200:
        # 保存图片到图片路径
        with open(filepath, "wb")as f:
            f.write(response.content)
    response.close()
    return filepath


def testrequest():
    data = request("资产设备管理系统", "产品页_顶部内容")
    # data = index_func("资产设备管理系统")
    print("记录数量：", len(data))
    for i in data:
        print(i)


if __name__ == "__main__":
    print(downfile(330, ".png"))
    # testrequest()
