# 配置信息

# 项目根目录,跟换地址先设置
path_root = "D:\\PythonProjects\\PPT\\"

# 下载图片路径
path_image = path_root + "file\\image\\"

# ppt生成路径
path_save = path_root + "file\\ppt_save\\"

# 使用模板路径
path_temp = {
    '模板1': path_root + r"file\ppt_temp\temp1.pptx",
    '模板2': path_root + r"file\ppt_temp\temp2.pptx"
}

# 站点
sites = {
    '资产设备管理系统': "eam.chenksoft.com",  # 资产设备管理系统
    '文件档案管理系统': "doc.chenksoft.com",  # 文件档案管理系统
    '图书管理系统': "lib.chenksoft.com",  # 图书管理系统
    '项目管理系统': "pms.chenksoft.com",  # 项目管理系统
    'ERP管理系统': "erp.chenksoft.com",  # ERP管理系统
    '售后工单管理系统': "tms.chenksoft.com"  # 售后工单管理系统
}

# 请求
urls = {
    '首页_功能': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_func.jsp",
    '首页_优点': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_adv.jsp",  # 0
    '首页_解决方案': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_solution.jsp",
    '轮播图': "http://hz.chenksoft.com:80/ckapi/api/1/v2/pic.jsp",
    '首页_功能标题': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_func_title.jsp",  # 1
    '首页_优点说明': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_adv_description.jsp",  # 0
    '首页_解决方案标题': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_solution_title.jsp",  # 1

    '产品页_顶部内容': "http://hz.chenksoft.com:80/ckapi/api/1/v2/product_top_content.jsp",  # 0
    '产品页_顶部标签': "http://hz.chenksoft.com:80/ckapi/api/1/v2/product_top_lable.jsp",  # 0
    '产品页_中间标题': "http://hz.chenksoft.com:80/ckapi/api/1/v2/product_center_title.jsp",  # 0
    '产品页_中间内容': "http://hz.chenksoft.com:80/ckapi/api/1/v2/product_center_content.jsp",  # 0
    '产品页_底部内容': "http://hz.chenksoft.com:80/ckapi/api/1/v2/pruduct_foot_content.jsp",  # 0
    '产品页_页底部内容': "http://hz.chenksoft.com:80/ckapi/api/1/v2/pro-foot.jsp",  # 0
    '获取产品ID及名称': "http://hz.chenksoft.com:80/ckapi/api/1/v2/select_pro_id.jsp",

    '获取系统名称': "http://hz.chenksoft.com:80/ckapi/api/1/v2/get_sys_name.jsp",  # 1
    '获取视频信息': "http://hz.chenksoft.com:80/ckapi/api/1/v2/get_video.jsp",
    '根据id获取视频url': "http://hz.chenksoft.com:80/ckapi/api/1/v2/get_videourl_byid.jsp",

    '下载': "http://hz.chenksoft.com/SK_CFW_Servlet.do",
    '产品简介': "http://hz.chenksoft.com/ckapi/api/1/v2/product_top_content.jsp?token=chenksoft!@!&domain=eam.chenksoft.com%2Fcol.jsp%3Fid%3D10"
}

# ppt的基础信息
pptinfo = {
    "公司名": "杭州晨科软件技术有限公司",
    "项目名": "ProNmme",
    "小标题": "管理软件定制专家",
    "网址": "www.chenksoft.com",
    "电话": "400-6990-220",
    "坐下角标": "www.chenksoft.com 400-6990-220"
}
