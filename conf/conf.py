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

# 站点主页
sites = {
    '资产设备管理系统': "eam.chenksoft.com",  # 资产设备管理系统
    '文件档案管理系统': "doc.chenksoft.com",  # 文件档案管理系统
    '图书管理系统': "lib.chenksoft.com",  # 图书管理系统
    '项目管理系统': "pms.chenksoft.com",  # 项目管理系统
    'ERP管理系统': "erp.chenksoft.com",  # ERP管理系统
    '售后工单管理系统': "tms.chenksoft.com"  # 售后工单管理系统
}

# 产品页面

products = {
    '晨科资产管理系统': sites['资产设备管理系统']+"/col.jsp?id=10",
    '晨科设备管理系统': sites['资产设备管理系统']+"/col.jsp?id=11",

    '晨科图书管理系统': sites['图书管理系统']+"/col.jsp?id=31",

    '晨科业务管理系统': sites['项目管理系统']+"/col.jsp?id=6",

    '晨科机械制造业ERP管理系统': sites['ERP管理系统']+"/col.jsp?id=17",
    '晨科鞋服生产ERP管理系统': sites['ERP管理系统']+"/col.jsp?id=18",
    '晨科化工行业ERP管理系统': sites['ERP管理系统']+"/col.jsp?id=17",
    '晨科家具生产ERP管理系统': sites['ERP管理系统']+"/col.jsp?id=17",

    '晨科售后工单管理系统': sites['售后工单管理系统']+"/col.jsp?id=8",
}

# 请求
urls = {
    '轮播图': "http://hz.chenksoft.com:80/ckapi/api/1/v2/pic.jsp",  # 有数据
    '下载': "http://hz.chenksoft.com/SK_CFW_Servlet.do",

    '首页_优点说明': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_adv_description.jsp",  # 有数据 1（有些系统没有写优点）
    '首页_优点': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_adv.jsp",  # 有数据（有些系统没有写优点）
    '首页_功能标题': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_func_title.jsp",  # 有数据 1
    '首页_功能': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_func.jsp",  # 有数据
    '首页_解决方案标题': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_solution_title.jsp",  # 1
    '首页_解决方案': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_solution.jsp",  # 有数据

    '产品页_顶部内容': "http://hz.chenksoft.com:80/ckapi/api/1/v2/product_top_content.jsp",  # 有数据
    '产品页_顶部标签': "http://hz.chenksoft.com:80/ckapi/api/1/v2/product_top_lable.jsp",  # 有数据
    '产品页_中间标题': "http://hz.chenksoft.com:80/ckapi/api/1/v2/product_center_title.jsp",  # 有数据 1
    '产品页_中间内容': "http://hz.chenksoft.com:80/ckapi/api/1/v2/product_center_content.jsp",  # 有数据
    '产品页_底部内容': "http://hz.chenksoft.com:80/ckapi/api/1/v2/pro-foot.jsp",   # 有数据 （产品功能和截图）

    '产品_底部内容': "http://hz.chenksoft.com:80/ckapi/api/1/v2/pruduct_foot_content.jsp",  # 0
    '获取产品ID及名称': "http://hz.chenksoft.com:80/ckapi/api/1/v2/select_pro_id.jsp",  # 有数据

    '获取系统名称': "http://hz.chenksoft.com:80/ckapi/api/1/v2/get_sys_name.jsp",  # 1
    '获取视频信息': "http://hz.chenksoft.com:80/ckapi/api/1/v2/get_video.jsp",
    '根据id获取视频url': "http://hz.chenksoft.com:80/ckapi/api/1/v2/get_videourl_byid.jsp",
}

# ppt的基础信息
info = {
    "公司": "杭州晨科软件技术有限公司",

    "系统名": "sysname",
    "项目名": "proname",

    "小标题": "管理软件定制专家",
    "网址": "www.chenksoft.com",
    "电话": "400-6990-220",
    "坐下角标": "www.chenksoft.com 400-6990-220"
}
