# 配置信息

# 生成路径
save_path = r"D:\PythonProjects\PPT\ppt_save"

# 使用模板
temps = {
    '模板1': r"D:\PythonProjects\PPT\ppt_template\template1.pptx",
    # '模板2': r"D:\PythonProjects\PPT\ppt_template\template2.pptx"
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
    '首页_优点': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_adv.jsp",
    '首页_解决方案': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_solution.jsp",
    '轮播图': "http://hz.chenksoft.com:80/ckapi/api/1/v2/pic.jsp",
    '首页_功能标题': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_func_title.jsp",
    '首页_优点说明': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_adv_description.jsp",
    '首页_解决方案标题': "http://hz.chenksoft.com:80/ckapi/api/1/v2/index_solution_title.jsp",

    '产品页_顶部内容': "http://hz.chenksoft.com:80/ckapi/api/1/v2/product_top_content.jsp",
    '产品页_顶部标签': "http://hz.chenksoft.com:80/ckapi/api/1/v2/product_top_lable.jsp",
    '产品页_中间标题': "http://hz.chenksoft.com:80/ckapi/api/1/v2/product_center_title.jsp",
    '产品页_中间内容': "http://hz.chenksoft.com:80/ckapi/api/1/v2/product_center_content.jsp",
    '产品页_底部内容': "http://hz.chenksoft.com:80/ckapi/api/1/v2/pruduct_foot_content.jsp",
    '产品页_页底部内容': "http://hz.chenksoft.com:80/ckapi/api/1/v2/pro-foot.jsp",
    '获取产品ID及名称': "http://hz.chenksoft.com:80/ckapi/api/1/v2/select_pro_id.jsp",

    '获取系统名称': "http://hz.chenksoft.com:80/ckapi/api/1/v2/get_sys_name.jsp",
    '获取视频信息': "http://hz.chenksoft.com:80/ckapi/api/1/v2/get_video.jsp",
    '根据id获取视频url': "http://hz.chenksoft.com:80/ckapi/api/1/v2/get_videourl_byid.jsp"
}

pptinfo = {
    "公司名": "杭州晨科软件技术有限公司",
    "小标题": "管理软件定制专家",
    "网址": "http://www.chenksoft.com/",
    "400电话": "400-6990-220"
}
