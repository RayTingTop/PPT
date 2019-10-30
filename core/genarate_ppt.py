import os
import time
import json
import win32com
import win32com.client
from core import get_data
from conf import conf


# 生成PPT
# site:主页
# producte：产品
# temp：使用模板
def genarate(sys, product, temp_name):
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    # 是否显示打开的文件
    ppt.Visible = True
    # 屏蔽错误弹框提示
    ppt.DisplayAlerts = False
    # 打开模板
    tempPPT = ppt.Presentations.Open(conf.path_temp[temp_name])

    # 保存信息
    conf.info["系统名"] = sys
    conf.info["项目名"] = product

    # 总页数
    slide_count = tempPPT.Slides.Count
    print('读取模板成功，模板页数', slide_count)

    first_page(tempPPT)  # 首页，封面页

    unit_2 = 3
    sys_adv(unit_2 + 1, tempPPT, conf.info["系统名"])  # 2.4问题分析-优点 （数据在首页）

    unit_3 = 5
    # pro_intro(unit_3 + 1, tempPPT, conf.info["项目名"])  # 3.1产品方案-产品简介
    # pro_function(unit_3 + 2, tempPPT, conf.info["项目名"])  # 3.2产品方案-功能特点
    # pro_solution(unit_3 + 3, tempPPT, conf.info["系统名"])  # 3.3产品方案-解决方案（行业）
    pro_function_show(unit_3 + 4, tempPPT, conf.info["项目名"])  # 3.4产品方案-功能展示（截图）

    # 保存
    save(tempPPT, conf.info["项目名"])
    # 退出ppt
    ppt.Quit()


# 首页封面
def first_page(tempPPT):
    #  处理小标题，大标题，公司，网址
    slide = tempPPT.Slides(1)
    print('\n处理首页,模型数量:', slide.Shapes.Count)
    page1_content = [conf.info['小标题'], conf.info["项目名"], conf.info['公司'], conf.info['网址']]
    for i in range(1, len(page1_content) + 1):
        if slide.Shapes(i).HasTextFrame:
            slide.Shapes(i).TextFrame.TextRange.Text = page1_content[i - 1]
            print(" ", i, page1_content[i - 1])


# 2.4问题分析-优点（数据在首页）
def sys_adv(index, tempPPT, sys):
    print("\n2.4问题分析-产品优点：")
    # 功能模板
    slide = tempPPT.Slides(index)
    # 查询数据
    adv_description = get_data.request("首页", sys, "首页_优点说明")  # 中标题，中内容
    adv = get_data.request("首页", sys, "首页_优点")  # 小标题

    slide.Shapes(1).TextFrame.TextRange.Text = "专业的" + sys  # 设置大标题
    slide.Shapes(7).TextFrame.TextRange.Text = adv_description[0]["index_adv_title"]  # 设置中标题
    slide.Shapes(8).TextFrame.TextRange.Text = adv_description[0]["index_adv_description"]  # 设置中内容

    slide.Shapes(11).TextFrame.TextRange.Text = "1." + adv[0]["index_adv_content"]  # 设置中标题
    slide.Shapes(15).TextFrame.TextRange.Text = "2." + adv[1]["index_adv_content"]  # 设置中内容
    slide.Shapes(13).TextFrame.TextRange.Text = "3." + adv[2]["index_adv_content"]  # 设置中标题
    slide.Shapes(17).TextFrame.TextRange.Text = "4." + adv[3]["index_adv_content"]  # 设置中内容


# 3.1产品方案-产品简介
def pro_intro(index, tempPPT, product):
    print("\n3.1产品方案-产品简介：")
    slide = tempPPT.Slides(index)
    top_content = get_data.request("产品", product, "产品页_顶部内容")  # 产品顶部信息
    slide.Shapes(7).TextFrame.TextRange.Text = top_content[0]["skt16.skf179"]  # 简介标题
    slide.Shapes(8).TextFrame.TextRange.Text = top_content[0]["skt16.skf181"]  # 简介内容
    slide.Shapes(10).TextFrame.TextRange.Text = top_content[0]["skt16.skf180"]  # 简介概括


# 3.2产品方案-功能特点
def pro_function(index, tempPPT, product):
    print("\n3.2产品方案-系统功能：")
    # 功能模板
    slide = tempPPT.Slides(index)
    # 查询数据
    center_title = get_data.request("产品", product, "产品页_中间标题")  # 大标题
    funcs = get_data.request("产品", product, "产品页_中间内容")  # 数据

    slide.Shapes(1).TextFrame.TextRange.Text = center_title[0]["skt18.skf196"]  # 设置大标题
    # slide.Shapes(6).TextFrame.TextRange.Text = center_title[0]["skt18.skf196"]  # 设置中标题
    # slide.Shapes(7).TextFrame.TextRange.Text = center_title[0]["skt18.skf196"]  # 设置中内容
    # 循环数据，每一条记录保存为一页
    for i in range(0, len(funcs)):
        sh_index = int(i / 3 + 1) * 10 + int(i % 3 + 1) * 3  # 下标

        slide.Shapes(sh_index).TextFrame.TextRange.Text = funcs[i]["skt19.skf205"]  # 设置内容
        slide.Shapes(sh_index + 1).TextFrame.TextRange.Text = funcs[i]["skt19.skf204"]  # 设置小标题

        # 设置图标保存尺寸，单位磅，1厘米=28.35磅
        shape_circle = slide.Shapes(sh_index - 1)  # 小圆圈
        size = 20  # 功能图标尺寸
        position = (shape_circle.Width - size) / 2

        # 下载图标,这里截取掉分号
        filepath = get_data.downfile(funcs[i]["skt19.skf203"][:-1], ".png")
        # 把图标插入到页面,并缩放，位移居中
        slide.Shapes.AddPicture(FileName=filepath, LinkToFile=False, SaveWithDocument=True,
                                Left=shape_circle.Left + position,
                                Top=shape_circle.Top + position, Width=size, Height=size)
        print("", i, slide.Shapes(sh_index + 1).TextFrame.TextRange.Text)


# 3.3产品方案-解决方案
def pro_solution(index, tempPPT, sys):
    print("\n3.3产品方案-解决方案：")
    # 功能模板
    slide = tempPPT.Slides(index)
    # 查询数据
    solution_title = get_data.request("首页", sys, "首页_解决方案标题")  # 大标题
    solutions = get_data.request("首页", sys, "首页_解决方案")  # 数据

    slide.Shapes(1).TextFrame.TextRange.Text = solution_title[0]["index_solution_title"]  # 设置大标题
    # 循环数据，每一条记录保存为一页
    for i in range(0, len(solutions)):
        # 图片与文本的下标
        shape_image = slide.Shapes(i + 4)
        shape_text = slide.Shapes(i + 13)

        filepath = get_data.downfile(solutions[i]["index_solution_logo"][:-1], ".png")  # 下载图标,这里截取掉分号
        # 插入图片,设置尺寸为模板
        slide.Shapes.AddPicture(FileName=filepath, LinkToFile=False, SaveWithDocument=True, Left=shape_image.Left,
                                Top=shape_image.Top, Width=shape_image.Width, Height=shape_image.Height)
        # 插入新的文本框
        newtext = slide.Shapes.AddShape(Type=1, Left=shape_text.Left, Top=shape_text.Top, Width=shape_text.Width,
                                        Height=shape_text.Height)
        # 填充文本
        newtext.TextFrame.TextRange.Text = solutions[i]["index_solution_name"]

        # 文本框样式
        shape_text.PickUp()
        newtext.Apply()

        print("", i, newtext.TextFrame.TextRange.Text)
    # 完成后删除模板自带图片部分shapes
    for i in range(4, 22):
        slide.Shapes(4).Delete()  # 删除后下标会往前移动，一直删除第一个图片即可


# 3.4产品方案-功能展示截图
def pro_function_show(index, tempPPT, product):
    print("\n3.4产品方案-解决方案：")
    slide_temp = tempPPT.Slides(index)
    function_show = get_data.request("产品", product, "产品页_底部内容")
    funs = json.loads(function_show[0]["result"])  # 数据字符串处理为json

    # 循环数据，每一条模块保存为一页
    for fun in funs:
        slide_temp.Copy()  # 复制模板并粘贴新出新页面以保存内容
        index += 1
        slide = tempPPT.Slides.Paste(index)  # 粘贴在模板的下一页

        slide.Shapes(1).TextFrame.TextRange.Text = "核心模块 - "+fun["title"]  # 设置大标题
        slide.Shapes(9).TextFrame.TextRange.Text = fun["title"]  # 中标题

        # 图片
        shape_image = slide.Shapes(7)
        filepath = get_data.downfile(fun["item"][0]['img'][:-1], ".png")  # 下载图标,这里截取掉分号
        # 插入图片,设置尺寸为模板
        slide.Shapes.AddPicture(FileName=filepath, LinkToFile=False, SaveWithDocument=True, Left=shape_image.Left,
                                Top=shape_image.Top, Width=shape_image.Width, Height=shape_image.Height)
        # 文本（小功能标题+内容）
        text = fun["item"][0]['text']
        # 第一组文本
        slide.Shapes(10).TextFrame.TextRange.Text = text[0]['subtitle']  # 小标题
        slide.Shapes(11).TextFrame.TextRange.Text = text[0]['subcontent']  # 小内容
        # 剩余的文本
        for t in range(1, len(text)):
            title = slide.Shapes(10)
            content = slide.Shapes(11)
            # 插入新的文本框(位置与样式)
            newtitle = slide.Shapes.AddShape(Type=1, Left=title.Left, Top=title.Top + t * 70,
                                             Width=title.Width, Height=title.Height)
            newcontent = slide.Shapes.AddShape(Type=1, Left=content.Left, Top=content.Top + t*70,
                                               Width=content.Width, Height=content.Height)
            # 填充文本
            newtitle.TextFrame.TextRange.Text = text[t]['subtitle']
            newcontent.TextFrame.TextRange.Text = text[t]['subcontent']
            # 文本框样式
            title.PickUp()
            newtitle.Apply()
            content.PickUp()
            newcontent.Apply()

        slide.Shapes(7).Delete()  # 删除图片模板
        print("", index, slide.Shapes(1).TextFrame.TextRange.Text)
    slide_temp.Delete()  # 删掉多余模板


# 保存ppt
def save(tempPPT, filename):
    # 文件保存名称
    saveName = filename + time.strftime('%Y%m%d', time.localtime(time.time())) + ".pptx"
    # 保存为指定ppt
    tempPPT.SaveAs(conf.path_save + saveName)
    # 删除缓存图片
    os.chdir(conf.path_image)
    os.system('del /Q *.png')
    print('\n保存成功:', saveName)


#  系统功能 index 5 ----  暂不使用
def sys_function(index, tempPPT, sys):
    print("\n3.2产品方案-产品简介：")
    # 功能模板
    slide = tempPPT.Slides(index)
    # 查询数据
    func_title = get_data.request("首页", sys, "首页_功能标题")  # 大标题
    funcs = get_data.request("首页", sys, "首页_功能")  # 数据

    slide.Shapes(1).TextFrame.TextRange.Text = func_title[0]["index_func_title"]  # 设置大标题
    slide.Shapes(6).TextFrame.TextRange.Text = func_title[0]["index_func_title"]  # 设置大标题
    slide.Shapes(7).TextFrame.TextRange.Text = func_title[0]["index_func_title"]  # 设置大标题
    # 循环数据，每一条记录保存为一页
    for i in range(0, len(funcs)):
        sh_index = int(i / 3 + 1) * 10 + int(i % 3 + 1) * 3  # 下标

        slide.Shapes(sh_index).TextFrame.TextRange.Text = funcs[i]["func_content"]  # 设置内容
        slide.Shapes(sh_index + 1).TextFrame.TextRange.Text = funcs[i]["func_title"]  # 设置小标题

        # 设置图标保存尺寸，单位磅，1厘米=28.35磅
        shape_circle = slide.Shapes(sh_index - 1)  # 小圆圈
        size = 25
        position = (shape_circle.Width - 25) / 2

        # 下载图标,这里截取掉分号
        filepath = get_data.downfile(funcs[i]["func_logo"][:-1], ".png")
        # 把图标插入到页面,并缩放，位移居中
        slide.Shapes.AddPicture(FileName=filepath, LinkToFile=False, SaveWithDocument=True,
                                Left=shape_circle.Left + position,
                                Top=shape_circle.Top + position, Width=size, Height=size)
        print("", i, slide.Shapes(sh_index + 1).TextFrame.TextRange.Text)


if __name__ == "__main__":
    # 系统名，产品名，模板
    genarate('图书管理系统', "晨科图书管理系统", "模板2")
