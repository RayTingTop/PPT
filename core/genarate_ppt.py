import os
import time
import win32com
import win32com.client
from core import get_data
from conf import conf


# 生成PPT
def genarate(site, temp):
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    # 是否显示打开的文件
    ppt.Visible = True
    # 屏蔽错误弹框提示
    ppt.DisplayAlerts = False
    # 打开模板
    tempPPT = ppt.Presentations.Open(conf.path_temps[temp])

    # 系统名称
    sys_neme = get_data.request("资产设备管理系统", "获取系统名称")[0]["sys_name"]
    conf.pptinfo["项目名"] = sys_neme

    # 总页数
    slide_count = tempPPT.Slides.Count
    print('读取模板成功，模板页数', slide_count)

    # 首页
    first_page(tempPPT)

    # 目录页由模板指定 不需要再更改
    # tempPPT.Slides(2).Shapes(1).TextFrame.TextRange.Text="目录测试1\n目录测试2"

    # 修改标题,Title不一定每页都有
    # tempPPT.Slides(4).Shapes.Title.TextFrame.TextRange.Text = "子页标题"
    # print(tempPPT.Slides(4).Shapes.Title.TextFrame.TextRange.Text)

    # 功能
    func_pages(tempPPT, site)
    # 解决方案
    solution(tempPPT, site)

    # 查找一页并且复制
    # tempPPT.Slides.FindBySlideID(270).Copy()
    # 粘贴到指定index之前,不写则追加到最后
    # tempPPT.Slides.Paste(5)

    # 保存
    save(tempPPT)
    # 退出ppt
    ppt.Quit()

    # 完成之后删除图片
    os.chdir(conf.path_images)
    os.system('del /Q *.png')


# 首页
def first_page(tempPPT):
    #  处理小标题，大标题，公司，网址
    slide = tempPPT.Slides(1)
    print('\n处理首页,模型数量:', slide.Shapes.Count)
    page1_content = [conf.pptinfo['小标题'], conf.pptinfo["项目名"], conf.pptinfo['公司名'], conf.pptinfo['网址']]
    for i in range(1, len(page1_content) + 1):
        if slide.Shapes(i).HasTextFrame:
            slide.Shapes(i).TextFrame.TextRange.Text = page1_content[i - 1]
            print(" ", i, page1_content[i - 1])

# 产品简介
def intro(tempPPt, site):
    print("\n生成产品简：")


# 功能页
def func_pages(tempPPt, site):
    print("\n生成功能列表：")
    # 功能模板
    index = 5  # 功能下标
    slide = tempPPt.Slides(index)
    # 查询数据
    func_title = get_data.request(site, "首页_功能标题")  # 大标题
    funcs = get_data.request(site, "首页_功能")  # 数据

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


# 解决方案页
def solution(tempPPt, site):
    print("\n生成解决方案列表：")
    # 功能模板
    index = 6  # 功能下标
    slide = tempPPt.Slides(index)
    # 查询数据
    solution_title = get_data.request(site, "首页_解决方案标题")  # 大标题
    solutions = get_data.request(site, "首页_解决方案")  # 数据

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
    # 完成后删除模板自带shapes
    for i in range(4, 22):
        slide.Shapes(4).Delete()


# 保存ppt
def save(tempPPT):
    # 文件保存名称
    saveName = conf.pptinfo["项目名"] + time.strftime('%Y%m%d', time.localtime(time.time())) + ".pptx"
    # 保存为指定ppt
    tempPPT.SaveAs(conf.path_save + saveName)
    print('\n保存成功:', saveName)


if __name__ == "__main__":
    genarate('资产设备管理系统', "模板2")
