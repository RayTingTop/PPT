import win32com
import win32com.client
import time
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
    tempPPT = ppt.Presentations.Open(conf.temps[temp])

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

    # 查找一页并且复制
    # tempPPT.Slides.FindBySlideID(270).Copy()
    # 粘贴到指定index之前,不写则追加到最后
    # tempPPT.Slides.Paste(5)

    # for i in range(4, slide_count + 1):
    #     slide = tempPPT.Slides(i)  # 页
    #
    #     shape_count = slide.Shapes.Count  # 模型数量
    #     print("\n[第%d页]模型数量：%d" % (slide.SlideIndex, shape_count))
    #
    #     for j in range(1, shape_count + 1):
    #         if slide.Shapes(j).HasTextFrame:
    #             shape = slide.Shapes(j)
    #             print(" ", j, shape.TextFrame.TextRange.Text)

    # 保存
    save(tempPPT)
    # 退出ppt
    ppt.Quit()


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


# 功能
def func_pages(tempPPt, site):
    # 查询数据
    func_title = get_data.request(site, "首页_功能标题")  # 标题
    funcs = get_data.request(site, "首页_功能")  # 数据

    slide = tempPPt.Slides(4)  # 方法模板
    # 循环数据，每一条记录保存为一页
    for i in range(1, len(funcs)):
        print("保存功能页面:", i, funcs[i]["func_title"])
        print(slide.Shapes(1).TextFrame.TextRange.Text)
        slide.Shapes(1).TextFrame.TextRange.Text = func_title  # 设置大标题
        slide.Shapes(5).TextFrame.TextRange.Text = funcs[i]["func_title"]  # 设置小标题
        slide.Shapes(6).TextFrame.TextRange.Text = funcs[i]["func_content"]  # 设置内容标题
        slide.Copy()
        slide.Paste(5)


# 解决方案
def solution(tempPPt, site):
    solution_title = get_data.index_solution_title(site)
    solutions = get_data.index_solution(site)


# 保存ppt
def save(tempPPT):
    # 文件保存名称
    saveName = conf.pptinfo["项目名"] + time.strftime('%Y%m%d', time.localtime(time.time())) + ".pptx"
    # 保存为指定ppt
    tempPPT.SaveAs(conf.save_path + "\\" + saveName)
    print('\n保存成功:', saveName)


if __name__ == "__main__":
    genarate('资产设备管理系统', "模板2")
