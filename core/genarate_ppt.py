import win32com
import win32com.client
import time
from core import get_data
from conf import conf


# 生成PPT
def genarate(site_name, temp):
    ppt = win32com.client.Dispatch('PowerPoint.Application')

    # 是否显示打开的文件
    ppt.Visible = True

    # 屏蔽错误弹框提示
    ppt.DisplayAlerts = False

    # 打开模板
    tempPPT = ppt.Presentations.Open(conf.temps[temp])

    # 总页数
    pagescount = tempPPT.Slides.Count

    # 首页处理 小标题，大标题，公司，网址
    page_1 = [conf.pptinfo['小标题'], site_name, conf.pptinfo['公司名'], conf.pptinfo['网址']]
    for i in range(1, len(page_1) + 1):
        tempPPT.Slides(1).Shapes(i).TextFrame.TextRange.text = page_1[i - 1]
        print(tempPPT.Slides(1).Shapes(i).TextFrame.TextRange.text)

    # 目录页由模板指定 不需要再更改
    # tempPPT.Slides(2).Shapes(1).TextFrame.TextRange.Text="目录测试1\n目录测试2"

    # 修改标题
    # tempPPT.Slides(4).Shapes.Title.TextFrame.TextRange.Text = "子页标题"
    # print(tempPPT.Slides(4).Shapes.Title.TextFrame.TextRange.Text)

    # 查找一页并且复制
    # tempPPT.Slides.FindBySlideID(270).Copy()
    # 粘贴到指定index之前,不写则追加到最后
    # tempPPT.Slides.Paste(5)

    for i in range(1, pagescount + 1):
        # print(tempPPT.Slides(i).Shapes.Title.TextFrame.TextRange.Text)
        print("[第%d页]：" % tempPPT.Slides(i).SlideId)

        # 模型数量
        shapescount = tempPPT.Slides(i).Shapes.Count
        # ranges = tempPPT.Slides(i).Shapes.ShapeRange
        # print(ranges)

        for j in range(1, shapescount + 1):
            print("-----%d：" % j)
            shape = tempPPT.Slides(i).Shapes(j)

        #     print(shape)

    # slide_count = tempPPT.Slides.Count
    # print(slide_count)
    # tempPPT.Slides(1).Shapes(4).TextFrame.TextRange.Text="详细2"

    # 文件保存名称
    # get_data.request(site, "获取系统名称")[0]['sys_name'] #请求查询系统名称
    saveName = site_name + time.strftime('%Y%m%d', time.localtime(time.time())) + ".pptx"
    # 保存为指定ppt
    tempPPT.SaveAs(conf.save_path + "\\" + saveName)
    # 退出ppt
    ppt.Quit()


if __name__ == "__main__":
    genarate('资产设备管理系统', "模板1")
