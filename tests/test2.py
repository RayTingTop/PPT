import win32com
from win32com.client import Dispatch, constants

ppt = win32com.client.Dispatch('PowerPoint.Application')
ppt.Visible = 1
pptSel = ppt.Presentations.Open(r"D:\PythonProjects\PPT\ppt_template\template2.pptx")
# pptSel = ppt.Presentations.Open(r"D:\PythonProjects\PPT\ppt_save\资产设备管理系统20191029.pptx")

win32com.client.gencache.EnsureDispatch('PowerPoint.Application')
# #get the ppt's pages

slide_count = pptSel.Slides.Count

for i in range(5, slide_count + 1):
    slide = pptSel.Slides(i)
    shape_count = slide.Shapes.Count
    print(i, "页")

    # print(slide.Shapes.Title.TextFrame.TextRange.Text)

    for j in range(1, shape_count + 1):
        if slide.Shapes(j).HasTextFrame:
            # 每一个内容   类型14 大标题
            s = slide.Shapes(j).TextFrame.TextRange.Text
            print(" ",j, slide.Shapes(j).Type, s)
    print("\n")
# ppt.Quit()
