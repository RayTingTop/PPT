import win32com
import win32com.client


def makeppt(path):

    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True

    # 增加一个文件
    pptFile = ppt.Presentations.Add()

    # 创建页  参数1为页数 参数2为主题类型
    page1 = pptFile.Slides.Add(1, 1)

    # 正标题副标题 就两个
    t1 = page1.Shapes(1).TextFrame.TextRange
    t1.Text = "Liuwang 111"
    t2 = page1.Shapes(2).TextFrame.TextRange
    # t2 = page1.Shapes.Item(2).TextFrame.TextRange 这两个表达是一样的
    t2.Text = "Liuwang is a good man  "

    # # 第二页
    page2 = pptFile.Slides.Add(2, 2)
    t3 = page2.Shapes(1).TextFrame.TextRange
    t3.Text = "LiuGE "
    t4 = page2.Shapes(2).TextFrame.TextRange
    t4.Text = "LiuGE is a good man  "
    # # 保存
    pptFile.SaveAs(path)
    pptFile.Close()
    ppt.Quit()


path = r"D:\\PythonProjects\PHP_Genarate\file\add_slide.pptx"
makeppt(path)

