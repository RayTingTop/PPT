import win32com
from win32com.client import Dispatch, constants

"""
包导入pywin32
然后还有pywin32-ctypes
不然会提示无模块
然后安装过WPS的要彻底卸载干净，注册列表中残留的要删除掉清理干净
"""
import os

def ppt_text(p):
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    ppt.Visible = 1

    # pptSel = ppt.Presentations.Open('d:\\code\\python\\ppt\\a.pptx')
    pptNew = ppt.Presentations.Open('D:\\PythonProjects\\test\\ppt\\template.pptx')

    win32com.client.gencache.EnsureDispatch('PowerPoint.Application')

    slide_count = pptNew.Slides.Count
    print(slide_count)

    pptOrg = ppt.Presentations.Open('D:\\PythonProjects\\test\\ppt\\a.pptx')
    r = pptOrg.Slides.Range((1, 2, 3, 4))
    r.Copy()
    # pptOrg.Slides.Item(1).Copy()
    pptNew.Slides.Paste(5)
    # pptNew.Slides.InsertFromFile('F:\\PycharmProjects\\PPT_win32\\aa.ppt', slide_count, 1, 2)

    """
         # InsertFromFile(_FileName_, _Index_, _SlideStart_, _SlideEnd_)
         FileName 必需  String  指示包含要插入的幻灯片的文件的文件名。
         Index    必需  Long   指定的幻灯片集合中您要在其后插入 ** ** 新幻灯片的幻灯片对象的索引号。
        SlideStart 可选  Long    按文件名表示的文件中的幻灯片集合中第一个幻灯片对象的索引号。
        SlideEnd  可选  Long     按文件名表示的文件中的幻灯片集合中最后一个幻灯片对象的索引号

    pptNew.Slides.InsertFromFile(r'F:\PycharmProjects\other\chenkppt1.pptx', 0, 1, 5)
    
    表示在活动演示文稿的第一张灯片后插入文件 F:\PycharmProjects\other\chenkppt1.pptx的第一张到第五张幻灯片
       
     """

    pptNew.SaveAs(r'D:\PythonProjects\test\ppt\insert_ppt.pptx')
    ppt.Quit()


if __name__ == '__main__':
    ppt_text('')
