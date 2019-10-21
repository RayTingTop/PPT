import win32com
from win32com.client import Dispatch, constants

ppt = win32com.client.Dispatch('PowerPoint.Application')
ppt.Visible = 1
pptSel = ppt.Presentations.Open("C:\\web\\phpStudy\\WWW\\ppt\\Russia\\1.pptx")

# win32com.client.gencache.EnsureDispatch('PowerPoint.Application')
# # #get the ppt's pages
#
# slide_count = pptSel.Slides.Count
#
# for i in range(1, slide_count + 1):
#         shape_count = pptSel.Slides(i).Shapes.Count
#         print(shape_count)
#
# for j in range(1, shape_count + 1):
#     if pptSel.Slides(i).Shapes(j).HasTextFrame:
#         s = pptSel.Slides(i).Shapes(j).TextFrame.TextRange.Text
#         print(s.encode('utf-8') + "\n")
# ppt.Quit()
