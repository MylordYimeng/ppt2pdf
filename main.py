import win32com.client as w32c


def ppt2pdf_api(ppt_path, pdf_path):
    powerpoint = w32c.Dispatch('PowerPoint.Application')
    ppt = powerpoint.Presentations.Open(ppt_path,1,0,0) # ReadOnly, titled, WithoutWindow
    ppt.SaveAs(pdf_path, 32)
    ppt.Close()
    powerpoint.Quit()

