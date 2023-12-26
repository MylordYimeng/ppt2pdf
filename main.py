import win32com.client as w32c
import  os
from typing import List
from io import BytesIO
from PyPDF2 import PdfFileMerger
import locale
import tempfile

def ppt2pdf_api(ppt_path, pdf_path):
    powerpoint = w32c.Dispatch('PowerPoint.Application')
    ppt = powerpoint.Presentations.Open(ppt_path,1,0,0) # ReadOnly, titled, WithoutWindow
    ppt.SaveAs(pdf_path, 32)
    ppt.Close()
    powerpoint.Quit()

def merge2pdf_api(ppt_path:List[str], pdf_path:str):
    powerpoint = w32c.Dispatch('PowerPoint.Application')
    total = PdfFileMerger()
    for path in ppt_path:
        ppt = powerpoint.Presentations.Open(path,1,0,0)
        # 创建一个临时文件
        temp_file = tempfile.mktemp(suffix='.pdf')
        # 保存演示文稿到临时文件
        ppt.SaveAs(temp_file, 32)
        ppt.Close()
        # 读取临时文件的内容到BytesIO对象中
        with open(temp_file, 'rb') as f:
            temp_pdf = BytesIO(f.read())
        total.append(temp_pdf)
        # 删除临时文件
        os.remove(temp_file)
    total.write(pdf_path)
    total.close()
    powerpoint.Quit()

def get_file_names_in_folder(upper_folder_path:str):
    file_names = []
    for file in os.listdir(upper_folder_path):
        if os.path.isfile(os.path.join(upper_folder_path, file)):
            file_names.append(upper_folder_path + '\\' + file)
    return file_names

def main(folder_path:str, target_path:str):
    locale.setlocale(locale.LC_COLLATE, 'zh_CN.UTF-8') # 中文排序
    file_list = get_file_names_in_folder(folder_path)
    file_list = sorted(file_list, key=locale.strxfrm)
    merge2pdf_api(file_list, target_path)

if __name__ == '__main__':
    upper = r'C:\Users\zhang\Desktop\test'
    target = r'C:\Users\zhang\Desktop\test\test.pdf'
    main(upper, target)






