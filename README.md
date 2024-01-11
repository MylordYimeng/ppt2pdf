- 主要就是调用windows中的office里面的库将ppt读取后保存为pdf, 再通过PyPDF2里面的函数将其合并, 在极少数的情况下有那么一点点作用, 比如想要将十几个(上课的)ppt整合在一起去打印.
- 默认将所谓"ppt文件夹"下的所以以.ppt和.pptx为后缀的文件合并成"目标文件夹"下的名为"Target.pdf"的文件
- 后悔用PySide6写gui, 打包出来50M...
# 库依赖:
## PySide6
## win32com
## PyPDF2
