import os #用于获取目标文件所在路径
from win32com import client as wc #导入模块

path = r"C:\Users\HuangMin\Desktop\temp"
files=[]
for file in os.listdir(path):
    if file.endswith(".doc"): #排除文件夹内的其它干扰文件，只获取".doc"后缀的word文件
        files.append(os.path.join(path,file))

word = wc.Dispatch("Word.Application") # 打开word应用程序
for file in files:
    doc = word.Documents.Open(file) #打开word文件
    doc.SaveAs("{}x".format(file), 12)#另存为后缀为".docx"的文件，其中参数12指docx文件
    doc.Close() #关闭原来word文件
    os.remove(file)
word.Quit()
print("完成！")