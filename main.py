#!/usr/bin/python3
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton,  QPlainTextEdit
from PyQt5 import uic
import requests,os
class UIClass:
    def __init__(self):
        self.ui = uic.loadUi("mariodownloader.ui")


def DownloadLevel():
    Log=""
    mainw.ui.labelLog.setText("")
    LevelID=mainw.ui.levelID.text().replace(" ","").replace("-","").upper()[0:9]
    if (len(LevelID)<9):
        Log="关卡 ID 错误，请重新输入。"
        mainw.ui.labelLog.setText(Log)
    else:
        LevelFolder=mainw.ui.levelFolder.text()
        if(LevelFolder[len(LevelFolder)-1]=="/" or LevelFolder[len(LevelFolder)-1]=="\\"):
            LevelFolder=LevelFolder[:-1]
            print(LevelFolder)
        LevelFolder=LevelFolder.replace("\n","")
        with open("config.txt","r+") as f:
            f.write(LevelFolder)
        APIBase="https://tgrcode.com/mm2/level_data/"
        Log=Log+"关卡 ID："+LevelID
        Log=Log+"\n工作目录："+LevelFolder
        LevelName=[x for x in os.listdir(LevelFolder) if "bcd" in x]
        LevelName.sort(key=lambda x:os.path.getmtime(os.path.join(LevelFolder,x)))
        Log=Log+"\n最新关卡 BCD 文件为："+LevelName[-1]
        LevelName=os.path.join(LevelFolder,LevelName[-1])
        Log=Log+"\n正在下载关卡 ..."
        mainw.ui.labelLog.setText(Log)
        reqData = requests.get(url=APIBase+LevelID)
        if("No course with that ID" in reqData.text):
            Log=Log+"\n发生错误，关卡不存在。"
        else:
            os.remove(LevelName)
            with open(LevelName,"wb") as f:
                f.write(reqData.content)
            Log=Log+"\n完成！ ..."
        mainw.ui.labelLog.setText(Log)
app = QApplication([])
mainw = UIClass() 
with open("config.txt","r") as f:
    LevelFolder=f.readline()
    mainw.ui.levelFolder.setText(LevelFolder)
mainw.ui.labelLog.setText("(日志区)")
mainw.ui.btnDownload.clicked.connect(DownloadLevel)
mainw.ui.show()
app.exec_()