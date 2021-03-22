# coding:utf-8
import tkinter
from tkinter import filedialog
import json
import sqlite3
import os
from tkinter import ttk
import os
import sys
import threading
import shutil#2020/11/02 ogawa
import zipfile #2020/11/04 shimada
import pathlib #2020/11/05 shimada
from distutils import dir_util #2020/11/05 shimada
import datetime
import glob

class Main_form(object):
    def __init__(self):
        self.dFile=""
        self.Mlist=[]
        self.list_A=["受信用","送信用"]
        self.list_B=[]
        #try:
        if os.path.isfile('Fileini.json'):
            with open('Fileini.json',encoding="utf-8") as f:
                
                dFile = json.load(f)
                dFile=dFile['FlleLink']
                conn = sqlite3.connect(dFile["SaveDB"])
                cur = conn.cursor()
                cur.execute("SELECT 機種名 FROM Model_table")
                for row in (cur):
                    self.Mlist.append(str(row[0]))
                cur.close()
                conn.close()
        
    def ask_folder(self):
        """　
        フォルダ指定
        """
        path = filedialog.askdirectory()
        folder_path.set(path)
        input_dir = folder_path.get()
        if not os.path.isdir(os.path.join(path,self.Mlist[0])):
            for a in self.Mlist:
                os.makedirs(os.path.join(path,a))
                self.list_B.append(os.path.join(path,a))
        if not os.path.isdir(os.path.join(self.list_B[0],self.list_A[0])):
            #os.path.join(path,list_B[0])
            #for b in self.list_B:
            for b in self.list_A:
                for c in self.list_B:
                    os.makedirs(os.path.join(c,b))
                    #os.makedirs(os.path.join(self.list_B[b],self.list_A[1]))

if __name__ == '__main__':
    Ma=Main_form()
    main_win = tkinter.Tk()
    TXRXLIST={'受信':0,'送信':1}
    MODELLIST=Ma.Mlist
    main_win.title("メール選別")
    main_win.geometry("540x100")
    main_win.resizable(0,0)
    v=tkinter.StringVar()
    
    ## ウィジェット作成
    #--ラベル作成
    #機種名
    Model_label=tkinter.Label(text="機種名:")
    Model_label.place(x=10,y=10)

    #機種名入力（Combobox)
    Model_List=ttk.Combobox(main_win,values=MODELLIST)
    Model_List.place(x=60,y=10)
    #メール種類
    mail_label=tkinter.Label(text="種類:")
    mail_label.place(x=220,y=10)
    #機種名入力（Combobox)
    mail_List=ttk.Combobox(textvariable=v,values=list(TXRXLIST.keys()))
    mail_List.place(x=270,y=10)

    folder_path = tkinter.StringVar()

    #Runボタン
    Run_but=tkinter.Button(text="実行",command=Ma.ask_folder)
    Run_but.place(x=450,y=10)

    Run_but=tkinter.Button(text="Cancel")
    Run_but.place(x=450,y=50)
    main_win.mainloop()
