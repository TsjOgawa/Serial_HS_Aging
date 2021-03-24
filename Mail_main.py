# coding:utf-8
import tkinter
from tkinter import filedialog
from tkinter import messagebox
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
import Mail_Check

class Main_form(object):
    def __init__(self):
        self.dFile=""
        self.conn_2=""
        self.Mlist=[]
        #self.list_A=["受信用","送信用"]
        self.list_B=""
        self.list_C=""
        #try:
        if os.path.isfile('Fileini.json'):
            with open('Fileini.json',encoding="utf-8") as f:
                
                dFile = json.load(f)
                dFile=dFile['FlleLink']
                self.dFile=dFile["SaveDB"]
                conn_1 = sqlite3.connect(dFile["SaveDB"])
                self.conn_2 = dFile["saveFolder"]
                cur_1 = conn_1.cursor()
                cur_1.execute("SELECT 機種名 FROM Model_table")
                for row in (cur_1):
                    self.Mlist.append(str(row[0]))
                cur_1.close()
                conn_1.close()

    def start(self):
        global stop_flag
        global thread

        # スレッドが無いなら生成してstart()する
        #if not thread:
        thread = threading.Thread(target=self.delete,name="Thread_1")
        thread.setDaemon(True)
        stop_flag=False
        thread.start()      


            
    def ask_folder(self):
        """　
        フォルダ指定
        """
        #path = filedialog.askdirectory()
        #folder_path.set(path)
        #input_dir = folder_path.get()
        self.list_B=(os.path.join(self.conn_2,Model_List.get()))
        
        if not os.path.isdir(self.list_B):
        #if not os.path.isdir(os.path.join(self.conn_2,self.Mlist[0])):#<=
        #    for a in self.Mlist:
            os.makedirs(self.list_B)
            #self.list_B=(os.path.join(self.conn_2,Model_List.get()))
        self.list_C=os.path.join(self.list_B,mail_List.get())
        
        self.list_B=os.path.join(self.list_B,"提出資料")
        if  not os.path.isdir(self.list_B):
            os.makedirs(self.list_B)
        self.list_C=os.path.join(self.list_B,mail_List.get())
        if not os.path.isdir(self.list_C):
            #os.path.join(path,list_B[0])
            #for b in self.list_B:
            #for b in self.list_A:
                #for c in self.list_B:
            os.makedirs(self.list_C)
                    #os.makedirs(os.path.join(self.list_B[b],self.list_A[1]))
        #Mailcheck
        print(mail_List.get())
        print(Model_List.get())
        result = Mail_Check.MailParser(self.dFile,self.conn_2,self.list_C,Model_List.get(),mail_List.get())
        thread = threading.Thread(target=result.app)
        Run_but_1.configure(state = "disabled")
        Run_but_2.configure(state = "disabled")
        Model_List.configure(state = "disabled")
        mail_List.configure(state = "disabled")
        pb2.start(100)
        thread.setDaemon(True)
        thread.start()
        thread.join()
        messagebox.showinfo('終了！', 'メールファイル完成致しました。')

        Run_but_1.configure(state = "normal")
        Run_but_2.configure(state = "normal")
        Model_List.configure(state = "normal")
        mail_List.configure(state = "normal")
        pb2.stop()
    
    def delete(self):
        main_win.destroy()

        
if __name__ == '__main__':
    Ma=Main_form()
    main_win = tkinter.Tk()
    TXRXLIST={'受信用':0,'送信用':1}
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
    Model_List.set(MODELLIST[0])
    #メール種類
    mail_label=tkinter.Label(text="種類:")
    mail_label.place(x=220,y=10)
    #機種名入力（Combobox)
    mail_List=ttk.Combobox(textvariable=v,values=list(TXRXLIST.keys()))
    mail_List.place(x=270,y=10)
    mail_List.set(list(TXRXLIST.keys())[0])

    folder_path = tkinter.StringVar()

    #プログレスバー
    pb2 = ttk.Progressbar(main_win,maximum=10,value=0,length=200,mode='indeterminate')
    pb2.place(x=60,y=50)

    #Runボタン
    Run_but_1=tkinter.Button(text="実行",command=Ma.ask_folder,width=10)
    Run_but_1.place(x=450,y=10)

    Run_but_2=tkinter.Button(text="Cancel",command=Ma.delete,width=10)
    Run_but_2.place(x=450,y=50)
    main_win.mainloop()
