# coding:utf-8
import tkinter 
import json
import sqlite3
import os
from tkinter import ttk
class Main_form(object):
    def __init__(self):
        self.dFile=""
        self.Mlist=[]
        #try:
        if os.path.isfile('Fileini.json'):
            with open('Fileini.json') as f:
                
                dFile = json.load(f)
                dFile=dFile['FlleLink']
                conn = sqlite3.connect(dFile["SaveDB"])
                cur = conn.cursor()
                cur.execute("SELECT 機種名 FROM Model_table")
                for row in (cur):
                    self.Mlist.append(str(row[0]))
                cur.close()
                conn.close()
                #print(self.Label_1)
        #except:
        #    return

if __name__ == "__main__":   
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

    #Runボタン
    Run_but=tkinter.Button(text="実行")
    Run_but.place(x=450,y=10)

    Run_but=tkinter.Button(text="Cancel")
    Run_but.place(x=450,y=50)
    main_win.mainloop()
