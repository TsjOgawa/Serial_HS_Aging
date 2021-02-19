# coding:utf-8
import tkinter
from tkinter import ttk

main_win = tkinter.Tk()
main_win.title("メール選別")
main_win.geometry("540x100")
main_win.resizable(0,0)

## ウィジェット作成
#--ラベル作成
#機種名
Model_label=tkinter.Label(text="機種名:")
Model_label.place(x=10,y=10)

#機種名入力（Combobox)
Model_List=ttk.Combobox()
Model_List.place(x=60,y=10)

#メール種類
mail_label=tkinter.Label(text="種類:")
mail_label.place(x=220,y=10)
#機種名入力（Combobox)
mail_List=ttk.Combobox()
mail_List.place(x=270,y=10)

#Runボタン
Run_but=tkinter.Button(text="実行")
Run_but.place(x=450,y=10)

Run_but=tkinter.Button(text="Cancel")
Run_but.place(x=450,y=50)
main_win.mainloop()
