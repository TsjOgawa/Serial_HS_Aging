#!/usr/bin/env python
# -*- coding: utf8 -*-
import sys
import tkinter
"""
---------------------------------------
[Function]
GUIを実験してみる
---------------------------------------
"""
root = tkinter.Tk()
root.title(u"Software Title")
root.geometry("400x300")
#ラベル
Static1 = tkinter.Label(text=u'test')
Static1.pack()

root.mainloop()