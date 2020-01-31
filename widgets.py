#!/usr/bin/python3
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import webbrowser

root = tk.Tk()

lbls = ['Worship Aid', 'Register', 'Cemetery Director', 'Filed']
offvars = []

waVar = tk.StringVar()
waVar.set('Not Printed')
cdVar = tk.StringVar()
rgVar = tk.StringVar()

chk_worship = tk.Checkbutton(root, 
							 text="Worship Aid", 
							 variable=waVar,
							 onvalue='Worship Aid',
							 offvalue='Not Printed',
							 command=lambda: print(waVar.get()))
							 
chk_worship.pack()

chk_cemdir = tk.Checkbutton(root, text="Cemetery Director", variable=cdVar, command=lambda: print(cdVar.get()))
chk_cemdir.pack()
	


root.mainloop()
