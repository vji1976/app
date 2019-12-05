#!/usr/bin/env python3
"""
	OFFICE MANAGER WIDGETS
	VERSION	:	1.0
	CODED	:	12.02.2019
	CODER	:	Vincent Iarocci
	ENV		: 	Win10 x64 python3
	
	Collection of custom widgets for Office
	assistant app.
"""
import datetime
import tkinter as tk
from tkinter import ttk
import webbrowser
# 3RD PARTY IMPORTS
from PIL import Image, ImageTk
from tkcalendar import DateEntry
from docx import Document

# HEX COLORS #
# ---------- #
# american river	#636e72
# blue nights		#353b48
# electromagnetic	#2f3640
# mazarine blue 	#273c75
# chain gang grey	#718093
# periwinkle		#9c88ff
# grey porcelain	#84817a
# hot stone			#aaa69d
# crocodile tooth	#d1ccc0
# olop blue			#2C5FAC
# olop green		#6D9A64
# olop light		#C7E4C2
# olop purple		#BCA597

# DEFAULT BASE FONT TUPLES #
# ------------------------ #
BaseFont = ('Calibri', 10, 'normal')
BoldFont = ('Calibri', 10, 'bold')
ItalFont = ('Calibri', 10, 'italic')
HeadFont = ('Segoe UI', 11, 'bold')

# CUSTOM WIDGET CLASS DEFINITIONS #
# ------------------------------- #
class wStyle(ttk.Style):
	def __init__(self):
		super().__init__()
		self.configure('.',
						foreground='#636e72',
						font=BaseFont)		
		self.configure('statusbar.TLabel',
					    borderwidth=2,
					    relief='groove')		
					    
class wMenu(tk.Menu):
	def __init__(self, parent):
		super().__init__(parent)
		parent.configure(menu=self)
		# create menus
		self.fm = tk.Menu(self, tearoff=0)
		self.om = tk.Menu(self, tearoff=0)
		self.hm = tk.Menu(self, tearoff=0)
		# add menus
		self.add_cascade(label="File", menu=self.fm)
		self.add_cascade(label="Options", menu=self.om)
		self.add_cascade(label="Help", menu=self.hm)
		
class wLabelFrame(tk.LabelFrame):
	def __init__(self, parent, **kw):
		super().__init__(parent, **kw)
		self['font'] = HeadFont
		self['foreground'] = '#84817a'
		
# UTILITY FUNCTIONS #
# ----------------- #
def webgo(url):
	webbrowser.open_new(url)
	
def websearch(query, engine='google'):
	httplink = ''
	if engine == 'google':
		httplink = 'www.google.com/search?q='
	elif engine == 'yahoo':
		httplink = 'https://search.yahoo.com/search?p='
	search_url = httplink + query
	webbrowser.open_new(search_url)
	
# WIDGET FACTORY FUNCTIONS #
# ------------------------ #	
# RAW ENTRY DICTIONARY FACTORIES #
# ------------------------------ #
def makeInfoDict(labels, var):
	infodict = {label:var for label in labels}
	return infodict
	
def getVals(datadict):
	for k, v in datadict.items():
		print("{} : {}".format(k, v.get()))
		
def drawEntry(parent, label, var=None, width=20):
	wfr = ttk.Frame(parent)
	lbl = ttk.Label(wfr, text=label, anchor=tk.W)
	ent = ttk.Entry(wfr, width=width, textvariable=var)
	lbl.pack(side=tk.LEFT)
	ent.pack()
	return wfr
		
def drawEntries(parent, dataDict):
	"""Build and return a framed custom widget set of labeled entries.
	
	Accepts:
		parent   tk/ttk frame placed in parent
		dataDict dictionary of built with labels
				 entries assigning tk.StringVars()		
	Returns:
		True	 boolean just a one off to return something
	"""
	for k, v in dataDict.items():
		drawEntry(parent, label=k, var=v).pack()
	return True
	
def drawRadios(parent, **kw):
	radfra = ttk.Frame(parent)
	labels = kw['labels']
	var = kw['var']
	rx = 0
	cy = 0
	for label in labels:
		rad = ttk.Radiobutton(radfra, text=label,
							  variable=var, value=label)
		rad.grid(row=rx, column=cy, sticky='nw')
		cy += 1 
		if cy % 3 == 0:
			rx += 1
			cy = 0
	return radfra
	
def drawCombo(parent, width=20, **kw):
	cbofra = ttk.Frame(parent)
	cbolbl = ttk.Label(cbofra, text=kw['label'])
	cbo = ttk.Combobox(cbofra, width=width, values=kw['vals'], 
					   textvariable=kw['var'])
	cbo.current(0)	# sets first item in dropdown
	cbo.bind('<<ComboboxSelected>>', lambda e: kw['var'].set(cbo.get()))
	cbolbl.grid(row=0, column=0, sticky='nw')
	cbo.grid(row=1, column=0)
	return cbofra
	
def drawDate(parent, width=10, **kw):
	year = datetime.datetime.now().year
	datefra = ttk.Frame(parent)
	datelbl = ttk.Label(datefra, text=kw['label'])
	de = DateEntry(datefra, width=width, year=year, 
				   foreground='white', background='darkblue',
				   borderwidth=2, textvariable=kw['var'],
				   cursor='hand1')
	datelbl.grid(row=0, column=0, sticky='nw')
	de.grid(row=1, column=0)
	return datefra
	
def drawImgLabel(parent, imgpath, size=(128,128)):
	raw = Image.open(imgpath)
	raw.thumbnail(size)
	pimg = ImageTk.PhotoImage(raw)
	imglbl = ttk.Label(parent, image=pimg)
	imglbl.image = pimg	# keep a reference to image so it doesn't
								# get garbage collected
	return imglbl
	
def drawImgWebLink(parent, imgpath, webpath, size=(32,32)):
	icolbl = drawImgLabel(parent, imgpath, size=size)
	icolbl["cursor"] = 'hand2'
	icolbl.bind('<Button-1>', lambda e: webgo(webpath))
	return icolbl
