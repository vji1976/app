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
# CUSTOM IMPORTS
import data

# HEX COLORS #
# ---------- #
# american river	#636e72
# blue nights		#353b48
# devils blue		#227093
# electromagnetic	#2f3640
# good samaritan    #3c6382
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
		self.configure('wLabel.TLabel',
						foreground='#21458C',
						font=BoldFont)
		self.configure('webLink.TLabel',
						foreground='#9c88ff')
		self.configure('webLinkHover.TLabel',
						foreground='#6D9A64')
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
		
class wEntry(tk.Entry):
	def __init__(self, parent, **kw):
		super().__init__(parent, **kw)
		self['borderwidth'] = 1
		self['relief'] = 'flat'
		self['background'] = '#d1ccc0'
		
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
	lbl = ttk.Label(wfr, text=label, anchor=tk.W, style='wLabel.TLabel')
	ent = wEntry(wfr, width=width, textvariable=var)
	lbl.pack(side=tk.LEFT)
	ent.pack(side=tk.RIGHT, fill=tk.X, expand=tk.TRUE, padx=8)
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
		drawEntry(parent, label=k, var=v).pack(side=tk.TOP, fill=tk.X, pady=4)
	return True
	
def drawSrvEntries(parent, dataDict):
	rx = 0
	cbodict = {"Service Time" : data.times,
			   "Service Day" :	data.days,
			   "Service Location" : data.fun_places,
			   "Celebrant" : data.celebrants,
			   "Funeral Home" : data.fun_homes}
	"""
	srvImgFrame = ttk.Frame(srvFrame)
	srvImgFrame.grid(row=0, column=1, sticky='news')
	srvImgLabel = wlib.drawImgLabel(srvImgFrame, 'img/servers.png')
	srvImgLabel.grid(row=0, column=0, sticky='nswe', pady=4)
	"""	
	return True	
	
def drawDpiEntries(parent, dataDict):
	msradlbls = ['Single', 'Married', 'Divorced', 'Widowed', 'Widower', 'Other']
	rx = 0	# row ctr
	for k, v in dataDict.items():
		if k == 'Date of Birth' or k == 'Date of Death':
			dfr = drawDate(parent, label=k, var=v, width=8)
			dfr.grid(row=rx, column=0, sticky='nw', padx=4, pady=4)
		elif k == 'Marital Status':
			rfr = ttk.Frame(parent)
			msrlbl = ttk.Label(rfr, text=k, style='wLabel.TLabel')
			msrlbl.grid(row=0, column=0, sticky='nw')
			msradsfra = drawRadios(rfr, labels=msradlbls, var=v)
			msradsfra.grid(row=1, column=0, sticky='nw')
			rfr.grid(row=rx, column=0, stick='nw', padx=4, pady=4)
		else:
			efr = drawEntry(parent, label=k, var=v)
			efr.grid(row=rx, column=0, sticky='new', padx=4, pady=4)
		rx += 1
	return True
	
def drawCemEntries(parent, dataDict):
	burlbls = ['Burial Date', 'Burial Time', 'Burial Day']
	burfra = ttk.Frame(parent)
	rx = 0 # row ctr
	cy = 0 # column ctr for burials
	for k, v in dataDict.items():
		if k in burlbls:
			if k == 'Burial Date':
				dfr = drawDate(burfra, label=k, var=v, width=8)
				dfr.grid(row=0, column=cy, sticky='nw', padx=4, pady=2)
			elif k == 'Burial Time':
				tfr = drawCombo(burfra, label=k, vals=data.times, var=v, width=7)
				tfr.grid(row=1, column=0, sticky='nw', padx=4, pady=2)
			elif k == 'Burial Day':
				yfr = drawCombo(burfra, label=k, vals=data.days, var=v, width=11)
				yfr.grid(row=0, column=cy, padx=4, pady=2)
			
			cy += 1			
			burfra.grid(row=rx, column=0, stick='new', padx=4, pady=4)
		else:
			efr = drawEntry(parent, label=k, var=v)
			efr.grid(row=rx, column=0, sticky='new', padx=4, pady=4)
		rx += 1
	return True
	
def drawDate(parent, width=10, **kw):
	rx = 0	# these are used to orient label
	cy = 1	# either above or next to entry
	year = datetime.datetime.now().year
	datefra = ttk.Frame(parent)
	datelbl = ttk.Label(datefra, text=kw['label'], style='wLabel.TLabel')
	de = DateEntry(datefra, width=width, year=year, 
				   foreground='white', background='darkblue',
				   borderwidth=2, textvariable=kw['var'],
				   cursor='hand1')
				   
	datelbl.grid(row=0, column=0, sticky='nw')
	de.grid(row=1, column=0)
	return datefra
	
def drawCombo(parent, width=20, **kw):
	cbofra = ttk.Frame(parent)
	cbolbl = ttk.Label(cbofra, text=kw['label'], style='wLabel.TLabel')
	cbo = ttk.Combobox(cbofra, width=width, values=kw['vals'], 
					   textvariable=kw['var'])
	cbo.current(0)	# sets first item in dropdown
	cbo.bind('<<ComboboxSelected>>', lambda e: kw['var'].set(cbo.get()))
	cbolbl.grid(row=0, column=0, sticky='nw')
	cbo.grid(row=1, column=0)
	return cbofra
	
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
	
def drawLinkLabel(parent, label, url):
	lbl = ttk.Label(parent, text=label, cursor='hand2',
				    style='weblink.TLabel')
	lbl.bind('<Enter>', lambda e: lbl.configure(style='webLinkHover.TLabel'))
	lbl.bind('<Leave>', lambda e: lbl.configure(style='webLink.TLabel'))
	lbl.bind('<Button-1>', lambda e: webgo(url))
	return lbl
	
def drawImgLabel(parent, imgpath, size=(128,128)):
	raw = Image.open(imgpath)
	raw.thumbnail(size)
	pimg = ImageTk.PhotoImage(raw)
	imglbl = ttk.Label(parent, image=pimg)
	imglbl.image = pimg	# keep a reference to image so it doesn't
	return imglbl		# get deleted in python garbage collection
	
def drawImgWebLink(parent, imgpath, webpath, size=(32,32)):
	icolbl = drawImgLabel(parent, imgpath, size=size)
	icolbl["cursor"] = 'hand2'
	icolbl.bind('<Button-1>', lambda e: webgo(webpath))
	return icolbl
