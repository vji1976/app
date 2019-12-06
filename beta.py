#!/usr/bin/env python3
"""OFFICE MANAGER v1.0

CODED ON: 
	11.24.2019	
CODED BY: 
	MosesLawn	
CODE ENV: 
	python 3.7 win10x64	
NOTES:
	tkinter.widget.winfo_children()
	tkinter.widget.children.values()
		- to access child widgets of a tkinter widget
		primarily a frame
"""
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
# 3RD PARTY IMPORTS
from tkcalendar import DateEntry
from PIL import Image, ImageTk
# CUSTOM IMPORTS
import wlib
import data

class App(tk.Tk):
	"""Basic tkinter gui interface inheriting from tk.Tk."""
	def __init__(self):
		super().__init__()
		self.iconbitmap('img/favicon.ico')	# uncomment for titlebar icon
		"""
		self.msgico = Image.open("img/favicon.ico")
		self.msgimg = ImageTk.PhotoImage(self.msgico)		
		self.tk.call('wm', 'iconphoto', self._w, self.msgimg)	# this line changes the icon for 	
																# tk messagebox
		"""
		self.title("Funeral Aid 1000")		
		self.drawmenu()
		self.columnconfigure(0, weight=0)
		self.columnconfigure(1, weight=1)
		self.columnconfigure(2, weight=1)
		self.rowconfigure(1, weight=1)
		self.initvars()
		self.drawgui()
		
	def initvars(self):
		self.dataDicts = []	# this main list will hold all
							# data dictionaries for access later
		# main data dictionary population
		self.srvDict = self.makeDataDict(data.srv_Labels)
		self.dpiDict = self.makeDataDict(data.dpi_Labels)
		self.nokDict = self.makeDataDict(data.nok_Labels)
		self.cemDict = self.makeDataDict(data.cem_Labels)
		self.srvDict["Service Type"].set("Full Funeral")
		# add data dicts
		self.dataDicts.append(self.srvDict)
		self.dataDicts.append(self.dpiDict)
		# status bar text variable for messages
		self.statusTxt = tk.StringVar()
		self.statusTxt.set("Program Running")
		# image paths
		self.fb_img_path = 'img/fb_tool.png'
		self.pa_img_path = 'img/mp_tool.png'
		self.tw_img_path = 'img/tw_tool.png'
		# url paths
		self.fb_url = data.social_paths["fb"]
		self.tw_url = data.social_paths["tw"]
		self.pa_url = data.social_paths["mp"]		
		
	def drawmenu(self):
		self.menubar = wlib.wMenu(self)
		# file menu commands
		self.menubar.fm.add_command(label="Create Document",
									command=lambda: self.createWordDoc(self.dataDicts))
		self.menubar.fm.add_separator()
		self.bind_all('<Control-q>', lambda e: sys.exit())
		self.menubar.fm.add_command(label="Quit", underline=0,
									command=sys.exit, accelerator="Ctrl+Q")		
		# options menu commands
		
		# help menu commands
		self.bind_all('<Control-h>', lambda e: self.m_help_about())
		self.menubar.hm.add_command(label="About", underline=0, 
									command=self.m_help_about, accelerator="Ctrl+H")
		
	def drawgui(self):
		self.style = wlib.wStyle()
		
		''' TOOL BAR FRAME START '''
		''' -------------------- '''
		tcol = ttk.Frame(self)
		tcol.grid(row=0, column=0, columnspan=3, sticky='new', padx=6, pady=2)		
		tbar = ttk.Frame(tcol)
		tbar.grid(row=0, column=0, sticky='nw')		
		
		wlib.drawImgWebLink(tbar, self.fb_img_path, self.fb_url).grid(row=0, column=0, sticky='nw', padx=2)
		wlib.drawImgWebLink(tbar, self.tw_img_path, self.tw_url).grid(row=0, column=1, sticky='nw', padx=2)
		wlib.drawImgWebLink(tbar, self.pa_img_path, self.pa_url).grid(row=0, column=2, sticky='nw', padx=2)
				
		''' LEFT COLUMN CONTROLS START '''
		''' -------------------------- '''
		lcol = ttk.Frame(self, width=320, style='column.TFrame')
		lcol.grid(row=1, column=0, sticky='nsew', padx=10, pady=10)
		lcol.grid_propagate(0)				# prevent column from resizing to inner widget size
		lcol.columnconfigure(0, weight=1)	# allows child widgets to fill space of column
		
		# - Funeral Information Frame
		srvFrame = wlib.wLabelFrame(lcol, text="Service Information")
		srvFrame.grid(row=0, column=0, sticky='new')
		wlib.drawSrvEntries(srvFrame, self.srvDict)
					
		
		''' MIDDLE COLUMN START '''
		''' ------------------- '''
		mcol = ttk.Frame(self, width=300, style='column.TFrame')
		mcol.grid(row=1, column=1, sticky='nsew', pady=10)
		mcol.grid_propagate(0)				# prevent column from resizing to inner widget size
		mcol.columnconfigure(0, weight=1)	# allows child widgets to fill space of column
		# - Deceased Personal Info Frame
		dpiFrame = wlib.wLabelFrame(mcol, text="Deceased Personal Info")
		dpiFrame.grid(row=0, column=0, sticky='new')
		dpiFrame.columnconfigure(0, weight=1)
		wlib.drawDpiEntries(dpiFrame, self.dpiDict)	
		# - Next of Kin Frame
		nokFrame = wlib.wLabelFrame(mcol, text="Next of Kin")
		nokFrame.grid(row=1, column=0, sticky='new')
		wlib.drawEntries(nokFrame, 	self.nokDict)
				
		''' RIGHT COLUMN START '''
		''' ------------------ '''
		rcol = ttk.Frame(self, width=300, style='column.TFrame')
		rcol.grid(row=1, column=2, sticky='nsew', padx=10, pady=10)
		rcol.grid_propagate(0)				# prevent column from resizing to inner widget size
		rcol.columnconfigure(0, weight=1)	# allows child widgets to fill space of column
		# - Cemetery Frame
		cemFrame = wlib.wLabelFrame(rcol, text="Cemetery ~ Burial Information")
		cemFrame.grid(row=0, column=0, sticky='new')
		cemFrame.columnconfigure(0, weight=1)
		wlib.drawCemEntries(cemFrame, self.cemDict)
		# resources frame
		resFrame = wlib.wLabelFrame(rcol, text="Resources")
		resFrame.grid(row=1, column=0, sticky='new')
		fhweblinkFrame = ttk.Frame(resFrame)
		for link in data.fhome_links:
			weblbl = wlib.drawLinkLabel(fhweblinkFrame, label=link[0], url=link[1])
			weblbl.pack()
		fhweblinkFrame.grid(row=0, column=0, sticky='nw')	
		
		''' STATUS BAR ROW START '''
		''' -------------------- '''
		statusbar = ttk.Label(self, 
							  textvariable=self.statusTxt,
							  style='statusbar.TLabel')
		statusbar.grid(row=2, column=0, columnspan=3, sticky='new')
		
	# - END GUI METHODS - #
	# - FUNCTIONALIT METHODS - #
	def makeDataDict(self, labels):
		dataDict = {k : tk.StringVar() for k in labels}
		return dataDict
		
	def createWordDoc(self, datadicts):
		for datadict in datadicts:
			for k, v in datadict.items():
				print("{} : {}".format(k, v.get()))
		
	def updateStatus(self, msg):
		self.statusTxt.set(msg)
		
	# menu command callback functions
	def m_help_about(self):
		messagebox.showinfo(title="About Money Penny App",
							message="Money Penny Office App v1.0")	
		
if __name__ == '__main__':
	root = App()
	root.geometry("960x720+80+80")
	root.minsize(960,720)
	"""
	root.update_idletasks()				# needed so widgets can be redrawn before next idle
	root.overrideredirect(True)			# removes window decorations (titlebar)
	w = root.winfo_screenwidth()		# get width of main window
	h = root.winfo_screenheight()		# get height of main window
	root.resizable(height=0, width=0)	# user cannot resize window
	"""
	root.mainloop()
