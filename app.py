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
		# main data dictionary population
		self.serviceDict = self.makeDataDict(data.srv_Labels)
		self.serviceDict["Service Type"].set("Full Funeral")
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
									command=lambda: self.createWordDoc(self.serviceDict))
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
		funInfoFrame = wlib.wLabelFrame(lcol, text="Service Information")
		funInfoFrame.grid(row=0, column=0, sticky='new')
		
		## funeral number
		funNumFrame = ttk.Frame(funInfoFrame)
		funNumFrame.grid(row=0, column=0, sticky='nw', padx=4, pady=6)
		funNumLabel = ttk.Label(funNumFrame, text="Service Number")
		funNumLabel.grid(row=0, column=0)
		funNumEnt = ttk.Entry(funNumFrame, width=8)
		funNumEnt["textvariable"] = self.serviceDict["Service Number"]
		funNumEnt.grid(row=0, column=1)
		
		## service type
		serviceTypeFrame = ttk.Frame(funInfoFrame)
		serviceTypeFrame.grid(row=1, column=0, sticky='nw', padx=4, pady=6)
		serviceTypeLbl = ttk.Label(serviceTypeFrame, text="Select Service Type")
		serviceTypeLbl.grid(row=0, column=0, sticky='nw')
		
		strad_frame = wlib.drawRadios(serviceTypeFrame,
									  labels=data.fun_types,
						              var=self.serviceDict["Service Type"])
		
		strad_frame.grid(row=1, column=0, sticky='nw')
		
		## funeral home info combo and entries
		fhomeFrame = ttk.Frame(funInfoFrame)
		fhomeFrame.grid(row=2, column=0, sticky='nw', padx=4, pady=6)
		
		cbofra = wlib.drawCombo(fhomeFrame, width=10, label="Funeral Home", vals=data.fun_homes,
								var=self.serviceDict["Funeral Home"])
		cbofra.grid(row=0, column=0, sticky='nw')
		
		fconFrame = ttk.Frame(fhomeFrame)
		fconFrame.grid(row=0, column=1, sticky='nw', padx=10)
		fconLabel = ttk.Label(fconFrame, text="F. H. Contact").grid(row=0, column=0, sticky='nw')
		self.fcontactEntry = ttk.Entry(fconFrame, width=28, 
									   textvariable=self.serviceDict["F. H. Contact"])
		self.fcontactEntry.grid(row=1, column=0, sticky='nw')
		
		fhomeNumFrame = ttk.Frame(fhomeFrame)
		fhomeNumFrame.grid(row=1, column=1, sticky='ne', padx=10)
		fhomeNumLabel = ttk.Label(fhomeNumFrame, text="Contact Phone #")
		fhomeNumLabel.grid(row=0, column=0)
		self.fconNumEntry = ttk.Entry(fhomeNumFrame, width=16,
									  textvariable=self.serviceDict["Contact Phone"])
		self.fconNumEntry.grid(row=1, column=0)		
		
		## date - time - place combo boxes
		dtpFrame = ttk.Frame(funInfoFrame)
		dtpFrame.grid(row=3, column=0, sticky='nw', padx=4, pady=6)
		wlib.drawDate(dtpFrame, label='Date', width=8,
					  var=self.serviceDict["Service Date"]).grid(row=0, column=0)
		wlib.drawCombo(dtpFrame, width=6, label='Time', vals=data.times, 
					   var=self.serviceDict["Service Time"]).grid(row=0, column=1, padx=6)
		wlib.drawCombo(dtpFrame, width=18, label='Location', vals=data.fun_places,
					   var=self.serviceDict["Service Location"]).grid(row=0, column=2)
					   
		## day - celebrant combo boxes
		dcFrame = ttk.Frame(funInfoFrame)
		dcFrame.grid(row=4, column=0, sticky='nw', padx=4, pady=6)
		wlib.drawCombo(dcFrame, width=8, label="Day",
					   var=self.serviceDict["Service Day"], 
					   vals=data.days).grid(row=0, column=0)
		wlib.drawCombo(dcFrame, width=16, label="Celebrant",
					   var=self.serviceDict["Celebrant"], 
					   vals=data.fun_celebrants).grid(row=0, column=1, padx=4)
					   
		## organist - cantor -servers
		ocsFrame = ttk.Frame(funInfoFrame)
		ocsFrame.grid(row=5, column=0, sticky='nw', padx=4, pady=6)
		organistLbl = ttk.Label(ocsFrame, text="Organist")
		organistLbl.grid(row=0, column=0, sticky='nw')
		self.organistEnt = ttk.Entry(ocsFrame, width=24,
						   textvariable=self.serviceDict["Organist"]).grid(row=1, column=0, sticky='nw')
		cantorLbl = ttk.Label(ocsFrame, text="Cantor")
		cantorLbl.grid(row=0, column=1, sticky='nw', padx=4)
		self.cantorEnt = ttk.Entry(ocsFrame, width=24,
						 textvariable=self.serviceDict["Cantor"]).grid(row=1, column=1, padx=4)
								   
		srvFrame = ttk.Frame(funInfoFrame)
		srvFrame.grid(row=6, column=0, sticky='nw', padx=4, pady=6)
		srvNameFrame = ttk.Frame(srvFrame)
		srvNameFrame.grid(row=0, column=0)
		lblctr = 0		# need to seed labels at 0
		entctr = 1		# seed entries at 1 and add 2 to both to
						# maintain column seperation and correct row numbers
		for i in range(3):
			keystr = "Server " + str(i+1)
			lbl = ttk.Label(srvNameFrame, text=keystr).grid(row=lblctr, column=0, sticky='nw')
			lblctr += 2
			self.serverEnt = ttk.Entry(srvNameFrame, width=24,
									   textvariable=self.serviceDict[keystr])
			self.serverEnt.grid(row=entctr, column=0, sticky='nw')
			entctr += 2
			
		srvImgFrame = ttk.Frame(srvFrame)
		srvImgFrame.grid(row=0, column=1, sticky='nsew')
		srvImgLabel = wlib.drawImgLabel(srvImgFrame, 'img/servers.png')
		srvImgLabel.grid(row=0, column=0, sticky='nsew', pady=4)			
		
		''' MIDDLE COLUMN START '''
		''' ------------------- '''
		mcol = ttk.Frame(self, width=300, style='column.TFrame')
		mcol.grid(row=1, column=1, sticky='nsew', pady=10)
		mcol.grid_propagate(0)				# prevent column from resizing to inner widget size
		mcol.columnconfigure(0, weight=1)	# allows child widgets to fill space of column
		# - Deceased Personal Info Frame
		dpiFrame = wlib.wLabelFrame(mcol, text="Deceased Personal Info")
		dpiFrame.grid(row=0, column=0, sticky='new')
		dpiDict = wlib.makeInfoDict(data.dpi_Labels, tk.StringVar())
		wlib.drawEntries(dpiFrame, dpiDict)
				
		''' RIGHT COLUMN START '''
		''' ------------------ '''
		rcol = ttk.Frame(self, width=300, style='column.TFrame')
		rcol.grid(row=1, column=2, sticky='nsew', padx=10, pady=10)
		rcol.grid_propagate(0)				# prevent column from resizing to inner widget size
		rcol.columnconfigure(0, weight=1)	# allows child widgets to fill space of column		
		
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
		
	def createWordDoc(self, datadict):
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
