from tkinter import *
from io import open
from os import startfile
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
from os.path import isfile
from util import splitFileFolderAndName
from excelSheet import excelSheet
from xmlTool import xmlTool

class first(Frame):
	"""The first interface which has options to read a XML file or an Excel spread sheet.

		Parameters
		----------
		parent: Tk
			A window for an application.
	"""
	def __init__(self, parent):
		""" Creat the first interface."""

		self.parent = parent # An attribute for its main window
		# Start to creat GUI 
		self.initGUI()

	def initGUI(self):
		"""Configure the entire GUI including buttons and their loctions."""

		# Name the title of the main window
		self.parent.title("Reference Copier")
		# Configure the frame that contains buttons and its loction
		self.buttonFrame = Frame(self.parent, width = 270, borderwidth = 1)
		self.buttonFrame.pack(fill = BOTH, expand = True)
		# Create buttons
		self.makeButtons()

	def makeButtons(self):
		"""Create all buttons."""

		# Cancel button with its command and location
		bCancel = Button(self.buttonFrame, text = "Cancel", width = 10 ,command = self.closeWindow)
		bCancel.pack(side = RIGHT, padx = 5, pady = 5)
		# Import button with its command and location
		bImport = Button(self.buttonFrame, text = "Import Sheet", width = 15,command = self.importSheet)
		bImport.pack(side = RIGHT, padx = 5, pady = 5)
		# Start button with its command and location
		bStart = Button(self.buttonFrame, text = "Start New", width = 10, command = self.startNew)
		bStart.pack(side = RIGHT, padx = 5, pady = 5)

	def closeWindow(self):
		"""Close a window."""

		self.parent.destroy()

	def hideWinodw(self):
		"""Hide a window."""

		self.parent.withdraw()

	def showWindow(self):
		"""Show a window that is hidden."""

		self.parent.deiconify()

	def startNew(self):
		"""Creat another interface that browses XML files."""

		# Hide the first interface
		self.hideWinodw()
		# Create a XML browse interface
		windowBX = Toplevel()
		startN = browse(windowBX, self, True)
		windowBX.mainloop()

	def importSheet(self):
		"""Creat another interface that browses Excel spreadsheets."""

		# Hide the first interface
		self.hideWinodw()
		# Create a spreadsheet browse interface
		windowBS = Toplevel()
		importS = browse(windowBS, self, False)
		windowBS.mainloop()

class errorMessage(Frame):
	"""A class of error message interface that presents saves(optional) error messages.

		Parameters
		----------
		parent: Tk
			A window for an application.
		message: string
			An error message.
		textFilePath: string
			An error message text file destination. 
	"""

	def __init__(self, parent, textFilePath, message):
		"""Creat an error message interface."""

		self.parent = parent				# Main window
		self.textFilePath = textFilePath	# File destination to save to
		self.message = message 				# Error message
		# Creat GUI
		self.initGUI()

	def initGUI(self):
		"""Creat GUI including a title, message, and buttons."""

		# Set window's title
		self.parent.title("Error Message")
		# Creat frames that contain messages and buttons 
		self.buttonFrame = Frame(self.parent)
		self.buttonFrame.pack(fill = BOTH, expand = True)
		messageFrame = Frame(self.buttonFrame, borderwidth = 1)
		messageFrame.pack(fill = BOTH, expand = True)
		# Creat buttons
		self.makeButtons()
		# Create and show an error message as an label
		var = StringVar()
		label = Message(messageFrame, textvariable=var, relief=RAISED, width = 1000)
		var.set(self.message)
		label.pack(fill = BOTH, expand = True)

	def makeButtons(self):
		"""Create all buttons."""

		# Create save and ok buttons and set their locations
		bSave = Button(self.buttonFrame, text = "Save", width = 5, command = self.writeToText)
		bSave.pack(side = RIGHT, padx=5, pady=2)
		bOk = Button(self.buttonFrame, text = "Ok", width = 5, command = self.parent.destroy)
		bOk.pack(side = RIGHT, padx=3, pady=2)

	def writeToText(self):
		"""Write an error message into a text file and save it in the current directory."""

		# Create a text file for write only mode in the current directory
		file = open(self.textFilePath,'w')
		# Write a message into file 
		file.write(self.message)
		# Close file
		file.close()
		# Open the message file in windows
		startfile(self.textFilePath)
		# Close the interface
		self.parent.destroy()

class browse(Frame):
	"""A class for browse interface that browses XML or spreadsheets.

		Parameters
		----------
		parent: Tk
			A window for applications.
		mainLevel: first
			The first interface (class)
		isXML: bool
			True if it's to browse XML. False otherwise.
	"""

	def __init__(self, parent, mainLevel, isXML):
		"""Create a browse interface."""

		# Create a main frame
		Frame.__init__(self, parent, width = 1000)
		self.parent = parent		# Window for interface
		self.mainLevel = mainLevel 	# first interface
		self.filePath = ""			# Final file path
		self.filePathEntry = None	# File path in browse entry
		self.isXmlNotXlsx = isXML 	# bool (if it's to browse XML files)
		self.withFocus = IntVar()	# IntVar for focus height 
		# Create GUI
		self.initGUI()

	def initGUI(self):
		"""Create GUI including buttons and path entry."""

		# Name window's title
		self.parent.title("Reference Copying")
		# Set main frame's location 
		self.pack(fill = BOTH, expand = True)
		# Set path entry frame and its location
		self.entryFrame = Frame(self, relief = RAISED, borderwidth = 1)
		self.entryFrame.pack(fill = BOTH, expand = True)
		# Create buttons
		self.makeButtons()
		# Set path entry and its location
		self.filePathEntry = Entry(self.entryFrame, bd = 4, width = 50)
		self.filePathEntry.grid(row = 0, column = 2, columnspan = 5, padx=2, pady=2)

	def makeButtons(self):
		"""Create buttons."""

		# If it's to browse XML files, name button 'Browse xml', Otherwise, 'Browse xlsx or xlsm'
		if self.isXmlNotXlsx:
			browseText = "Browse xml"
			width = 11
			# Add checkbox for focus height
			checkbox = Checkbutton(self, text = "Focus Height", variable = self.withFocus, onvalue = 1, offvalue = 0, height=1, width =12)
			checkbox.pack(side = LEFT, padx=4, pady=2)
		else:
			browseText = "Browse xlsx or xlsm"
			width = 17
		# Set browse button and location
		bBrowse = Button(self.entryFrame, text = browseText, width = width, command = self.getFilePath)	
		bBrowse.grid(row = 0, column = 1, padx=3, pady=3)
		# Set cancel button and location
		bCancel = Button(self, text = "Cancel", width = 10 ,command = self.closeMainAndToplevelWindow)
		bCancel.pack(side = RIGHT,padx=4, pady=2)
		# Set back button and location
		bBack = Button(self, text = "Back", width = 7, command = self.back)
		bBack.pack(side = RIGHT,padx=4, pady=2)
		# Set ok button and lcation
		bOk = Button(self, text = "Ok", width = 5, command = self.OK)
		bOk.pack(side = RIGHT, padx=4, pady=2)

	def getFilePath(self):
		"""Get file Path from file path entry."""

		# If it's to browse XML, give file types to .xml, otherwise .xlsx and .xlsm and get a filepath from file dialog
		if self.isXmlNotXlsx:
			fileType = ("XML file", "*.xml")
			# Open file dialog 
			self.filePath = askopenfilename(filetypes = (fileType, ("All files", "*.*")), parent = self.parent)
		else:
			fileType1 = ("Excel Workbook", "*.xlsx")
			fileType2 = ("Excel Macro-Enabled Workbook", "*.xlsm")
			self.filePath = askopenfilename(filetypes = (fileType2, fileType1, ("All files", "*.*")), parent = self.parent)

		# Once self.filePath gets a filepath, delete what's in the entry and put self.filePath into the entry
		self.filePathEntry.delete(0, 'end')
		self.filePathEntry.insert(0, self.filePath)

	def closeMainAndToplevelWindow(self):
		"""Close main window and toplevel window."""

		# Closing the main window automatically closes toplevel window
		self.mainLevel.closeWindow()

	def closeWindow(self):
		"""Close the toplevel window."""

		self.parent.destroy()

	def OK(self):
		"""A set of instructions are ran when ok button is clicked. 

			First get a file path from browse entry and make sure it is valid. If it is a XML file,
			read it and creat an Excel spread sheet if no error occurs. Otherwise, file is an Excel
			sheet. Then read it and generate a new XML file if no error occurs. Error messages are 
			presented based on their types such string, list, and None. See detail descriptions in 
			each methods.		
		"""

		# Get a file path from browe entry
		self.filePath = self.filePathEntry.get()
		# When entry is empty						
		if self.filePath == "":
			self.emptyFileNameWarning()
		# When file path is invaild meaning not a legit file
		elif not isfile(self.filePath):
			self.incorrectFileNameWarning()
		else:
			# If file is a XML
			if self.isXmlNotXlsx:
				# Read xml and create an Excel sheet
				result = readXMLAndStartSheet(self.filePath, self.withFocus.get())
			else:
				# Read an Excel sheet and create a new XML file
				result = readSheetAndModifyXML(self.filePath)
			if result[0]:
				# Close both windows
				self.closeMainAndToplevelWindow()
				# Creat a window
				errorWindow = Tk()
				# Create an error message interface
				errorMessage(errorWindow, result[0], result[1])
				# Launch window
				errorWindow.mainloop()
			elif result[1]:
				# Show an error message in a pop up window
				self.popErrorMessage(result[1])
			else:
				# No error and close both windows
				self.closeMainAndToplevelWindow()
			
	def back(self):
		"""Go back to the first interface."""

		# Close the current window and unhide the first interface
		self.closeWindow()
		self.mainLevel.showWindow()

	def incorrectFileNameWarning(self):
		"""Warning when file path is incorrect(file does not exist)."""

		messagebox.showinfo("Warning", "File does not exist!", parent = self.parent)

	def emptyFileNameWarning(self):
		"""Warning when file path entry is empty but ok is clicked."""

		messagebox.showinfo("Warning", "No files selected!", parent = self.parent)

	def popErrorMessage(self, message):
		"""Show error message in a pop up window.

			Parameters
			----------
			message: string
				An error message.
		"""

		messagebox.showinfo("Warning", message, parent = self.parent)

def readXMLAndStartSheet(filePath, withFocus):
	"""Get data from XML and present them in a Excel spreadsheet.

		Parameters
		----------
		filePath: string
			A xml file path.
		withFocus: int
			1: with focus. 0: without focus

		Returns
		-------
		file path: string
			An error message for incorrect file format.

		or

		(errorFilePath, info): tuple
			When there is any repeating referenece names. 
				errorFilePath: string
					Info text file path.
				info: string
					Reference systems information.
	"""

	# Split file path into folder path and file name without extension
	folderPath, fileName = splitFileFolderAndName(filePath)
	# Read xml and get all references' names, missing references' names, types, dependon, and  wire count information
	refInfo, wireInfo = xmlTool.readXML(filePath)
	# If reference name list is None or wireInfo dictionary is None, file format is incorrect
	if not refInfo and not wireInfo:
		return ("", "File: " + fileName + " - format incorrect!")
	if not refInfo and 0 in wireInfo:
		return ("", wireInfo[0])
	# If there is repeating reference name
	if refInfo['repeats']: 
		# Generate a text of information of xml data       
		info = xmlTool.XMLInfo(filePath, refInfo['repeats'], refInfo['name'], refInfo['gap'], wireInfo['total'])
		# Create a error text file path to save to
		errorFilePath = folderPath + '/' +  fileName + '_info.txt'
		return (errorFilePath, info)
	else:  
		# Create an excelSheet object
		if withFocus:
			excelWrite = excelSheet(True)
		else:
			excelWrite = excelSheet(False)
		# Write data into an Excel spreadsheet
		error = excelWrite.startNewExcelSheet(filePath, refInfo, wireInfo)
		return ("", error)

def readSheetAndModifyXML(filePath):
	""" Function that reads a Excel sheet and modify a XML file.

		Parameters
		----------
		filePath: string
			xlsm or xlsx file path.

		Returns
		-------
		(errorFilePath, error): tuple
			Errors occur when reading Excel sheets or writing data into xml files.
			error: string
				Error message.
			errorFilePath: string
				File path of an error text file.
	"""

	# Split file path into folder path and file name without extension
	folderPath, fileName = splitFileFolderAndName(filePath)
	# Create an excelRead object
	excelRead = excelSheet(None)
	# Read excel spreadsheet
	xmlPath, refExcelDict, error = excelRead.readExcelSheet(filePath)
	# If xmlPath and refExcelDict are not None
	if xmlPath and refExcelDict:
		# If error is not None, meaning there is an error
		if error:
			# Create an error text file path
			errorFilePath = folderPath + '/' + fileName + '_error.txt'
			return (errorFilePath, error)
		else:
			# Create a new xml file with modified data
			newXmlFilePathOrError = xmlTool.modifier(xmlPath, refExcelDict)
			# If newXmlFilePathOrError path is a file
			if isfile(newXmlFilePathOrError):
				# Open the new xml file in windows
				startfile(newXmlFilePathOrError)
				# No errors
				return ("", error)
			# There is an error--newXmlFilePathOrError is an error message
			return ("", newXmlFilePathOrError)
	# Errors occur when reading Excel sheet
	return ("", error)


def main():
	"""Function that starts the first interface."""
	# Create a window
	window = Tk()
	# Create a first object
	firstW = first(window)
	# Launch
	window.mainloop()

# Start program
main()