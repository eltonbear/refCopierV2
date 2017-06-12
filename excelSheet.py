import xlsxwriter
from tkinter import Tk
from os import startfile
import openpyxl as op
from util import splitFileFolderAndName
import re

class excelSheet():
	def __init__(self):
		### Set columns
		self.statusC = 'A'
		self.refC = 'B'
		self.copyC = 'C'
		self.typeC = 'D'
		self.depC = 'E'
		self.wireSCountC = 'F'
		self.wireDCountC = 'G'
		self.wireNewDcountC = 'H'
		self.hiddenRefC='U'
		self.vbaButtonC = 'J'
		self.pseudoRefC = 'L'
		self.realRefC = 'M'
		self.pseudoCountC = 'N'
		self.wirePseudoCountSC = 'O'
		self.wirePseudoCountDC = 'P'
		### Set rows
		self.titleRow = '1'
		self.firstInputRow = str(int(self.titleRow) + 1)
		self.pseudoTitleRow = '6'
		### Set cell address
		self.xmlFilePathCell ='M1'
		self.wireTagCell = 'M3'
		self.wireCountCell = 'M4'
		self.hiddenRowsC = 'I'
		self.lastAppendRowCell = self.hiddenRowsC + '4'
		self.appendRowCountCell = self.hiddenRowsC + '3' 
		self.lastRefRowBeforeMacroCell = self.hiddenRowsC + '2'
		self.hiddenLastExistingRefRowCell = self.hiddenRowsC + '1'
		### Set rag names
		self.mTag = 'missing'
		self.eTag = 'existing'
		self.aTag = 'appending'
		self.workSheetName = 'Reference_copying'
		self.copyBlockedText = 'BLOCKED'

	def startNewExcelSheet(self, xmlFilePath, refInfo, wireSDInfo):
		"""

			parameters
			----------
			refGap: list
			
		"""
		
		refNumList = refInfo['name']
		refGap = refInfo['gap']
		numOfGap = len(refGap)
		### if the number of gaps > the number of existing references, it's an error
		if numOfGap > len(refNumList):
			return "The number of missing refs: " + str(numOfGap) + " > the number of existing refs: " + str(len(refNumList))

		### get folder name and xml file name without extension/ name xlsm file path
		xmlFolderPath, xmlFileName = splitFileFolderAndName(xmlFilePath)
		xlsxFileName = xmlFileName + '_instruction.xlsm'
		xlsxFilePath = xmlFolderPath + '/' + xlsxFileName
		### creat workbook and worksheet
		workbook = xlsxwriter.Workbook(xlsxFilePath)
		worksheet = workbook.add_worksheet(self.workSheetName)

		### add cell format
		unlocked = workbook.add_format({'locked': 0, 'valign': 'vcenter', 'align': 'center'})
		centerF = workbook.add_format({'valign': 'vcenter', 'align': 'center'})
		centerHiddenF = workbook.add_format({'valign': 'vcenter', 'hidden': 1, 'align': 'center'})
		titleF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#b8cce0', 'font_color': '#1f497d', 'bold': True, 'bottom': 2, 'bottom_color': '#82a5d0'})
		topBorderF = workbook.add_format({'top': 2, 'top_color': '#82a5d0'})
		copyBlockedF = workbook.add_format({'bg_color': '#a6a6a6', 'font_color': '#a6a6a6'})
		missingTagAndRefF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'border_color': '#b2b2b2'})
		missingUnblockedF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#FFC7CE', 'locked': 0, 'border': 1, 'border_color': '#b2b2b2'})
		missingDepBlockedF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#FFC7CE', 'locked': 1, 'hidden': 1,'border': 1, 'border_color': '#b2b2b2'})
		missingDepBlockedBlankF = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#FFC7CE', 'locked': 1, 'hidden': 1,'border': 1, 'border_color': '#b2b2b2'})
		missingWireCountF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#FFC7CE', 'border': 1, 'border_color': '#b2b2b2'})
		missingWireCountHiddenF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#FFC7CE', 'hidden': 1,'border': 1, 'border_color': '#b2b2b2'})
		existingWhiteBlockedF = workbook.add_format({'font_color': 'white', 'locked': 1, 'hidden': 1})
		appendTagAndRefF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'font_color': 'white', 'bg_color': '#92cddc', 'locked': 1, 'border': 1, 'border_color': '#b2b2b2'})
		appendUnblockedF =  workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#92cddc', 'locked': 0,'border': 1, 'border_color': '#b2b2b2'})
		appendBlockedF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#92cddc', 'locked': 1, 'hidden': 1, 'border': 1, 'border_color': '#b2b2b2'})
		appendHiddenZeroBlockedF = workbook.add_format({'bg_color': '#92cddc', 'font_color': '#92cddc', 'locked': 1, 'hidden': 1, 'border': 1, 'border_color': '#b2b2b2'})
		pseudoRefLetter = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#c6efce', 'font_color': '#006100'})
		pseudoCounts = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#c6efce'})

		### activate protection with password "elton"
		worksheet.protect('elton')

		### set column width and protection
		wireTagC = re.findall("[a-zA-Z]+", self.wireTagCell)[0]
		worksheet.set_column(self.statusC + ':' + self.statusC, 9)
		worksheet.set_column(self.refC + ':' + self.refC, 20)
		worksheet.set_column(self.typeC + ':' + self.typeC, 14)
		worksheet.set_column(self.depC + ':' + self.depC, 17)
		worksheet.set_column(self.wireSCountC + ':' + self.wireSCountC, 12)
		worksheet.set_column(self.wireDCountC + ':' + self.wireDCountC, 12)
		worksheet.set_column(self.wireNewDcountC + ':' + self.wireNewDcountC, 16)
		worksheet.set_column(self.hiddenRowsC + ':' + self.hiddenRowsC, 5)
		worksheet.set_column(wireTagC + ':' + wireTagC, 10)
		worksheet.set_column(wireTagC + ':' + wireTagC, 10)

		### write title
		worksheet.write(self.statusC + self.titleRow, 'Status', titleF)
		worksheet.write(self.refC + self.titleRow, 'Reference Number (R)', titleF)
		worksheet.write(self.copyC + self.titleRow, 'Copy (R)', titleF)
		worksheet.write(self.typeC + self.titleRow, 'Reference Type', titleF)
		worksheet.write(self.depC + self.titleRow, 'Dependent On (R)', titleF)
		worksheet.write(self.wireSCountC + self.titleRow, 'Wire Count S', titleF)
		worksheet.write(self.wireDCountC + self.titleRow, 'Wire Count D', titleF)
		worksheet.write(self.wireNewDcountC + self.titleRow, 'Wire New Count D', titleF)
		worksheet.write(self.wireTagCell, "Wire Count", centerF)
		worksheet.write(self.wireCountCell, wireSDInfo['total'], centerF)
		worksheet.write(self.xmlFilePathCell, 'XML: ' + xmlFilePath)

		### write rows
		lastRefRow = int(self.titleRow) + int(refNumList[-1])
		fstAppendRow = str(int(lastRefRow) + 1)
		lastAppendRow = str(int(lastRefRow) + len(refNumList) - numOfGap)
		lastHiddenRefRow = str(len(refNumList))

		# Get pseudo Reference 
		pseudo = refInfo['pseudo'] # A dictionary
		if pseudo:
			worksheet.set_column(self.pseudoRefC + ':' + self.pseudoRefC, 18)
			worksheet.set_column(self.realRefC + ':' + self.realRefC, 19)
			worksheet.set_column(self.pseudoCountC + ':' + self.pseudoCountC, 6)
			worksheet.set_column(self.wirePseudoCountSC + ':' + self.wirePseudoCountSC, 12)
			worksheet.set_column(self.wirePseudoCountDC + ':' + self.wirePseudoCountSC, 12)

			worksheet.write(self.pseudoRefC + self.pseudoTitleRow, 'Pseudo Reference (R)', titleF)
			worksheet.write(self.realRefC + self.pseudoTitleRow, 'Reference Number (R)', titleF)
			worksheet.write(self.pseudoCountC + self.pseudoTitleRow, 'Count', titleF)
			worksheet.write(self.wirePseudoCountSC + self.pseudoTitleRow, 'Wire Count S', titleF)
			worksheet.write(self.wirePseudoCountDC + self.pseudoTitleRow, 'Wire Count D', titleF)

			pseudoRefRowS = str(int(self.pseudoTitleRow) + 1)
			sortedPseudo = sorted(pseudo.keys())
			numPseudo = len(sortedPseudo)
			for pseudoRef in sortedPseudo:
				worksheet.write(self.pseudoRefC + pseudoRefRowS , pseudoRef, pseudoRefLetter)
				worksheet.write(self.realRefC + pseudoRefRowS, None, unlocked)
				worksheet.write(self.pseudoCountC + pseudoRefRowS , pseudo[pseudoRef], pseudoCounts)
				worksheet.write(self.wirePseudoCountSC + pseudoRefRowS, len(wireSDInfo[pseudoRef]['s']), pseudoCounts)
				worksheet.write(self.wirePseudoCountDC + pseudoRefRowS, len(wireSDInfo[pseudoRef]['d']), pseudoCounts)
				f1 = 'COUNTIF($' + self.realRefC + '$' + str(int(self.pseudoTitleRow)+1) + ':$' + self.realRefC + '$' + str(int(self.pseudoTitleRow)+numPseudo) + ',' + self.realRefC + pseudoRefRowS + ')=1'
				f2 = 'COUNTIF($' + self.hiddenRefC + '$1' + ':$' + self.hiddenRefC + '$' + lastHiddenRefRow + ',' + self.realRefC + pseudoRefRowS + ')=1'
				pseudoRefFormula = '=AND(' + f1 + ', ' + f2 + ')'
				worksheet.data_validation(self.realRefC + pseudoRefRowS, {'validate': 'custom', 'value': pseudoRefFormula, 'error_title': 'Warning', 'error_message': 'Reference does not exist or Duplicates!', 'error_type': 'stop'})
				pseudoRefRowS = str(int(pseudoRefRowS) + 1)

		typeList = ['Default','NEWBENCH', 'AVX_S', 'AVX_D', 'PRSID_S', 'PRSID_D', 'TECDIA_S', 'TECDIA_D', 'PKGFLOOR', 'IC_S', 'IC_D', 'SIGE', 'RESIST_S', 'RESIST_D', 'TF_S', 'TF_D', 'COIN_S', 'COIN_D']
		refGapSet = set(refGap)
		refNumber = 1	
		refListIndex = 0

		for rowN in range(int(self.firstInputRow), int(lastAppendRow) + 1):
			rowS = str(rowN)
			if rowN < int(fstAppendRow):
				if str(refNumber) in refGapSet: ### missing ref row
					worksheet.write(self.statusC + rowS, self.mTag,  missingTagAndRefF)
					worksheet.write(self.refC + rowS, refNumber,  missingTagAndRefF)
					worksheet.write(self.copyC + rowS, None,  missingUnblockedF)
					worksheet.write(self.typeC + rowS, None,  missingUnblockedF)
					worksheet.write(self.wireSCountC + rowS, 0, missingWireCountF)
					worksheet.write(self.wireDCountC + rowS, 0, missingWireCountF)
					### formulas for dependon cells
					worksheet.write_formula(self.depC + rowS, '=' + self.copyC + rowS, missingDepBlockedF)
					worksheet.conditional_format(self.depC + rowS, {'type': 'cell', 'criteria': 'equal to', 'value': 0, 'format': missingDepBlockedBlankF})
					### formulas for wire new D Count cell
					wireNewDFormula = '=IF(ISBLANK(' + self.copyC + rowS + '), 0, INDIRECT("' + self.wireDCountC + '"& ' + self.copyC + rowS + '+1))'
					worksheet.write_formula(self.wireNewDcountC + rowS, wireNewDFormula, missingWireCountHiddenF)
					### wire formula for data validation. it prevents duplicates and anything outside the list
					f1 = 'COUNTIF($' + self.copyC + '$' + self.firstInputRow + ':$' + self.copyC + '$' + lastAppendRow + ',' + self.copyC + rowS + ')=1'
					f2 = 'COUNTIF($' + self.hiddenRefC + '$1' + ':$' + self.hiddenRefC + '$' + lastHiddenRefRow + ',' + self.copyC + rowS + ')=1'
					countFormula = '=AND(' + f1 + ', ' + f2 + ')'
					worksheet.data_validation(self.copyC + rowS, {'validate': 'custom', 'value': countFormula, 'error_title': 'Warning', 'error_message': 'Reference number does not exist or Duplicates!', 'error_type': 'stop'}) 
					# Data validation for types
					worksheet.data_validation(self.typeC + rowS, {'validate': 'list', 'source': typeList, 'error_title': 'Warning', 'error_message': 'Type does not exist in the library!', 'error_type': 'stop'}) 
				else:  ### existing ref row
					worksheet.write(self.statusC + rowS, self.eTag, existingWhiteBlockedF)
					worksheet.write(self.refC + rowS, refNumber, centerF)
					worksheet.write(self.hiddenRefC + str(refListIndex+1), int(refNumList[refListIndex]), existingWhiteBlockedF)
					worksheet.write(self.copyC + rowS, self.copyBlockedText, copyBlockedF)
					worksheet.write(self.typeC + rowS, refInfo['type'][refListIndex],  unlocked)
					worksheet.write(self.depC + rowS, refInfo['dependon'][refListIndex],  centerF)
					### data validation for dep
					listF = 'COUNTIF($' + self.hiddenRefC + '$1' + ':$' + self.hiddenRefC + '$' + lastHiddenRefRow + ',' + self.depC + rowS + ')=1'
					worksheet.data_validation(self.depC + rowS, {'validate': 'custom', 'value': listF, 'error_title': 'Warning', 'error_message': 'Reference does not exist!', 'error_type': 'stop'})
					# Data validation for types
					worksheet.data_validation(self.typeC + rowS, {'validate': 'list', 'source': typeList, 'error_title': 'Warning', 'error_message': 'Type does not exist in the library!', 'error_type': 'stop'}) 
					### formulas for Wire new D Count cell
					wireRefSCount = len(wireSDInfo[str(refNumber)]['s'])
					wireRefDCount = len(wireSDInfo[str(refNumber)]['d'])
					if pseudo:
						wireCountSFormula = '=IF(COUNTIF(' + self.realRefC +  str(int(self.pseudoTitleRow) + 1) + ':' + self.realRefC + str(int(self.pseudoTitleRow) + numPseudo) + ', ' + str(refNumber) + ') > 0, '\
											+ str(wireRefSCount) + ' + INDIRECT("' + self.wirePseudoCountSC + '" & MATCH(' + str(refNumber) + ', ' + self.realRefC + str(int(self.pseudoTitleRow) + 1) + ':' + self.realRefC + str(int(self.pseudoTitleRow) + numPseudo)\
											+ ', 0) + ' + self.pseudoTitleRow +'), ' + str(wireRefSCount) + ')'
						wireCountDFormula = '=IF(COUNTIF(' + self.realRefC +  str(int(self.pseudoTitleRow) + 1) + ':' + self.realRefC + str(int(self.pseudoTitleRow) + numPseudo) + ', ' + str(refNumber) + ') > 0, '\
											+ str(wireRefDCount) + ' + INDIRECT("' + self.wirePseudoCountDC + '" & MATCH(' + str(refNumber) + ', ' + self.realRefC + str(int(self.pseudoTitleRow) + 1) + ':' + self.realRefC + str(int(self.pseudoTitleRow) + numPseudo)\
											+ ', 0) + ' + self.pseudoTitleRow +'), ' + str(wireRefDCount) + ')'
						wireNewDFormula = '=IF(COUNTIF(' + self.copyC + self.firstInputRow + ':' + self.copyC + lastAppendRow + ', ' + str(refNumber) + ') > 0, 0, '\
											+ 'IF(COUNTIF(' + self.realRefC +  str(int(self.pseudoTitleRow) + 1) + ':' + self.realRefC + str(int(self.pseudoTitleRow) + numPseudo) + ', ' + str(refNumber) + ') > 0, '\
											+ str(wireRefDCount) + ' + INDIRECT("' + self.wirePseudoCountDC + '" & MATCH(' + str(refNumber) + ', ' + self.realRefC + str(int(self.pseudoTitleRow) + 1) + ':' + self.realRefC + str(int(self.pseudoTitleRow) + numPseudo)\
											+ ', 0) + ' + self.pseudoTitleRow +'), ' + str(wireRefDCount) + '))'										
						worksheet.write_formula(self.wireSCountC + rowS, wireCountSFormula, centerHiddenF)
						worksheet.write_formula(self.wireDCountC + rowS, wireCountDFormula, centerHiddenF)
						worksheet.write_formula(self.wireNewDcountC + rowS, wireNewDFormula, centerHiddenF)
					else:
						worksheet.write(self.wireSCountC + rowS, wireRefSCount, centerF)
						worksheet.write(self.wireDCountC + rowS, wireRefDCount, centerF)
						wireNewDFormula = '=IF(COUNTIF(' + self.copyC + self.firstInputRow + ':' + self.copyC + lastAppendRow + ', ' + str(refNumber) + ') > 0, 0, ' + str(wireRefDCount) + ')'
						worksheet.write_formula(self.wireNewDcountC + rowS, wireNewDFormula, centerHiddenF)

					refListIndex += 1
			else: ### append section
				if not refGap or rowS == fstAppendRow:
					worksheet.write(self.statusC + rowS, self.aTag, appendTagAndRefF)
					worksheet.write(self.refC + rowS, refNumber, appendTagAndRefF)
					worksheet.write(self.copyC + rowS, None, appendUnblockedF)
					worksheet.write(self.typeC + rowS, None,  appendUnblockedF)
					worksheet.write(self.wireSCountC + rowS, 0, appendBlockedF)
					worksheet.write(self.wireDCountC + rowS, 0, appendBlockedF)	
					### formulas for dep and Wire new D Count cell
					worksheet.write_formula(self.depC + rowS, '=' + self.copyC + rowS, appendBlockedF)
					wireNewDFormula = '=IF(ISBLANK(' + self.copyC + rowS + '), 0, INDIRECT("' + self.wireDCountC + '"& ' + self.copyC + rowS + '+1))'
					worksheet.write_formula(self.wireNewDcountC + rowS, wireNewDFormula, appendBlockedF)
				
				### conditional formats for wire counts --> dont show values if there is no input in copy column
				wireConditionalFormula = '=AND(ISBLANK(' + self.copyC + rowS + '), NOT(ISBLANK(' + self.statusC + rowS + ')))'
				worksheet.conditional_format(self.depC + rowS, {'type': 'formula', 'criteria': wireConditionalFormula, 'format': appendHiddenZeroBlockedF})
				worksheet.conditional_format(self.wireSCountC + rowS, {'type': 'formula', 'criteria': wireConditionalFormula, 'format': appendHiddenZeroBlockedF})
				worksheet.conditional_format(self.wireDCountC + rowS, {'type': 'formula', 'criteria': wireConditionalFormula, 'format': appendHiddenZeroBlockedF})
				worksheet.conditional_format(self.wireNewDcountC + rowS, {'type': 'formula', 'criteria': wireConditionalFormula, 'format': appendHiddenZeroBlockedF})

				### wire formula for datavalidation. it prevents duplicates and anything outside the list(it writes every row til the last appendable row becuase i dont want to data validation in VBA)
				f1 = 'COUNTIF($' + self.copyC + '$' + self.firstInputRow + ':$' + self.copyC + '$' + lastAppendRow + ',' + self.copyC + rowS + ')=1'
				f2 = 'COUNTIF($' + self.hiddenRefC + '$1' + ':$' + self.hiddenRefC + '$' + lastHiddenRefRow + ',' + self.copyC + rowS + ')=1'
				countFormula = '=AND(' + f1 + ', ' + f2 + ')'
				worksheet.data_validation(self.copyC + rowS, {'validate': 'custom', 'value': countFormula, 'error_title': 'Warning', 'error_message': 'Reference number does not exist or Duplicates!', 'error_type': 'stop'})
				# Data validation for types
				worksheet.data_validation(self.typeC + rowS, {'validate': 'list', 'source': typeList, 'error_title': 'Warning', 'error_message': 'Type does not exist in the library!', 'error_type': 'stop'}) 	
			
			refNumber = refNumber + 1
		### hidden info in excel sheet
		if not refGap or int(fstAppendRow) > int(lastAppendRow): ### meaning no gaps or no appending section:
			worksheet.write(self.lastRefRowBeforeMacroCell, rowN,  existingWhiteBlockedF)
			worksheet.write(self.appendRowCountCell, rowN,  existingWhiteBlockedF)
			if refGap: ### add a topborder color if there is no appending section
				worksheet.write(self.statusC + fstAppendRow, None, topBorderF)
				worksheet.write(self.refC + fstAppendRow, None, topBorderF)
				worksheet.write(self.copyC + fstAppendRow, None, topBorderF)
				worksheet.write(self.typeC + fstAppendRow, None, topBorderF)
				worksheet.write(self.depC + fstAppendRow, None, topBorderF)
				worksheet.write(self.wireSCountC + fstAppendRow, None, topBorderF)
				worksheet.write(self.wireDCountC + fstAppendRow, None, topBorderF)
				worksheet.write(self.wireNewDcountC + fstAppendRow, None, topBorderF)
		else:
			worksheet.write(self.lastRefRowBeforeMacroCell, int(fstAppendRow),  existingWhiteBlockedF)
			worksheet.write(self.appendRowCountCell, int(fstAppendRow),  existingWhiteBlockedF)
		worksheet.write(self.lastAppendRowCell, int(lastAppendRow),  existingWhiteBlockedF)
		worksheet.write(self.hiddenLastExistingRefRowCell, int(refNumList[-1]) + int(self.titleRow),  existingWhiteBlockedF)

		### import VBA
		workbook.add_vba_project('vbaProject.bin')
		workbook.set_vba_name("ThisWorkbook")
		worksheet.set_vba_name("Sheet1")
		### add VBA buttons
		worksheet.insert_button(self.vbaButtonC + str(lastRefRow - 1), {'macro': 'appendARow',
		                               								 	'caption': 'Append',
		                               								 	'width': 128,
		                              								 	'height': 40})

		worksheet.insert_button(self.vbaButtonC + str(int(fstAppendRow)+1), {'macro': 'undoRow',
		                               								 		 'caption': 'UnAppend',
		                               								 		 'width': 128,
		                              								 		 'height': 40})
		### merger two cells beteen the two buttons
		worksheet.merge_range(self.vbaButtonC + str(lastRefRow + 1) + ':' +  chr(ord(self.vbaButtonC)+1) + str(lastRefRow + 1), None)

		### add comment
		copyTitleComment = 'Input a name of any existing refernces from the XML file (number only).\nAll gaps need to be filled out'
		worksheet.write_comment(self.copyC + self.titleRow, copyTitleComment, {'author': 'Elton', 'width': 250, 'height': 50})
		worksheet.write_comment(self.depC + self.titleRow, "Double click to unlock or lock cells", {'author': 'Elton', 'width': 173, 'height': 16})
		if int(fstAppendRow) <= int(lastAppendRow):
			worksheet.write_comment(self.statusC + fstAppendRow, 'Optional Section', {'author': 'Elton', 'width': 100, 'height': 15})
		if refGap:
			worksheet.write_comment(self.statusC + str(int(refGap[0]) + int(self.titleRow)), 'Reference gaps in xml file' , {'author': 'Elton', 'width': 130, 'height': 15})

		### close workbook. error if there is same workbook open
		try:
			workbook.close()
		except PermissionError:
			message = "Please close the existing Excel Workbook!"
			return message

		startfile(xlsxFilePath)
		return ""

	def readExcelSheet(self, xlsxFilePath):
		try:
			workbook = op.load_workbook(filename = xlsxFilePath, read_only = True, data_only=True)
			worksheet = workbook.get_sheet_by_name(self.workSheetName)
		except op.utils.exceptions.InvalidFileException:
			_, fileName = splitFileFolderAndName(xlsxFilePath)
			message = "File: " + fileName + " - format incorrect!"
			return None, None, message
		except KeyError:
			message = "Cannot find excel sheet - " + self.workSheetName + "!"
			return None, None, message
			
		xmlFilePath = worksheet[self.xmlFilePathCell].value[5:]
		lastRow = worksheet[self.appendRowCountCell].value # int 
		### excelReference data structure --> {'og': {'refNum':[type, dependon]}, 'add': {'refNum': [copyNum, type]}, 'newRefName': [str(refNum)]}
		excelReference = {'og': {}, 'add': {}, 'newRefName': []}
		missingRef = []
		missingCopy = []
		missingType = []
		missingDep = []
		wrongSeqRow = []
		newRefName = []
		checkRepeatRef = set()
		allCopy = {}
		repeat = {}
		row  = self.firstInputRow
		prevAllExist = True
		error = False
		while int(row) <= lastRow:
			status = worksheet[self.statusC + row].value
			ref = str(worksheet[self.refC + row].value)
			copy = str(worksheet[self.copyC + row].value)
			typ = str(worksheet[self.typeC + row].value)
			dep = str(worksheet[self.depC + row].value)

			refExists = ref and ref != 'None'
			copyExists = copy and copy !='None'
			typeExists = typ and typ !='None'
			depExists = dep and dep != 'None' and dep != '0' ### with formula 
			depCellEmpty = dep == None or dep == 'None'      ### if gets modified by users and left empty
			if dep == 'None':
				dep = None
			if status == self.eTag: 
				if refExists and copyExists and typeExists and not error:
					excelReference['og'][ref] = [typ, dep]
				else:
					if not refExists:
						missingRef.append(row)
					if not typeExists:
						missingType.append(row)
					error = True
			elif status == self.mTag:
				if refExists and copyExists and typeExists and depExists and not error:
					excelReference['add'][ref] = [copy, typ]
					excelReference['newRefName'].append(ref)
				else:
					if not refExists:
						missingRef.append(row)
					if not copyExists:
						missingCopy.append(row)
					if not typeExists:
						missingType.append(row)
					if depCellEmpty:
						missingDep.append(row)
					error = True
			else: ### append
				if prevAllExist:					
					if refExists and copyExists and typeExists and depExists and not error:
						excelReference['add'][ref] = [copy, typ]
						excelReference['newRefName'].append(ref)
					elif refExists and not copyExists and not typeExists and not depExists:
						prevAllExist = False
					else:
						if not refExists:
							missingRef.append(row)
						if not copyExists:
							missingCopy.append(row)
						if not typeExists:
							missingType.append(row)
						if depCellEmpty:
							missingDep.append(row)
						error = True
				elif copyExists or typeExists or depExists:
					wrongSeqRow.append(row)
					if not refExists:
						missingRef.append(row)
					if not copyExists:
						missingCopy.append(row)
					if not typeExists:
						missingType.append(row)
					if depCellEmpty:
						missingDep.append(row) 
					prevAllExist = True
					error = True
			### chcek repeats
			if copy == self.copyBlockedText:
				copy = None
			if copy != 'None' and copy:
				if copy in allCopy and not copy in repeat:
					repeat[copy] = allCopy[copy]
					repeat[copy].append(row)
					error = True
				elif copy in allCopy and copy in repeat:
					repeat[copy].append(row)
				else:
					allCopy[copy] = [row]

			row = str(int(row) + 1)

		missingRealRefNum = []
		if worksheet[self.pseudoRefC + self.pseudoTitleRow].value:
			pseudo2Real = {}
			exist = True
			pseudoRefRow = str(int(self.pseudoTitleRow) + 1)

			while exist:
				try:
					pseudoRef = worksheet[self.pseudoRefC + pseudoRefRow].value
					realRef = str(worksheet[self.realRefC + pseudoRefRow].value)
					if not pseudoRef or not realRef:
						exist = False
				except IndexError:
					exist = False

				if exist:
					if not realRef or realRef == 'None':
						missingRealRefNum.append(self.realRefC + pseudoRefRow)
					else:
						pseudo2Real[pseudoRef] = realRef
				pseudoRefRow = str(int(pseudoRefRow) + 1)

			excelReference['pseudo2Real'] = pseudo2Real

		errorText = ""
		if missingRef or missingCopy or missingType or missingDep or repeat or wrongSeqRow or missingRealRefNum:
			errorText = writeErrorMessage(missingRef, missingCopy, missingType, missingDep, repeat, wrongSeqRow, missingRealRefNum)
			
		return xmlFilePath, excelReference, errorText

def writeErrorMessage(missingRefRow, missingCopyRow, missingTypeRow, missingDepRow, repeatRefRow, wrongSequenceRow, missingRealRef):
	message = ""
	if missingRefRow:
		message = message + "\nMissing Reference Number at Row: "
		for i in range(0, len(missingRefRow) - 1):
			message = message + missingRefRow[i] + ", "
		message = message + missingRefRow[-1] + "\n"

	if missingCopyRow:
		message = message + "\nMissing Copying Number at Row: "
		for i in range(0, len(missingCopyRow) - 1):
			message = message + missingCopyRow[i] + ", "
		message = message + missingCopyRow[-1] + "\n"

	if missingTypeRow:
		message = message + "\nMissing Reference Type at Row: "
		for i in range(0, len(missingTypeRow) - 1):
			message = message + missingTypeRow[i] + ", "
		message = message + missingTypeRow[-1] + "\n"

	if missingDepRow:
		message = message + "\nMissing Dependent Number at Row: "
		for i in range(0, len(missingDepRow) - 1):
			message = message + missingDepRow[i] + ", "
		message = message + missingDepRow[-1] + "\n"
 
	if repeatRefRow:
		for ref in sorted(repeatRefRow.keys()):
			message = message + "\nR" + ref + " is repeated at Row: "
			for i in range(0, len(repeatRefRow[ref]) -1):
				message = message + repeatRefRow[ref][i] + ", "
			message = message + repeatRefRow[ref][-1]
		message = message + "\n"

	if wrongSequenceRow:
		message = message + "\nSequence Incorrect at Row: "
		for i in range(0, len(wrongSequenceRow) - 1):
			message = message + wrongSequenceRow[i] + ", "
		message = message + wrongSequenceRow[-1] + "\n"

	if missingRealRef:
		message = message + "\nMissing Reference Number at Cell: "
		for i in range(0, len(missingRealRef) - 1):
			message = message + missingRealRef[i] + ", "
		message = message + missingRealRef[-1] + "\n"

	return message