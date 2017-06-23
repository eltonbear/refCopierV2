import xlsxwriter
from tkinter import Tk
from os import startfile
import openpyxl as op
from util import splitFileFolderAndName
import re

class excelSheet():
	def __init__(self, withFocus):
		### Device
		self.withFocus = withFocus
		### Set tag names
		self.mTag = 'missing'
		self.eTag = 'existing'
		self.aTag = 'appending'
		self.workSheetName = 'Reference_copying'
		self.copyBlockedText = 'BLOCKED'
		### Set column
		self.hiddenRefC='U'
		### Set rows
		self.titleRow = '1'
		self.firstInputRow = str(int(self.titleRow) + 1)
		self.pseudoTitleRow = '6'
		### Set cell address
		self.hiddenIfFocusHeightCell = 'V1'
		self.hiddenIfFirstTimeOpenCell = 'W1'

	def valueInitialization(self, ifWithFocus):
		 ### Set columns
		self.statusC = 'A'
		self.refC = 'B'
		self.copyC = 'C'
		self.typeC = 'D'
		self.deviceC = 'E'
		self.streDeviceC = 'F'
		if(ifWithFocus):
			self.focusHC = 'G'
			self.depC = 'H'
			self.wireSCountC = 'I'
			self.wireDCountC = 'J'
			self.wireNewDcountC = 'K'
			self.warningC = 'L'
			### Others
			self.hiddenRowsC = 'M'
			self.vbaButtonC = 'N'
			### Pseudo	
			self.pseudoRefC = 'P'
			self.realRefC = 'Q'
			self.pseudoCountC = 'R'
			self.wirePseudoCountSC = 'S'
			self.wirePseudoCountDC = 'T'
		else:
			self.depC = 'G'
			self.wireSCountC = 'H'
			self.wireDCountC = 'I'
			self.wireNewDcountC = 'J'
			self.warningC = 'K'
			### Others
			self.hiddenRowsC = 'L'
			self.vbaButtonC = 'M'
			### Pseudo	
			self.pseudoRefC = 'O'
			self.realRefC = 'P'
			self.pseudoCountC = 'Q'
			self.wirePseudoCountSC = 'R'
			self.wirePseudoCountDC = 'S'

		### Set cell address
		self.xmlFilePathCell = self.realRefC + '1'
		self.wireTagCell = self.realRefC + '3'
		self.wireCountCell = self.realRefC + '4'		
		self.lastAppendRowCell = self.hiddenRowsC + '4'
		self.appendRowCountCell = self.hiddenRowsC + '3' 
		self.lastRefRowBeforeMacroCell = self.hiddenRowsC + '2'
		self.hiddenLastExistingRefRowCell = self.hiddenRowsC + '1'

	def startNewExcelSheet(self, xmlFilePath, refInfo, wireSDInfo):
		"""

			parameters
			----------
			refGap: list
			
		"""
		### Set up column and cell addres 
		self.valueInitialization(self.withFocus)
		### Initialize data structures
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
		centerBlankF = workbook.add_format({'valign': 'vcenter', 'hidden': 1, 'font_color': '#FFFFFF'})
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
		appendUnblockedF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#92cddc', 'locked': 0,'border': 1, 'border_color': '#b2b2b2'})
		appendBlockedF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#92cddc', 'locked': 1, 'hidden': 1, 'border': 1, 'border_color': '#b2b2b2'})
		appendHiddenZeroBlockedF = workbook.add_format({'bg_color': '#92cddc', 'font_color': '#92cddc', 'locked': 1, 'hidden': 1, 'border': 1, 'border_color': '#b2b2b2'})
		pseudoRefLetterF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#c6efce', 'font_color': '#006100'})
		pseudoCountsF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#c6efce'})
		warningGoodF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#c6efce', 'font_color': '#006100', 'hidden': 1, 'right': 0, 'top': 0})
		warningCheckF = workbook.add_format({'valign': 'vcenter', 'align': 'center', 'bg_color': '#FFEB9C', 'font_color': '#9c5700', 'hidden': 1, 'right': 0, 'top': 0})
		### activate protection with password "elton"
		worksheet.protect('elton')

		### set column width and protection
		wireTagC = re.findall("[a-zA-Z]+", self.wireTagCell)[0]
		worksheet.set_column(self.statusC + ':' + self.statusC, 9)
		worksheet.set_column(self.refC + ':' + self.refC, 20)
		worksheet.set_column(self.typeC + ':' + self.typeC, 14)
		worksheet.set_column(self.deviceC + ':' + self.deviceC, 10)
		worksheet.set_column(self.streDeviceC + ':' + self.streDeviceC, 12)
		if (self.withFocus):
			worksheet.set_column(self.focusHC + ':' + self.focusHC, 12)
		worksheet.set_column(self.depC + ':' + self.depC, 17)
		worksheet.set_column(self.wireSCountC + ':' + self.wireSCountC, 12)
		worksheet.set_column(self.wireDCountC + ':' + self.wireDCountC, 12)
		worksheet.set_column(self.wireNewDcountC + ':' + self.wireNewDcountC, 16)
		worksheet.set_column(self.warningC + ':' + self.warningC, 10)
		worksheet.set_column(self.hiddenRowsC + ':' + self.hiddenRowsC, 5)
		worksheet.set_column(wireTagC + ':' + wireTagC, 10)

		### write title
		worksheet.write(self.statusC + self.titleRow, 'Status', titleF)
		worksheet.write(self.refC + self.titleRow, 'Reference Number (R)', titleF)
		worksheet.write(self.copyC + self.titleRow, 'Copy (R)', titleF)
		worksheet.write(self.typeC + self.titleRow, 'Reference Type', titleF)
		worksheet.write(self.deviceC + self.titleRow, 'Device', titleF)
		worksheet.write(self.streDeviceC + self.titleRow, 'Stretch Device', titleF)
		if (self.withFocus):
			worksheet.write(self.focusHC + self.titleRow, 'Focus Height', titleF)
		worksheet.write(self.depC + self.titleRow, 'Dependent On (R)', titleF)
		worksheet.write(self.wireSCountC + self.titleRow, 'Wire Count S', titleF)
		worksheet.write(self.wireDCountC + self.titleRow, 'Wire Count D', titleF)
		worksheet.write(self.wireNewDcountC + self.titleRow, 'Wire New Count D', titleF)
		worksheet.write(self.warningC + self.titleRow, 'Warning', titleF)
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
				worksheet.write(self.pseudoRefC + pseudoRefRowS , pseudoRef, pseudoRefLetterF)
				worksheet.write(self.realRefC + pseudoRefRowS, None, unlocked)
				worksheet.write(self.pseudoCountC + pseudoRefRowS , pseudo[pseudoRef], pseudoCountsF)
				worksheet.write(self.wirePseudoCountSC + pseudoRefRowS, len(wireSDInfo[pseudoRef]['s']), pseudoCountsF)
				worksheet.write(self.wirePseudoCountDC + pseudoRefRowS, len(wireSDInfo[pseudoRef]['d']), pseudoCountsF)
				f1 = 'COUNTIF($' + self.realRefC + '$' + str(int(self.pseudoTitleRow)+1) + ':$' + self.realRefC + '$' + str(int(self.pseudoTitleRow)+numPseudo) + ',' + self.realRefC + pseudoRefRowS + ')=1'
				f2 = 'COUNTIF($' + self.hiddenRefC + '$1' + ':$' + self.hiddenRefC + '$' + lastHiddenRefRow + ',' + self.realRefC + pseudoRefRowS + ')=1'
				pseudoRefFormula = '=AND(' + f1 + ', ' + f2 + ')'
				worksheet.data_validation(self.realRefC + pseudoRefRowS, {'validate': 'custom', 'value': pseudoRefFormula, 'error_title': 'Warning', 'error_message': 'Reference does not exist or Duplicates!', 'error_type': 'stop'})
				pseudoRefRowS = str(int(pseudoRefRowS) + 1)
		### Type drop down list
		typeList = ['CAP_D', 'CAP_S', 'COIN_D', 'COIN_S', 'CURTWRIG', 'DCFEEDTC', 'Default', 'IC_D', 'IC_S', 'ORION', 'PKGFLOOR', 'RESIST_D', 'RESIST_S', 'TF_D', 'TF_S']
		### Device drop down lsist
		deviceList = ['CAP', 'TF', 'IC', 'DCFEED']
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
					####
					worksheet.write(self.deviceC + rowS, None,  missingUnblockedF)
					worksheet.write(self.streDeviceC + rowS, 0,  missingDepBlockedBlankF) ### format?
					if (self.withFocus):
						worksheet.write(self.focusHC + rowS, None,  missingUnblockedF)
						worksheet.data_validation(self.focusHC + rowS, {'validate': 'integer', 'criteria': 'between','minimum': -20,'maximum': 20, 'error_title': 'Warning', 'error_message': 'Value not in the range of -20 and 20!', 'error_type': 'stop'})
					####
					worksheet.write(self.wireSCountC + rowS, 0, missingWireCountF)
					worksheet.write(self.wireDCountC + rowS, 0, missingWireCountF)
					### formulas for dependon cells
					worksheet.write_formula(self.depC + rowS, '=' + self.copyC + rowS, missingDepBlockedF)
					worksheet.conditional_format(self.depC + rowS, {'type': 'cell', 'criteria': 'equal to', 'value': 0, 'format': missingDepBlockedBlankF})
					### formulas for wire new D Count cell
					wireNewDFormula = '=IF(ISBLANK(' + self.copyC + rowS + '), 0, INDIRECT("' + self.wireDCountC + '"& ' + self.copyC + rowS + '+1))'
					worksheet.write_formula(self.wireNewDcountC + rowS, wireNewDFormula, missingWireCountHiddenF)
					### Formulas for warning
					warningFormula = 'IF(OR('+ self.wireSCountC + rowS + '=0,'+ self.wireNewDcountC + rowS + '=0), "good","check")'
					worksheet.write_formula(self.warningC + rowS, warningFormula, warningCheckF)
					worksheet.conditional_format(self.warningC + rowS, {'type': 'text', 'criteria': 'containing', 'value': 'good', 'format': warningGoodF})
					### wire formula for data validation. it prevents duplicates and anything outside the list
					f1 = 'COUNTIF($' + self.copyC + '$' + self.firstInputRow + ':$' + self.copyC + '$' + lastAppendRow + ',' + self.copyC + rowS + ')=1'
					f2 = 'COUNTIF($' + self.hiddenRefC + '$1' + ':$' + self.hiddenRefC + '$' + lastHiddenRefRow + ',' + self.copyC + rowS + ')=1'
					countFormula = '=AND(' + f1 + ', ' + f2 + ')'
					worksheet.data_validation(self.copyC + rowS, {'validate': 'custom', 'value': countFormula, 'error_title': 'Warning', 'error_message': 'Reference number does not exist or Duplicates!', 'error_type': 'stop'}) 
					### Data validation for types
					worksheet.data_validation(self.typeC + rowS, {'validate': 'list', 'source': typeList, 'error_title': 'Warning', 'error_message': 'Type does not exist in the library!', 'error_type': 'warning'}) 
					### Data validation for devices
					worksheet.data_validation(self.deviceC + rowS, {'validate': 'list', 'source': deviceList, 'error_title': 'Warning', 'error_message': 'Device does not exist in the library!', 'error_type': 'warning'}) 
				else:  ### existing ref row
					worksheet.write(self.statusC + rowS, self.eTag, existingWhiteBlockedF)
					worksheet.write(self.refC + rowS, refNumber, centerF)
					worksheet.write(self.hiddenRefC + str(refListIndex+1), int(refNumList[refListIndex]), existingWhiteBlockedF)
					worksheet.write(self.copyC + rowS, self.copyBlockedText, copyBlockedF)
					worksheet.write(self.typeC + rowS, refInfo['type'][refListIndex],  unlocked)
					####
					worksheet.write(self.deviceC + rowS, None,  unlocked)
					worksheet.write(self.streDeviceC + rowS, 0,  centerBlankF) ## format?
					if (self.withFocus):
						worksheet.write(self.focusHC + rowS, None,  unlocked) ## format?
						worksheet.data_validation(self.focusHC + rowS, {'validate': 'integer', 'criteria': 'between','minimum': -20,'maximum': 20, 'error_title': 'Warning', 'error_message': 'Value not in the range of -20 and 20!', 'error_type': 'stop'})
					####
					worksheet.write(self.depC + rowS, refInfo['dependon'][refListIndex],  centerF)
					### data validation for dep
					listF = 'COUNTIF($' + self.hiddenRefC + '$1' + ':$' + self.hiddenRefC + '$' + lastHiddenRefRow + ',' + self.depC + rowS + ')=1'
					worksheet.data_validation(self.depC + rowS, {'validate': 'custom', 'value': listF, 'error_title': 'Warning', 'error_message': 'Reference does not exist!', 'error_type': 'stop'})
					### Data validation for types
					worksheet.data_validation(self.typeC + rowS, {'validate': 'list', 'source': typeList, 'error_title': 'Warning', 'error_message': 'Type does not exist in the library!', 'error_type': 'warning'}) 
					### Data validation for devices
					worksheet.data_validation(self.deviceC + rowS, {'validate': 'list', 'source': deviceList, 'error_title': 'Warning', 'error_message': 'Device does not exist in the library!', 'error_type': 'warning'}) 
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
					### Formulas for warning
					warningFormula = 'IF(OR('+ self.wireSCountC + rowS + '=0,'+ self.wireNewDcountC + rowS + '=0), "good","check")'
					worksheet.write_formula(self.warningC + rowS, warningFormula, warningCheckF)
					worksheet.conditional_format(self.warningC + rowS, {'type': 'text', 'criteria': 'containing', 'value': 'good', 'format': warningGoodF})
					refListIndex += 1
			else: ### append section
				if not refGap or rowS == fstAppendRow:
					worksheet.write(self.statusC + rowS, self.aTag, appendTagAndRefF)
					worksheet.write(self.refC + rowS, refNumber, appendTagAndRefF)
					worksheet.write(self.copyC + rowS, None, appendUnblockedF)
					worksheet.write(self.typeC + rowS, None,  appendUnblockedF)
					####
					worksheet.write(self.deviceC + rowS, None,  appendUnblockedF)
					worksheet.write(self.streDeviceC + rowS, 0,  appendHiddenZeroBlockedF) ## format?
					if (self.withFocus):
						worksheet.write(self.focusHC + rowS, None,  appendUnblockedF)
					####
					worksheet.write(self.wireSCountC + rowS, 0, appendBlockedF)
					worksheet.write(self.wireDCountC + rowS, 0, appendBlockedF)	
					### formulas for dep, Wire new D Count, and warning cell
					worksheet.write_formula(self.depC + rowS, '=' + self.copyC + rowS, appendBlockedF)
					wireNewDFormula = '=IF(ISBLANK(' + self.copyC + rowS + '), 0, INDIRECT("' + self.wireDCountC + '"& ' + self.copyC + rowS + '+1))'
					worksheet.write_formula(self.wireNewDcountC + rowS, wireNewDFormula, appendBlockedF)
					warningFormula = 'IF(ISBLANK(' + self.copyC + rowS + '), "", IF(OR('+ self.wireSCountC + rowS + '=0,'+ self.wireNewDcountC + rowS + '=0), "good","check"))'
					worksheet.write_formula(self.warningC + rowS, warningFormula, centerHiddenF)
									
				### conditional formats for wire counts --> dont show values if there is no input in copy column
				wireConditionalFormula = '=AND(ISBLANK(' + self.copyC + rowS + '), NOT(ISBLANK(' + self.statusC + rowS + ')))'
				worksheet.conditional_format(self.depC + rowS, {'type': 'formula', 'criteria': wireConditionalFormula, 'format': appendHiddenZeroBlockedF})
				worksheet.conditional_format(self.wireSCountC + rowS, {'type': 'formula', 'criteria': wireConditionalFormula, 'format': appendHiddenZeroBlockedF})
				worksheet.conditional_format(self.wireDCountC + rowS, {'type': 'formula', 'criteria': wireConditionalFormula, 'format': appendHiddenZeroBlockedF})
				worksheet.conditional_format(self.wireNewDcountC + rowS, {'type': 'formula', 'criteria': wireConditionalFormula, 'format': appendHiddenZeroBlockedF})
				### conditional formats for warning
				worksheet.conditional_format(self.warningC + rowS, {'type': 'text', 'criteria': 'containing', 'value': 'check', 'format': warningCheckF}) ## format?
				worksheet.conditional_format(self.warningC + rowS, {'type': 'text', 'criteria': 'containing', 'value': 'good', 'format': warningGoodF}) ##########???

				### wire formula for datavalidation. it prevents duplicates and anything outside the list(it writes every row til the last appendable row becuase i dont want to do data validation in VBA)
				f1 = 'COUNTIF($' + self.copyC + '$' + self.firstInputRow + ':$' + self.copyC + '$' + lastAppendRow + ',' + self.copyC + rowS + ')=1'
				f2 = 'COUNTIF($' + self.hiddenRefC + '$1' + ':$' + self.hiddenRefC + '$' + lastHiddenRefRow + ',' + self.copyC + rowS + ')=1'
				countFormula = '=AND(' + f1 + ', ' + f2 + ')'
				worksheet.data_validation(self.copyC + rowS, {'validate': 'custom', 'value': countFormula, 'error_title': 'Warning', 'error_message': 'Reference number does not exist or Duplicates!', 'error_type': 'stop'})
				### Data validation for types
				worksheet.data_validation(self.typeC + rowS, {'validate': 'list', 'source': typeList, 'error_title': 'Warning', 'error_message': 'Type does not exist in the library!', 'error_type': 'warning'}) 	
				### Data validation for devices
				worksheet.data_validation(self.deviceC + rowS, {'validate': 'list', 'source': deviceList, 'error_title': 'Warning', 'error_message': 'Device does not exist in the library!', 'error_type': 'warning'}) 
				if (self.withFocus):
					worksheet.data_validation(self.focusHC + rowS, {'validate': 'integer', 'criteria': 'between','minimum': -20,'maximum': 20, 'error_title': 'Warning', 'error_message': 'Value not in the range of -20 and 20!', 'error_type': 'stop'})
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
		if (self.withFocus):
			worksheet.write(self.hiddenIfFocusHeightCell, 1,  existingWhiteBlockedF)
		else:
			worksheet.write(self.hiddenIfFocusHeightCell, 0,  existingWhiteBlockedF)
		worksheet.write(self.hiddenIfFirstTimeOpenCell, 1, existingWhiteBlockedF)
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
		### Setup column and cell addresses 
		self.withFocus = worksheet[self.hiddenIfFocusHeightCell].value
		self.valueInitialization(self.withFocus)

		xmlFilePath = worksheet[self.xmlFilePathCell].value[5:]
		lastRow = worksheet[self.appendRowCountCell].value # int 
		### excelReference data structure --> {'og': {'refNum':[type, device, stretch, focus, dependon]}, 'add': {'refNum': [copyNum, type, device, stretch, focus]}, 'newRefName': [str(refNum)]}
		excelReference = {'og': {}, 'add': {}, 'newRefName': []}
		missingRef = []
		missingCopy = []
		missingType = []
		###
		missingDevice = []
		missingFocus = []
		###
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
			###
			device = str(worksheet[self.deviceC + row].value)
			streDevice = str(worksheet[self.streDeviceC + row].value)
			###

			refExists = ref and ref != 'None'
			copyExists = copy and copy != 'None'
			typeExists = typ and typ != 'None'
			depExists = dep and dep != 'None' and dep != '0' ### with formula 
			depCellEmpty = dep == None or dep == 'None'      ### if gets modified by users and left empty
			###
			deviceExists = device and device != 'None'
			if self.withFocus:
				focus = str(worksheet[self.focusHC + row].value)
				focusExists = focus and focus != 'None'
			else:
				focus = None
				focusExists = True
			###
			
			if dep == 'None':
				dep = None
			if status == self.eTag: 
				if refExists and copyExists and typeExists and deviceExists and focusExists and not error:
					excelReference['og'][ref] = [typ, device, streDevice, focus, dep]
				else:
					if not refExists:
						missingRef.append(row)
					if not typeExists:
						missingType.append(row)
					if not deviceExists:
						missingDevice.append(row)
					if not focusExists:
						missingFocus.append(row)
					error = True
			elif status == self.mTag:
				if refExists and copyExists and typeExists and deviceExists and focusExists and depExists and not error:
					excelReference['add'][ref] = [copy, typ, device, streDevice, focus]
					excelReference['newRefName'].append(ref)
				else:
					if not refExists:
						missingRef.append(row)
					if not copyExists:
						missingCopy.append(row)
					if not typeExists:
						missingType.append(row)
					if not deviceExists:
						missingDevice.append(row)
					if not focusExists:
						missingFocus.append(row)
					if depCellEmpty:
						missingDep.append(row)
					error = True
			else: ### append
				if prevAllExist:					
					if refExists and copyExists and typeExists and deviceExists and focusExists and depExists and not error:
						excelReference['add'][ref] = [copy, typ, device, streDevice, focus]
						excelReference['newRefName'].append(ref)
					elif refExists and not copyExists and not typeExists and not deviceExists and streDevice == '0' and not (focusExists and self.withFocus) and not depExists:
						prevAllExist = False
					else:
						if not refExists:
							missingRef.append(row)
						if not copyExists:
							missingCopy.append(row)
						if not typeExists:
							missingType.append(row)
						if not deviceExists:
							missingDevice.append(row)
						if not focusExists:
							missingFocus.append(row)						
						if depCellEmpty:
							missingDep.append(row)
						error = True
				elif copyExists or typeExists or deviceExists or streDevice == '1' or (self.withFocus and focusExists) or depExists:
					wrongSeqRow.append(row)
					if not refExists:
						missingRef.append(row)
					if not copyExists:
						missingCopy.append(row)
					if not typeExists:
						missingType.append(row)
					if not deviceExists:
							missingDevice.append(row)
					if not focusExists:
						missingFocus.append(row)
					if depCellEmpty:
						missingDep.append(row) 
					prevAllExist = True
					error = True
			### check repeats
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
		if missingRef or missingCopy or missingType or missingDevice or missingFocus or missingDep or repeat or wrongSeqRow or missingRealRefNum:
			errorText = writeErrorMessage(missingRef, missingCopy, missingType, missingDevice, missingFocus, missingDep, repeat, wrongSeqRow, missingRealRefNum)
			
		return xmlFilePath, excelReference, errorText

def writeErrorMessage(missingRefRow, missingCopyRow, missingTypeRow, missingDeviceRow, missingFocusRow, missingDepRow, repeatRefRow, wrongSequenceRow, missingRealRef):
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

	if missingDeviceRow:
		message = message + "\nMissing Device at Row: "
		for i in range(0, len(missingDeviceRow) - 1):
			message = message + missingDeviceRow[i] + ", "
		message = message + missingDeviceRow[-1] + "\n"

	if missingFocusRow:
		message = message + "\nMissing Focus Height at Row: "
		for i in range(0, len(missingFocusRow) - 1):
			message = message + missingFocusRow[i] + ", "
		message = message + missingFocusRow[-1] + "\n"

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
		message = message + "\nIncorrect Sequence at Row: "
		for i in range(0, len(wrongSequenceRow) - 1):
			message = message + wrongSequenceRow[i] + ", "
		message = message + wrongSequenceRow[-1] + "\n"

	if missingRealRef:
		message = message + "\nMissing Reference Number at Cell: "
		for i in range(0, len(missingRealRef) - 1):
			message = message + missingRealRef[i] + ", "
		message = message + missingRealRef[-1] + "\n"

	return message