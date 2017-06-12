import xml.etree.ElementTree as ET
from xml.etree.ElementTree import ElementTree, Element, SubElement
# import re
from os.path import exists
from util import splitFileFolderAndName

class xmlTool():
	
	prefix = 'R'
	prefixLen = len(prefix)
	xMin = -295
	xMax = 1
	yMin = 0
	yMax = 275

	@classmethod
	def readXML(cls, xmlFilePath):
		"""Read a XML data file and collect reference and wire information.

			Parameters
			----------
			xmlFilePath: string

			Returns
			-------
			referenceInfo: dictionary
				referenceInfo consists of lists of reference names, types, dependencies, gaps, and repeats.
				referenceInfo --> {'name': ['refNum'], 'type': ['type'], 'dependon': ['refNum'], 'gap': ['refNum'],'repeats': [['refname', number]]}

			wireSDCount: dictionary
				keys: All reference names. 
				values: Dictionaries of wire source and destination count. 
		"""

		parseFailure = False
		try:
			# Parse xml content
			tree = ET.parse(xmlFilePath)                                    
		except ET.ParseError: 
			# True if it fails parsing
			parseFailure  = True
		# Get root of the xml data tree and all reference and wire elements if xml is parsed successfully
		if not parseFailure:
			root = tree.getroot() 
			referenceE = root.findall('ReferenceSystem')
			wireE = root.findall('Wire')
		# if there is an error in parsing or getting elements, return empty dictionaries
		if parseFailure or not referenceE or not wireE:
			return {}, {}

		isLetter = False
		refName = [] 				 # A list of reference numbers in string
		refNameGap = [] 			 # A list of missing reference numbers in string
		typ = [] 					 # A list of refernece types in string
		dependon = [] 				 # A list of reference number that references depend on in string
		pseudoRef = {}  			 # A dictionary of pseudo reference name with its counts
		numOfRef = len(referenceE)	 # The number of reference elements
		prevNum = 0
		pointOutOfBounds = {}

		# Obtain wire source and destination information and total count 
		wireSDCount = xmlTool.readWireSDInfo(wireE)
				
		# Obtain ref, type, dep, and gaps
		for i in range(0, numOfRef):
			# Get reference name
			name = referenceE[i].find('Name').text
			pointEs = referenceE[i].findall('Point')

			if name[0] != cls.prefix:
				return {}, {0: "Reference [" + name + "] has an incorrect naming format! A reference name consists of a prefix ""R"" + a number or letter."}
			if name.isalpha():
				# Add to dictionary if name only contains letters
				letter = name[cls.prefixLen:]
				if letter in pseudoRef:
					pseudoRef[letter] = pseudoRef[letter] + 1 
				else:
					pseudoRef[letter] = 1
				if not isLetter:
					isLetter = True
			elif isLetter:
				return {}, {0: "Pseduo reference must be placed in the end! Check the reference after the last R" + letter}
			else:
				# Get reference numbers in string without the prefix('R')
				# numberS = re.findall('\d+', name)[0]
				numberS = name[cls.prefixLen:]
				# Get dependent values in string
				depS = referenceE[i].find('Dependon')
				# If dependent values exist, get dependent numbers in string without the prefix('R')
				if depS != None:
					# depS = re.findall('\d+', depS.text)[0]
					depS = depS.text[cls.prefixLen:]
				# Append reference name, type, and dependency to lists
				refName.append(numberS)
				typ.append(referenceE[i].find('Type').text)
				dependon.append(depS)
				# Obtain gaps if difference between current reference number and previous is greater than 1
				try:
					currNum = int(numberS)
				except ValueError:
					return {}, {0: "Reference [" + name + "] has an incorrect number format! The number needs to be an integer."}

				if currNum - prevNum > 1:
					for missing in range(prevNum + 1, currNum):
						# Append refernce name in string to list
						refNameGap.append(str(missing))
				prevNum = currNum

				# Add on references has no wires attached to
				if 0 not in wireSDCount:
					if not numberS in wireSDCount:
						wireSDCount[numberS] = {'s': [], 'd': []}

			# Check if reference is out of bounds
			for pointE in pointEs:
				x = pointE.find('XPosition').text
				y = pointE.find('YPosition').text
				if float(x) <= cls.xMin or float(x) >= cls.xMax  or float(y) <= cls.yMin  or float(y) >= cls.yMax:
					return {}, {0: name +" location out of bounds!\n-295 mm < X < 1 mm\n   0 mm < Y < 275 mm"}

		# Check location error
		if 0 in wireSDCount:
			return {}, wireSDCount

		# Compress references information into a dictionary
		referenceInfo = {'name': refName,'type': typ, 'dependon': dependon, 'gap': refNameGap, 'repeats': checkRepeats(refName), 'pseudo': pseudoRef}
		return referenceInfo, wireSDCount

	@classmethod
	def readWireSDInfo(cls, wireElements):
		"""Obtain wire source and destination counts.

			Parameters
			----------
			wireElements: list
				A list of wire elements from a parsed XML file

			Returns
			-------
		 	wireSDInfo: dictionary
		 		A dictionary contains information such as total number of wires and dictionaries for lists of wire indices for sources and destinations
		 		that references are treated as. It does not include references that has no wire attached to. 
		 		wireSDInfo --> {totalWireCount: number in int, 'reference number in str': {s:[wire index], d:[wire index]}}
		"""
		# Get total number of wires
		wireSDInfo = {'total': len(wireElements)}
		for wireIndex in range(0, len(wireElements)):
			# Check if location out of bounds
			bondEs = wireElements[wireIndex].findall('Bond')
			for bondE in bondEs:
				x = bondE.find('XPosition').text
				y = bondE.find('YPosition').text
				if float(x) <= cls.xMin or float(x) >= cls.xMax  or float(y) <= cls.yMin  or float(y) >= cls.yMax:
					return {0: 'Wire ' + str((wireIndex + 1)) + ' location out of bounds!\n-295 mm < X < 1 mm\n   0 mm < Y < 275 mm'}
			# Get reference number in string without prefix('R') for source
			# source = re.findall('\d+', wireElements[wireIndex].findall('Bond')[0].find('Refsys').text)[0] 
			source = bondEs[0].find('Refsys').text[cls.prefixLen:]
			# Get reference number in string without prefix('R') for desination
			# destination = re.findall('\d+', wireElements[wireIndex].findall('Bond')[1].find('Refsys').text)[0]
			destination = bondEs[1].find('Refsys').text[cls.prefixLen:]
			# Add reference name as source into the dictionary with wire index
			if source in wireSDInfo:
				wireSDInfo[source]['s'].append(wireIndex)
			else:
				wireSDInfo[source] = {'s': [wireIndex], 'd': []}
			# Add reference name as destination into the dictionary with wire index
			if destination in wireSDInfo:
				wireSDInfo[destination]['d'].append(wireIndex)
			else:
				wireSDInfo[destination] = {'s': [], 'd': [wireIndex]}
		return wireSDInfo

	@classmethod
	def XMLInfo(cls, xmlFilePath, repRef, refName, refGap, wireCount):
		"""Creat a message of all information of the XML data. It's typically used when there is an error in the XML file.

			Parameters
			----------
			xmlFilePath: string
				XML file path.
			repRef: list
				A list of lists of repeating reference names and their count.
			refName: list
				A list of reference name in string.
			refGap: list
				A list of missing refernece name(gap) in string.
			wireCount: int
				The number of wires.

			Returns
			-------
			info: string
				Information of the XML file.
		"""

		if exists(xmlFilePath):
			info = ""
			# Write file path		
			info = info + "#Input XML File: " + xmlFilePath + '\n\n'
			# Write repeating ref name if there is any
			info = info + "#Repeating Reference:\n"
			if repRef:
				for r in repRef:
					info = info + "There are " + str(r[1]) + " R" + r[0] + '\n'
			else:
				info  = info + "None\n"
			# Write first and last ref name
			info = info + "\n#First Reference: R" + refName[0] + '\n'
			info = info + "#Last Reference:  R" + refName[-1] + '\n'
			# Write refernce gaps
			info = info + "\n#Range of Gaps (included):\n"
			if refGap:
				for g in refGap:
					if len(g) == 1:
						info = info + cls.prefix + str(g[0]) + '\n'
					else:
						info = info + cls.prefix + str(g[0]) + ' - ' + cls.prefix + str(g[1]) + '\n'
			else:
				info = info + "None\n"
			# Write wire count
			info = info + "\n#Number of Wires: " + str(wireCount) + "\n"
		else:
			info = "File does not exist!"

		return info
	
	@classmethod
	def modifier(cls, xmlFilePath, referenceDictDFromExc):
		"""Create a new XML file with modified information.
			
			Parameters
			----------
			xmlFilePath: string
				XML file path.
			referenceDictDFromExc: dictionary
				A collection of data from an Excel spreadsheet.
				referenceDictDFromExc data structure --> {'og': {'refNum':[type, dependon]}, 'add': {'refNum': [copyNum, type]}, 'newRefName': ['refNum'], 'pseudo2Real': {'A': '1', 'B': '2'}}


			Returns
			-------
			message: string
				An error message.

			or

			newXmlFilePath: string
				A new XML file path.
		"""

		# Split a XML file path into folder path and file name without the extension
		xmlFolderPath, xmlFileName = splitFileFolderAndName(xmlFilePath)
		# Try to parse XML file. If there is an error, file format is incorrect.
		try:
			tree = ET.parse(xmlFilePath)                                    
		except ET.ParseError: 
			message = "File: " + xmlFileName + " - format incorrect!"
			return message
		except FileNotFoundError:
			message = "File Not Found! File Path: " + xmlFilePath
			return message
		# Create ElementTree object and find its root (highest node)
		root = tree.getroot() 
		# Create lists of reference elements and wire elements
		referenceE = root.findall('ReferenceSystem')
		wireE = root.findall('Wire') 
		# Get numbers of references and wires
		numOfRef = len(referenceE)
		numOfWire = len(wireE)
		referenceEDict = {}

		for r in referenceE: 
			# Get name of the reference
			ref = r.find('Name').text
			if ref.isalpha():
				# Remove pseudo reference elements when the name contains only letters
				root.remove(r)
			else:
				# Get number in string from name
				refNumber = ref[cls.prefixLen:]
				# Get type and dependency from Excel inputs
				typ, dep = referenceDictDFromExc['og'][refNumber]
				# If types dont match, modify type
				if r.find('Type').text != typ:
					r.find('Type').text = typ
				if dep:
					# If there is no dependon originally, create and dependon element in the reference
					if r.find('Dependon') == None:
						newDepEle = Element('Dependon')
						newDepEle.text = cls.prefix + dep
						r.insert(2, newDepEle)
						indent(newDepEle, 2)
					# If dependon values dont match, modify the original value
					elif r.find('Dependon').text != cls.prefix + dep:
						r.find('Dependon').text = cls.prefix + dep
				else:
					# If there is depenon originally, but not in Excel sheet, remove dependon element in XML
					if r.find('Dependon') != None:
						r.remove(r.find('Dependon'))
				referenceEDict[refNumber] = r

		# Read wire source and destination information, see function readWieSDInfo
		wireSDInfo = xmlTool.readWireSDInfo(wireE)
		# If there is pseudo references 
		if 'pseudo2Real' in referenceDictDFromExc:
			# Get converion information dictionary --> {'A': '1', 'B': '2'}
			pseudoTrans = referenceDictDFromExc['pseudo2Real']
			# Get a list of pseudo letters
			letters = pseudoTrans.keys()
			# Translate pseudo reference name to real reference name in wire's source or destination
			for letter in letters:
				if letter in wireSDInfo:
					translation = pseudoTrans[letter]
					# Translate reference name
					wirePseudoSrcIndex = wireSDInfo[letter]['s']
					wirePseudoDesIndex = wireSDInfo[letter]['d']
					if wirePseudoSrcIndex:
						modifyWireRef(letter, translation, wireE, wirePseudoSrcIndex, cls.prefix, 0)
					if wirePseudoDesIndex:
						modifyWireRef(letter, translation, wireE, wirePseudoDesIndex, cls.prefix, 1)
					# Modify wire source and destination index
					if translation in wireSDInfo:
						wireSDInfo[translation]['s'] = wireSDInfo[translation]['s'] + wirePseudoSrcIndex
						wireSDInfo[translation]['d'] = wireSDInfo[translation]['d'] + wirePseudoDesIndex
					else:
						wireSDInfo[translation] = {'s': wirePseudoSrcIndex, 'd': wirePseudoDesIndex}

		# Get a dictionary references to add, addRefDict --> {'refNum': ['copyNum', type]}
		addRefDict = referenceDictDFromExc['add']
		# referenceDictDFromExc['newRefName'] --> ['ref num in str']
		for nName in referenceDictDFromExc['newRefName']:
			# Get reference name to be copied
			refNameToCopy = addRefDict[nName][0]
			# Get a new reference element 
			copy = writeARefCopy(referenceEDict[refNameToCopy], refNameToCopy, nName, addRefDict[nName][1], cls.prefix)
			# Insert into reference tree
			root.insert(int(nName)-1, copy)
			# Change wire destination
			if refNameToCopy in wireSDInfo:
				modifyWireRef(refNameToCopy, nName, wireE, wireSDInfo[refNameToCopy]['d'], cls.prefix, 1)
		# Create new XML file path and write data into it
		newXmlFilePath = xmlFolderPath + "/" + xmlFileName + "_new.xml"
		tree.write(newXmlFilePath)
		return newXmlFilePath

def writeARefCopy(refEToCopy, oldName, newName, typ, prefix): 
	"""Create a referece element (no points).

		Parameters
		----------
		oldName: string
			The reference name to copy
		newNmea: string
			Name of the new reference.
		typ: string
			Type of new reference.
		prefix: string
			String before number in reference name.

		Returns
		-------
		newRefEle: Element
			New reference element.
	"""
	# Creat a new referenceSystem node
	newRefEle = Element('ReferenceSystem')
	# Creat a sub-element for name in reference
	newNameEle = SubElement(newRefEle, 'Name')
	newNameEle.text = prefix + newName
	# Creat a sub-element for type in reference
	newTypeEle = SubElement(newRefEle, 'Type')
	newTypeEle.text = typ
	# Creat a sub-element for dependon in reference
	newDepEle = SubElement(newRefEle, 'Dependon')
	newDepEle.text = prefix + oldName
	# Copy over the first point
	newRefEle.append(refEToCopy.findall('Point')[0])
	# Formatting xml text so it prints nicly 
	indent(newRefEle, 1)
	# Return the reference(address) of the new reference element
	return newRefEle

def modifyWireRef(old, new, wireElements, wireIndex, prefix, srcOrDes): 
	"""Modify wire source or destination.

		Parameters
		----------
		old: string
			The original src or destination (reference number in string). 
		new: string
			The new src ordestination (reference number in string). 
		wireElements: Element
			Wire elements in the XML file
		wireIndex: list
			A list of wire indices of wire elements whose destinations need to be modified.
		prefix: string
			String before number in reference name.
		srcOrDes: int
			0: Source, 1: Destination
	"""

	for index in wireIndex:
		# Find destination element in a wire element 
		bond = wireElements[index].findall('Bond')[srcOrDes].find('Refsys')
		# Change the name from oldDes to newDes
		if bond.text == prefix + old:
			bond.text = prefix + new

def checkRepeats(refNameList):
		"""Check for any repeating references.
			
			Parameters
			----------
			refNameList: list
				A list of reference name in the XML file.

			Returns
			-------
			repeat: list
				1. A list of lists of reference names in str and number of times they repeat. 
				2. A empty list when there is no repeats
				repeat --> [['reference number', count]]
		"""

		repeat = []
		# Get a set of reference name which has no duplicates
		singles = set(refNameList)
		# If both has the same number of elements, there is no repeats
		if len(refNameList) == len(singles):
			return repeat 
		# Append to list when the count of reference name is more than one
		for s in singles:
			count = refNameList.count(s)
			if count > 1:
				repeat.append([s, count])
		return repeat

def indent(elem, level=0):
	"""In-place prettyprint formatter found online --> http://effbot.org/zone/element-lib.htm #prettyprint."""

	i = "\n" + level*"  "
	if len(elem):
		if not elem.text or not elem.text.strip():
			elem.text = i + "  "
		if not elem.tail or not elem.tail.strip():
			elem.tail = i
		for elem in elem:
			indent(elem, level+1)
		if not elem.tail or not elem.tail.strip():
			elem.tail = i
	else:
		if level and (not elem.tail or not elem.tail.strip()):
			elem.tail = i