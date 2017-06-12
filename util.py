from os.path import split, splitext

def splitFileFolderAndName(filePath):
	folderPath, fileName = split(filePath)
	return folderPath, splitext(fileName)[0] ### No extension