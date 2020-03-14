
'''
Prints out the live sets from a Montage file.

Based on the excellent work done by Chris Webb, who did a lot of helpful
reverse engineering on the Motif file format, and wrote Python code
based on that. I used his work as a starting point for this code.
Link: http://www.motifator.com/index.php/forum/viewthread/460307/

@author:  Michael Trigoboff
@contact: mtrigoboff@comcast.net
@contact: http://spot.pcc.edu/~mtrigobo

Copyright 2012, 2013 Michael Trigoboff.

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program as file gpl.txt.
If not, see <http://www.gnu.org/licenses/>.
'''

import collections, os.path, struct, sys

VERSION = '1.01'

SONG_ABBREV =		'Sg'
PATTERN_ABBREV =	'Pt'

FILE_HDR_ID =		b'YAMAHA-YSFC'
FILE_VERSION_MIN =	(4, 0, 4)

ENTRY_BLOCK_ID =	b'Entr'
DATA_BLOCK_ID =		b'Data'

BLOCK_HDR_LGTH =				   12
CATALOG_ENTRY_LGTH =			 	8
DATA_HDR_LGTH =						8	# block header length
DLST_DATA_LGTH =			   0x1C69
DLST_DATA_HDR_LGTH =				8
DLST_PAGE_LGTH =				0x1C5
ENTRY_HDR_LGTH =				   16
FILE_HDR_LGTH =					   64
MONTAGE_NAME_MAX_LGTH =			   20
PERF_DATA_LGTH =				   27

def strFromBytes(bytes):
	return bytes.decode('ascii').rstrip('\x00').split('\x00')[0]

def doPerformance(entryName, entryData, dataBlock):
	global userPerfNames

	userPerfNames[entryData[2] - 32][entryData[3]] = entryName.split(':')[1]

def doLiveSetBlock(entryName, entryData, dataBlock):
	global userPerfNames

	assert len(dataBlock) == DLST_DATA_LGTH
	print(entryName + '\n')
	pageOffset = 25
	pages = []			# each page will be of form: [page name, [[performance data (5 bytes)], ...]
	while pageOffset < len(dataBlock):
		page = [strFromBytes(dataBlock[pageOffset : pageOffset + MONTAGE_NAME_MAX_LGTH])]
		perfOffset = pageOffset + MONTAGE_NAME_MAX_LGTH + 23
		pageEmpty = True
		for _ in range(0, 16):
			perfData = struct.unpack('> B B B B ?', dataBlock[perfOffset : perfOffset + 5])
			if perfData[4]:
				pageEmpty = False
			page.append(perfData)
			perfOffset += PERF_DATA_LGTH
		if not pageEmpty:
			pages.append(page)
		pageOffset += DLST_PAGE_LGTH
	for page in pages:
		print('   ' + page[0] + '\n')
		for perfData in page[1:]:
			perfBank = perfData[1]
			perfNum = perfData[2]
			perfPresent = perfData[4]
			print('      ', end='')
			if perfPresent:
				printNum = True
				printName = False
				if perfBank >= 0 and perfBank < 32:
					perfStr = 'PRE{:02}'.format(perfBank + 1)
				elif perfBank >= 32 and perfBank < 37:
					perfBank -= 32
					perfStr = 'USR{:02}'.format(perfBank + 1)
					printName = True
				elif perfBank >= 40 and perfBank < 76:
					bank = perfBank - 40
					perfStr = 'LIB{}({})'.format(int(bank / 5) + 1, (bank % 5) + 1)
				else:
					perfStr = '???'
					printNum = False
				print('{:5}'.format(perfStr), end=' ')
				if printNum:
					print('{:03}'.format(perfNum + 1), end='')
				if printName:
					print(' ' + userPerfNames[perfBank][perfNum])
				else:
					print()
				#print(': {0[0]:3} {0[1]:3} {0[2]:3} {0[3]:3} {0[4]:3}'.format(perfData))
			else:
				print('---')
		print()

class BlockSpec:
	def __init__(self, ident, doFn, needsData):
		self.ident =			ident
		self.doFn =				doFn			# what to do with each item of this type
		self.needsData =		needsData

# when printing out all blocks, they will print out in this order
blockSpecs = collections.OrderedDict((
	('pf',  BlockSpec(b'EPFM',	doPerformance,	False)),		\
	('ls',  BlockSpec(b'ELST',	doLiveSetBlock,	True)),			\
	))

def doBlock(blockSpec):
	global catalog
	
	try:
		inputStream.seek(catalog[blockSpec.ident])
	except:
		print('no data of type: {}\n'.format(blockSpec.name))
		return

	blockHdr = inputStream.read(BLOCK_HDR_LGTH)
	blockIdData, nEntries = struct.unpack('> 4s 4x I', blockHdr)

	assert blockIdData == blockSpec.ident, blockSpec.ident
	
	for i in range(0, nEntries):
		entryHdr = inputStream.read(ENTRY_HDR_LGTH)
		entryId, entryDataLgth, dataOffset = \
			struct.unpack('> 4s I 4x I', entryHdr)
		entryDataLgth -= 8
		entryData = inputStream.read(entryDataLgth)
		assert entryId == ENTRY_BLOCK_ID, ENTRY_BLOCK_ID
		entryNameBytes = entryData[14:].lstrip(b'\xFF')
		entryName = entryNameBytes.decode('ascii').rstrip('\x00').split('\x00')[0]
		if blockSpec.needsData:
			entryPosn = inputStream.tell()
			dataIdent = bytearray(blockSpec.ident)
			dataIdent[0] = ord('D')
			dataIdent = bytes(dataIdent)
			inputStream.seek(catalog[dataIdent] + dataOffset)
			dataHdr = inputStream.read(DATA_HDR_LGTH)
			dataId, dataBlockLgth = struct.unpack('> 4s I', dataHdr)
			assert dataId == DATA_BLOCK_ID, DATA_BLOCK_ID
			dataBlock = inputStream.read(dataBlockLgth)
			inputStream.seek(entryPosn)
		else:
			dataBlock = None
		blockSpec.doFn(entryName, entryData, dataBlock)

def printLiveSets(fileName, selectedItems):
	# globals
	global catalog, fileVersion, inputStream, userPerfNames

	catalog =		{}
	userPerfNames =	[['' for _ in range(128)] for _ in range(5)]
	
	# open file
	try:
		inputStream = open(fileName, 'rb')
	except IOError:
		errStr = 'could not open file: %s' % fileName
		print(errStr)
		raise Exception(errStr)

	# read file header
	fileHdr = inputStream.read(FILE_HDR_LGTH)
	fileHdrId, fileVersionBytes, catalogSize = struct.unpack('> 16s 16s I 28x', fileHdr)
	assert fileHdrId[0:len(FILE_HDR_ID)] == FILE_HDR_ID, FILE_HDR_ID
	fileVersionStr = fileVersionBytes.decode('ascii').rstrip('\x00')
	fileVersion = tuple(map(int, fileVersionStr.split('.')))
	
	if fileVersion[0] < FILE_VERSION_MIN[0] or \
	   fileVersion[1] < FILE_VERSION_MIN[1] or \
	   fileVersion[2] < FILE_VERSION_MIN[2]:
		raise Exception('bad Montage file version {}.{}.{}, needs to be at least {}.{}.{}'
				  .format(fileVersion[0], fileVersion[1], fileVersion[2],
						  FILE_VERSION_MIN[0], FILE_VERSION_MIN[1], FILE_VERSION_MIN[2]))

	# build catalog
	for _ in range(0, int(catalogSize / CATALOG_ENTRY_LGTH)):
		entry = inputStream.read(CATALOG_ENTRY_LGTH)
		entryId, offset = struct.unpack('> 4s I', entry)
		catalog[entryId] = offset

	print('{} (livesets v{}, Montage file v{})\n'.format(os.path.basename(fileName), VERSION, fileVersionStr))

	if len(selectedItems) == 0:					# print everything
		for blockSpec in blockSpecs.values():
			doBlock(blockSpec)
	else:										# print selectedItems
		# cmd line specifies what to print
		for blockAbbrev in selectedItems:
			try:
				doBlock(blockSpecs[blockAbbrev])
			except KeyError:
				print('unknown data type: %s\n' % blockAbbrev)
	
	inputStream.close()
	print()

help1Str = \
'''
To print Live Sets, type:

   python livesets.py montageFileName

If you want to save the output into a text file, do this:

   python livesets.py montageFileName > textFileName
'''

help2Str = \
'''Copyright 2012-2018 Michael Trigoboff.

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.'''

if len(sys.argv) == 1:
	# print help information
	print('livesets version {}\n'.format(VERSION))
	print('by Michael Trigoboff\nmtrigoboff@comcast.net\nhttp://spot.pcc.edu/~mtrigobo')
	print(help1Str)
	#for blockFlag, blockSpec in blockSpecs.items():
	#	print('   {}    {}'.format(blockFlag, blockSpec.name.lower()))
	print(help2Str)
	print()
else:
	# process file
	if len(sys.argv) > 2:
		itemFlags = sys.argv[1:-1]
	else:
		itemFlags = ()
	try:
		printLiveSets(sys.argv[-1], itemFlags)
	except Exception as e:
		print('*** {}\n'.format(e), file=sys.stderr)

