from PIL import Image, ImageDraw, ImageFont
import glob, sys
import xlrd

# -------------------------------------------------------------
# Image Resizer and Watermark Marker
# -------------------------------------------------------------
# Enironment
# 1) Python3 (https://www.python.org/download/releases/3.0/)
# 2) Python Image Library (pip3 install pillow)
# 3) xlrd (pip3 install xlrd) (used to read excel files)
# -------------------------------------------------------------
# How to use
# 1) Arrange the images in the following structure
#    .
#    +-- imageProcess.py
#    +-- fontFile.ttf
#    +-- yourExcel.xlsx
#    +-- folderName 
#        +-- All your images put here
#    +-- resize
#        +-- folderName
#            +-- All your output resized images will put here
#    Note:
#    1) folderName in parent directory and the one in resize 
#       should have the same name
#    2) All the images in folderName should be backed up first
#       as they will be directly modified by the program
# 2) Format the excel files 
#    The first row is always the field names
#    Required: 1) File name
#              2) Person who takes this photo
#    Optional: 1) Location where this photo is taken
# 3) Change some codes in the function defaultProgram()
#    to suit your uses
#    Line 134: Get the files name in folderName
#    Line 135-137: Resize the photos according to the names 
#                  queried
#    Line 139-147: Get the information from excel files. 
#                  The first argument is always the name of 
#                  excel files.
#                  The rest of argument is the custom field names
#                  in the excel file. More details in line 48.
#                  You can learn the function call in this link.
#                  https://www.tutorialspoint.com/python3/python_functions.htm
#    line 148-153: Add the label with respect to the information
#                  extracted in excel file
# -------------------------------------------------------------


# Read the data from excel to form a string
def labelBuilder(url, fileLabel = 'New File name', personLabel = 'Taken by', locationLabel = 'Taken at'):
	try:
		table = xlrd.open_workbook(url).sheet_by_index(0)
		firstRow = table.row_values(0)
		# extract the columns from the labels in first row
		fileName = table.col_values(firstRow.index(fileLabel), start_rowx = 1)
		person = table.col_values(firstRow.index(personLabel), start_rowx = 1)
		location = table.col_values(firstRow.index(locationLabel), start_rowx = 1)
		output = []
		for index, item in enumerate(fileName):
			if item != '':
				tempList = [item]
				temp = 'Photo taken by ' + person[index]
				if location[index] != '?':
					temp += ' in ' + location[index]
				tempList.append(temp)
				if person[index] != 'Christina YM CHAN' and person[index] != '流浪攝' and person[index] != 'Ralph F.Kresge' and person[index] != '?':
					output.append(tempList)
		return output
	except FileNotFoundError:
		print('File not exist or path has some problems at ' + url)
		sys.exit(1)

def addLabelToImage(url, label):
	try:
		# covert the image into Red-Green-Blue-Alpha channel
		im = Image.open(url).convert('RGBA')
		# load the font for text
		try:
			font = ImageFont.truetype('chawp.ttf', int(im.size[1] * 3 / 100))
		except FileNotFoundError:
			print('Can\'t open the font file')
			sys.exit(1)
		drawingContainer = Image.new('RGBA', im.size, (255,255,255,0))
		text = ImageDraw.Draw(drawingContainer)
		# get the width and height of the text in watermark
		textSize = font.getsize(label)
		textLocation = (im.size[0] * 98 / 100 - textSize[0], im.size[1] * 97.5 / 100 - textSize[1])
		# the transparent box for better illustration
		boxLocation = (textLocation[0] - im.size[0] * 1 / 100, textLocation[1] - im.size[1] * 1 / 100, textLocation[0] + textSize[0] + im.size[0] * 0.5 / 100, textLocation[1] + textSize[1] + im.size[1] * 1 / 100)
		text.rectangle(boxLocation, fill=(255,255,255,128))
		text.text(textLocation, label, font=font, fill=(0,0,0,255))
		# Combine two images
		Image.alpha_composite(im, drawingContainer).convert("RGB").save(url)
	except FileNotFoundError:
		print('File not exist or path has some problems at ' + url)
		sys.exit(1)

def resizeImage(url):
	try:
		im = Image.open(url)
		# two different handling methods for landscape or portrait image
		if im.size[0] > im.size[1]:
			# Crop the image into desired ratio (6:4 in this case)
			cropSize = (im.size[1] * 6 / 4 , im.size[1] )
			cropLocation = (im.size[0]/2 - cropSize[0] / 2 , 0 , im.size[0]/2 + cropSize[0] / 2 , im.size[1] )
			nim = im.crop(cropLocation)
			# resize the cropped images
			nim.thumbnail((230, 153))
			nim.save('resize\\'+url)
		else: 
			# Crop the image into desired ratio (6:4 in this case)
			cropSize = (im.size[0] , im.size[0] * 4 / 6 )
			cropLocation = (0 , im.size[1]/2 - cropSize[1] / 2 , im.size [0] , im.size[1]/2 + cropSize[1] / 2 )
			nim = im.crop(cropLocation)
			# resize the cropped images
			nim.thumbnail((230, 153))
			nim.save('resize\\'+url)
	except FileNotFoundError:
		print('File not exist or path has some problems at ' + url)
		sys.exit(1)

# Get the file names in folder
def getFileName():
	# specify the data types
	filePath = glob.glob('**/*.jpg')
	filePath += glob.glob('**/*.png')
	filePath += glob.glob('**/*.tif')
	return filePath

def defaultProgram():
	# resize
	fileName = getFileName()
	for url in fileName:
		print("resizing " + url)
		resizeImage(url)
	# add label
	excelPath = glob.glob('*.xls')
	if len(excelPath) == 0:
		print('Please place one excel file under the directory where program is executed')
		sys.exit(1)
	elif len(excelPath) > 1:
		print('Don\'t place more than one excel files under the directory where program is executed')
		sys.exit(1)
	else:
		labels = labelBuilder(excelPath[0])
		for label in labels:
			for url in fileName:
				if url.find(label[0]) != -1:
					print("adding label to " + url)
					addLabelToImage(url, label[1])
					break


# main method
if __name__ == "__main__":
	defaultProgram()