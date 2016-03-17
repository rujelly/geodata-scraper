#==============================================================================
# Python 2.7
#
# This program scrapes HTML pages for data about ASU facilities, including
# Facility Number (e.g. 001), Facility Name, Abbreviation (e.g. SCOB), Address,
# Campus, and Longitude/Latitude. Data is saved to a spreadsheet (.xls).
#
# Andrew Roman meroman1@asu.edu
#==============================================================================

from HTMLParser import HTMLParser  # unescape characters
from lxml import html 		   # parse HTML
import xlwt			   # write to final spreadsheet
import urllib			   # get page HTML


def createSpreadsheet():
	"""
	Create spreadsheet with headers

	Returns:
		Workbook object of newly created spreadsheet
	"""
	book = xlwt.Workbook()
	sheet1 = book.add_sheet('Sheet 1', cell_overwrite_ok=True)

	sheetHeaders = ['Facility Number', 'Facility Name', 'Abbreviation',
		'Address', 'Campus/Site Location', 'Longitude/Latitude']
	col = 0
	while(col < 6):
		sheet1.write(0,col, sheetHeaders[col])
		col += 1

	return book

def getFacilityPage(facNum):
	"""
	Scrape html from a facility webpage

	Args:
		String facility number (e.g. 025, X72, etc.)
	Returns: 
		String of HTML from the webpage
	"""
	BASE_URL = 'https://fdm-apps.asu.edu/UFRM/FDS/FacilityData.aspx?bNum='
	url = BASE_URL + facNum 					
	pageHTML = urllib.urlopen(url).read()

	return pageHTML

def parseFacilityPage(pageHTML):
	"""
	Parse page HTML for facility data

	Args:
		String of HTML from the webpage
	Returns:
		Dict of facility data (number, common name, abbreviation, address, campus, long/lat)
	"""
	facility = {	# initialized to 0, since some facility pages have no data available
		'NUMBER' : 0,
		'COMMON_NAME' : 0,
		'ABBREVIATION' : 0,
		'ADDRESS' : 0,
		'CAMPUS' : 0,
		'LONG_LAT' : 0
		}
	htmlList = pageHTML.split('\r\n')
	i = 0

	# parse lines for keywords and add to facility dict
	while(i < len(htmlList) - 4): # -4 guards against index errors (HTML footer doesn't matter)
		if '<td>FACILITY COMMON NAME' in htmlList[i]:
			commonName = htmlList[i + 2].strip()
			commonName = commonName.replace('<',':').replace('>',':').split(':')
			commonName = commonName[4].strip()
			facility['COMMON_NAME'] = commonName
			i += 3
		elif '<td>FACILITY ABBREVIATION' in htmlList[i]:
			abbrev = htmlList[i + 2].strip()
			abbrev = abbrev.replace('<',':').replace('>',':').split(':')
			abbrev = abbrev[4].strip()
			facility['ABBREVIATION'] = abbrev
			i += 3			
		elif '<td>FACILITY ADDRESS' in htmlList[i]:
			address = htmlList[i + 2].strip()
			address = address.replace('<',':').replace('>',':').split(':')
			address = address[4].strip()
			facility['ADDRESS'] = address
			i += 3
		elif '<td>CAMPUS/SITE LOCATION' in htmlList[i]:
			campus = htmlList[i + 2].strip()
			campus = campus.replace('<',':').replace('>',':').split(':')
			campus = campus[4].strip()
			facility['CAMPUS'] = campus
			i += 3
		elif '<td>FACILITY LATITUDE,' in htmlList[i]:
			longLat = htmlList[i + 3].strip()
			longLat = longLat.replace('<',':').replace('>',':').split(':')
			longLat = longLat[2].strip() + ',' + longLat[6].strip()
			longLat = HTMLParser().unescape(unicode(longLat)) # unescape html chars
			facility['LONG_LAT'] = longLat
			i += 3
		else:
			i += 1

	return facility

def appendToSheet(facility, sheet, row):
	"""
	Write facility data to the spreadsheet
	
	Args:
		Dict of facility data to be written
	Returns:
		Boolean, True if write is successful, False otherwise
	"""
	try:
		sheet.write(row, 0, facility['NUMBER'])
		sheet.write(row, 1, facility['COMMON_NAME'])
		sheet.write(row, 2, facility['ABBREVIATION'])
		sheet.write(row, 3, facility['ADDRESS'])
		sheet.write(row, 4, facility['CAMPUS'])
		sheet.write(row, 5, facility['LONG_LAT'])
		return True
	except:
		return False

def main():
	facilityNumFile = 'ASUfacilityNumbers.txt'

	# get facility numbers (e.g. 025, X72, etc.)
	with open(facilityNumFile) as f:
		facNums = f.readlines()

	# strip newline chars from facility codes
	i = 0
	while(i < len(facNums)):
		facNums[i] = facNums[i].rstrip()
		i += 1

	# create spreadsheet with headers
	book = createSpreadsheet()
	sheet = book.get_sheet(0)
	row = 1 # keep track of spreadsheet row to write to; starts after header row

	# scrape/parse html pages and write to spreadsheet
	for facNum in facNums:	
		print('#####################################')
		print('Parsing facility: ' + facNum)

		pageHTML = getFacilityPage(facNum);		# scrape page
		facility = parseFacilityPage(pageHTML);	# parse page

		facility['NUMBER'] = facNum
		
		# print facility data
		print('COMMON NAME:\t\t' + str(facility['COMMON_NAME']))
		print('ABBREVIATION:\t\t' + str(facility['ABBREVIATION']))
		print('ADDRESS:\t\t' + str(facility['ADDRESS']))
		print('CAMPUS:\t\t\t' + str(facility['CAMPUS']))
		print('LONG/LAT:\t\t' + unicode(facility['LONG_LAT'])) # long/lat needs Unicode encoding

		# write to facility dict to Sheet 1
		print('Writing facility %s to Sheet 1' % facNum)
		if(appendToSheet(facility, sheet, row) != 1):
			print('Could not write facility %s to Sheet1' % facNum)

		row += 1

	book.save('ASU_Facilities_Geo_Data.xls')
	print('Spreadsheet saved.')

	print('Process complete.')
# end main

if __name__ == '__main__':
	main()



