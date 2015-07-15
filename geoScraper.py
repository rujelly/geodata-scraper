# ASU building geo data scraper
# python 2.7

from HTMLParser import HTMLParser	# unescape characters
from lxml import html 				# parse HTML
import xlwt							# write to final spreadsheet
import urllib						# get page HTML
import re 							# string splitting

BASE_URL = 'https://fdm-apps.asu.edu/UFRM/FDS/FacilityData.aspx?bNum='
FACILITY_NUM_FILE = 'ASUfacilityNumbers.txt'

# get facility numbers
with open(FACILITY_NUM_FILE) as f:
	facNums = f.readlines()

# strip newline chars from facility codes
i = 0
while(i < len(facNums)):
	facNums[i] = facNums[i].rstrip()
	i += 1

# create spreadsheet
book = xlwt.Workbook()
sheet1 = book.add_sheet('Sheet 1', cell_overwrite_ok=True)
sheetHeaders = ['Facility Number', 'Facility Name', 'Abbreviation',
	'Address', 'Campus/Site Location', 'Longitude/Latitude']
col = 0
while(col < 6):
	sheet1.write(0,col, sheetHeaders[col])
	col += 1
row = 1

# scrape/parse html pages and write to spreadsheet
for facNum in facNums:	
	print('#####################################')
	print("Parsing facility: " + facNum)

	facility = {
	'NUMBER' : 0,
	'COMMON_NAME' : 0,
	'ABBREVIATION' : 0,
	'ADDRESS' : 0,
	'CAMPUS' : 0,
	'LONG_LAT' : 0
	}

	url = BASE_URL + facNum 					
	pageHTML = urllib.urlopen(url).read()	# get facility page

	htmlList = pageHTML.split('\r\n')		# split html into lines

	facility['NUMBER'] = facNum

	# parse lines for keywords and add to facility dict
	i = 0
	while(i < len(htmlList) - 4): # - 4 guards against index errors
		if '<td>FACILITY COMMON NAME' in htmlList[i]:
			commonName = htmlList[i + 2].strip()
			commonName = commonName.replace('<',':').replace('>',':').split(':')
			commonName = commonName[4].strip()
			facility['COMMON_NAME'] = commonName
			print("COMMON NAME:\t\t" + str(commonName)) # debug	
			i += 3
		elif '<td>FACILITY ABBREVIATION' in htmlList[i]:
			abbrev = htmlList[i + 2].strip()
			abbrev = abbrev.replace('<',':').replace('>',':').split(':')
			abbrev = abbrev[4].strip()
			facility['ABBREVIATION'] = abbrev
			print("ABBREV:\t\t\t" + str(abbrev)) # debug	
			i += 3			
		elif '<td>FACILITY ADDRESS' in htmlList[i]:
			address = htmlList[i + 2].strip()
			address = address.replace('<',':').replace('>',':').split(':')
			address = address[4].strip()
			facility['ADDRESS'] = address
			print("ADDRESS:\t\t" + str(address)) # debug	
			i += 3
		elif '<td>CAMPUS/SITE LOCATION' in htmlList[i]:
			campus = htmlList[i + 2].strip()
			campus = campus.replace('<',':').replace('>',':').split(':')
			campus = campus[4].strip()
			facility['CAMPUS'] = campus
			print("CAMPUS:\t\t\t" + str(campus)) # debug	
			i += 3
		elif '<td>FACILITY LATITUDE,' in htmlList[i]:
			longLat = htmlList[i + 3].strip()
			longLat = longLat.replace('<',':').replace('>',':').split(':')
			longLat = longLat[2].strip() + ',' + longLat[6].strip()
			longLat = HTMLParser().unescape(unicode(longLat))
			facility['LONG_LAT'] = longLat
			print("LONG/LAT:\t\t" + longLat) # unescape
			i += 3
		else:
			i += 1

	# write to facility dict to Sheet 1
	print("Writing facility %s to Sheet 1" % facNum)
	sheet1.write(row, 0, facility['NUMBER'])
	sheet1.write(row, 1, facility['COMMON_NAME'])
	sheet1.write(row, 2, facility['ABBREVIATION'])
	sheet1.write(row, 3, facility['ADDRESS'])
	sheet1.write(row, 4, facility['CAMPUS'])
	sheet1.write(row, 5, facility['LONG_LAT'])

	row += 1

book.save('ASU_Facilities_Geo_Data.xls')



