import csv
import xlrd
from datetime import datetime
from re import sub

from lxml import etree
from lxml import objectify

start = datetime.now()
time = start.strftime('%Y-%m-%dT%H:%M:%S.%f')
client_instance = 'testing'

# file_path = 'edit_bchang_catalog.xls'
# workbook = xlrd.open_workbook(file_path)
# sheet = workbook.sheet_by_index(0)
# delimiter_main, delimiter_sub = ',', ';'

# broken_gtins = 'TESTING_broken_gtins.csv'
# broken_gtins = open(broken_gtins, 'w', encoding='utf-8')
# broken_gtins = csv.writer(broken_gtins)
# broken_gtins.writerow(['ProductId','GTIN','Problem'])

header_string = '<?xml version="1.0" encoding="utf-8" ?>' + '\n'
xfeed = objectify.Element("Feed",xmlns="http://www.bazaarvoice.com/xs/PRR/ProductFeed/5.6", name="coty-inc", incremental="false", encoding="utf-8", extractDate="2017-04-2T05:17:33.945-06:00")

def id_clean(element_string):
	# need unicode handling
	if element_string[0] == ' ':
		element_string = element_string[1:]
	if element_string[-1] == ' ':
		element_string = element_string[:-1]
	if '&' in element_string:
		element_string = element_string.replace('&', 'and')
	if any(char in element_string for char in ('>', '<', ',', '.', '/', '', '', '',
		'®', '©', '™')):
		element_string = sub(r'[>|<|,|\.|\/||||®|©|™]', '', element_string)
	if ' ' in element_string:
		element_string = element_string.replace(' ', '-')
	return element_string

def xml_clean(element_string):
	if '&' in element_string:
		element_string = sub(r'&(?!amp;)(?!gt;)(?!lt;)(?!#)', '&amp;', element_string)
	if '<' in element_string:
		element_string = sub(r'<', '&lt;', element_string)
	if '>' in element_string:
		element_string = sub(r'>', '&gt;', element_string)
	if any(char in element_string for char in ('', '', '',)):
		element_string = sub(r'[|||]', '', element_string)
	if element_string[-1] == ' ':
		element_string = element_string[:-1]
	if element_string[0] == ' ':
		element_string = element_string[1:]
	return element_string

#Function reads excel file that was sent by the main function
def xls_to_array(raw_file_path):
	wb = xlrd.open_workbook(raw_file_path)
	sh = wb.sheet_by_index(0)
	row_array = []
	#Goes through each cell and adds data to an array
	for row in range(1,sh.nrows):
		vals=[]
		for col in range(sh.ncols):
			cell = sh.cell(row,col)
			cell_value = cell.value
			#Checks to see if the cell is an int data type
			if cell.ctype in (2,3) and int(cell_value) == cell_value:
				cell_value = int(cell_value)
			#Checks to see if it needs to import a date cell to keep the format correct
			if cell.ctype == xlrd.XL_CELL_DATE:
				date_value = xlrd.xldate_as_tuple(cell_value,wb.datemode)
				date_as_tuple = datetime.date(*date_value[:3])
				cell_value = date_as_tuple.strftime("%Y/%m/%d")
				cell_value = str(cell_value)+"T00:00:00.000-00:00"
			#Removes any unacceptle symbols or html tags
			cell_value = re.sub('<[^>]*>', '', str(cell_value))
			vals.append(str(cell_value))

		row_array.append(vals)
	return row_array



x_products = objectify.Element("Products")
x_categories = objectify.Element("Categories")

x_product = objectify.SubElement(x_products, "Product")
x_product.append(objectify.Element('ExternalId'))
x_product.append(objectify.Element('Name'))
x_product.append(objectify.Element('Description'))


x_products.append(x_product)
xfeed.append(x_products)

objectify.deannotate(xfeed, cleanup_namespaces=True)
newxml = etree.tostring(xfeed, pretty_print=True)

xml_path = client_instance+'_product_catalog.xml'
xml_file = open(xml_path, 'w', encoding='utf-8')
xml_file.write(header_string)
xml_file.close()

xml_file = open(xml_path, 'ab')
xml_file.write(newxml)
xml_file.close()



print (newxml)