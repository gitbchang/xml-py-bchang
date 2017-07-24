import csv
import xlrd
from datetime import datetime
from math import ceil as roundup
from re import sub
from lxml import etree, objectify

start = datetime.now()
time = start.strftime('%Y-%m-%dT%H:%M:%S.%f')
client_instance = 'bc-tutorial'

file_path = 'edit_bchang_catalog_noclient.xls'
workbook = xlrd.open_workbook(file_path)
sheet = workbook.sheet_by_index(0)
delimiter_main, delimiter_sub = ',', ';'

broken_gtins = 'bc-tutorial_broken_gtins.csv'
broken_gtins = open(broken_gtins, 'w', encoding='utf-8')
broken_gtins = csv.writer(broken_gtins)
broken_gtins.writerow(['ProductId','GTIN','Problem'])

"""
**SET COLUMN START**
"""
product_id_index = 0
product_name_index = 1
product_locale_names = None

product_description_index = 2
product_locale_descriptions = None

brand_id_index = None
brand_name_index = 3

category_id_index = None
category_name_index = None

product_page_url_index = None
locale_product_page_urls = None
image_url_index = None
locale_image_urls = None

eans_index = 5
upcs_index = 4

model_numbers_index = None
manufacturer_part_numbers_index = None

product_families_index = None
product_families_expand = None
vendor_ids = None
custom_attributes = None
"""
**SET COLUMN END**
"""

header_string = '<?xml version="1.0" encoding="UTF-8" ?>' + '\n'
f = objectify.E
xfeed = objectify.Element("Feed",xmlns="http://www.bazaarvoice.com/xs/PRR/ProductFeed/5.6", name="coty-inc", incremental="false", extractDate=time)

"""
** VALIDATE AND CLEAN FUNCTIONS START **
"""
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

def validate_eans(product_id, eans, broken_gtins):
	valid_eans = []
	for ean in eans:
		if len(ean) in (8, 13):
			# if ean length is 8 or 13, validate
			ean_digits = [int(x) for x in ean]
			for i in range(len(ean_digits[:-1])):
				if len(ean) == 8:
					# xxx - what is going on here?
					if i % 2 == 0:
						ean_digits[i] = ean_digits[i]*3
				else:
					if i % 2 != 0:
						ean_digits[i] = ean_digits[i]*3
			round_nearest_ten = int(roundup(float(sum(ean_digits[:-1]))/10)*10)
			if round_nearest_ten - sum(ean_digits[:-1]) == ean_digits[-1]:
				valid_eans.append(ean)
			else: 
				broken_gtins.writerow([product_id, ean,'EAN: Invalid Checkdigit'])
				# if ean length between 9 - 13, add leading zeros
		elif len(ean) in range(9, 13):
			zeroes = ''
			add = 13 - len(ean)
			for i in range(add):
				zeroes += '0'
			ean = zeroes + ean
			try:
				ean_digits = [int(x) for x in ean]
				for i in range(len(ean_digits[:-1])):
					if i % 2 != 0:
						ean_digits[i] = ean_digits[i]*3
				round_nearest_ten = int(roundup(float(sum(ean_digits[:-1]))/10)*10)
				if round_nearest_ten - sum(ean_digits[:-1]) == ean_digits[-1]:
					valid_eans.append(ean)
				else:
					ean = ean[len(zeroes):]
					broken_gtins.writerow([product_id, ean,'EAN: Invalid Length'])

			except ValueError:
				print(ean, 'too many zeroes?')
				ean = ean[len(zeroes):]
				broken_gtins.writerow([product_id, ean,'EAN: Invalid Length'])
		else:
			broken_gtins.writerow([product_id, ean,'EAN: Invalid Length'])
	
	if len(valid_eans) >= 1:
		x_eans = objectify.SubElement(x_Product, 'EANs')
		# product_string_array.append('      ' + '<EANs>' + '\n')
		for index, ean in enumerate(valid_eans):
			x_ean = objectify.SubElement(x_eans, 'EAN')
			x_ean[index] = ean

# xls validate
def validate_upcs(product_id, upcs, broken_gtins, broken_gtin_count):
	valid_upcs = []
	for upc in upcs:
		if len(upc) in (6, 12):
			upc_digits = [int(x) for x in upc]
			for i in range(len(upc_digits[:-1])):
				if i % 2 == 0:
					upc_digits[i] = upc_digits[i]*3
			round_nearest_ten = int(roundup(float(sum(upc_digits[:-1]))/10)*10)
			if round_nearest_ten - sum(upc_digits[:-1]) == upc_digits[-1]:
				valid_upcs.append(upc)
			else:
				broken_gtins.writerow([product_id, upc,'UPC: Invalid Checkdigit'])
				broken_gtin_count += 1
		elif len(upc) in range(7, 12):
			zeroes = ''
			add = 12 - len(upc)
			for i in range(add):
				zeroes += '0'
			upc = zeroes + upc
			try:
				upc_digits = [int(x) for x in upc]
				for i in range(len(upc_digits[:-1])):
					if i % 2 == 0:
						upc_digits[i] = upc_digits[i]*3
				round_nearest_ten = int(roundup(float(sum(upc_digits[:-1]))/10)*10)
				if round_nearest_ten - sum(upc_digits[:-1]) == upc_digits[-1]:
					valid_upcs.append(upc)
				else:
					upc = upc[len(zeroes):]
					broken_gtins.writerow([product_id, upc,'UPC: Invalid Length'])
					broken_gtin_count += 1

			except ValueError:
				upc = upc[len(zeroes):]
				broken_gtins.writerow([product_id, upc,'UPC: Invalid Length'])
				broken_gtin_count += 1

		else:
			broken_gtins.writerow([product_id, upc,'UPC: Invalid Length'])
			broken_gtin_count += 1

	if len(valid_upcs) >= 1:
		x_upcs = objectify.SubElement(x_Product, "UPCs")
		# product_string_array.append('      ' + '<UPCs>' + '\n')
		for upc in valid_upcs:
			objectify.SubElement(x_upcs, "UPC")
			x_upcs.UPC = upc
		# 	product_string_array.append('        ' + '<UPC>' + upc + '</UPC>' + '\n')
		# product_string_array.append('      ' + '</UPCs>' + '\n')
	print(broken_gtin_count)
	return broken_gtin_count
"""
** VALIDATE AND CLEAN FUNCTIONS END **
"""

"""
** START READING ROWS **
"""
broken_gtin_count = 0
urls = []
row_number = 0
x_Products = objectify.SubElement(xfeed, 'Products')
for row in range(1, sheet.nrows):
	if sheet.cell(row, product_id_index).value != '':
		x_Product = objectify.SubElement(x_Products, 'Product')
		if type(sheet.cell(row, product_id_index).value) == str:
			product_id = sheet.cell(row,product_id_index).value
		if type(sheet.cell(row, product_id_index).value) == float:
			product_id = str(int(sheet.cell(row, product_id_index).value))
		else:
			pass
		
		objectify.SubElement(x_Product, 'ExternalId')
		product_id = id_clean(product_id)
		x_Product.ExternalId = product_id

		if product_name_index is not None:
			product_name = str(sheet.cell(row, product_name_index).value)
			if product_name != '':
				objectify.SubElement(x_Product, 'Name')
				clean_name = xml_clean(product_name)
				x_Product.Name = clean_name

		if product_description_index is not None:
			product_description = str(sheet.cell(row, product_description_index).value)
			if product_description != '':
				objectify.SubElement(x_Product, 'Description')
				product_description = xml_clean(product_description)
				x_Product.Description = product_description
		
		# product_locale_names = xml_clean(row[product_locale_names])

		if brand_id_index is not None:
			brand_id = sheet.cell(row, brand_id_index).value
			if brand_id != '':
				objectify.SubElement(x_Product, 'BrandExternalId')
				brand_id = id_clean(brand_id)
				x_Product.BrandExternalId = brand_id

		if brand_name_index is not None:
			brand_name = sheet.cell(row, brand_name_index).value
			if brand_name != '':
				brand_outer = objectify.SubElement(x_Product, "Brand")
				objectify.SubElement(brand_outer, "Name")
				brand_name = xml_clean(brand_name)
				brand_outer.Name = brand_name

		if category_id_index is not None:
			category_id = sheet.cell(row, category_id_index).value
			if category_id != '':
				objectify.SubElement(x_Product, 'CategoryExternalId')
				category_id = id_clean(category_id)
				x_Product.CategoryExternalId = category_id
		
		if category_name_index is not None:
			category_name = sheet.cell(row, category_name_index).value
			if category_name != '':
				cat_path = objectify.SubElement(x_Product, 'CategoryPath')
				# product_string_array.append('      ' + '<CategoryPath>' + '\n')
		
				if delimiter_sub in category_name:
					categories = category_name.split(delimiter_sub)
					for category in categories:
						objectify.SubElement(cat_path, 'CategoryName')
						category = xml_clean(category)
						cat_path.CategoryName = category
				
				else:
					objectify.SubElement(cat_path, 'CategoryName')
					category = xml_clean(category_name)
					cat_path.CategoryName = category

		if product_page_url_index is not None:
			product_page_url = sheet.cell(row, product_page_url_index).value
			if product_page_url != '':
				objectify.SubElement(x_Product, "ProductPageUrl")
				product_page_url = xml_clean(product_page_url)
				x_Product.ProductPageUrl = product_page_url
			# locale_product_page_urls = row[locale_product_page_urls]

		if image_url_index is not None:
			image_url = sheet.cell(row, image_url_index).value
			if image_url != '':
				objectify.SubElement(x_Product, "ImageUrl")
				image_url = xml_clean(image_url)
				x_Product.ImageUrl = image_url
			# locale_image_urls = row[locale_image_urls]
		
		if upcs_index is not None:
			upcs = sheet.cell(row, upcs_index).value
			if upcs != '':
				if type(upcs) == float:
					upcs = [str(int(upcs))]
				elif delimiter_sub in upcs:
					upcs = upcs.split(delimiter_sub)
				else:
					upcs = [str(upcs)]
				broken_gtin_count = validate_upcs(product_id, upcs, broken_gtins, broken_gtin_count)
			else:
				broken_gtins.writerow([product_id, upcs, 'UPC: Missing UPC'])
				broken_gtin_count +=1

		if eans_index is not None:
			eans = sheet.cell(row, eans_index).value
			if eans != '':
				if type(eans) == float:
					eans = [str(int(eans))]
				elif delimiter_main in eans:
					eans = eans.split(delimiter_main)
				else:
					eans = [str(eans)]
				validate_eans(product_id, eans, broken_gtins)
			else:
				broken_gtins.writerow([product_id, eans, 'EAN: Missing EAN'])
				broken_gtin_count +=1
		
		# if mixed_gtins_index is not None:
		# 	mixed_gtins = sheet.cell(row, mixed_gtins_index).value
		# 	if mixed_gtins != '':
		# 		run upcs or ean check on each item.

		if model_numbers_index is not None:
			model_numbers = sheet.cell(row, model_numbers_index).value
			if model_numbers != '':
				if delimiter_sub in model_numbers:
					model_numbers = model_numbers.split(delimiter_sub)
				else:
					model_numbers = [str(model_numbers)]
				product_string_array.append('      ' + '<ModelNumbers>' + '\n')
				for number in model_numbers:
					number = xml_clean(number)
					product_string_array.append('        ' + '<ModelNumber>' + number + '</ModelNumber>' + '\n')
				product_string_array.append('      ' + '</ModelNumbers>' + '\n')

		if manufacturer_part_numbers_index is not None:
			manufacturer_part_numbers = sheet.cell(row, manufacturer_part_numbers_index).value
			if manufacturer_part_numbers != '':
				if delimiter_sub in manufacturer_part_numbers:
					manufacturer_part_numbers = manufacturer_part_numbers.split(delimiter_sub)
				else:
					manufacturer_part_numbers = [str(manufacturer_part_numbers)]
				product_string_array.append('      ' + '<ManufacturerPartNumbers>' + '\n')
				for number in manufacturer_part_numbers:
					number = xml_clean(number)
					product_string_array.append('        ' + '<ManufacturerPartNumber>' + number + '</ManufacturerPartNumber>' + '\n')
				product_string_array.append('      ' + '</ManufacturerPartNumbers>' + '\n')
		
		# vendor_ids = row[vendor_ids]
		# not sure how to handle multiple custom attributes:
		# custom_attributes = row[custom_attributes]

		# if any attributes: product_string_array += '    <Attributes>' + '\n'
		# if product_families != '':
		# 	product_families = product_families.split(delimiter_sub)
		# 	if len(product_families) >= 1:
		# 		product_string_array.append('    ' + '<Attributes>' + '\n')
		# 		for family in product_families:
		# 			product_string_array.append('      ' + '<Attribute id="BV_FE_FAMILY">' + '\n')
		# 			product_string_array.append('        ' + '<Value>' + family + '</Value>' + '\n')
		# 			product_string_array.append('      ' + '</Attribute>' + '\n')
		# 		product_string_array.append('    ' + '</Attributes>' + '\n')

		if product_families_index is not None:
			product_families = sheet.cell(row, product_families_index).value
			if product_families != '':
				product_string_array.append('    ' + '<Attributes>' + '\n')
				if delimiter_sub in product_families:
					product_families = product_families_expand.split(delimiter_sub)
					if len(product_families) >= 1:
						for family in product_families:
							product_string_array.append('      ' + '<Attribute id="BV_FE_FAMILY">' + '\n')
							product_string_array.append('        ' + '<Value>' + family + '</Value>' + '\n')
							product_string_array.append('      ' + '</Attribute>' + '\n')
							product_string_array.append('      ' + '<Attribute id="BV_FE_EXPAND">' + '\n')
							product_string_array.append('        ' + '<Value>BV_FE_FAMILY:' + family + '</Value>' + '\n')
							product_string_array.append('      ' + '</Attribute>' + '\n')
				else:
					product_string_array.append('      ' + '<Attribute id="BV_FE_FAMILY">' + '\n')
					product_string_array.append('        ' + '<Value>' + family + '</Value>' + '\n')
					product_string_array.append('      ' + '</Attribute>' + '\n')
					product_string_array.append('      ' + '<Attribute id="BV_FE_EXPAND">' + '\n')
					product_string_array.append('        ' + '<Value>BV_FE_FAMILY:' + family + '</Value>' + '\n')
					product_string_array.append('      ' + '</Attribute>' + '\n')
				product_string_array.append('    ' + '</Attributes>' + '\n')

		# product_string_array.append('    </Product>' + '\n')

		row_number += 1
		if row_number % 1000 == 0:
			print('row_number', row_number)



"""
** END READING ROWS **
"""

objectify.deannotate(xfeed, cleanup_namespaces=True)
newxml = etree.tostring(xfeed, pretty_print=True)

xml_path = client_instance+'_product_catalog.xml'
xml_file = open(xml_path, 'w', encoding='UTF-8')
xml_file.write(header_string)
xml_file.close()

xml_file = open(xml_path, 'ab')
xml_file.write(newxml)
xml_file.close()



print (newxml)