# encoding: utf-8
# csv to bvxml
# Python 3.5

import csv
from datetime import datetime
from math import ceil as roundup
from re import sub
import requests
import xlrd

# for tracking execution speed
start = datetime.now()

time = start.strftime('%Y-%m-%dT%H:%M:%S.%f')
client_instance = 'cotyinc'

file_path = 'edit_bchang_catalog.xls'
workbook = xlrd.open_workbook(file_path)
sheet = workbook.sheet_by_index(0)
delimiter_main, delimiter_sub = ',', ';'

broken_gtins = 'broken_gtins.csv'
broken_gtins = open(broken_gtins, 'w', encoding='utf-8')
broken_gtins = csv.writer(broken_gtins)
broken_gtins.writerow(['ProductId','GTIN','Problem'])

xml_path = client_instance+'_product_catalog.xml'
xml_file = open(xml_path, 'w', encoding='utf-8')
header_string = ('<?xml version="1.0" encoding="utf-8" ?>' + '\n' +
	'<Feed xmlns="http://www.bazaarvoice.com/xs/PRR/ProductFeed/5.6" ' +
	'name="' + client_instance + '" ' + 
	'incremental="false" ' +
	'extractDate="' + time + '">' + 
	'\n'
)

# identify which sheet is which, if multiple sheets in the file.

product_string_array = ['  <Products>' + '\n']

# columns. App needs to set these based on the dropdown selections.
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

def validate_eans(product_id, eans, product_string, broken_gtins):
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
		product_string_array.append('      ' + '<EANs>' + '\n')
		for ean in valid_eans:
			product_string_array.append('        ' + '<EAN>' + ean + '</EAN>' + '\n')
		product_string_array.append('      ' + '</EANs>' + '\n')

# xls validate
def validate_upcs(product_id, upcs, product_string_array, broken_gtins, broken_gtin_count):
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
		product_string_array.append('      ' + '<UPCs>' + '\n')
		for upc in valid_upcs:
			product_string_array.append('        ' + '<UPC>' + upc + '</UPC>' + '\n')
		product_string_array.append('      ' + '</UPCs>' + '\n')
	print(broken_gtin_count)
	return broken_gtin_count

broken_gtin_count = 0
urls = []
row_number = 0
for row in range(1, sheet.nrows):
	if sheet.cell(row, product_id_index).value != '':
		product_string_array.append('    <Product>' + '\n')
		if type(sheet.cell(row, product_id_index).value) == str:
			product_id = sheet.cell(row,product_id_index).value
		if type(sheet.cell(row, product_id_index).value) == float:
			product_id = str(int(sheet.cell(row, product_id_index).value))
		else:
			pass
		
		product_id = id_clean(product_id)
		product_string_array.append('      ' + '<ExternalId>' + product_id + '</ExternalId>' + '\n')
		
		if product_name_index is not None:
			product_name = str(sheet.cell(row, product_name_index).value)
			if product_name != '':
				product_name = xml_clean(product_name)
				product_string_array.append('      ' + '<Name>' + product_name + '</Name>' + '\n')

		if product_description_index is not None:
			product_description = str(sheet.cell(row, product_description_index).value)
			if product_description != '':
				product_description = xml_clean(product_description)
				product_string_array.append('      ' + '<Description>' + product_description + '</Description>' + '\n')
		
		# product_locale_names = xml_clean(row[product_locale_names])

		if brand_id_index is not None:
			brand_id = sheet.cell(row, brand_id_index).value
			if brand_id != '':
				brand_id = id_clean(brand_id)
				product_string_array.append('      ' + '<BrandExternalId>' + brand_id + '</BrandExternalId>' + '\n')

		if brand_name_index is not None:
			brand_name = sheet.cell(row, brand_name_index).value
			if brand_name != '':
				brand_name = xml_clean(brand_name)
				product_string_array.append('      ' + '<Brand>' + '\n')
				product_string_array.append('        ' + '<Name>' + brand_name + '</Name>' + '\n')
				product_string_array.append('      ' + '</Brand>' + '\n')

		if category_id_index is not None:
			category_id = sheet.cell(row, category_id_index).value
			if category_id != '':
				category_id = id_clean(category_id)
				product_string_array.append('      ' + '<CategoryExternalId>' + category_id + '</CategoryExternalId>' + '\n')
		
		if category_name_index is not None:
			category_name = sheet.cell(row, category_name_index).value
			if category_name != '':
				product_string_array.append('      ' + '<CategoryPath>' + '\n')
		
				if delimiter_sub in category_name:
					categories = category_name.split(delimiter_sub)
					for category in categories:
						category = xml_clean(category)
						product_string_array.append('        ' + '<CategoryName>' + category + '</CategoryName>' + '\n')
				
				else:
					category = xml_clean(category_name)
					product_string_array.append('        ' + '<CategoryName>' + category + '</CategoryName>' + '\n')
				
				product_string_array.append('      ' + '</CategoryPath>' + '\n')

		if product_page_url_index is not None:
			product_page_url = sheet.cell(row, product_page_url_index).value
			if product_page_url != '':
				product_page_url = xml_clean(product_page_url)
				product_string_array.append('      ' + '<ProductPageUrl>' + product_page_url + '</ProductPageUrl>' + '\n')
			# locale_product_page_urls = row[locale_product_page_urls]

		if image_url_index is not None:
			image_url = sheet.cell(row, image_url_index).value
			if image_url != '':
				image_url = xml_clean(image_url)
				product_string_array.append('      ' + '<ImageUrl>' + image_url + '</ImageUrl>' + '\n')
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
				broken_gtin_count = validate_upcs(product_id, upcs, product_string_array, broken_gtins, broken_gtin_count)
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
				validate_eans(product_id, eans, product_string_array, broken_gtins)
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

		product_string_array.append('    </Product>' + '\n')
		row_number += 1
		if row_number % 1000 == 0:
			print('row_number', row_number)

product_string_array.append('  </Products>' + '\n')
product_string_array.append('</Feed>' + '\n')	

xml_file.write(header_string)
xml_file.write(''.join(product_string_array))
xml_file.close()
print(broken_gtin_count, 'bad gtins')
print(datetime.now() - start)

# Url tests
# for url in urls:
# 	print('new')
# 	r = requests.get(url)
# 	if r.status_code != 200:
# 		print(url, r.status_code)
# 	else:
# 		print('good')