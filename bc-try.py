import csv
import xlrd
from datetime import datetime
from re import sub

from lxml import etree
from lxml import objectify

start = datetime.now()
time = start.strftime('%Y-%m-%dT%H:%M:%S.%f')
client_instance = 'bc-try'

Feed = objectify.Element('Feed')
xProducts = objectify.Element('Products')
xProduct = objectify.Element('Product')

objectify.SubElement(xProduct, 'Name')
xProduct.Name = 'BCHANG'





"""
# initial object for E-factory tree generation
p = objectify.E
f = objectify.E
ps = objectify.E



Product = p.Product(
  p.name("Indian Motorcycle Balance Bike"),
  p.description('Indian Motorcycle riders of the future can get started riding their favorite brand with this Balance Bike'),
  p.imageUrl("military-dev.polarisindcms.com/globalassets/pga/apparel/youth/2863878.jpg/SmallThumbnail"),
  p.ProductPageUrl("store.indianmotorcycle.com/en-us/shop/apparel/gifts/juniors/2863878")
)

#append product to products
Products = ps.Products(
  Product
)

# append Products to Feed
Feed = f.Feed(
  Products
)
"""


xProducts.append(xProduct)
Feed.append(xProducts)

Feed.set("xmlns", "http://www.bazaarvoice.com/xs/PRR/ProductFeed/5.6")
Feed.set("name", client_instance)
Feed.set("incremental", "false")
Feed.set("encoding", "utf-8")
Feed.set("extractDate", time)

objectify.deannotate(Feed, cleanup_namespaces=True)
newxml = etree.tostring(Feed, pretty_print=True)
header_string = '<?xml version="1.0" encoding="utf-8" ?>' + '\n'
# xfeed = objectify.Element("Feed",xmlns="http://www.bazaarvoice.com/xs/PRR/ProductFeed/5.6", name="coty-inc", incremental="false", encoding="utf-8", extractDate="2017-04-2T05:17:33.945-06:00")

xml_path = client_instance+'_product_catalog.xml'
xml_file = open(xml_path, 'w', encoding='utf-8')
xml_file.write(header_string)
xml_file.close()

xml_file = open(xml_path, 'ab')
xml_file.write(newxml)
xml_file.close()


print (etree.tostring(Feed, pretty_print=True))



