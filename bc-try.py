import csv
import xlrd
from datetime import datetime
from re import sub

from lxml import etree
from lxml import objectify


t1 = objectify.E

root = t1.root(
  t1.name("Indian Motorcycle Balance Bike"),
  t1.description('Indian Motorcycle riders of the future can get started riding their favorite brand with this Balance Bike'),
  t1.imageUrl("military-dev.polarisindcms.com/globalassets/pga/apparel/youth/2863878.jpg/SmallThumbnail"),
  t1.ProductPageUrl("store.indianmotorcycle.com/en-us/shop/apparel/gifts/juniors/2863878")
)

print (etree.tostring(root, pretty_print=True))



