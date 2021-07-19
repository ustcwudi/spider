
import os
import time
from scrapy.cmdline import execute
os.system("scrapy crawl mdpi -a p=2071-1050 -a v=13 -a n=14")
os.system("scrapy crawl mdpi -a p=2071-1050 -a v=13 -a n=13")
os.system("scrapy crawl mdpi -a p=2071-1050 -a v=13 -a n=12")
