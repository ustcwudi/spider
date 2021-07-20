
import os
from scrapy.cmdline import execute

for num in range(1, 14):
    os.system("scrapy crawl mdpi -a p=2071-1050 -a v=13 -a n="+str(num))
