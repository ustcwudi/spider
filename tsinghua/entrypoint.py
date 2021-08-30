
import os
from scrapy.cmdline import execute

for num in range(1, 9):
    os.system("scrapy crawl mdpi -a p=1911-8074 -a v=14 -a n="+str(num))
