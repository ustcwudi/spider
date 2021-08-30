import sys
import os

if __name__ == '__main__':
    arg1 = sys.argv[1]
    arg2 = sys.argv[2]
    arg3 = sys.argv[3]
    os.system("scrapy crawl mdpi -a p="+arg1+" -a v="+arg2+" -a n=" + arg3)

